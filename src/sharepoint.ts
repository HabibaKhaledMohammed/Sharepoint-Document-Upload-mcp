/**
 * SharePoint upload via Microsoft Graph + interactive browser sign-in.
 * See `constants.ts` for Graph scopes, redirect URI, and defaults.
 */

import {
  InteractiveBrowserCredential,
  deserializeAuthenticationRecord,
  serializeAuthenticationRecord,
  useIdentityPlugin,
  type AuthenticationRecord,
} from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import {
  AADSTS65004_SNIPPET,
  CONSENT_DECLINED_SNIPPET,
  ENTRA_DEFAULT_REDIRECT_URI,
  ENTRA_TENANT_FALLBACK_ORGANIZATIONS,
  ERR_CONSENT_DECLINED_HINT,
  ERR_GRAPH_CREATE_FOLDER,
  ERR_GRAPH_FOLDER_CHECK,
  ERR_INTERACTIVE_NO_GRAPH_TOKEN,
  ERR_MICROSOFT_CLIENT_ID_REQUIRED,
  ERR_SHAREPOINT_NO_DRIVES,
  DOCUMENT_HASH_ALGORITHM,
  GRAPH_DELEGATED_SCOPES,
  HTTP_CONTENT_TYPE_JSON,
  MCP_SERVER_NAME,
  MICROSOFT_AUTH_RECORD_BASE_DIR,
  DEBUG_SHAREPOINT_MY_URL_ENV,
  MICROSOFT_GRAPH_CONFLICT_BEHAVIOR_FAIL,
  MICROSOFT_TOKEN_CACHE_DEFAULT_NAME,
  SHAREPOINT_DEFAULT_DOCUMENT_LIBRARY_NAME,
  MICROSOFT_REDIRECT_URI,
} from "./constants.js";
import {
  encodeDriveRelativePath,
  type GraphDriveItemSummary,
  graphBetaGetJson,
  graphGetJson,
  graphPutStream,
  graphRequest,
} from "./graph-client.js";

/** Re-export for callers that imported redirect from this module. */
export const DEFAULT_REDIRECT_URI = ENTRA_DEFAULT_REDIRECT_URI;

let msalPersistencePluginAttempted = false;
let msalPersistencePluginRegistered = false;

function ensureMsalPersistencePlugin(): boolean {
  if (msalPersistencePluginRegistered) {
    return true;
  }
  if (msalPersistencePluginAttempted) {
    return false;
  }
  msalPersistencePluginAttempted = true;

  const disabled = process.env.MICROSOFT_DISABLE_TOKEN_CACHE;
  if (disabled === "1" || disabled === "true") {
    return false;
  }

  try {
    useIdentityPlugin(cachePersistencePlugin);
    msalPersistencePluginRegistered = true;
    return true;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(
      `[${MCP_SERVER_NAME}] MSAL token cache plugin failed (${msg}). Browser sign-in may be required every run.`
    );
    return false;
  }
}

function resolveTokenCachePersistenceOptions():
  | { enabled: true; name: string; unsafeAllowUnencryptedStorage?: boolean }
  | undefined {
  if (!ensureMsalPersistencePlugin()) {
    return undefined;
  }
  const name =
    process.env.MICROSOFT_TOKEN_CACHE_NAME?.trim() ||
    MICROSOFT_TOKEN_CACHE_DEFAULT_NAME;
  const unsafe =
    process.env.MICROSOFT_TOKEN_CACHE_UNSAFE_UNENCRYPTED === "1" ||
    process.env.MICROSOFT_TOKEN_CACHE_UNSAFE_UNENCRYPTED === "true";
  return {
    enabled: true,
    name,
    ...(unsafe ? { unsafeAllowUnencryptedStorage: true as const } : {}),
  };
}

export type SharePointUploadOverrides = {
  MICROSOFT_TENANT_ID?: string;
  MICROSOFT_CLIENT_ID?: string;
  SHAREPOINT_SITE_HOST?: string;
  SHAREPOINT_SITE_PATH?: string;
  SHAREPOINT_UPLOAD_BASE_PATH?: string;
};

type ResolvedSiteConfig = {
  tenantId?: string;
  clientId?: string;
  redirectUri: string;
  siteHost: string;
  sitePath: string;
  folderPath: string;
  driveName: string;
};

function pick(
  override: string | undefined,
  ...envKeys: string[]
): string | undefined {
  if (override !== undefined) {
    return override;
  }
  for (const key of envKeys) {
    const v = process.env[key];
    if (v !== undefined && v !== "") {
      return v;
    }
  }
  return undefined;
}

function normalizeSitePath(path: string): string {
  const p = path.trim();
  return p.startsWith(":") ? p.slice(1) : p;
}

function resolveSiteConfig(
  overrides?: SharePointUploadOverrides
): ResolvedSiteConfig | null {
  const siteHost = pick(overrides?.SHAREPOINT_SITE_HOST, "SHAREPOINT_SITE_HOST");
  if (!siteHost) {
    return null;
  }

  const rawPath =
    pick(overrides?.SHAREPOINT_SITE_PATH, "SHAREPOINT_SITE_PATH") ?? "";

  return {
    tenantId: pick(overrides?.MICROSOFT_TENANT_ID, "MICROSOFT_TENANT_ID"),
    clientId: pick(overrides?.MICROSOFT_CLIENT_ID, "MICROSOFT_CLIENT_ID"),
    redirectUri:
      pick(undefined, "MICROSOFT_REDIRECT_URI") ?? MICROSOFT_REDIRECT_URI,
    siteHost,
    sitePath: normalizeSitePath(rawPath),
    folderPath:
      pick(
        overrides?.SHAREPOINT_UPLOAD_BASE_PATH,
        "SHAREPOINT_UPLOAD_BASE_PATH",
        "SHAREPOINT_FOLDER_PATH"
      ) ?? "",
    driveName:
      pick(undefined, "SHAREPOINT_DRIVE_NAME") ??
      SHAREPOINT_DEFAULT_DOCUMENT_LIBRARY_NAME,
  };
}

const GRAPH_SITE_WEB_GUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * `GET /sites/{id}` returns composite ids `hostname,{siteCollectionId},{webId}`.
 * List subpaths like `/lists/{listId}/views` return 400 with the composite form; use the web GUID.
 */
function siteIdForListScopedRequests(siteId: string): string {
  const segments = siteId.split(",").map((s) => s.trim());
  if (segments.length === 3) {
    const webId = segments[2];
    if (webId && GRAPH_SITE_WEB_GUID_RE.test(webId)) {
      return webId;
    }
  }
  const trimmed = siteId.trim();
  if (GRAPH_SITE_WEB_GUID_RE.test(trimmed)) {
    return trimmed;
  }
  return siteId;
}

function credentialBindingKey(
  tenantId: string,
  clientId: string,
  redirectUri: string
): string {
  return `${tenantId}|${clientId}|${redirectUri}`;
}

function authenticationRecordFilePath(bindingKey: string): string {
  const dir =
    process.env.MICROSOFT_AUTH_RECORD_DIR?.trim() ||
    path.join(os.homedir(), MICROSOFT_AUTH_RECORD_BASE_DIR);
  const hash = crypto
    .createHash(DOCUMENT_HASH_ALGORITHM)
    .update(bindingKey)
    .digest("hex")
    .slice(0, 24);
  return path.join(dir, `auth-record-${hash}.json`);
}

async function loadSavedAuthenticationRecord(
  bindingKey: string
): Promise<AuthenticationRecord | undefined> {
  try {
    const raw = await fs.readFile(authenticationRecordFilePath(bindingKey), "utf-8");
    return deserializeAuthenticationRecord(raw);
  } catch {
    return undefined;
  }
}

async function saveAuthenticationRecord(
  bindingKey: string,
  record: AuthenticationRecord
): Promise<void> {
  const filePath = authenticationRecordFilePath(bindingKey);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, serializeAuthenticationRecord(record), "utf-8");
}

function rethrowWithConsentMessageIfNeeded(e: unknown): never {
  const msg = e instanceof Error ? e.message : String(e);
  if (
    msg.includes(AADSTS65004_SNIPPET) ||
    msg.includes(CONSENT_DECLINED_SNIPPET)
  ) {
    throw new Error(`${msg}${ERR_CONSENT_DECLINED_HINT}`);
  }
  throw e;
}

async function acquireGraphToken(
  tenantId: string | undefined,
  clientId: string,
  redirectUri: string
): Promise<string> {
  const tenant = tenantId?.trim() || ENTRA_TENANT_FALLBACK_ORGANIZATIONS;
  const bindingKey = credentialBindingKey(tenant, clientId, redirectUri);
  let record = await loadSavedAuthenticationRecord(bindingKey);
  const persistence = resolveTokenCachePersistenceOptions();

  const credential = new InteractiveBrowserCredential({
    tenantId: tenant,
    clientId,
    redirectUri,
    ...(record ? { authenticationRecord: record } : {}),
    ...(persistence ? { tokenCachePersistenceOptions: persistence } : {}),
  });

  if (!record) {
    try {
      const interactiveRecord = await credential.authenticate([
        ...GRAPH_DELEGATED_SCOPES,
      ]);
      if (!interactiveRecord) {
        throw new Error(ERR_INTERACTIVE_NO_GRAPH_TOKEN);
      }
      await saveAuthenticationRecord(bindingKey, interactiveRecord);
    } catch (e) {
      rethrowWithConsentMessageIfNeeded(e);
    }
  }

  try {
    const result = await credential.getToken([...GRAPH_DELEGATED_SCOPES]);
    if (!result?.token) {
      throw new Error(ERR_INTERACTIVE_NO_GRAPH_TOKEN);
    }
    return result.token;
  } catch (e) {
    rethrowWithConsentMessageIfNeeded(e);
  }
}

function sitePathToGraphKey(host: string, sitePath: string): string {
  const seg = sitePath.trim();
  if (seg === "" || seg === "/") {
    return encodeURIComponent(`${host}:/`);
  }
  const normalized = seg.startsWith("/") ? seg : `/${seg}`;
  return encodeURIComponent(`${host}:${normalized}`);
}

async function resolveSiteAndDriveIds(
  token: string,
  host: string,
  sitePath: string,
  driveName: string
): Promise<{ siteId: string; driveId: string }> {
  const siteId = (
    await graphGetJson<{ id: string }>(
      token,
      `/sites/${sitePathToGraphKey(host, sitePath)}`
    )
  ).id;

  const drives = await graphGetJson<{ value: { id: string; name: string }[] }>(
    token,
    `/sites/${encodeURIComponent(siteId)}/drives`
  );
  const drive =
    drives.value.find((d) => d.name === driveName) ?? drives.value[0];
  if (!drive) {
    throw new Error(ERR_SHAREPOINT_NO_DRIVES);
  }
  return { siteId, driveId: drive.id };
}

function mergeDriveItemMeta(
  a: GraphDriveItemSummary,
  b: GraphDriveItemSummary
): GraphDriveItemSummary {
  const sharepointIds =
    a.sharepointIds || b.sharepointIds
      ? { ...a.sharepointIds, ...b.sharepointIds }
      : undefined;
  return {
    id: b.id ?? a.id,
    webUrl: b.webUrl ?? a.webUrl,
    sharepointIds,
  };
}

async function getDriveItemBySiteDrivePath(
  token: string,
  siteId: string,
  driveId: string,
  encodedRelativePath: string
): Promise<GraphDriveItemSummary> {
  const select = "id,webUrl,sharepointIds";
  const metaPath = `/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/root:/${encodedRelativePath}?$select=${select}`;
  return graphGetJson<GraphDriveItemSummary>(token, metaPath);
}

function sharePointHostnameWithoutScheme(siteHost: string): string {
  return siteHost.replace(/^https?:\/\//i, "").replace(/\/.*$/, "");
}

/** e.g. `integrantincorp` from `integrantincorp.sharepoint.com`. */
function tenantKeyFromSharePointHost(hostname: string): string | undefined {
  const lower = hostname.toLowerCase();
  const suffix = ".sharepoint.com";
  if (!lower.endsWith(suffix)) {
    return undefined;
  }
  const sub = lower.slice(0, -suffix.length);
  const first = sub.split(".")[0];
  return first || undefined;
}

function serverRelativeItemPathFromGraphWebUrl(itemWebUrl: string): string {
  const u = new URL(itemWebUrl);
  const path = u.pathname.replace(/\/+$/, "") || u.pathname;
  return decodeURIComponent(path);
}

function parentServerRelativePathFromItem(itemPath: string): string {
  const i = itemPath.lastIndexOf("/");
  if (i <= 0) {
    return itemPath;
  }
  return itemPath.slice(0, i);
}

/** Decode URL until stable so `.../Shared%2520Documents` becomes a single-encoded path for `listurl`. */
function normalizeUrlForSharedListParam(url: string): string {
  let s = url.trim();
  let prev = "";
  while (s !== prev && /%[0-9a-f]{2}/i.test(s)) {
    prev = s;
    try {
      s = decodeURIComponent(s);
    } catch {
      break;
    }
  }
  return s;
}

function buildMySharedQueryString(params: {
  listurl: string;
  viewid?: string;
  id: string;
  parent: string;
  ovuser?: string;
}): string {
  const listurl = normalizeUrlForSharedListParam(params.listurl);
  const parts: string[] = [`listurl=${encodeURIComponent(listurl)}`];
  if (params.viewid?.trim()) {
    parts.push(`viewid=${encodeURIComponent(params.viewid.trim())}`);
  }
  parts.push(`id=${encodeURIComponent(params.id)}`);
  parts.push(`parent=${encodeURIComponent(params.parent)}`);
  if (params.ovuser?.trim()) {
    parts.push(`ovuser=${encodeURIComponent(params.ovuser.trim())}`);
  }
  return parts.join("&");
}

async function fetchListWebRootUrl(
  token: string,
  siteId: string,
  listId: string
): Promise<string | undefined> {
  const sid = siteIdForListScopedRequests(siteId);
  const list = await graphGetJson<{ webUrl?: string }>(
    token,
    `/sites/${encodeURIComponent(sid)}/lists/${encodeURIComponent(
      listId
    )}?$select=webUrl`
  );
  const u = list.webUrl?.trim();
  return u || undefined;
}

function logMySharedUrlDebug(message: string, err?: unknown): void {
  if (process.env[DEBUG_SHAREPOINT_MY_URL_ENV] !== "1") {
    return;
  }
  const detail = err instanceof Error ? err.message : err;
  console.error(`[${MCP_SERVER_NAME}] my/shared URL: ${message}`, detail ?? "");
}

/**
 * Graph list `id` + library webUrl for the document library drive.
 * Prefer site-scoped `/sites/.../drives/.../list` (works for SharePoint libraries).
 */
async function fetchDriveDocumentLibraryList(
  token: string,
  siteId: string,
  driveId: string
): Promise<{ id: string; webUrl?: string } | undefined> {
  const select = "$select=id,webUrl";
  const attempts: { label: string; path: string }[] = [
    {
      label: "site+drive /list",
      path: `/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(
        driveId
      )}/list?${select}`,
    },
    {
      label: "/drives /list",
      path: `/drives/${encodeURIComponent(driveId)}/list?${select}`,
    },
    {
      label: "/drives $expand=list",
      path: `/drives/${encodeURIComponent(driveId)}?$expand=list`,
    },
  ];

  for (const { label, path } of attempts) {
    try {
      if (path.includes("$expand=list")) {
        const drive = await graphGetJson<{
          list?: { id: string; webUrl?: string };
        }>(token, path);
        if (drive.list?.id) {
          return { id: drive.list.id, webUrl: drive.list.webUrl };
        }
      } else {
        const list = await graphGetJson<{ id: string; webUrl?: string }>(
          token,
          path
        );
        if (list?.id) {
          return { id: list.id, webUrl: list.webUrl };
        }
      }
    } catch (e) {
      logMySharedUrlDebug(`${label} failed`, e);
    }
  }
  logMySharedUrlDebug("no list id from site/drives list endpoints or expand");
  return undefined;
}

type GraphListViewRow = {
  id: string;
  name?: string;
  displayName?: string;
};

function pickPreferredLibraryViewId(views: GraphListViewRow[]): string | undefined {
  if (!views.length) {
    return undefined;
  }
  const preferred =
    views.find((v) => v.name === "All Documents") ??
    views.find((v) => /all\s*documents/i.test(v.name ?? "")) ??
    views.find((v) => /all\s*documents/i.test(v.displayName ?? "")) ??
    views[0];
  return preferred?.id;
}

/**
 * Resolves `viewid` for `*-my.sharepoint.com/shared`. Document libraries often reject
 * `.../lists/{id}/views` on v1; we try `$expand=views`, then v1 `/views`, then beta `/views`.
 */
async function fetchLibraryViewIdForSharedLink(
  token: string,
  siteId: string,
  listId: string
): Promise<string | undefined> {
  const sid = siteIdForListScopedRequests(siteId);
  const listBase = `/sites/${encodeURIComponent(sid)}/lists/${encodeURIComponent(
    listId
  )}`;

  try {
    const expanded = await graphGetJson<{
      views?: GraphListViewRow[];
    }>(
      token,
      `${listBase}?$expand=views($select=id,name,displayName)&$select=id`
    );
    const vid = pickPreferredLibraryViewId(expanded.views ?? []);
    if (vid) {
      return vid;
    }
  } catch (e) {
    logMySharedUrlDebug("list $expand=views (v1) failed", e);
  }

  try {
    const data = await graphGetJson<{ value: GraphListViewRow[] }>(
      token,
      `${listBase}/views`
    );
    const vid = pickPreferredLibraryViewId(data.value ?? []);
    if (vid) {
      return vid;
    }
  } catch (e) {
    logMySharedUrlDebug("list /views (v1) failed", e);
  }

  try {
    const data = await graphBetaGetJson<{ value: GraphListViewRow[] }>(
      token,
      `${listBase}/views`
    );
    const vid = pickPreferredLibraryViewId(data.value ?? []);
    if (vid) {
      return vid;
    }
  } catch (e) {
    logMySharedUrlDebug("list /views (beta) failed", e);
  }

  return undefined;
}

/**
 * Browser often opens library files via
 * `https://{tenant}-my.sharepoint.com/shared?listurl=&viewid=&id=&parent=` (and optional `ovuser`).
 * Graph `webUrl` alone points at `*.sharepoint.com/...`, which may not match what users copy from the address bar.
 * Omits `xsdata` / `sdata` (session tokens); those are not available server-side.
 */
async function tryBuildSharePointMySharedDocumentUrl(options: {
  token: string;
  siteId: string;
  driveId: string;
  /** From driveItem.sharepointIds; used only if `/drives/{id}/list` is unavailable. */
  listIdFromSharepointIds?: string;
  itemWebUrl: string;
  sharePointSiteHost: string;
  tenantId?: string;
}): Promise<string | undefined> {
  const {
    token,
    siteId,
    driveId,
    listIdFromSharepointIds,
    itemWebUrl,
    sharePointSiteHost,
    tenantId,
  } = options;

  const host = sharePointHostnameWithoutScheme(sharePointSiteHost);
  const tenantKey = tenantKeyFromSharePointHost(host);
  if (!tenantKey) {
    logMySharedUrlDebug("could not derive tenant prefix", host);
    return undefined;
  }

  const libraryList =
    (await fetchDriveDocumentLibraryList(token, siteId, driveId)) ??
    (listIdFromSharepointIds
      ? { id: listIdFromSharepointIds, webUrl: undefined }
      : undefined);
  if (!libraryList?.id) {
    logMySharedUrlDebug("no library list id (drive list + sharepointIds exhausted)");
    return undefined;
  }
  const listId = libraryList.id;

  let listWebUrl: string | undefined = libraryList.webUrl?.trim();
  if (!listWebUrl) {
    try {
      listWebUrl = await fetchListWebRootUrl(token, siteId, listId);
    } catch {
      return undefined;
    }
  }
  if (!listWebUrl) {
    logMySharedUrlDebug("no list webUrl", { listId });
    return undefined;
  }

  const viewId = await fetchLibraryViewIdForSharedLink(token, siteId, listId);
  if (!viewId) {
    logMySharedUrlDebug(
      "no view id discovered; continuing without viewid query param",
      { listId }
    );
  }

  let itemPath: string;
  try {
    itemPath = serverRelativeItemPathFromGraphWebUrl(itemWebUrl);
    if (!itemPath.startsWith("/")) {
      logMySharedUrlDebug("item path not server-relative", itemPath);
      return undefined;
    }
  } catch (e) {
    logMySharedUrlDebug("parse item webUrl failed", e);
    return undefined;
  }

  const parentPath = parentServerRelativePathFromItem(itemPath);
  const base = `https://${tenantKey}-my.sharepoint.com/shared`;

  let ovuser: string | undefined;
  const tid = tenantId?.trim();
  if (tid) {
    try {
      const me = await graphGetJson<{ userPrincipalName?: string }>(
        token,
        "/me?$select=userPrincipalName"
      );
      const upn = me.userPrincipalName?.trim();
      if (upn) {
        ovuser = `${tid},${upn}`;
      }
    } catch {
      /* ovuser is optional */
    }
  }

  const q = buildMySharedQueryString({
    listurl: listWebUrl,
    viewid: viewId,
    id: itemPath,
    parent: parentPath,
    ovuser,
  });
  return `${base}?${q}`;
}

/**
 * Creates an org-scoped **view** link via Graph; `link.webUrl` generally opens in the
 * browser even when the raw driveItem `webUrl` does not.
 */
async function tryCreateOrganizationViewLink(
  token: string,
  siteId: string,
  driveId: string,
  itemId: string
): Promise<string | undefined> {
  const path = `/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(
    driveId
  )}/items/${encodeURIComponent(itemId)}/createLink`;
  try {
    const res = await graphRequest(token, path, {
      method: "POST",
      headers: { "Content-Type": HTTP_CONTENT_TYPE_JSON },
      body: JSON.stringify({
        type: "view",
        scope: "organization",
      }),
    });
    if (!res.ok) {
      logMySharedUrlDebug(
        "createLink (organization view) failed",
        `${res.status} ${await res.text()}`
      );
      return undefined;
    }
    const data = (await res.json()) as { link?: { webUrl?: string } };
    const url = data.link?.webUrl?.trim();
    return url || undefined;
  } catch (e) {
    logMySharedUrlDebug("createLink error", e);
    return undefined;
  }
}

async function ensureFolderPath(
  token: string,
  driveId: string,
  folderPath: string
): Promise<void> {
  const segments = folderPath.split("/").filter(Boolean);
  if (segments.length === 0) {
    return;
  }

  let pathSoFar = "";
  for (const segment of segments) {
    pathSoFar = pathSoFar ? `${pathSoFar}/${segment}` : segment;
    const encoded = encodeDriveRelativePath(pathSoFar);
    const itemPath = `/drives/${encodeURIComponent(driveId)}/root:/${encoded}`;

    const check = await graphRequest(token, itemPath);
    if (check.ok) {
      continue;
    }
    if (check.status !== 404) {
      throw new Error(
        ERR_GRAPH_FOLDER_CHECK(pathSoFar, check.status, await check.text())
      );
    }

    const parentPath = pathSoFar.includes("/")
      ? pathSoFar.slice(0, pathSoFar.lastIndexOf("/"))
      : "";
    const childrenSegment = parentPath
      ? `root:/${encodeDriveRelativePath(parentPath)}:/children`
      : "root/children";
    const createPath = `/drives/${encodeURIComponent(driveId)}/${childrenSegment}`;

    const created = await graphRequest(token, createPath, {
      method: "POST",
      headers: { "Content-Type": HTTP_CONTENT_TYPE_JSON },
      body: JSON.stringify({
        name: segment,
        folder: {},
        "@microsoft.graph.conflictBehavior": MICROSOFT_GRAPH_CONFLICT_BEHAVIOR_FAIL,
      }),
    });

    if (created.status === 409) {
      continue;
    }
    if (!created.ok) {
      throw new Error(
        ERR_GRAPH_CREATE_FOLDER(
          segment,
          parentPath,
          created.status,
          await created.text()
        )
      );
    }
  }
}

/**
 * Uploads when `SHAREPOINT_SITE_HOST` resolves (params and/or env).
 * Skips when the site host is missing.
 */
export type SharePointUploadSaved = {
  skipped: false;
  webUrl?: string;
  /** Microsoft Graph driveItem id (opaque). */
  driveItemId?: string;
  /** Present for items in SharePoint document libraries. */
  sharepointIds?: GraphDriveItemSummary["sharepointIds"];
};

export async function uploadBufferToSharePointIfConfigured(
  buffer: Buffer,
  filename: string,
  overrides?: SharePointUploadOverrides
): Promise<{ skipped: true } | SharePointUploadSaved> {
  const cfg = resolveSiteConfig(overrides);
  if (!cfg) {
    return { skipped: true };
  }

  if (!cfg.clientId) {
    throw new Error(ERR_MICROSOFT_CLIENT_ID_REQUIRED(ENTRA_DEFAULT_REDIRECT_URI));
  }

  const host = cfg.siteHost.replace(/^https?:\/\//, "");
  const folder = cfg.folderPath
    .replaceAll("\\", "/")
    .replace(/^\/+|\/+$/g, "");

  const token = await acquireGraphToken(
    cfg.tenantId,
    cfg.clientId,
    cfg.redirectUri
  );
  const { siteId, driveId } = await resolveSiteAndDriveIds(
    token,
    host,
    cfg.sitePath,
    cfg.driveName
  );

  await ensureFolderPath(token, driveId, folder);

  const relativeFile = folder ? `${folder}/${filename}` : filename;
  const encodedFile = encodeDriveRelativePath(relativeFile);
  const uploadPath = `/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/root:/${encodedFile}:/content`;

  const putItem = await graphPutStream(token, uploadPath, buffer);
  const metaItem = await getDriveItemBySiteDrivePath(
    token,
    siteId,
    driveId,
    encodedFile
  );
  const merged = mergeDriveItemMeta(putItem, metaItem);

  let webUrl = merged.webUrl;
  const itemId = merged.id;
  if (itemId) {
    const openUrl = await tryCreateOrganizationViewLink(
      token,
      siteId,
      driveId,
      itemId
    );
    if (openUrl) {
      webUrl = openUrl;
    } else if (merged.webUrl) {
      const modern = await tryBuildSharePointMySharedDocumentUrl({
        token,
        siteId,
        driveId,
        listIdFromSharepointIds: merged.sharepointIds?.listId,
        itemWebUrl: merged.webUrl,
        sharePointSiteHost: cfg.siteHost,
        tenantId: cfg.tenantId,
      });
      if (modern) {
        webUrl = modern;
      }
    }
  } else if (merged.webUrl) {
    const modern = await tryBuildSharePointMySharedDocumentUrl({
      token,
      siteId,
      driveId,
      listIdFromSharepointIds: merged.sharepointIds?.listId,
      itemWebUrl: merged.webUrl,
      sharePointSiteHost: cfg.siteHost,
      tenantId: cfg.tenantId,
    });
    if (modern) {
      webUrl = modern;
    }
  }

  return {
    skipped: false,
    webUrl,
    driveItemId: merged.id,
    sharepointIds: merged.sharepointIds,
  };
}
