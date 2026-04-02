/** Minimal Microsoft Graph v1 HTTP helpers (Bearer token). */

import {
  ERR_GRAPH_GET,
  ERR_GRAPH_PUT,
  GRAPH_BETA_BASE_URL,
  GRAPH_V1_BASE_URL,
  HTTP_CONTENT_TYPE_OCTET_STREAM,
} from "./constants.ts";

export const GRAPH_V1_ROOT = GRAPH_V1_BASE_URL;
export const GRAPH_BETA_ROOT = GRAPH_BETA_BASE_URL;

/** Subset of DriveItem fields used after upload / metadata reads. */
export type GraphDriveItemSummary = {
  id?: string;
  webUrl?: string;
  sharepointIds?: {
    listId?: string;
    listItemId?: string;
    listItemUniqueId?: string;
    siteId?: string;
  };
};

/** Drive item path: encode each path segment (handles `/`, `&`, spaces). */
export function encodeDriveRelativePath(relativePath: string): string {
  return relativePath
    .split("/")
    .filter(Boolean)
    .map((s) => encodeURIComponent(s))
    .join("/");
}

export async function graphRequest(
  token: string,
  graphPath: string,
  init?: RequestInit
): Promise<Response> {
  const headers = new Headers(init?.headers ?? undefined);
  headers.set("Authorization", `Bearer ${token}`);
  return fetch(`${GRAPH_V1_ROOT}${graphPath}`, { ...init, headers });
}

export async function graphGetJson<T>(token: string, graphPath: string): Promise<T> {
  const res = await graphRequest(token, graphPath);
  if (!res.ok) {
    throw new Error(ERR_GRAPH_GET(graphPath, res.status, await res.text()));
  }
  return res.json() as Promise<T>;
}

export async function graphBetaGetJson<T>(token: string, graphPath: string): Promise<T> {
  const headers = new Headers();
  headers.set("Authorization", `Bearer ${token}`);
  const res = await fetch(`${GRAPH_BETA_ROOT}${graphPath}`, { headers });
  if (!res.ok) {
    throw new Error(ERR_GRAPH_GET(`[beta]${graphPath}`, res.status, await res.text()));
  }
  return res.json() as Promise<T>;
}

export async function graphPutStream(
  token: string,
  graphPath: string,
  body: Buffer
): Promise<GraphDriveItemSummary> {
  const res = await graphRequest(token, graphPath, {
    method: "PUT",
    headers: { "Content-Type": HTTP_CONTENT_TYPE_OCTET_STREAM },
    body: new Uint8Array(body),
  });
  if (!res.ok) {
    throw new Error(ERR_GRAPH_PUT(graphPath, res.status, await res.text()));
  }
  const raw = await res.text();
  if (!raw.trim()) {
    return {};
  }
  try {
    return JSON.parse(raw) as GraphDriveItemSummary;
  } catch {
    return {};
  }
}
