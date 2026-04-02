/** Central configuration literals (no secrets). */

/* Microsoft Graph */
export const GRAPH_V1_BASE_URL = "https://graph.microsoft.com/v1.0" as const;
export const GRAPH_BETA_BASE_URL = "https://graph.microsoft.com/beta" as const;
export const MICROSOFT_AUTH_MODE="interactive" as const;
export const MICROSOFT_REDIRECT_URI="http://localhost" as const;
/** Delegated scopes requested for interactive SharePoint upload. */
export const GRAPH_DELEGATED_SCOPES = [
  "https://graph.microsoft.com/Sites.Read.All",
  "https://graph.microsoft.com/Files.ReadWrite.All",
] as const;

/* Microsoft Entra (interactive public client) */
export const ENTRA_DEFAULT_REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient" as const;

/** Used when MICROSOFT_TENANT_ID is omitted. */
export const ENTRA_TENANT_FALLBACK_ORGANIZATIONS = "organizations" as const;

/**
 * MSAL disk cache folder name under the OS store (e.g. Windows:
 * %LocalAppData%\\.IdentityService\\<name>). Override with MICROSOFT_TOKEN_CACHE_NAME.
 * Set MICROSOFT_DISABLE_TOKEN_CACHE=true to skip persistence (browser every run).
 * If cache init fails without OS encryption, set MICROSOFT_TOKEN_CACHE_UNSAFE_UNENCRYPTED=true.
 */
export const MICROSOFT_TOKEN_CACHE_DEFAULT_NAME = "document-upload-mcp-graph" as const;

/**
 * Account binding file for silent auth (Azure Identity needs this in addition to MSAL token cache).
 * Default directory: ~/.document-upload-mcp (override MICROSOFT_AUTH_RECORD_DIR).
 * Delete `auth-record-*.json` there to force sign-in again.
 */
export const MICROSOFT_AUTH_RECORD_BASE_DIR = ".document-upload-mcp" as const;

/* SharePoint */
export const SHAREPOINT_DEFAULT_DOCUMENT_LIBRARY_NAME = "Documents" as const;

/* HTTP (Graph requests) */
export const HTTP_CONTENT_TYPE_JSON = "application/json" as const;
export const HTTP_CONTENT_TYPE_OCTET_STREAM = "application/octet-stream" as const;

/* MCP server metadata */
export const MCP_SERVER_NAME = "sharepoint-document-upload-mcp" as const;
export const MCP_SERVER_VERSION = "1.0.0" as const;
export const MCP_SERVER_DESCRIPTION = "A MCP server for uploading document" as const;

/* Local document handling */
export const ALLOWED_UPLOAD_EXTENSIONS = [
  ".txt",
  ".md",
  ".pdf",
  ".docx",
] as const;

export const LOCAL_UPLOAD_DIR = "uploads" as const;

/** Prefix length from SHA-256 hex used in SharePoint upload filenames. */
export const SHAREPOINT_FILENAME_HASH_PREFIX_LENGTH = 12;

/** Used with `crypto.createHash` for document deduplication and filenames. */
export const DOCUMENT_HASH_ALGORITHM = "sha256" as const;

/* MCP tools — names, titles, descriptions, responses */
export const TOOL_NAME_HELLO_WORLD = "hello_world" as const;
export const TOOL_NAME_UPLOAD_DOCUMENT = "upload_document" as const;

export const TOOL_HELLO_WORLD_TITLE = "Hello World" as const;
export const TOOL_HELLO_WORLD_DESCRIPTION =
  "This tool will return Hello World" as const;
export const TOOL_HELLO_WORLD_RESPONSE_TEXT = "Hello World" as const;

export const TOOL_UPLOAD_DOCUMENT_TITLE = "Upload a document" as const;
export const TOOL_UPLOAD_DOCUMENT_DESCRIPTION =
  "Upload a document locally and to SharePoint when configured. Uses interactive Microsoft sign-in with Graph scopes Sites.Read.All and Files.ReadWrite.All only — remove other API permissions on the Entra app for SharePoint-only consent. After the first sign-in, account metadata and tokens are stored on disk (see ~/.document-upload-mcp and OS .IdentityService cache) so later runs usually skip the browser. Set MICROSOFT_DISABLE_TOKEN_CACHE=true only to disable the MSAL file cache." as const;

/** Zod field descriptions for `upload_document` input schema */
export const SCHEMA_DESC_DOCUMENT_FILENAME =
  "Original filename including extension (e.g. report.pdf)" as const;
export const SCHEMA_DESC_DOCUMENT_CONTENT =
  "File body as UTF-8 text or base64-encoded bytes" as const;
export const SCHEMA_DESC_MICROSOFT_TENANT_ID =
  "Entra tenant ID (optional; env MICROSOFT_TENANT_ID; omit for organizations / multi-tenant)" as const;
export const SCHEMA_DESC_MICROSOFT_CLIENT_ID =
  "App registration client ID (required for interactive browser sign-in; env MICROSOFT_CLIENT_ID)" as const;
export const SCHEMA_DESC_SHAREPOINT_SITE_HOST =
  "SharePoint hostname, e.g. contoso.sharepoint.com (or set SHAREPOINT_SITE_HOST env)" as const;
export const SCHEMA_DESC_SHAREPOINT_SITE_PATH =
  "Server-relative site path (e.g. /sites/Team). Use empty string for the root site" as const;
export const SCHEMA_DESC_SHAREPOINT_UPLOAD_BASE_PATH =
  "Folder under the document library (e.g. Incoming/2026). Omit or empty for library root" as const;

/* Server messages */
export const ERR_UNSUPPORTED_FILE_TYPE = "Unsupported file type" as const;
export const MSG_UPLOAD_SUCCESS_TEMPLATE = (sharePointNote: string) =>
  `Upload successful.${sharePointNote}`;
/** Appended when Graph returns a browser URL for the uploaded item. */
export const MSG_SHAREPOINT_DOC_URL_LINE = (url: string) =>
  `\n\nSharePoint document URL:\n${url}` as const;
/** Opaque Microsoft Graph `driveItem` id (for `/drives/.../items/{id}`). */
export const MSG_SHAREPOINT_DRIVE_ITEM_ID_LINE = (id: string) =>
  `\n\nMicrosoft Graph drive item id:\n${id}` as const;
/** SharePoint list row id (often what people mean by "item id" in libraries). */
export const MSG_SHAREPOINT_LIST_ITEM_ID_LINE = (id: string) =>
  `\n\nSharePoint list item id:\n${id}` as const;
export const MSG_SHAREPOINT_LIST_ITEM_UNIQUE_ID_LINE = (id: string) =>
  `\n\nSharePoint list item unique id:\n${id}` as const;
export const SP_RESULT_NOTE_COMPLETED =
  "\n\nSharePoint upload completed (no document URL or ids in API response)." as const;

/* SharePoint / Graph runtime errors (user-facing) */
export const ERR_INTERACTIVE_NO_GRAPH_TOKEN =
  "Interactive sign-in did not return a Microsoft Graph token." as const;
export const ERR_CONSENT_DECLINED_HINT =
  " Grant only delegated Sites.Read.All and Files.ReadWrite.All on the app registration, then accept consent." as const;
export const AADSTS65004_SNIPPET = "AADSTS65004" as const;
export const CONSENT_DECLINED_SNIPPET = "declined to consent" as const;

export const ERR_SHAREPOINT_NO_DRIVES = "No drives found on SharePoint site" as const;

export const ERR_MICROSOFT_CLIENT_ID_REQUIRED = (redirectUri: string) =>
  `MICROSOFT_CLIENT_ID is required. Use a public Entra client with redirect "${redirectUri}" and delegated Sites.Read.All + Files.ReadWrite.All.`;

export const ERR_GRAPH_FOLDER_CHECK = (
  pathSoFar: string,
  status: number,
  body: string
) => `Graph folder check ${pathSoFar}: ${status} ${body}`;

export const ERR_GRAPH_CREATE_FOLDER = (
  segment: string,
  parentPath: string,
  status: number,
  body: string
) =>
  `Create folder "${segment}" under "${parentPath || "root"}": ${status} ${body}`;

export const MICROSOFT_GRAPH_CONFLICT_BEHAVIOR_FAIL = "fail" as const;

/** Set to `1` to log why `*-my.sharepoint.com/shared` URL construction was skipped. */
export const DEBUG_SHAREPOINT_MY_URL_ENV = "SHAREPOINT_DEBUG_MY_URL" as const;

/* Graph client error prefixes */
export const ERR_GRAPH_GET = (graphPath: string, status: number, body: string) =>
  `Graph GET ${graphPath}: ${status} ${body}`;
export const ERR_GRAPH_PUT = (graphPath: string, status: number, body: string) =>
  `Graph PUT ${graphPath}: ${status} ${body}`;

/* Integration test script (run-mcp-upload-test) */
export const MCP_UPLOAD_TEST_CLIENT_NAME = "upload-test" as const;
export const MCP_UPLOAD_TEST_CLIENT_VERSION = "1.0.0" as const;
export const MCP_CALL_TIMEOUT_ENV = "MCP_CALL_TIMEOUT_MS" as const;
export const MCP_CALL_TIMEOUT_DEFAULT_MS = 900_000;
/** Default fixture under `fixtures/`; override with env MCP_UPLOAD_TEST_FIXTURE. */
export const FIXTURE_UPLOAD_TEST_FILENAME = "mcp-random-test.md" as const;
export const MCP_UPLOAD_TEST_FIXTURE_ENV = "MCP_UPLOAD_TEST_FIXTURE" as const;

/** Project-relative paths (repo root). */
export const PROJECT_REL_TSX_CLI = "node_modules/tsx/dist/cli.mjs" as const;
export const PROJECT_REL_SERVER_ENTRY = "src/server.ts" as const;
