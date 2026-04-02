import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import dotenv from "dotenv";
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

const __serverDir = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__serverDir, "..", ".env") });
import {
  ALLOWED_UPLOAD_EXTENSIONS,
  DOCUMENT_HASH_ALGORITHM,
  ERR_UNSUPPORTED_FILE_TYPE,
  ERR_UNSUPPORTED_FORMAT,
  FAKE_EMBEDDING_DIMENSION,
  LOCAL_UPLOAD_DIR,
  MCP_SERVER_DESCRIPTION,
  MCP_SERVER_NAME,
  MCP_SERVER_VERSION,
  MOCK_PARSED_DOCX_CONTENT,
  MOCK_PARSED_PDF_CONTENT,
  MSG_DOCUMENT_ALREADY_PROCESSED,
  MSG_SHAREPOINT_DOC_URL_LINE,
  MSG_SHAREPOINT_DRIVE_ITEM_ID_LINE,
  MSG_SHAREPOINT_LIST_ITEM_ID_LINE,
  MSG_SHAREPOINT_LIST_ITEM_UNIQUE_ID_LINE,
  MSG_UPLOAD_SUCCESS_TEMPLATE,
  SCHEMA_DESC_DOCUMENT_CONTENT,
  SCHEMA_DESC_DOCUMENT_FILENAME,
  SCHEMA_DESC_MICROSOFT_CLIENT_ID,
  SCHEMA_DESC_MICROSOFT_TENANT_ID,
  SCHEMA_DESC_SHAREPOINT_SITE_HOST,
  SCHEMA_DESC_SHAREPOINT_SITE_PATH,
  SCHEMA_DESC_SHAREPOINT_UPLOAD_BASE_PATH,
  SHAREPOINT_FILENAME_HASH_PREFIX_LENGTH,
  SP_RESULT_NOTE_COMPLETED,
  TEXT_CHUNK_OVERLAP_CHARS,
  TEXT_CHUNK_SIZE_CHARS,
  TOOL_HELLO_WORLD_DESCRIPTION,
  TOOL_HELLO_WORLD_RESPONSE_TEXT,
  TOOL_HELLO_WORLD_TITLE,
  TOOL_NAME_HELLO_WORLD,
  TOOL_NAME_UPLOAD_DOCUMENT,
  TOOL_UPLOAD_DOCUMENT_DESCRIPTION,
  TOOL_UPLOAD_DOCUMENT_TITLE,
} from "./constants.ts";
import { uploadBufferToSharePointIfConfigured } from "./sharepoint.ts";

const server = new McpServer({
  name: MCP_SERVER_NAME,
  version: MCP_SERVER_VERSION,
  description: MCP_SERVER_DESCRIPTION,
});
server.registerTool(
  TOOL_NAME_HELLO_WORLD,
  {
    title: TOOL_HELLO_WORLD_TITLE,
    description: TOOL_HELLO_WORLD_DESCRIPTION,
    inputSchema: {}
  },
  async () => {
    return {  content: [
        {
          type: "text",
          text: TOOL_HELLO_WORLD_RESPONSE_TEXT,
        },
      ],
    };
  }
);

server.registerTool(
  TOOL_NAME_UPLOAD_DOCUMENT,
  {
    title: TOOL_UPLOAD_DOCUMENT_TITLE,
    description: TOOL_UPLOAD_DOCUMENT_DESCRIPTION,
    inputSchema: z.object({
      document: z.object({
        filename: z.string().describe(SCHEMA_DESC_DOCUMENT_FILENAME),
        content: z.string().describe(SCHEMA_DESC_DOCUMENT_CONTENT),
      }),
      MICROSOFT_TENANT_ID: z
        .string()
        .describe(SCHEMA_DESC_MICROSOFT_TENANT_ID),
      MICROSOFT_CLIENT_ID: z
        .string()
        .describe(SCHEMA_DESC_MICROSOFT_CLIENT_ID),
      SHAREPOINT_SITE_HOST: z
        .string()
        .describe(SCHEMA_DESC_SHAREPOINT_SITE_HOST),
      SHAREPOINT_SITE_PATH: z
        .string()
        .describe(SCHEMA_DESC_SHAREPOINT_SITE_PATH),
      SHAREPOINT_UPLOAD_BASE_PATH: z
        .string()
        .describe(SCHEMA_DESC_SHAREPOINT_UPLOAD_BASE_PATH),
    }),
  },
  async ({
    document,
    MICROSOFT_TENANT_ID,
    MICROSOFT_CLIENT_ID,
    SHAREPOINT_SITE_HOST,
    SHAREPOINT_SITE_PATH,
    SHAREPOINT_UPLOAD_BASE_PATH,
  }) => {
    const { filename, content } = document;

    const ext = path.extname(filename).toLowerCase();
    if (!(ALLOWED_UPLOAD_EXTENSIONS as readonly string[]).includes(ext)) {
      throw new Error(ERR_UNSUPPORTED_FILE_TYPE);
    }

    const uploadDir = path.resolve(LOCAL_UPLOAD_DIR);
    await fs.mkdir(uploadDir, { recursive: true });

    const filePath = path.join(uploadDir, filename);

    const buffer = isBase64(content)
      ? Buffer.from(content, "base64")
      : Buffer.from(content, "utf-8");

    await fs.writeFile(filePath, buffer);

    const hash = crypto.createHash(DOCUMENT_HASH_ALGORITHM).update(buffer).digest("hex");

    if (await alreadyProcessed(hash)) {
      return {
        content: [
          {
            type: "text" as const,
            text: MSG_DOCUMENT_ALREADY_PROCESSED,
          },
        ],
      };
    }

    const parsed = await parseDocument(filePath);

    const chunks = chunkText(parsed);

    const vectors = await embedChunks(chunks);

    await saveToVectorDB(vectors, chunks, hash);

    const sp = await uploadBufferToSharePointIfConfigured(
      buffer,
      `${hash.slice(0, SHAREPOINT_FILENAME_HASH_PREFIX_LENGTH)}_${filename}`,
      {
        MICROSOFT_TENANT_ID,
        MICROSOFT_CLIENT_ID,
        SHAREPOINT_SITE_HOST,
        SHAREPOINT_SITE_PATH,
        SHAREPOINT_UPLOAD_BASE_PATH,
      }
    );
    const spNote = formatSharePointResultNote(sp);

    return {
      content: [
        {
          type: "text" as const,
          text: MSG_UPLOAD_SUCCESS_TEMPLATE(chunks.length, spNote),
        },
      ],
    };
  }
);

function formatSharePointResultNote(
  sp: Awaited<ReturnType<typeof uploadBufferToSharePointIfConfigured>>
): string {
  if (sp.skipped) {
    return "";
  }
  const parts: string[] = [];
  if (sp.webUrl) {
    parts.push(MSG_SHAREPOINT_DOC_URL_LINE(sp.webUrl));
  }
  if (sp.driveItemId) {
    parts.push(MSG_SHAREPOINT_DRIVE_ITEM_ID_LINE(sp.driveItemId));
  }
  const listItemId = sp.sharepointIds?.listItemId;
  if (listItemId) {
    parts.push(MSG_SHAREPOINT_LIST_ITEM_ID_LINE(listItemId));
  }
  const uniqueId = sp.sharepointIds?.listItemUniqueId;
  if (uniqueId) {
    parts.push(MSG_SHAREPOINT_LIST_ITEM_UNIQUE_ID_LINE(uniqueId));
  }
  if (parts.length > 0) {
    return parts.join("");
  }
  return SP_RESULT_NOTE_COMPLETED;
}

function isBase64(str: string): boolean {
  try {
    return Buffer.from(str, "base64").toString("base64") === str;
  } catch {
    return false;
  }
}

async function parseDocument(filePath: string): Promise<string> {
  const ext = path.extname(filePath);

  if (ext === ".txt" || ext === ".md") {
    return fs.readFile(filePath, "utf-8");
  }

  if (ext === ".pdf") {
    return MOCK_PARSED_PDF_CONTENT;
  }

  if (ext === ".docx") {
    return MOCK_PARSED_DOCX_CONTENT;
  }

  throw new Error(ERR_UNSUPPORTED_FORMAT);
}

function chunkText(
  text: string,
  size = TEXT_CHUNK_SIZE_CHARS,
  overlap = TEXT_CHUNK_OVERLAP_CHARS
): string[] {
  const chunks: string[] = [];

  for (let i = 0; i < text.length; i += size - overlap) {
    chunks.push(text.slice(i, i + size));
  }

  return chunks;
}

type EmbedItem = { vector: number[]; text: string };

async function embedChunks(chunks: string[]): Promise<EmbedItem[]> {
  return chunks.map((chunk) => ({
    vector: fakeEmbedding(chunk),
    text: chunk,
  }));
}

function fakeEmbedding(text: string): number[] {
  return Array.from({ length: FAKE_EMBEDDING_DIMENSION }, () => Math.random());
}

type VectorRow = {
  id: string;
  hash: string;
  vector: number[];
  text: string;
};

const inMemoryVectorDB: VectorRow[] = [];

async function saveToVectorDB(
  vectors: EmbedItem[],
  chunks: string[],
  hash: string
): Promise<void> {
  vectors.forEach((v, i) => {
    inMemoryVectorDB.push({
      id: crypto.randomUUID(),
      hash,
      vector: v.vector,
      text: chunks[i],
    });
  });
}

async function alreadyProcessed(hash: string): Promise<boolean> {
  return inMemoryVectorDB.some((x) => x.hash === hash);
}

const transport = new StdioServerTransport();
await server.connect(transport);
