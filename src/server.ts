import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  MCP_SERVER_DESCRIPTION,
  MCP_SERVER_NAME,
  MCP_SERVER_VERSION,
  MSG_UPLOAD_SUCCESS_TEMPLATE,
  SCHEMA_DESC_DOCUMENT_CONTENT,
  SCHEMA_DESC_DOCUMENT_FILENAME,
  SCHEMA_DESC_MICROSOFT_CLIENT_ID,
  SCHEMA_DESC_MICROSOFT_TENANT_ID,
  SCHEMA_DESC_SHAREPOINT_SITE_HOST,
  SCHEMA_DESC_SHAREPOINT_SITE_PATH,
  SCHEMA_DESC_SHAREPOINT_UPLOAD_BASE_PATH,
  TOOL_HELLO_WORLD_RESPONSE_TEXT,
  TOOL_HELLO_WORLD_TITLE,
  TOOL_NAME_HELLO_WORLD,
  TOOL_NAME_UPLOAD_DOCUMENT,
  TOOL_UPLOAD_DOCUMENT_DESCRIPTION,
  TOOL_UPLOAD_DOCUMENT_TITLE,
  TOOL_HELLO_WORLD_DESCRIPTION
} from "./constants.ts";
import { uploadDocument } from "./document-upload.ts";

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
    const spNote = await uploadDocument( document,
      MICROSOFT_TENANT_ID,
      MICROSOFT_CLIENT_ID,
      SHAREPOINT_SITE_HOST,
      SHAREPOINT_SITE_PATH,
      SHAREPOINT_UPLOAD_BASE_PATH,
    );
    return {
      content: [
        {
          type: "text" as const,
          text: MSG_UPLOAD_SUCCESS_TEMPLATE(spNote),
        },
      ],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
