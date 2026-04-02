import {
    ALLOWED_UPLOAD_EXTENSIONS,
    DOCUMENT_HASH_ALGORITHM,
    ERR_UNSUPPORTED_FILE_TYPE,
    LOCAL_UPLOAD_DIR,
    MSG_SHAREPOINT_DOC_URL_LINE,
    MSG_SHAREPOINT_DRIVE_ITEM_ID_LINE,
    MSG_SHAREPOINT_LIST_ITEM_ID_LINE,
    MSG_SHAREPOINT_LIST_ITEM_UNIQUE_ID_LINE,
    SHAREPOINT_FILENAME_HASH_PREFIX_LENGTH,
    SP_RESULT_NOTE_COMPLETED,
  } from "./constants.ts";
import { uploadBufferToSharePointIfConfigured } from "./sharepoint.ts";
import path from "node:path";
import fs from "node:fs/promises";
import crypto from "node:crypto";

export async function  uploadDocument(
    document: {
        filename: string;
        content: string;
    },
    MICROSOFT_TENANT_ID: string,
    MICROSOFT_CLIENT_ID: string,
    SHAREPOINT_SITE_HOST: string,
    SHAREPOINT_SITE_PATH: string,
    SHAREPOINT_UPLOAD_BASE_PATH: string
){
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
    return spNote;

}
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