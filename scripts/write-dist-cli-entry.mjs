import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "dist", "server.js");
writeFileSync(
  out,
  `#!/usr/bin/env node
import "./src/server.js";
`,
  "utf8"
);
