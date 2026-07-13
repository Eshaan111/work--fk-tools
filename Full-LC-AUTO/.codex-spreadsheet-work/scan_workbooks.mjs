import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = process.argv[2];

async function walk(dir) {
  const out = [];
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    if (entry.name === "dist" || entry.name === ".git" || entry.name === "node_modules") continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await walk(full));
    else if (/\.xlsx$/i.test(entry.name)) out.push(full);
  }
  return out;
}

for (const file of await walk(root)) {
  try {
    const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
    const beige = await wb.inspect({
      kind: "match",
      searchTerm: "Beige-Baggy",
      options: { useRegex: false, maxResults: 100 },
      maxChars: 12000,
    });
    const white = await wb.inspect({
      kind: "match",
      searchTerm: "White-Baggy",
      options: { useRegex: false, maxResults: 100 },
      maxChars: 12000,
    });
    if (!beige.ndjson.includes('"matches":[]') || !white.ndjson.includes('"matches":[]')) {
      console.log(JSON.stringify({ file, beige: beige.ndjson, white: white.ndjson }));
    }
  } catch (error) {
    console.error(JSON.stringify({ file, error: String(error) }));
  }
}
