import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const file = process.argv[2];
const outDir = process.argv[3];
await fs.mkdir(outDir, { recursive: true });
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const summary = await wb.inspect({
  kind: "workbook,sheet,table",
  maxChars: 12000,
  tableMaxRows: 8,
  tableMaxCols: 15,
  tableMaxCellChars: 100,
});
const beige = await wb.inspect({
  kind: "match",
  searchTerm: "Beige-Baggy",
  options: { useRegex: false, maxResults: 200 },
  maxChars: 16000,
});
console.log(JSON.stringify({ summary: summary.ndjson, beige: beige.ndjson }));
for (const sheet of wb.worksheets.items) {
  const preview = await wb.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outDir, `${sheet.name.replace(/[^a-z0-9_-]+/gi, "_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}
