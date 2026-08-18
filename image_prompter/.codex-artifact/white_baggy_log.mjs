import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "C:/Users/ESHAAN/HAKUR/work--fk-tools/image_prompter/USED-IMAGE-DESIGNS.xlsx";
const previewPath = "C:/Users/ESHAAN/HAKUR/work--fk-tools/image_prompter/.codex-artifact/white-baggy-log-preview.png";
const mode = process.argv[2] ?? "inspect";

const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(0);
const values = sheet.getRange("A1:CR30").values;
const match = values.map((row, index) => ({ row: index + 1, values: row }))
  .find(({ values: row }) => String(row[0] ?? "").toUpperCase().includes("WHITE"));

if (!match) throw new Error("White Baggy row not found");

if (mode === "inspect") {
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
  console.log(JSON.stringify({ sheet: sheet.name, row: match.row, values: match.values, previewPath }, null, 2));
} else if (mode === "update") {
  const concepts = JSON.parse(process.argv[3]);
  const firstEmpty = match.values.findIndex((value, index) => index > 0 && (value === null || value === ""));
  if (firstEmpty < 0) throw new Error("No empty cells available in White Baggy row");
  sheet.getRangeByIndexes(match.row - 1, firstEmpty, 1, concepts.length).values = [concepts];

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(workbookPath);
  console.log(JSON.stringify({ row: match.row, firstEmpty, concepts, errors: errors.ndjson, previewPath }, null, 2));
} else if (mode === "verify") {
  console.log(JSON.stringify({ row: match.row, values: match.values }, null, 2));
} else {
  throw new Error(`Unknown mode: ${mode}`);
}
