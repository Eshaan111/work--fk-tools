import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = process.argv[2];
const outputFile = process.argv[3];
const previewFile = process.argv[4];

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputFile));
const sheet = wb.worksheets.getItem("Product Inputs");
const source = sheet.getRange("A2:AA6");
const destination = sheet.getRange("A32:AA36");

const before = await wb.inspect({
  kind: "table,formula,computedStyle",
  sheetId: "Product Inputs",
  range: "A1:AA42",
  maxChars: 16000,
  tableMaxRows: 42,
  tableMaxCols: 27,
  options: { maxResults: 200 },
});

destination.copyFrom(source, "all");
sheet.getRange("A32:A36").values = Array.from({ length: 5 }, () => ["WHITE"]);
sheet.getRange("C32:C36").values = source.values.map((row) => [String(row[2]).replace(/Beige-Baggy/gi, "White-Baggy")]);

const verification = await wb.inspect({
  kind: "table",
  sheetId: "Product Inputs",
  range: "A30:AA38",
  include: "values,formulas",
  maxChars: 12000,
  tableMaxRows: 9,
  tableMaxCols: 27,
});
const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 6000,
});

await fs.mkdir(path.dirname(outputFile), { recursive: true });
const preview = await wb.render({ sheetName: "Product Inputs", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(previewFile, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputFile);

console.log(JSON.stringify({ before: before.ndjson, verification: verification.ndjson, errors: errors.ndjson }));
