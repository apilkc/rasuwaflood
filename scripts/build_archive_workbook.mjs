import fs from "node:fs/promises";
import crypto from "node:crypto";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = process.argv[2] || "data/archive.json";
const outputPath = process.argv[3] || "outputs/rasuwa-archive-metadata/rasuwa-flood-source-metadata.xlsx";
const previewPath = process.argv[4] || "outputs/rasuwa-archive-metadata/preview.png";
const readmePreviewPath = previewPath.replace(/\.png$/i, "-readme.png");
const archive = JSON.parse(await fs.readFile(inputPath, "utf8"));
const items = archive.items || [];
const workbook = Workbook.create();
const summary = workbook.worksheets.add("Read Me");
const sheet = workbook.worksheets.add("Archive Metadata");
summary.showGridLines = false;
sheet.showGridLines = false;

summary.getRange("A1:J2").merge();
summary.getRange("A1").values = [["Rasuwa Flood — Source Metadata Archive"]];
summary.getRange("A1:J2").format = { fill: "#233B31", font: { bold: true, color: "#FFFFFF", size: 22 }, verticalAlignment: "center" };
summary.getRange("A4:B8").values = [
  ["Archive generated", archive.generated_at ? new Date(archive.generated_at) : "Not recorded"],
  ["Retained records", null],
  ["Update frequency", "Every three hours"],
  ["Public website", "https://rasuwaflood.org/archive.html"],
  ["Retention policy", "Metadata and source-provided excerpts are retained. Full copyrighted article text is not republished without permission or an open licence."],
];
summary.getRange("B4").format.numberFormat = "yyyy-mm-dd hh:mm";
summary.getRange("B5").formulas = [["=COUNTA('Archive Metadata'!A6:A1005)"]];
summary.getRange("A4:A8").format = { font: { bold: true, color: "#355747" }, fill: "#E9EEE9" };
summary.getRange("A4:B8").format.borders = { preset: "insideHorizontal", style: "thin", color: "#D9D7D1" };
summary.getRange("A4:A8").format.columnWidth = 24;
summary.getRange("B4:B8").format.columnWidth = 90;
summary.getRange("B8").format.wrapText = true;
summary.getRange("A10:J11").merge();
summary.getRange("A10").values = [["Fields include publication date, first-indexed timestamp, publisher, headline, retained excerpt, original URL, archived-copy lookup, retention status and rights note."]];
summary.getRange("A10:J11").format = { fill: "#F3EBDD", font: { color: "#5F4B32", italic: true }, wrapText: true, verticalAlignment: "center" };

const headers = ["Record ID", "Published date", "First indexed (UTC)", "Publisher", "Headline", "Retained text / excerpt", "Original URL", "Archived-copy lookup", "Content retention", "Rights note"];
sheet.getRange("A1:J2").merge();
sheet.getRange("A1").values = [["Rasuwa Flood source metadata"]];
sheet.getRange("A1:J2").format = { fill: "#233B31", font: { bold: true, color: "#FFFFFF", size: 20 }, verticalAlignment: "center" };
sheet.getRange("A3:J3").merge();
sheet.getRange("A3").values = [[`Generated ${archive.generated_at || ""} · ${items.length} records · original URLs remain the authoritative source`]];
sheet.getRange("A3:J3").format = { fill: "#E9EEE9", font: { color: "#355747", italic: true } };
sheet.getRange("A5:J5").values = [headers];
sheet.getRange("A5:J5").format = { fill: "#355747", font: { bold: true, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "center" };
const rows = items.map(item => {
  const url = String(item.url || "");
  return [
    item.record_id || crypto.createHash("sha256").update(url).digest("hex").slice(0, 16),
    item.date ? new Date(`${item.date}T00:00:00Z`) : null,
    item.first_seen_at || archive.generated_at ? new Date(item.first_seen_at || archive.generated_at) : null,
    item.source || "",
    item.title || "",
    item.retained_excerpt || item.summary || "",
    url,
    item.archive_lookup_url || `https://web.archive.org/web/*/${url}`,
    item.content_retention || "Metadata and source-provided excerpt retained",
    item.rights_note || "Full text is not republished unless reuse permission or an open licence is documented.",
  ];
});
if (rows.length) sheet.getRangeByIndexes(5, 0, rows.length, headers.length).values = rows;
sheet.getRange(`B6:B${Math.max(6, rows.length + 5)}`).format.numberFormat = "yyyy-mm-dd";
sheet.getRange(`C6:C${Math.max(6, rows.length + 5)}`).format.numberFormat = "yyyy-mm-dd hh:mm";
sheet.getRange(`A6:J${Math.max(6, rows.length + 5)}`).format = { verticalAlignment: "top", wrapText: false };
sheet.getRange(`D6:F${Math.max(6, rows.length + 5)}`).format.wrapText = true;
sheet.getRange(`I6:J${Math.max(6, rows.length + 5)}`).format.wrapText = true;
sheet.getRange(`A5:J${Math.max(6, rows.length + 5)}`).format.borders = { insideHorizontal: { style: "thin", color: "#E2E0DA" } };
const widths = [19, 15, 22, 24, 44, 56, 38, 38, 32, 42];
widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, rows.length + 5, 1).format.columnWidth = width);
sheet.getRange("A5:J5").format.rowHeight = 34;
if (rows.length) sheet.getRange(`A6:J${rows.length + 5}`).format.rowHeight = 48;
sheet.freezePanes.freezeRows(5);
sheet.freezePanes.freezeColumns(2);
if (rows.length) {
  const table = sheet.tables.add(`A5:J${rows.length + 5}`, true, "ArchiveMetadataTable");
  table.style = "TableStyleMedium4";
  table.showFilterButton = true;
}

await fs.mkdir(outputPath.slice(0, outputPath.lastIndexOf("/")), { recursive: true });
const inspect = await workbook.inspect({ kind: "table", range: "Archive Metadata!A1:J12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 10, maxChars: 7000 });
console.log(inspect.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "Archive Metadata", range: "A1:J14", scale: 0.8, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const readmePreview = await workbook.render({ sheetName: "Read Me", range: "A1:J11", scale: 1, format: "png" });
await fs.writeFile(readmePreviewPath, new Uint8Array(await readmePreview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}, ${previewPath}, and ${readmePreviewPath}`);
