import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const OUTPUT_HEADERS = ["时间", "分类", "二级分类", "类型", "金额", "备注", "标签 1", "标签 2"];
const REQUIRED_HEADERS = ["时间", "分类", "二级分类", "类型", "金额", "备注", "账单标记", "标签"];
const ALLOWED_TYPES = new Set(["收入", "支出"]);

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) throw new Error(`无法识别的参数：${token}`);
    const key = token.slice(2);
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) throw new Error(`参数 --${key} 缺少值`);
    args[key] = value;
    i += 1;
  }
  args.mode = args.mode ?? "clean";
  if (!new Set(["inspect", "clean"]).has(args.mode)) throw new Error("--mode 只能是 inspect 或 clean");
  if (!args.input) throw new Error("必须提供 --input");
  if (args.mode === "clean" && (!args.output || !args["start-date"] || !args["end-date"])) {
    throw new Error("clean 模式必须提供 --output、--start-date 和 --end-date");
  }
  return args;
}

function normalizeHeader(value) {
  return String(value ?? "").replace(/^\uFEFF/, "").trim();
}

function dateOnly(value, sourceRow) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Math.round((value - 25569) * 86400000));
    if (!Number.isNaN(date.getTime())) {
      return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    }
  }
  const match = String(value ?? "").trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/);
  if (!match) throw new Error(`第 ${sourceRow} 行「时间」无法解析：${String(value ?? "")}`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error(`第 ${sourceRow} 行「时间」不是有效日期：${String(value ?? "")}`);
  }
  return date;
}

function dateArgument(value, name) {
  const match = String(value ?? "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error(`${name} 必须使用 YYYY-MM-DD 格式`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error(`${name} 不是有效日期：${value}`);
  }
  return date;
}

function asAmount(value, sourceRow) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const normalized = String(value ?? "").trim().replace(/,/g, "");
  const amount = Number(normalized);
  if (!normalized || !Number.isFinite(amount)) {
    throw new Error(`第 ${sourceRow} 行「金额」不是有效数值：${String(value ?? "")}`);
  }
  return amount;
}

function splitTag(value, sourceRow) {
  const text = String(value ?? "").trim();
  const parts = text.split("#");
  if (parts.length > 2) {
    throw new Error(`第 ${sourceRow} 行「标签」包含超过 1 个 #：${text}`);
  }
  return [parts[0] ?? "", parts[1] ?? ""];
}

async function loadSource(inputPath) {
  const extension = path.extname(inputPath).toLowerCase();
  if (extension === ".csv") {
    const csvText = (await fs.readFile(inputPath, "utf8")).replace(/^\uFEFF/, "");
    return Workbook.fromCSV(csvText, { sheetName: "原始数据" });
  }
  if (extension === ".xlsx") {
    const input = await FileBlob.load(inputPath);
    return SpreadsheetFile.importXlsx(input);
  }
  throw new Error(`暂不支持 ${extension || "无扩展名"}；确定性脚本只接受 CSV 或 XLSX`);
}

const args = parseArgs(process.argv.slice(2));
const inputPath = path.resolve(args.input);
await fs.access(inputPath);

const sourceWorkbook = await loadSource(inputPath);
const sourceSheet = sourceWorkbook.worksheets.getItemAt(0);
const usedRange = sourceSheet.getUsedRange(true);
const matrix = usedRange?.values ?? [];
if (matrix.length < 1) throw new Error("输入表格为空");

const headers = matrix[0].map(normalizeHeader);
const indexes = Object.fromEntries(headers.map((header, index) => [header, index]));
const missing = REQUIRED_HEADERS.filter((header) => indexes[header] === undefined);
if (missing.length) throw new Error(`缺少必要列：${missing.join("、")}`);

const records = matrix.slice(1).map((row, index) => {
  const sourceRow = index + 2;
  return {
    row,
    sourceRow,
    sourceIndex: index + 1,
    date: dateOnly(row[indexes["时间"]], sourceRow),
    type: String(row[indexes["类型"]] ?? "").trim(),
    marked: Boolean(String(row[indexes["账单标记"]] ?? "").trim()),
  };
});

const sourceDates = records.map((record) => record.date).sort((left, right) => left - right);
const eligibleRecords = records.filter((record) => !record.marked && ALLOWED_TYPES.has(record.type));
const eligibleDates = eligibleRecords.map((record) => record.date).sort((left, right) => left - right);
const typeCounts = {};
for (const record of records) typeCounts[record.type] = (typeCounts[record.type] ?? 0) + 1;

if (args.mode === "inspect") {
  console.log(JSON.stringify({
    inputPath,
    sourceRows: records.length,
    headers,
    sourceFirstDate: sourceDates[0]?.toISOString().slice(0, 10) ?? null,
    sourceLastDate: sourceDates.at(-1)?.toISOString().slice(0, 10) ?? null,
    markedRows: records.filter((record) => record.marked).length,
    typeCounts,
    removedOtherTypeRows: records.filter((record) => !record.marked && !ALLOWED_TYPES.has(record.type)).length,
    eligibleRows: eligibleRecords.length,
    eligibleFirstDate: eligibleDates[0]?.toISOString().slice(0, 10) ?? null,
    eligibleLastDate: eligibleDates.at(-1)?.toISOString().slice(0, 10) ?? null,
    dateRangePromptExample: "2026-07-01 至 2026-07-31（包含起止日期）",
  }, null, 2));
  process.exit(0);
}

const outputPath = path.resolve(args.output);
if (inputPath === outputPath) throw new Error("输出路径不得与源文件相同");
try {
  await fs.access(outputPath);
  throw new Error(`输出文件已存在，拒绝覆盖：${outputPath}`);
} catch (error) {
  if (error?.code !== "ENOENT") throw error;
}
const startDate = dateArgument(args["start-date"], "--start-date");
const endDate = dateArgument(args["end-date"], "--end-date");
if (startDate > endDate) throw new Error("--start-date 不得晚于 --end-date");

let removedMarkedRows = 0;
let removedOtherTypeRows = 0;
let removedOutsideDateRows = 0;
const cleaned = [];
for (const record of records) {
  const { row, sourceRow, sourceIndex, date, type, marked } = record;
  if (marked) {
    removedMarkedRows += 1;
    continue;
  }
  if (!ALLOWED_TYPES.has(type)) {
    removedOtherTypeRows += 1;
    continue;
  }
  if (date < startDate || date > endDate) {
    removedOutsideDateRows += 1;
    continue;
  }
  const [tag1, tag2] = splitTag(row[indexes["标签"]], sourceRow);
  cleaned.push({
    date,
    sourceIndex,
    values: [
      date,
      String(row[indexes["分类"]] ?? ""),
      String(row[indexes["二级分类"]] ?? ""),
      type,
      asAmount(row[indexes["金额"]], sourceRow),
      String(row[indexes["备注"]] ?? ""),
      tag1,
      tag2,
    ],
  });
}

if (!cleaned.length) throw new Error("指定日期范围内没有可保留的收入或支出记录");

cleaned.sort((left, right) => left.date - right.date || left.sourceIndex - right.sourceIndex);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("清理结果");
const values = [OUTPUT_HEADERS, ...cleaned.map((item) => item.values)];
const lastRow = values.length;
sheet.getRange(`A1:H${lastRow}`).values = values;
sheet.getRange("A1:H1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
  verticalAlignment: "center",
};
sheet.getRange("A1:H1").format.rowHeight = 24;
sheet.getRange(`A2:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
sheet.getRange(`E2:E${lastRow}`).format.numberFormat = "0.00";
sheet.getRange(`A2:A${lastRow}`).format.horizontalAlignment = "center";
sheet.getRange(`D2:D${lastRow}`).format.horizontalAlignment = "center";
sheet.getRange(`A1:H${lastRow}`).format.verticalAlignment = "center";
sheet.getRange(`A1:H${lastRow}`).format.autofitColumns();
sheet.getRange("A:A").format.columnWidth = 13;
sheet.getRange("B:D").format.columnWidth = 14;
sheet.getRange("E:E").format.columnWidth = 12;
sheet.getRange("F:F").format.columnWidth = 28;
sheet.getRange("G:H").format.columnWidth = 18;
sheet.getRange(`F2:F${lastRow}`).format.wrapText = true;
sheet.freezePanes.freezeRows(1);
sheet.showGridLines = false;
const table = sheet.tables.add(`A1:H${lastRow}`, true, "CleanedBillsTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

if (args.preview) {
  const preview = await workbook.render({
    sheetName: "清理结果",
    range: `A1:H${Math.min(lastRow, 25)}`,
    scale: 1.5,
    format: "png",
  });
  await fs.mkdir(path.dirname(path.resolve(args.preview)), { recursive: true });
  await fs.writeFile(path.resolve(args.preview), new Uint8Array(await preview.arrayBuffer()));
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await fs.rm(`${outputPath}.inspect.ndjson`, { force: true });

const sourceRows = Math.max(0, matrix.length - 1);
const dates = cleaned.map((item) => item.date.toISOString().slice(0, 10));
console.log(JSON.stringify({
  inputPath,
  outputPath,
  sourceRows,
  removedMarkedRows,
  removedOtherTypeRows,
  removedOutsideDateRows,
  outputRows: cleaned.length,
  startDate: startDate.toISOString().slice(0, 10),
  endDate: endDate.toISOString().slice(0, 10),
  outputHeaders: OUTPUT_HEADERS,
  firstDate: dates[0] ?? null,
  lastDate: dates.at(-1) ?? null,
}, null, 2));
