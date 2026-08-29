#!/usr/bin/env python3
"""钱迹账单检查、清理，以及安全导入现有财务统计表。仅依赖 Python 标准库。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"m": MAIN, "r": REL, "p": PKG_REL}
ET.register_namespace("", MAIN)
ET.register_namespace("r", REL)

CLEAN_HEADERS = ["时间", "分类", "二级分类", "类型", "金额", "备注", "标签 1", "标签 2"]
REQUIRED_HEADERS = ["时间", "分类", "二级分类", "类型", "金额", "备注", "账单标记", "标签"]
TARGET_HEADERS = ["日期", "一级分类ｼｭｯﾋﾟ", "二级分类", "类型", "金额", "备注", "支出类型", "支出必要性"]
TARGET_HEADER_ALIASES = [
    {"日期", "时间"}, {"一级分类ｼｭｯﾋﾟ", "一级分类", "分类"}, {"二级分类"}, {"类型"},
    {"金额"}, {"备注"}, {"支出类型", "标签 1", "标签1"}, {"支出必要性", "标签 2", "标签2"},
]
ALLOWED_TYPES = {"收入", "支出"}
BACKUP_RETENTION_DAYS = 30


def fail(message: str) -> None:
    raise ValueError(message)


def norm(value) -> str:
    return str(value if value is not None else "").lstrip("\ufeff").strip()


def parse_day(value, label: str) -> date:
    text = norm(value)
    try:
        parsed = datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        fail(f"{label} 无法解析为 YYYY-MM-DD 日期：{text}")
    if len(text) < 10 or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text[:10]):
        fail(f"{label} 无法解析为 YYYY-MM-DD 日期：{text}")
    return parsed


def excel_day(serial: str, date1904: bool = False) -> date:
    try:
        number = float(serial)
    except ValueError:
        fail(f"Excel 日期序列无效：{serial}")
    origin = date(1904, 1, 1) if date1904 else date(1899, 12, 30)
    return origin + timedelta(days=int(number))


def excel_serial(day: date, date1904: bool = False) -> int:
    origin = date(1904, 1, 1) if date1904 else date(1899, 12, 30)
    return (day - origin).days


def parse_amount(value, row_number: int) -> Decimal:
    text = norm(value).replace(",", "")
    try:
        amount = Decimal(text)
    except InvalidOperation:
        fail(f"第 {row_number} 行「金额」不是有效数值：{value}")
    if not amount.is_finite():
        fail(f"第 {row_number} 行「金额」不是有限数值：{value}")
    return amount


def split_tag(value, row_number: int) -> tuple[str, str]:
    text = norm(value)
    parts = text.split("#")
    if len(parts) > 2:
        fail(f"第 {row_number} 行「标签」包含超过 1 个 #：{text}")
    return parts[0], parts[1] if len(parts) == 2 else ""


def col_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref)
    if not letters:
        fail(f"无法解析单元格地址：{cell_ref}")
    value = 0
    for char in letters.group(0):
        value = value * 26 + ord(char) - 64
    return value


def first_sheet_path(book: ZipFile) -> tuple[str, bool]:
    workbook = ET.fromstring(book.read("xl/workbook.xml"))
    props = workbook.find("m:workbookPr", NS)
    date1904 = props is not None and props.attrib.get("date1904") in {"1", "true"}
    first = workbook.find("m:sheets/m:sheet", NS)
    if first is None:
        fail("XLSX 中没有工作表")
    rid = first.attrib[f"{{{REL}}}id"]
    rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    target = next((x.attrib["Target"] for x in rels if x.attrib.get("Id") == rid), None)
    if not target:
        fail("无法找到第一个工作表文件")
    return "xl/" + target.lstrip("/").removeprefix("xl/"), date1904


def sheet_path_by_name(book: ZipFile, name: str) -> tuple[str, bool]:
    workbook = ET.fromstring(book.read("xl/workbook.xml"))
    props = workbook.find("m:workbookPr", NS)
    date1904 = props is not None and props.attrib.get("date1904") in {"1", "true"}
    item = next((x for x in workbook.findall("m:sheets/m:sheet", NS) if x.attrib.get("name") == name), None)
    if item is None:
        fail(f"目标工作簿缺少工作表「{name}」")
    rid = item.attrib[f"{{{REL}}}id"]
    rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    target = next((x.attrib["Target"] for x in rels if x.attrib.get("Id") == rid), None)
    if not target:
        fail(f"无法定位工作表「{name}」")
    return "xl/" + target.lstrip("/").removeprefix("xl/"), date1904


def shared_strings(book: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in book.namelist():
        return []
    root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.iter(f"{{{MAIN}}}t")) for item in root]


def cell_value(cell: ET.Element, strings: list[str]):
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        inline = cell.find("m:is", NS)
        return "" if inline is None else "".join(node.text or "" for node in inline.iter(f"{{{MAIN}}}t"))
    value = cell.find("m:v", NS)
    if value is None or value.text is None:
        return ""
    if kind == "s":
        return strings[int(value.text)]
    if kind == "b":
        return value.text == "1"
    return value.text


def xlsx_matrix(path: Path, sheet_name: str | None = None) -> tuple[list[list], bool]:
    with ZipFile(path) as book:
        sheet_path, date1904 = sheet_path_by_name(book, sheet_name) if sheet_name else first_sheet_path(book)
        strings = shared_strings(book)
        root = ET.fromstring(book.read(sheet_path))
        rows = root.findall("m:sheetData/m:row", NS)
        max_col = max((col_number(cell.attrib["r"]) for row in rows for cell in row.findall("m:c", NS)), default=0)
        matrix = []
        for row in rows:
            values = [""] * max_col
            for cell in row.findall("m:c", NS):
                values[col_number(cell.attrib["r"]) - 1] = cell_value(cell, strings)
            matrix.append(values)
        return matrix, date1904


def read_source(path: Path) -> tuple[list[str], list[dict]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            matrix = list(csv.reader(handle))
        date1904 = False
    elif suffix == ".xlsx":
        matrix, date1904 = xlsx_matrix(path)
    else:
        fail(f"暂不支持 {suffix or '无扩展名'}；只接受 UTF-8 CSV 或 XLSX")
    if not matrix:
        fail("输入表格为空")
    headers = [norm(item) for item in matrix[0]]
    index = {header: pos for pos, header in enumerate(headers)}
    missing = [header for header in REQUIRED_HEADERS if header not in index]
    if missing:
        fail("缺少必要列：" + "、".join(missing))
    records = []
    for offset, row in enumerate(matrix[1:], start=2):
        row += [""] * (len(headers) - len(row))
        raw_date = row[index["时间"]]
        if isinstance(raw_date, str) and re.fullmatch(r"-?\d+(?:\.\d+)?", raw_date.strip()) and suffix == ".xlsx":
            day = excel_day(raw_date, date1904)
        else:
            day = parse_day(raw_date, f"第 {offset} 行「时间」")
        records.append({
            "row": row, "row_number": offset, "day": day,
            "type": norm(row[index["类型"]]),
            "marked": bool(norm(row[index["账单标记"]])), "index": index,
        })
    return headers, records


def inspect_source(path: Path) -> dict:
    headers, records = read_source(path)
    eligible = [r for r in records if not r["marked"] and r["type"] in ALLOWED_TYPES]
    source_days = sorted(r["day"] for r in records)
    eligible_days = sorted(r["day"] for r in eligible)
    types = Counter(r["type"] for r in records)
    return {
        "inputPath": str(path.resolve()), "sourceRows": len(records), "headers": headers,
        "sourceFirstDate": source_days[0].isoformat() if source_days else None,
        "sourceLastDate": source_days[-1].isoformat() if source_days else None,
        "markedRows": sum(r["marked"] for r in records), "typeCounts": dict(types),
        "removedOtherTypeRows": sum(not r["marked"] and r["type"] not in ALLOWED_TYPES for r in records),
        "eligibleRows": len(eligible),
        "eligibleFirstDate": eligible_days[0].isoformat() if eligible_days else None,
        "eligibleLastDate": eligible_days[-1].isoformat() if eligible_days else None,
        "dateRangePromptExample": "2026-07-01 至 2026-07-31（包含起止日期）",
    }


def clean_records(path: Path, start: date, end: date) -> tuple[list[list], dict]:
    if start > end:
        fail("起始日期不得晚于结束日期")
    _, records = read_source(path)
    cleaned = []
    removed_marked = removed_type = removed_date = 0
    for record in records:
        row, index = record["row"], record["index"]
        if record["marked"]:
            removed_marked += 1
            continue
        if record["type"] not in ALLOWED_TYPES:
            removed_type += 1
            continue
        if not start <= record["day"] <= end:
            removed_date += 1
            continue
        tag1, tag2 = split_tag(row[index["标签"]], record["row_number"])
        cleaned.append([
            record["day"], str(row[index["分类"]] or ""), str(row[index["二级分类"]] or ""),
            record["type"], parse_amount(row[index["金额"]], record["row_number"]),
            str(row[index["备注"]] or ""), tag1, tag2, record["row_number"],
        ])
    if not cleaned:
        fail("指定日期范围内没有可保留的收入或支出记录")
    cleaned.sort(key=lambda item: (item[0], item[8]))
    rows = [item[:8] for item in cleaned]
    stats = {
        "sourceRows": len(records), "removedMarkedRows": removed_marked,
        "removedOtherTypeRows": removed_type, "removedOutsideDateRows": removed_date,
        "outputRows": len(rows), "startDate": start.isoformat(), "endDate": end.isoformat(),
        "outputHeaders": CLEAN_HEADERS, "firstDate": rows[0][0].isoformat(), "lastDate": rows[-1][0].isoformat(),
    }
    if len(records) != removed_marked + removed_type + removed_date + len(rows):
        fail("内部行数守恒校验失败")
    return rows, stats


def xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def parse_xml_for_rewrite(payload: bytes) -> tuple[ET.Element, dict[str, str]]:
    namespaces: dict[str, str] = {}
    for _, item in ET.iterparse(io.BytesIO(payload), events=("start-ns",)):
        prefix, uri = item
        prefix = prefix or ""
        namespaces[prefix] = uri
        if prefix not in {"xml", "xmlns"} and not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)
    return ET.fromstring(payload), namespaces


def serialize_rewritten_xml(root: ET.Element, namespaces: dict[str, str], part_name: str) -> bytes:
    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    text = payload.decode("utf-8")
    root_match = re.search(r"<(?:[A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*(?=[\s>])", text)
    if root_match is None:
        fail(f"无法定位 XML 根元素：{part_name}")
    root_end = text.find(">", root_match.start())
    if root_end < 0:
        fail(f"XML 根元素未闭合：{part_name}")
    root_tag = text[root_match.start():root_end + 1]
    declared = set(re.findall(r"xmlns:([A-Za-z_][\w.-]*)=", root_tag))
    additions = []
    for prefix, uri in namespaces.items():
        if prefix and prefix not in {"xml", "xmlns"} and prefix not in declared:
            additions.append(f' xmlns:{prefix}="{xml_escape(uri)}"')
    if additions:
        insert_at = root_end - 1 if text[root_end - 1] == "/" else root_end
        text = text[:insert_at] + "".join(additions) + text[insert_at:]
        root_end += sum(len(item) for item in additions)
        root_tag = text[root_match.start():root_end + 1]
        declared = set(re.findall(r"xmlns:([A-Za-z_][\w.-]*)=", root_tag))
    ignorable = re.findall(r"(?:[A-Za-z_][\w.-]*:)?Ignorable=\"([^\"]+)\"", root_tag)
    missing = sorted(set(" ".join(ignorable).split()) - declared)
    if missing:
        fail(f"{part_name} 的 mc:Ignorable 存在未声明前缀：{', '.join(missing)}")
    return text.encode("utf-8")


def element_signature(element: ET.Element):
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        element.text or "",
        tuple(element_signature(child) for child in list(element)),
    )


def cell_xml(ref: str, value, style: int) -> str:
    if isinstance(value, date):
        return f'<c r="{ref}" s="{style}"><v>{excel_serial(value)}</v></c>'
    if isinstance(value, Decimal):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = str(value)
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{preserve}>{xml_escape(text)}</t></is></c>'


def write_clean_xlsx(path: Path, rows: list[list]) -> None:
    if path.exists():
        fail(f"输出文件已存在，拒绝覆盖：{path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    row_xml = []
    header_cells = "".join(cell_xml(f"{chr(65+i)}1", value, 1) for i, value in enumerate(CLEAN_HEADERS))
    row_xml.append(f'<row r="1" ht="24" customHeight="1">{header_cells}</row>')
    for row_num, row in enumerate(rows, start=2):
        cells = []
        for col, value in enumerate(row):
            style = 2 if col == 0 else 3 if col == 4 else 0
            cells.append(cell_xml(f"{chr(65+col)}{row_num}", value, style))
        row_xml.append(f'<row r="{row_num}">{"".join(cells)}</row>')
    last = len(rows) + 1
    sheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{MAIN}"><dimension ref="A1:H{last}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="13" customWidth="1"/><col min="2" max="5" width="14" customWidth="1"/><col min="6" max="6" width="28" customWidth="1"/><col min="7" max="8" width="18" customWidth="1"/></cols><sheetData>{''.join(row_xml)}</sheetData><autoFilter ref="A1:H{last}"/></worksheet>'''
    files = {
        "[Content_Types].xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>''',
        "_rels/.rels": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''',
        "xl/workbook.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="清理结果" sheetId="1" r:id="rId1"/></sheets><calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>''',
        "xl/_rels/workbook.xml.rels": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/styles.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{MAIN}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles></styleSheet>''',
    }
    with ZipFile(path, "w", ZIP_DEFLATED) as book:
        for name, content in files.items():
            book.writestr(name, content.encode("utf-8"))


def target_rows(path: Path, sheet_name: str) -> tuple[list[dict], int | None, list[str]]:
    matrix, date1904 = xlsx_matrix(path, sheet_name)
    if not matrix:
        fail(f"工作表「{sheet_name}」为空")
    headers = [norm(item) for item in matrix[0][:8]]
    if len(headers) < 8 or any(headers[i] not in TARGET_HEADER_ALIASES[i] for i in range(8)):
        fail(f"目标表头不匹配；实际为：{headers}")
    nonblank = [i + 1 for i, row in enumerate(matrix[1:], start=1) if any(norm(v) for v in row[:8])]
    reserved_start = None
    previous = 1
    for row_num in nonblank:
        if row_num - previous >= 100:
            reserved_start = row_num
            break
        previous = row_num
    limit = reserved_start or 1048577
    records = []
    for row_num, row in enumerate(matrix[1:limit - 1], start=2):
        row += [""] * (8 - len(row))
        if not any(norm(v) for v in row[:8]):
            continue
        raw_day = row[0]
        day = excel_day(str(raw_day), date1904) if re.fullmatch(r"-?\d+(?:\.\d+)?", norm(raw_day)) else parse_day(raw_day, f"目标表第 {row_num} 行日期")
        values = [
            day, str(row[1] or ""), str(row[2] or ""), norm(row[3]),
            parse_amount(row[4], row_num), str(row[5] or ""), str(row[6] or ""), str(row[7] or ""),
        ]
        records.append({"row": row_num, "day": day, "values": values})
    return records, reserved_start, headers


def inspect_target(path: Path, sheet_name: str) -> dict:
    records, reserved_start, headers = target_rows(path, sheet_name)
    next_row = max((item["row"] for item in records), default=1) + 1
    capacity_end = (reserved_start - 1) if reserved_start else 1048576
    return {
        "targetPath": str(path.resolve()), "sheetName": sheet_name, "headers": headers,
        "existingRows": len(records), "nextAppendRow": next_row,
        "reservedStartRow": reserved_start, "capacityEndRow": capacity_end,
        "availableRows": capacity_end - next_row + 1,
        "firstDate": min((r["day"] for r in records), default=None).isoformat() if records else None,
        "lastDate": max((r["day"] for r in records), default=None).isoformat() if records else None,
    }


def analyze_import(existing: list[dict], rows: list[list]) -> dict:
    incoming_first = min(row[0] for row in rows)
    incoming_last = max(row[0] for row in rows)
    existing_first = min((item["day"] for item in existing), default=None)
    existing_last = max((item["day"] for item in existing), default=None)
    incoming_days = {row[0] for row in rows}
    overlaps = [item for item in existing if item["day"] in incoming_days]
    gap_days = 0
    historical = False
    if existing_last is not None:
        if incoming_first > existing_last:
            gap_days = max(0, (incoming_first - existing_last).days - 1)
        else:
            historical = True
    return {
        "existingFirstDate": existing_first,
        "existingLastDate": existing_last,
        "incomingFirstDate": incoming_first,
        "incomingLastDate": incoming_last,
        "gapDays": gap_days,
        "historicalBackfill": historical,
        "overlaps": overlaps,
    }


def read_cleaned(path: Path) -> list[list]:
    matrix, date1904 = xlsx_matrix(path)
    if not matrix:
        fail("清理版表格为空")
    headers = [norm(value) for value in matrix[0][:8]]
    if headers != CLEAN_HEADERS:
        fail(f"清理版表头错误：{headers}")
    rows = []
    for row_number, row in enumerate(matrix[1:], start=2):
        row += [""] * (8 - len(row))
        if not any(norm(value) for value in row[:8]):
            continue
        raw_day = row[0]
        day = excel_day(str(raw_day), date1904) if re.fullmatch(r"-?\d+(?:\.\d+)?", norm(raw_day)) else parse_day(raw_day, f"清理版第 {row_number} 行日期")
        row_type = norm(row[3])
        if row_type not in ALLOWED_TYPES:
            fail(f"清理版第 {row_number} 行类型无效：{row_type}")
        rows.append([
            day, str(row[1] or ""), str(row[2] or ""), row_type,
            parse_amount(row[4], row_number), str(row[5] or ""), str(row[6] or ""), str(row[7] or ""),
        ])
    if not rows:
        fail("清理版没有可导入记录")
    return rows


def clear_cell_value(cell: ET.Element) -> None:
    for child in list(cell):
        if child.tag in {f"{{{MAIN}}}f", f"{{{MAIN}}}v", f"{{{MAIN}}}is"}:
            cell.remove(child)
    cell.attrib.pop("t", None)


def set_inline(cell: ET.Element, value) -> None:
    clear_cell_value(cell)
    if isinstance(value, date):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, f"{{{MAIN}}}v").text = str(excel_serial(value))
    elif isinstance(value, Decimal):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, f"{{{MAIN}}}v").text = str(value)
    else:
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, f"{{{MAIN}}}is")
        text = ET.SubElement(inline, f"{{{MAIN}}}t")
        value = str(value)
        if value != value.strip():
            text.attrib[f"{{{XML}}}space"] = "preserve"
        text.text = value


def formatting_snapshot(path: Path, sheet_name: str) -> dict:
    with ZipFile(path) as book:
        sheet_path, _ = sheet_path_by_name(book, sheet_name)
        root = ET.fromstring(book.read(sheet_path))
        style_hash = hashlib.sha256(book.read("xl/styles.xml")).hexdigest()
        row_formatting = []
        cell_styles = []
        for row in root.findall("m:sheetData/m:row", NS):
            row_formatting.append((row.attrib.get("r"), tuple(sorted(row.attrib.items()))))
            for cell in row.findall("m:c", NS):
                cell_styles.append((cell.attrib.get("r"), cell.attrib.get("s")))
        protected_parts = {}
        for tag in [
            "sheetFormatPr", "cols", "mergeCells", "conditionalFormatting",
            "dataValidations", "sheetProtection", "drawing", "legacyDrawing",
        ]:
            protected_parts[tag] = [
                element_signature(item) for item in root.findall(f"m:{tag}", NS)
            ]
        return {
            "stylesXmlSha256": style_hash,
            "rowFormatting": row_formatting,
            "cellStyles": cell_styles,
            "protectedParts": protected_parts,
        }


def build_validated_target(
    target: Path,
    output: Path,
    rows: list[list],
    sheet_name: str,
    overlap_policy: str,
    gap_policy: str,
    historical_policy: str,
) -> dict:
    if output.exists():
        fail(f"目标输出已存在，拒绝覆盖：{output}")
    if target.resolve() == output.resolve():
        fail("校验阶段的临时输出路径不得与原统计表相同")
    formatting_before = formatting_snapshot(target, sheet_name)
    existing, reserved_start, _ = target_rows(target, sheet_name)
    analysis = analyze_import(existing, rows)
    overlaps = analysis["overlaps"]
    if overlaps and overlap_policy == "stop":
        days = sorted({item["day"] for item in overlaps})
        fail(f"目标表已有 {len(overlaps)} 行落在导入日期中（{days[0]} 至 {days[-1]}）；请明确使用 --overlap-policy replace 或 append")
    if analysis["gapDays"] and gap_policy == "stop":
        fail(
            f"目标表最新日期为 {analysis['existingLastDate']}，本次数据从 {analysis['incomingFirstDate']} 开始，"
            f"中间相隔 {analysis['gapDays']} 天；日期空档不等于缺少账单，请确认后使用 --gap-policy continue"
        )
    if analysis["historicalBackfill"] and historical_policy == "stop":
        fail(
            f"本次数据日期为 {analysis['incomingFirstDate']} 至 {analysis['incomingLastDate']}，"
            f"目标表最新日期为 {analysis['existingLastDate']}；这是历史补录，请确认后使用 --historical-policy sort"
        )

    remove_rows = {item["row"] for item in overlaps} if overlap_policy == "replace" else set()
    kept = [item for item in existing if item["row"] not in remove_rows]
    historical_sort = bool(analysis["historicalBackfill"] and historical_policy == "sort")
    if historical_sort:
        combined = [(item["day"], 0, item["row"], item["values"]) for item in kept]
        combined.extend((values[0], 1, index, values) for index, values in enumerate(rows))
        combined.sort(key=lambda item: (item[0], item[1], item[2]))
        values_to_write = [item[3] for item in combined]
        first_written_row = 2
        rows_to_clear = {item["row"] for item in existing}
    else:
        values_to_write = rows
        first_written_row = max((item["row"] for item in kept), default=1) + 1
        rows_to_clear = remove_rows

    capacity_end = reserved_start - 1 if reserved_start else 1048576
    final_count = len(kept) + len(rows)
    required_last_row = (1 + final_count) if historical_sort else (first_written_row + len(rows) - 1)
    if required_last_row > capacity_end:
        fail(f"目标表可写空间不足：导入后需要写至第 {required_last_row} 行，可写区截止第 {capacity_end} 行")

    with ZipFile(target) as source:
        sheet_path, _ = sheet_path_by_name(source, sheet_name)
        root, sheet_namespaces = parse_xml_for_rewrite(source.read(sheet_path))
        data = root.find("m:sheetData", NS)
        if data is None:
            fail("目标工作表缺少 sheetData")
        row_map = {int(row.attrib["r"]): row for row in data.findall("m:row", NS)}
        for row_num in rows_to_clear:
            target_row = row_map.get(row_num)
            if target_row is None:
                continue
            for cell in target_row.findall("m:c", NS):
                if col_number(cell.attrib["r"]) <= 8:
                    clear_cell_value(cell)

        for offset, values in enumerate(values_to_write):
            row_num = first_written_row + offset
            target_row = row_map.get(row_num)
            if target_row is None:
                fail(f"目标表第 {row_num} 行不存在预设格式，无法保证仅粘贴值")
            cells = {col_number(cell.attrib["r"]): cell for cell in target_row.findall("m:c", NS)}
            for col, value in enumerate(values, start=1):
                cell = cells.get(col)
                if cell is None:
                    fail(f"目标单元格 {chr(64 + col)}{row_num} 不存在预设格式，无法保证仅粘贴值")
                set_inline(cell, value)
        data[:] = sorted(data, key=lambda row: int(row.attrib["r"]))
        sheet_bytes = serialize_rewritten_xml(root, sheet_namespaces, sheet_path)

        workbook, workbook_namespaces = parse_xml_for_rewrite(source.read("xl/workbook.xml"))
        calc = workbook.find("m:calcPr", NS)
        if calc is None:
            calc = ET.SubElement(workbook, f"{{{MAIN}}}calcPr")
        calc.attrib.update({"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"})
        workbook_bytes = serialize_rewritten_xml(workbook, workbook_namespaces, "xl/workbook.xml")

        replacements = {sheet_path: sheet_bytes, "xl/workbook.xml": workbook_bytes}
        for name in source.namelist():
            if name.startswith("xl/pivotCache/pivotCacheDefinition") and name.endswith(".xml"):
                cache, cache_namespaces = parse_xml_for_rewrite(source.read(name))
                cache.attrib["refreshOnLoad"] = "1"
                replacements[name] = serialize_rewritten_xml(cache, cache_namespaces, name)
        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w") as destination:
            for item in source.infolist():
                destination.writestr(item, replacements.get(item.filename, source.read(item.filename)))

    formatting_after = formatting_snapshot(output, sheet_name)
    if formatting_after != formatting_before:
        output.unlink(missing_ok=True)
        fail("导入前后格式签名不一致；已删除输出，未交付可能改变格式的文件")
    verification = inspect_target(output, sheet_name)
    if verification["existingRows"] != final_count:
        output.unlink(missing_ok=True)
        fail(f"导入后回读行数不符：预期 {final_count}，实际 {verification['existingRows']}")
    if historical_sort:
        verified_rows, _, _ = target_rows(output, sheet_name)
        verified_days = [item["day"] for item in verified_rows]
        if verified_days != sorted(verified_days):
            output.unlink(missing_ok=True)
            fail("历史补录后日期升序回读校验失败")
    return {
        "targetInputPath": str(target.resolve()), "targetOutputPath": str(output.resolve()),
        "sheetName": sheet_name, "overlapPolicy": overlap_policy,
        "gapPolicy": gap_policy, "historicalPolicy": historical_policy,
        "gapDays": analysis["gapDays"], "historicalBackfill": analysis["historicalBackfill"],
        "replacedRows": len(remove_rows), "importedRows": len(rows),
        "firstWrittenRow": first_written_row, "lastWrittenRow": first_written_row + len(values_to_write) - 1,
        "existingRowsAfterImport": verification["existingRows"], "chronologicallySorted": historical_sort,
        "pivotRefreshOnLoad": True, "sheetProtectionPreserved": True,
        "pasteValuesOnly": True, "formattingSignaturePreserved": True, "sourceWorkbookPreserved": True,
    }


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def default_backup_path(target: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = target.with_name(f"{target.stem}_写入前备份_{stamp}{target.suffix}")
    counter = 2
    while candidate.exists() or backup_trash_path(candidate, date.today()).exists():
        candidate = target.with_name(f"{target.stem}_写入前备份_{stamp}_{counter}{target.suffix}")
        counter += 1
    return candidate


def backup_trash_path(backup: Path, archived_on: date) -> Path:
    trash = backup.parent / ".trash"
    return trash / f"{archived_on.isoformat()}_{backup.name}"


def archive_verified_backup(backup: Path, archived_on: date) -> Path:
    archived = backup_trash_path(backup, archived_on)
    archived.parent.mkdir(parents=True, exist_ok=True)
    if archived.exists():
        fail(f"安全删除目标已存在，拒绝覆盖：{archived}")
    os.replace(backup, archived)
    return archived


def cleanup_expired_backups(trash: Path, target: Path, today: date) -> list[str]:
    if not trash.exists():
        return []
    pattern = re.compile(
        rf"^(\d{{4}}-\d{{2}}-\d{{2}})_{re.escape(target.stem)}_写入前备份_"
        rf"\d{{8}}_\d{{6}}(?:_\d+)?{re.escape(target.suffix)}$"
    )
    removed = []
    for candidate in trash.iterdir():
        if not candidate.is_file():
            continue
        match = pattern.fullmatch(candidate.name)
        if not match:
            continue
        archived_on = datetime.strptime(match.group(1), "%Y-%m-%d").date()
        if (today - archived_on).days <= BACKUP_RETENTION_DAYS:
            continue
        candidate.unlink()
        removed.append(str(candidate.resolve()))
    return removed


def write_target_in_place(
    target: Path,
    rows: list[list],
    sheet_name: str,
    overlap_policy: str,
    gap_policy: str,
    historical_policy: str,
    backup_output: Path | None = None,
) -> dict:
    if not target.exists():
        fail(f"统计表不存在：{target}")
    backup = backup_output or default_backup_path(target)
    if backup.resolve() == target.resolve():
        fail("备份路径不得与原统计表相同")
    if backup.exists():
        fail(f"备份文件已存在，拒绝覆盖：{backup}")
    archive_candidate = backup_trash_path(backup, date.today())
    if archive_candidate.exists():
        fail(f"安全删除目标已存在，拒绝覆盖：{archive_candidate}")

    # 不使用点号开头的临时文件名。macOS，尤其是 iCloud Drive，可能在文件
    # 改名到可见目标路径后，再次给由点号文件创建的 inode 加上 UF_HIDDEN。
    fd, temp_name = tempfile.mkstemp(prefix=f"{target.stem}_导入临时_", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    temp_output = Path(temp_name)
    temp_output.unlink()
    backup_temp = backup.with_name(f"{backup.name}.tmp")
    if backup_temp.exists():
        fail(f"备份临时文件已存在，拒绝覆盖：{backup_temp}")

    original_hash = file_sha256(target)
    original_flags = getattr(target.stat(), "st_flags", None)
    archived_backup = None
    replaced = False
    try:
        imported = build_validated_target(
            target, temp_output, rows, sheet_name,
            overlap_policy, gap_policy, historical_policy,
        )
        updated_hash = file_sha256(temp_output)

        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(target, backup_temp)
        if file_sha256(backup_temp) != original_hash:
            fail("写入前备份哈希与原统计表不一致")
        os.replace(backup_temp, backup)
        if original_flags is not None and hasattr(os, "chflags"):
            os.chflags(backup, original_flags)

        shutil.copystat(target, temp_output)
        if original_flags is not None and hasattr(os, "chflags"):
            os.chflags(temp_output, original_flags)
        os.replace(temp_output, target)
        replaced = True
        if original_flags is not None and hasattr(os, "chflags"):
            os.chflags(target, original_flags)
        if file_sha256(target) != updated_hash:
            fail("原子替换后统计表哈希与已校验临时文件不一致")
        verification = inspect_target(target, sheet_name)
        if verification["existingRows"] != imported["existingRowsAfterImport"]:
            fail("原地写入后回读行数与临时文件校验结果不一致")

        archived_on = date.today()
        archived_backup = archive_verified_backup(backup, archived_on)
        if file_sha256(archived_backup) != original_hash:
            fail("移入 .trash 后的写入前备份哈希与原统计表不一致")
        expired_backups = cleanup_expired_backups(archived_backup.parent, target, archived_on)
        final_flags = getattr(target.stat(), "st_flags", None)
        if original_flags is not None and final_flags != original_flags:
            fail(
                f"原地更新后文件系统标志不一致：更新前 {original_flags}，更新后 {final_flags}"
            )

        imported.update({
            "targetInputPath": str(target.resolve()),
            "targetOutputPath": str(target.resolve()),
            "backupPath": str(archived_backup.resolve()),
            "backupMovedToTrash": True,
            "backupRetentionDays": BACKUP_RETENTION_DAYS,
            "expiredBackupsDeleted": len(expired_backups),
            "inPlaceUpdated": True,
            "originalWorkbookBackedUp": True,
            "atomicReplace": True,
            "fileSystemFlagsPreserved": final_flags == original_flags,
            "fileSystemFlagsBefore": original_flags,
            "fileSystemFlagsAfter": final_flags,
            "sourceWorkbookPreserved": False,
        })
        return imported
    except Exception:
        recovery_backup = archived_backup if archived_backup and archived_backup.exists() else backup
        if replaced and recovery_backup.exists():
            restore_fd, restore_name = tempfile.mkstemp(
                prefix=f"{target.stem}_回退临时_", suffix=target.suffix, dir=target.parent
            )
            os.close(restore_fd)
            restore_temp = Path(restore_name)
            try:
                shutil.copy2(recovery_backup, restore_temp)
                if original_flags is not None and hasattr(os, "chflags"):
                    os.chflags(restore_temp, original_flags)
                os.replace(restore_temp, target)
                if original_flags is not None and hasattr(os, "chflags"):
                    os.chflags(target, original_flags)
            finally:
                restore_temp.unlink(missing_ok=True)
        raise
    finally:
        temp_output.unlink(missing_ok=True)
        backup_temp.unlink(missing_ok=True)


def verify_clean(path: Path) -> dict:
    matrix, date1904 = xlsx_matrix(path)
    headers = [norm(v) for v in matrix[0][:8]] if matrix else []
    if headers != CLEAN_HEADERS:
        fail(f"清理结果表头错误：{headers}")
    previous = None
    for row_num, row in enumerate(matrix[1:], start=2):
        row += [""] * (8 - len(row))
        day = excel_day(str(row[0]), date1904) if re.fullmatch(r"-?\d+(?:\.\d+)?", norm(row[0])) else parse_day(row[0], f"第 {row_num} 行日期")
        if previous and day < previous:
            fail("清理结果日期没有按升序排列")
        previous = day
        if norm(row[3]) not in ALLOWED_TYPES:
            fail(f"第 {row_num} 行类型无效：{row[3]}")
        parse_amount(row[4], row_num)
    return {"verified": True, "rows": max(0, len(matrix) - 1), "headers": headers}


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    sub = result.add_subparsers(dest="mode", required=True)
    inspect = sub.add_parser("inspect")
    inspect.add_argument("--input", required=True, type=Path)
    clean = sub.add_parser("clean")
    clean.add_argument("--input", required=True, type=Path)
    clean.add_argument("--output", required=True, type=Path)
    clean.add_argument("--start-date", required=True)
    clean.add_argument("--end-date", required=True)
    target = sub.add_parser("inspect-target")
    target.add_argument("--target", required=True, type=Path)
    target.add_argument("--sheet", default="5.每日收入支出明细表")
    combined = sub.add_parser("clean-import")
    combined.add_argument("--input", required=True, type=Path)
    combined.add_argument("--clean-output", required=True, type=Path)
    combined.add_argument("--target", required=True, type=Path)
    combined.add_argument("--backup-output", type=Path)
    combined.add_argument("--start-date", required=True)
    combined.add_argument("--end-date", required=True)
    combined.add_argument("--sheet", default="5.每日收入支出明细表")
    combined.add_argument("--overlap-policy", choices=["stop", "replace", "append"], default="stop")
    combined.add_argument("--gap-policy", choices=["stop", "continue"], default="stop")
    combined.add_argument("--historical-policy", choices=["stop", "sort"], default="stop")
    import_cleaned = sub.add_parser("import")
    import_cleaned.add_argument("--cleaned-input", required=True, type=Path)
    import_cleaned.add_argument("--target", required=True, type=Path)
    import_cleaned.add_argument("--backup-output", type=Path)
    import_cleaned.add_argument("--sheet", default="5.每日收入支出明细表")
    import_cleaned.add_argument("--overlap-policy", choices=["stop", "replace", "append"], default="stop")
    import_cleaned.add_argument("--gap-policy", choices=["stop", "continue"], default="stop")
    import_cleaned.add_argument("--historical-policy", choices=["stop", "sort"], default="stop")
    return result


def main() -> None:
    args = parser().parse_args()
    if args.mode == "inspect":
        result = inspect_source(args.input)
    elif args.mode == "inspect-target":
        result = inspect_target(args.target, args.sheet)
    elif args.mode == "import":
        rows = read_cleaned(args.cleaned_input)
        imported = write_target_in_place(
            args.target, rows, args.sheet,
            args.overlap_policy, args.gap_policy, args.historical_policy,
            args.backup_output,
        )
        result = {
            "cleanedInputPath": str(args.cleaned_input.resolve()),
            "cleanedRows": len(rows),
            "import": imported,
        }
    else:
        start, end = parse_day(args.start_date, "--start-date"), parse_day(args.end_date, "--end-date")
        rows, stats = clean_records(args.input, start, end)
        if args.mode == "clean":
            write_clean_xlsx(args.output, rows)
            stats.update({"inputPath": str(args.input.resolve()), "outputPath": str(args.output.resolve()), "verification": verify_clean(args.output)})
            result = stats
        else:
            if args.clean_output.resolve() == args.target.resolve():
                fail("--clean-output 与 --target 不得相同")
            write_clean_xlsx(args.clean_output, rows)
            clean_verification = verify_clean(args.clean_output)
            try:
                imported = write_target_in_place(
                    args.target, rows, args.sheet,
                    args.overlap_policy, args.gap_policy, args.historical_policy,
                    args.backup_output,
                )
            except Exception:
                args.clean_output.unlink(missing_ok=True)
                raise
            stats.update({"inputPath": str(args.input.resolve()), "cleanOutputPath": str(args.clean_output.resolve()), "cleanVerification": clean_verification, "import": imported})
            result = stats
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError, PermissionError) as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(2)
