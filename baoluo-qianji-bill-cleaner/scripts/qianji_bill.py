#!/usr/bin/env python3
"""钱迹账单检查、清理，以及安全导入现有财务统计表。仅依赖 Python 标准库。"""

from __future__ import annotations

import argparse
import csv
import json
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
        records.append({"row": row_num, "day": day, "values": row[:8]})
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


def set_inline(cell: ET.Element, value) -> None:
    for child in list(cell):
        cell.remove(child)
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


def write_target_copy(target: Path, output: Path, rows: list[list], sheet_name: str, overlap_policy: str) -> dict:
    if output.exists():
        fail(f"目标输出已存在，拒绝覆盖：{output}")
    if target.resolve() == output.resolve():
        fail("默认禁止原地覆盖财务统计表；请提供新的 --target-output 路径")
    existing, reserved_start, _ = target_rows(target, sheet_name)
    incoming_days = {row[0] for row in rows}
    overlaps = [r for r in existing if r["day"] in incoming_days]
    if overlaps and overlap_policy == "stop":
        days = sorted({r["day"] for r in overlaps})
        fail(f"目标表已有 {len(overlaps)} 行落在导入日期中（{days[0]} 至 {days[-1]}）；请明确使用 --overlap-policy replace 或 append")
    remove_rows = {r["row"] for r in overlaps} if overlap_policy == "replace" else set()
    kept = [r for r in existing if r["row"] not in remove_rows]
    next_row = max((r["row"] for r in kept), default=1) + 1
    capacity_end = reserved_start - 1 if reserved_start else 1048576
    if next_row + len(rows) - 1 > capacity_end:
        fail(f"目标表可写空间不足：需要 {len(rows)} 行，仅剩 {max(0, capacity_end-next_row+1)} 行")

    with ZipFile(target) as source:
        sheet_path, _ = sheet_path_by_name(source, sheet_name)
        root = ET.fromstring(source.read(sheet_path))
        data = root.find("m:sheetData", NS)
        if data is None:
            fail("目标工作表缺少 sheetData")
        row_map = {int(r.attrib["r"]): r for r in data.findall("m:row", NS)}
        template_row = row_map.get(2)
        if template_row is None:
            fail("目标工作表缺少第 2 行格式模板")
        template_styles = {}
        for cell in template_row.findall("m:c", NS):
            template_styles[col_number(cell.attrib["r"])] = cell.attrib.get("s")
        for row_num in remove_rows:
            row = row_map.get(row_num)
            if row is not None:
                for cell in row.findall("m:c", NS):
                    if col_number(cell.attrib["r"]) <= 8:
                        for child in list(cell):
                            cell.remove(child)
                        cell.attrib.pop("t", None)
        for offset, values in enumerate(rows):
            row_num = next_row + offset
            row = row_map.get(row_num)
            if row is None:
                row = ET.Element(f"{{{MAIN}}}row", {"r": str(row_num), "ht": template_row.attrib.get("ht", "30"), "customHeight": "1"})
                data.append(row)
                row_map[row_num] = row
            cells = {col_number(c.attrib["r"]): c for c in row.findall("m:c", NS)}
            for col, value in enumerate(values, start=1):
                cell = cells.get(col)
                if cell is None:
                    cell = ET.SubElement(row, f"{{{MAIN}}}c", {"r": f"{chr(64+col)}{row_num}"})
                if template_styles.get(col) is not None:
                    cell.attrib["s"] = template_styles[col]
                set_inline(cell, value)
        data[:] = sorted(data, key=lambda r: int(r.attrib["r"]))
        sheet_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        workbook = ET.fromstring(source.read("xl/workbook.xml"))
        calc = workbook.find("m:calcPr", NS)
        if calc is None:
            calc = ET.SubElement(workbook, f"{{{MAIN}}}calcPr")
        calc.attrib.update({"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"})
        workbook_bytes = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)

        replacements = {sheet_path: sheet_bytes, "xl/workbook.xml": workbook_bytes}
        for name in source.namelist():
            if name.startswith("xl/pivotCache/pivotCacheDefinition") and name.endswith(".xml"):
                cache = ET.fromstring(source.read(name))
                cache.attrib["refreshOnLoad"] = "1"
                replacements[name] = ET.tostring(cache, encoding="utf-8", xml_declaration=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w") as dest:
            for item in source.infolist():
                dest.writestr(item, replacements.get(item.filename, source.read(item.filename)))

    verification = inspect_target(output, sheet_name)
    expected_count = len(existing) - len(remove_rows) + len(rows)
    if verification["existingRows"] != expected_count:
        output.unlink(missing_ok=True)
        fail(f"导入后回读行数不符：预期 {expected_count}，实际 {verification['existingRows']}")
    return {
        "targetInputPath": str(target.resolve()), "targetOutputPath": str(output.resolve()),
        "sheetName": sheet_name, "overlapPolicy": overlap_policy,
        "replacedRows": len(remove_rows), "importedRows": len(rows),
        "firstWrittenRow": next_row, "lastWrittenRow": next_row + len(rows) - 1,
        "existingRowsAfterImport": verification["existingRows"], "pivotRefreshOnLoad": True,
        "sourceWorkbookPreserved": True,
    }


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
    combined.add_argument("--target-output", required=True, type=Path)
    combined.add_argument("--start-date", required=True)
    combined.add_argument("--end-date", required=True)
    combined.add_argument("--sheet", default="5.每日收入支出明细表")
    combined.add_argument("--overlap-policy", choices=["stop", "replace", "append"], default="stop")
    return result


def main() -> None:
    args = parser().parse_args()
    if args.mode == "inspect":
        result = inspect_source(args.input)
    elif args.mode == "inspect-target":
        result = inspect_target(args.target, args.sheet)
    else:
        start, end = parse_day(args.start_date, "--start-date"), parse_day(args.end_date, "--end-date")
        rows, stats = clean_records(args.input, start, end)
        if args.mode == "clean":
            write_clean_xlsx(args.output, rows)
            stats.update({"inputPath": str(args.input.resolve()), "outputPath": str(args.output.resolve()), "verification": verify_clean(args.output)})
            result = stats
        else:
            if args.clean_output.resolve() == args.target_output.resolve():
                fail("--clean-output 与 --target-output 不得相同")
            write_clean_xlsx(args.clean_output, rows)
            clean_verification = verify_clean(args.clean_output)
            try:
                imported = write_target_copy(args.target, args.target_output, rows, args.sheet, args.overlap_policy)
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
