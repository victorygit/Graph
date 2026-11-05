#!/usr/bin/env python3
"""Extract table names from SELECT ... FROM statements in ABAP routine columns."""

import argparse
import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"x": MAIN_NS, "r": REL_NS}
ROUTINE_COLUMN_OUTPUTS = {
    "start_routine_code": "Start_Routine_Lookup_Tables",
    "end_routine_code": "End_Routine_Lookup_Tables",
    "expert_routine_code": "Expert_Routine_Lookup_Tables",
    "routine_code": "Routine_Lookup_Tables",
}
SELECT_RE = re.compile(r"\bSELECT\b", re.IGNORECASE)
FROM_RE = re.compile(
    r"\bFROM\s+(/?[A-Za-z_][A-Za-z0-9_/]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)",
    re.IGNORECASE,
)


def remove_abap_comments_and_literals(text):
    cleaned = []
    for line in text.splitlines():
        if line.startswith("*"):
            cleaned.append("")
            continue

        result = []
        index = 0
        while index < len(line):
            char = line[index]
            if char == '"':
                break
            if char in ("'", "`", "|"):
                delimiter = char
                index += 1
                while index < len(line):
                    if (
                        delimiter == "'"
                        and line[index] == "'"
                        and index + 1 < len(line)
                        and line[index + 1] == "'"
                    ):
                        index += 2
                        continue
                    if (
                        delimiter == "`"
                        and line[index] == "`"
                        and index + 1 < len(line)
                        and line[index + 1] == "`"
                    ):
                        index += 2
                        continue
                    if line[index] == delimiter:
                        index += 1
                        break
                    if (
                        delimiter == "|"
                        and line[index] == "\\"
                        and index + 1 < len(line)
                    ):
                        index += 2
                        continue
                    index += 1
                result.append(" ")
                continue
            result.append(char)
            index += 1
        cleaned.append("".join(result))
    return "\n".join(cleaned)


def extract_tables(text):
    code = remove_abap_comments_and_literals(text)
    selects = list(SELECT_RE.finditer(code))
    tables = []
    seen = set()
    for index, select in enumerate(selects):
        end = code.find(".", select.end())
        if end == -1:
            end = len(code)
        if index + 1 < len(selects):
            end = min(end, selects[index + 1].start())

        match = FROM_RE.search(code, select.end(), end)
        if match:
            table = match.group(1)
            normalized = table.casefold()
            if normalized not in seen:
                tables.append(table)
                seen.add(normalized)
    return tables


def cell_text(cell, shared_strings):
    if cell is None:
        return ""
    if cell.attrib.get("t") == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//x:t", NS))
    value = cell.find("x:v", NS)
    if value is None or value.text is None:
        return ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(value.text)]
    return value.text


def column_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref).group()
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - ord("A") + 1
    return index


def column_name(index):
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def resolve_sheet_path(workbook, relationships):
    sheet = workbook.find("x:sheets/x:sheet", NS)
    if sheet is None:
        raise ValueError("Workbook does not contain a worksheet")
    relationship_id = sheet.attrib["{" + REL_NS + "}id"]
    target = relationships[relationship_id]
    return posixpath.normpath(posixpath.join("xl", target))


def replace_or_append_cell(row_xml, cell_ref, replacement):
    pattern = re.compile(
        r"<c\b(?=[^>]*\br=[\"']"
        + re.escape(cell_ref)
        + r"[\"'])[^>]*(?:/>|>.*?</c>)",
        re.DOTALL,
    )
    updated, count = pattern.subn(replacement, row_xml, count=1)
    if count:
        return updated
    close_tag = re.search(r"</row\s*>", row_xml)
    if close_tag is None:
        raise ValueError("Could not locate worksheet row closing tag")
    return row_xml[: close_tag.start()] + replacement + row_xml[close_tag.start() :]


def add_inline_string(row_xml, cell_ref, value):
    cell = (
        f'<c r="{cell_ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    )
    return replace_or_append_cell(row_xml, cell_ref, cell)


def process_workbook(input_path, output_path):
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Input and output paths must be different")

    with zipfile.ZipFile(input_path, "r") as source:
        workbook = ET.fromstring(source.read("xl/workbook.xml"))
        relationships_root = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
        relationships = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in relationships_root.findall(
                "{" + PKG_REL_NS + "}Relationship"
            )
        }
        sheet_path = resolve_sheet_path(workbook, relationships)

        shared_strings = []
        if "xl/sharedStrings.xml" in source.namelist():
            strings_root = ET.fromstring(source.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(node.text or "" for node in item.findall(".//x:t", NS))
                for item in strings_root.findall("x:si", NS)
            ]

        sheet_xml = source.read(sheet_path).decode("utf-8")
        sheet_root = ET.fromstring(sheet_xml)
        rows = sheet_root.findall(".//x:sheetData/x:row", NS)
        if not rows:
            raise ValueError("Worksheet contains no rows")

        header_cells = {
            cell_text(cell, shared_strings): cell.attrib["r"]
            for cell in rows[0].findall("x:c", NS)
        }
        headers_by_normalized_name = {
            name.casefold(): name for name in header_cells
        }
        routine_columns = {
            headers_by_normalized_name[source_header]: output_header
            for source_header, output_header in ROUTINE_COLUMN_OUTPUTS.items()
            if source_header in headers_by_normalized_name
        }
        if not routine_columns:
            expected = ", ".join(ROUTINE_COLUMN_OUTPUTS)
            raise ValueError(f"No supported routine-code column found; expected one of: {expected}")

        existing_output_columns = {}
        for header, cell_ref in header_cells.items():
            if header in ROUTINE_COLUMN_OUTPUTS.values():
                existing_output_columns[header] = column_index(cell_ref)

        next_column = max(
            (column_index(cell.attrib["r"]) for row in rows for cell in row.findall("x:c", NS)),
            default=0,
        ) + 1
        output_columns = {}
        for header in routine_columns.values():
            if header in existing_output_columns:
                output_columns[header] = existing_output_columns[header]
            else:
                output_columns[header] = next_column
                next_column += 1

        updates = {}
        row_hits = Counter()
        reference_counts = Counter()
        for row in rows:
            row_number = int(row.attrib["r"])
            cells = {cell.attrib["r"]: cell for cell in row.findall("x:c", NS)}
            row_xml_match = re.search(
                r"<row\b(?=[^>]*\br=[\"']"
                + str(row_number)
                + r"[\"'])[^>]*>.*?</row>",
                sheet_xml,
                re.DOTALL,
            )
            if row_xml_match is None:
                raise ValueError(f"Could not locate worksheet row {row_number}")
            row_xml = row_xml_match.group(0)

            for source_header, output_header in routine_columns.items():
                output_ref = f"{column_name(output_columns[output_header])}{row_number}"
                if row_number == int(rows[0].attrib["r"]):
                    value = output_header
                else:
                    source_ref = header_cells[source_header]
                    source_column = re.match(r"[A-Z]+", source_ref).group()
                    source_cell_ref = f"{source_column}{row_number}"
                    names = extract_tables(cell_text(cells.get(source_cell_ref), shared_strings))
                    value = "; ".join(names)
                    if names:
                        row_hits[source_header] += 1
                        reference_counts[source_header] += len(names)
                row_xml = add_inline_string(row_xml, output_ref, value)

            updates[row_xml_match.span()] = row_xml

        for (start, end), replacement in sorted(updates.items(), reverse=True):
            sheet_xml = sheet_xml[:start] + replacement + sheet_xml[end:]

        last_column = max(output_columns.values())
        dimension_pattern = re.compile(
            r'(<dimension\b[^>]*\bref=")([A-Z]+)(\d+):([A-Z]+)(\d+)(")'
        )
        dimension_match = dimension_pattern.search(sheet_xml)
        if dimension_match:
            dimension = (
                dimension_match.group(1)
                + dimension_match.group(2)
                + dimension_match.group(3)
                + ":"
                + column_name(max(last_column, column_index(dimension_match.group(4))))
                + dimension_match.group(5)
                + dimension_match.group(6)
            )
            sheet_xml = (
                sheet_xml[: dimension_match.start()]
                + dimension
                + sheet_xml[dimension_match.end() :]
            )

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as result:
            for item in source.infolist():
                contents = (
                    sheet_xml.encode("utf-8")
                    if item.filename == sheet_path
                    else source.read(item.filename)
                )
                result.writestr(item, contents)

    return row_hits, reference_counts


def main():
    parser = argparse.ArgumentParser(
        description="Add lookup-table names extracted from ABAP routine SELECT statements."
    )
    parser.add_argument("input", type=Path, help="Source .xlsx workbook")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output workbook (default: '<input name> tables.xlsx')",
    )
    args = parser.parse_args()
    input_path = args.input
    output_path = args.output or input_path.with_name(input_path.stem + " tables.xlsx")
    if not input_path.is_file():
        parser.error(f"Input workbook not found: {input_path}")

    row_hits, reference_counts = process_workbook(input_path, output_path)
    print(f"Created: {output_path}")
    print("Routine cells with extracted tables:", dict(row_hits))
    print("Extracted table references:", dict(reference_counts))


if __name__ == "__main__":
    main()
