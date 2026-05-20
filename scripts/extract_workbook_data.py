from __future__ import annotations

import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "[MHNow] Damage Calculator v3.6.3.xlsx"
OUTPUT = ROOT / "data" / "workbook-data.json"
OUTPUT_JS = ROOT / "data" / "workbook-data.js"
SHEET_VERSION = "3.6.3"
CALCULATOR_SHEET_CANDIDATES = ("Calculator", "Calculator1")
WEAPON_FIELD_ROWS = range(3, 8)


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(num: int) -> str:
    out = []
    while num:
        num, rem = divmod(num - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def parse_ref(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)", ref)
    if not match:
        raise ValueError(f"invalid cell reference: {ref}")
    return col_to_num(match.group(1)), int(match.group(2))


def normalize_scalar(text: str | None, cell_type: str | None, shared_strings: list[str]):
    if text is None:
        return None
    if cell_type == "s":
        return shared_strings[int(text)]
    if cell_type == "b":
        return bool(int(text))
    if cell_type == "str":
        return text
    try:
        value = float(text)
    except ValueError:
        return text
    if value.is_integer():
        return int(value)
    return value


def semantic_key(label: str) -> str:
    label = label.replace("&", " and ")
    label = re.sub(r"\bAdv\.", "Advanced", label)
    label = label.replace("'", "")
    parts = re.findall(r"[A-Za-z0-9]+", label)
    if not parts:
        raise ValueError(f"cannot build semantic key from label: {label!r}")
    return parts[0].lower() + "".join(part[:1].upper() + part[1:] for part in parts[1:])


def field_key(ref: str, label: object, field_kind: str) -> str:
    if field_kind == "build":
        if ref == "B3":
            return "elementalAttack"
        if ref == "B4":
            return "advancedElementalAttack"
        if isinstance(label, str) and label.startswith("=") and "Vital " in label:
            return "vitalElement"
    if field_kind == "weapon":
        overrides = {
            "E4": "weaponElement",
            "E5": "damageType",
            "E6": "weaponAffinity",
            "E7": "weaponType",
        }
        if ref in overrides:
            return overrides[ref]

    if not isinstance(label, str) or label.startswith("="):
        return semantic_key(ref)
    return semantic_key(label)


def sanitize_formula(sheet: str, ref: str, formula: str) -> str:
    if sheet == "Backyard" and ref == "S2":
        formula = (
            'IF($B$5="Raw",INDEX(Riftborne!$B$4:$B$24,MATCH(Calculator!$E$18,Riftborne!$A$4:$A$24,0)),'
            'IF($B$5="Element",INDEX(Riftborne!$C$4:$C$24,MATCH(Calculator!$E$18,Riftborne!$A$4:$A$24,0)),'
            'INDEX(Riftborne!$E$4:$E$24,MATCH(Calculator!$E$18,Riftborne!$A$4:$A$24,0))))'
        )
    elif sheet == "Backyard" and ref == "S3":
        formula = (
            'COUNTIF(Calculator!$E$19:$E$21,"Attack")*'
            'IF($B$5="Raw",INDEX(Riftborne!$I$3:$I$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'IF($B$5="Element",INDEX(Riftborne!$J$3:$J$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'INDEX(Riftborne!$K$3:$K$6,MATCH("Attack",Riftborne!$H$3:$H$6,0))))'
        )
    elif sheet == "Backyard" and ref == "AX21":
        formula = "SUMIF($AX$3:$AY$9,0.75,$AX$12:$AY$18)"
    elif sheet == "Backyard" and ref == "AX23":
        formula = "SUMIF($AX$3:$AY$9,1.25,$AX$12:$AY$18)"
    if sheet == "Calculator1" and ref == "AX86":
        formula = 'SUMIF($AX$68:$AY$74,0.75,$AX$77:$AY$83)'
    elif sheet == "Calculator1" and ref == "AX88":
        formula = 'SUMIF($AX$68:$AY$74,1.25,$AX$77:$AY$83)'
    if sheet == "Calculator1" and ref == "S67":
        formula = (
            'IF($B$70="Raw",INDEX(Riftborne!$B$4:$B$24,MATCH($E$18,Riftborne!$A$4:$A$24,0)),'
            'IF($B$70="Element",INDEX(Riftborne!$C$4:$C$24,MATCH($E$18,Riftborne!$A$4:$A$24,0)),'
            'INDEX(Riftborne!$E$4:$E$24,MATCH($E$18,Riftborne!$A$4:$A$24,0))))'
        )
    elif sheet == "Calculator1" and ref == "S68":
        formula = (
            'COUNTIF($E$19:$E$21,"Attack")*'
            'IF($B$70="Raw",INDEX(Riftborne!$I$3:$I$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'IF($B$70="Element",INDEX(Riftborne!$J$3:$J$6,MATCH("Attack",Riftborne!$H$3:$H$6,0)),'
            'INDEX(Riftborne!$K$3:$K$6,MATCH("Attack",Riftborne!$H$3:$H$6,0))))'
        )

    formula = re.sub(r"\bTRUE\b(?!\s*\()", "TRUE()", formula, flags=re.IGNORECASE)
    formula = re.sub(r"\bFALSE\b(?!\s*\()", "FALSE()", formula, flags=re.IGNORECASE)
    return f"={formula}"


def parse_sqref(sqref: str) -> list[str]:
    refs: list[str] = []
    for token in sqref.split():
        if ":" not in token:
            refs.append(token)
            continue
        start, end = token.split(":")
        start_col, start_row = parse_ref(start)
        end_col, end_row = parse_ref(end)
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                refs.append(f"{num_to_col(col)}{row}")
    return refs


def iter_range(range_ref: str) -> list[str]:
    start, end = range_ref.split(":")
    start_col, start_row = parse_ref(start.replace("$", ""))
    end_col, end_row = parse_ref(end.replace("$", ""))
    refs: list[str] = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            refs.append(f"{num_to_col(col)}{row}")
    return refs


def resolve_validation_options(
    formula: str,
    workbook_values: dict[str, dict[str, object]],
    default_sheet_name: str,
) -> list[object]:
    if formula.startswith('"') and formula.endswith('"'):
        return [item.strip() for item in formula[1:-1].split(",")]

    if "!" in formula:
        sheet_name, range_ref = formula.split("!", 1)
    else:
        sheet_name, range_ref = default_sheet_name, formula

    sheet_name = sheet_name.strip("'")
    refs = iter_range(range_ref)
    values = workbook_values[sheet_name]
    return [values[ref] for ref in refs if ref in values and values[ref] not in (None, "")]


def matrix_cell(matrix: list[list[object | None]], ref: str) -> object | None:
    col, row = parse_ref(ref)
    if row - 1 >= len(matrix) or col - 1 >= len(matrix[row - 1]):
        return None
    return matrix[row - 1][col - 1]


def find_build_rows(calculator_matrix: list[list[object | None]]) -> list[int]:
    rows: list[int] = []
    for row in range(3, len(calculator_matrix) + 1):
        label = matrix_cell(calculator_matrix, f"A{row}")
        value = matrix_cell(calculator_matrix, f"B{row}")
        if label in (None, "") and value in (None, ""):
            break
        if label not in (None, ""):
            rows.append(row)
    return rows


def build_uptime_fields(
    calculator_matrix: list[list[object | None]],
    workbook_values: dict[str, dict[str, object]],
    calculator_sheet_name: str,
) -> list[dict[str, object]]:
    fields: list[dict[str, object]] = []

    for row in range(1, len(calculator_matrix) + 1):
        if matrix_cell(calculator_matrix, f"D{row}") == "Remaining Health":
            fields.append(
                {
                    "ref": f"E{row}",
                    "key": "remainingHealth",
                    "label": "Remaining Health",
                    "defaultValue": 100,
                    "displayScale": 1,
                    "maxValue": 160,
                }
            )
            break

    uptime_header_row = None
    for row in range(1, len(calculator_matrix) + 1):
        if matrix_cell(calculator_matrix, f"D{row}") == "Uptime":
            uptime_header_row = row
            break

    if uptime_header_row is None:
        return fields

    for row in range(uptime_header_row + 1, len(calculator_matrix) + 1):
        label = matrix_cell(calculator_matrix, f"D{row}")
        default_value = workbook_values[calculator_sheet_name].get(f"E{row}")
        if label in (None, "") or default_value in (None, ""):
            break
        fields.append(
            {
                "ref": f"E{row}",
                "key": semantic_key(str(label)),
                "label": str(label),
                "defaultValue": default_value,
                "displayScale": 100,
            }
        )

    return fields


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(SOURCE) as archive:
        shared_strings_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        shared_strings = [
            "".join(t.text or "" for t in si.iterfind(".//a:t", NS))
            for si in shared_strings_root.findall("a:si", NS)
        ]

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}
        sheet_files = {
            sheet.attrib["name"]: f"xl/{rel_map[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]}"
            for sheet in workbook_root.find("a:sheets", NS)
        }
        calculator_sheet_name = next(
            (name for name in CALCULATOR_SHEET_CANDIDATES if name in sheet_files),
            None,
        )
        if calculator_sheet_name is None:
            raise RuntimeError("Could not find a calculator sheet.")

        sheet_matrices: dict[str, list[list[object | None]]] = {}
        workbook_values: dict[str, dict[str, object]] = {}

        for sheet_name in sheet_files:
            root = ET.fromstring(archive.read(sheet_files[sheet_name]))
            cells = root.findall(".//a:c", NS)
            max_row = 0
            max_col = 0
            workbook_values[sheet_name] = {}
            parsed_cells: list[tuple[int, int, object]] = []

            for cell in cells:
                ref = cell.attrib["r"]
                col, row = parse_ref(ref)
                max_col = max(max_col, col)
                max_row = max(max_row, row)

                formula = cell.find("a:f", NS)
                value = cell.find("a:v", NS)
                if formula is not None and formula.text:
                    content: object = sanitize_formula(sheet_name, ref, formula.text)
                else:
                    content = normalize_scalar(
                        None if value is None else value.text,
                        cell.attrib.get("t"),
                        shared_strings,
                    )

                if value is not None:
                    workbook_values[sheet_name][ref] = normalize_scalar(
                        value.text,
                        cell.attrib.get("t"),
                        shared_strings,
                    )
                elif formula is None:
                    workbook_values[sheet_name][ref] = None

                parsed_cells.append((row, col, content))

            matrix: list[list[object | None]] = [
                [None for _ in range(max_col)] for _ in range(max_row)
            ]
            for row, col, content in parsed_cells:
                matrix[row - 1][col - 1] = content

            sheet_matrices[sheet_name] = matrix

        calculator_root = ET.fromstring(archive.read(sheet_files[calculator_sheet_name]))
        data_validations = calculator_root.find("a:dataValidations", NS)
        validation_map: dict[str, list[object]] = {}
        if data_validations is not None:
            for validation in data_validations.findall("a:dataValidation", NS):
                formula1 = validation.find("a:formula1", NS)
                if formula1 is None or not formula1.text:
                    continue
                options = resolve_validation_options(
                    formula1.text,
                    workbook_values,
                    calculator_sheet_name,
                )
                for ref in parse_sqref(validation.attrib["sqref"]):
                    validation_map[ref] = options

    calculator_matrix = sheet_matrices[calculator_sheet_name]
    build_rows = find_build_rows(calculator_matrix)
    build_fields = [
        {
            "ref": f"B{row}",
            "labelRef": f"A{row}",
            "key": field_key(f"B{row}", matrix_cell(calculator_matrix, f"A{row}"), "build"),
            "options": validation_map.get(f"B{row}"),
            "defaultValue": workbook_values[calculator_sheet_name].get(f"B{row}"),
        }
        for row in build_rows
    ]

    weapon_fields = [
        {
            "ref": f"E{row}",
            "labelRef": f"D{row}",
            "key": field_key(f"E{row}", matrix_cell(calculator_matrix, f"D{row}"), "weapon"),
            "options": validation_map.get(f"E{row}"),
            "defaultValue": workbook_values[calculator_sheet_name].get(f"E{row}"),
        }
        for row in WEAPON_FIELD_ROWS
    ]
    uptime_fields = build_uptime_fields(
        calculator_matrix,
        workbook_values,
        calculator_sheet_name,
    )

    payload = {
        "formatVersion": 2,
        "sheetVersion": SHEET_VERSION,
        "calculatorSheet": calculator_sheet_name,
        "sheets": sheet_matrices,
        "buildFields": build_fields,
        "weaponFields": weapon_fields,
        "uptimeFields": uptime_fields,
        "resultCell": "H12",
    }

    serialized = json.dumps(payload, separators=(",", ":"))
    OUTPUT.write_text(serialized)
    OUTPUT_JS.write_text(f"window.WORKBOOK_DATA={serialized};")
    print(f"wrote {OUTPUT}")
    print(f"wrote {OUTPUT_JS}")


if __name__ == "__main__":
    main()
