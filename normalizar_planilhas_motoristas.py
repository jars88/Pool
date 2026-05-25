from __future__ import annotations

import datetime as dt
import html
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

FILES = [
    "MOB_PROGEN.xlsx",
    "PW_APPS_SPOT.xlsx",
    "SRD_ATUAL.xlsx",
    "BM_MEDICAO_ABRIL_2026.xlsx",
]


def norm(value: object) -> str:
    text = "" if value is None else str(value)
    text = "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", text.upper()).strip()


def key(value: object) -> str:
    return re.sub(r"[^A-Z0-9]+", " ", norm(value)).strip()


def col_idx(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref or "A")
    total = 0
    for char in match.group(1):
        total = total * 26 + ord(char) - 64
    return total - 1


def read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values = []
    for item in root.findall("m:si", NS):
        values.append("".join(t.text or "" for t in item.iter(f"{{{NS['m']}}}t")))
    return values


def sheet_map(zf: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_by_id = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    result = {}
    for sheet in workbook.find("m:sheets", NS):
        rid = sheet.attrib[f"{{{NS['r']}}}id"]
        target = rel_by_id[rid].lstrip("/")
        result[sheet.attrib["name"]] = target if target.startswith("xl/") else "xl/" + target
    return result


def read_workbook(path: Path) -> dict[str, list[list[str]]]:
    with ZipFile(path) as zf:
        strings = read_shared_strings(zf)
        sheets = sheet_map(zf)
        out = {}
        for name, sheet_path in sheets.items():
            root = ET.fromstring(zf.read(sheet_path))
            rows = []
            for row in root.findall(".//m:sheetData/m:row", NS):
                current = []
                for cell in row.findall("m:c", NS):
                    index = col_idx(cell.attrib.get("r", "A1"))
                    while len(current) <= index:
                        current.append("")
                    value_node = cell.find("m:v", NS)
                    cell_type = cell.attrib.get("t")
                    value = ""
                    if cell_type == "s" and value_node is not None:
                        value = strings[int(value_node.text)]
                    elif cell_type == "inlineStr":
                        value = "".join(t.text or "" for t in cell.iter(f"{{{NS['m']}}}t"))
                    elif value_node is not None:
                        value = value_node.text or ""
                    current[index] = value
                rows.append(current)
            out[name] = rows
        return out


def row_score(row: list[str]) -> int:
    normalized = [norm(c) for c in row]
    joined = " | ".join(normalized)
    score = sum(1 for c in row if str(c).strip())
    keywords = [
        "MOTORISTA",
        "NOME EMPREGADO",
        "EMPREGADO",
        "FORNECEDOR",
        "DATA",
        "HORARIO",
        "STATUS",
        "LOCALIDADE",
        "VALOR",
        "PROTOCOLO",
        "DIRETORIA",
        "EMPRESA",
    ]
    return score + sum(8 for word in keywords if word in joined)


def header_start(rows: list[list[str]]) -> int:
    if not rows:
        return 0
    best = max(range(min(len(rows), 12)), key=lambda i: row_score(rows[i]))
    return best if row_score(rows[best]) >= 12 else 0


def hmap(header: list[str]) -> dict[str, int]:
    return {norm(name): index for index, name in enumerate(header)}


def first_idx(header: dict[str, int], *names: str) -> int | None:
    for name in names:
        index = header.get(norm(name))
        if index is not None:
            return index
    return None


def cell(row: list[str], index: int | None) -> str:
    return row[index] if index is not None and index < len(row) else ""


def is_driver_header(name: str) -> bool:
    n = norm(name)
    return n in {
        "MOTORISTA",
        "NOME EMPREGADO",
        "NOME DO EMPREGADO",
        "EMPREGADO",
        "NOME MOTORISTA",
        "NOME DO MOTORISTA",
    }


def add_official(official: dict[str, str], name: str) -> None:
    clean = re.sub(r"\s+", " ", str(name or "").strip())
    k = key(clean)
    if not k:
        return
    if len(clean) > len(official.get(k, "")):
        official[k] = clean


def build_official_names(root: Path) -> dict[str, str]:
    official: dict[str, str] = {}

    mob = read_workbook(root / "MOB_PROGEN.xlsx")["Mobilizados"]
    header = hmap(mob[0])
    name_idx = first_idx(header, "NOME EMPREGADO", "MOTORISTA", "EMPREGADO", "NOME DO MOTORISTA")
    status_idx = first_idx(header, "STATUS DO TERCEIRO", "STATUS")
    cargo_idx = first_idx(header, "CARGO")
    for row in mob[1:]:
        if norm(cell(row, cargo_idx)) == "MOTORISTA":
            add_official(official, cell(row, name_idx))

    bm = read_workbook(root / "BM_MEDICAO_ABRIL_2026.xlsx")["EFETIVO"]
    start = header_start(bm)
    header = hmap(bm[start])
    name_idx = first_idx(header, "NOME EMPREGADO", "MOTORISTA", "EMPREGADO", "NOME DO MOTORISTA")
    desc_idx = first_idx(header, "DESCRICAO", "DESCRIÇÃO", "DESCRICAO DO SERVICO")
    regime_idx = first_idx(header, "REGIME")
    if name_idx is not None:
        for row in bm[start + 1 :]:
            description = " ".join(
                row[i] for i in [desc_idx, regime_idx] if i is not None and i < len(row)
            )
            if "MOTORISTA" in norm(description) or not description.strip():
                add_official(official, cell(row, name_idx))

    return official


def resolve_name(raw: str, official: dict[str, str]) -> str:
    raw = str(raw or "").strip()
    raw_key = key(raw)
    if not raw_key:
        return ""
    if raw_key in official:
        return official[raw_key]

    raw_tokens = set(raw_key.split())
    best_name = raw
    best_score = 0.0
    for official_key, official_name in official.items():
        if raw_key in official_key or official_key in raw_key:
            score = 80 + min(len(raw_key), len(official_key)) / max(len(official_key), 1)
        else:
            official_tokens = set(official_key.split())
            shared = len(raw_tokens & official_tokens)
            score = shared * 10
            if raw_key.split()[:1] == official_key.split()[:1]:
                score += 5
        if score > best_score:
            best_name = official_name
            best_score = score
    return best_name if best_score >= 15 else raw


def normalize_sheet(rows: list[list[str]], official: dict[str, str]) -> list[list[str]]:
    if not rows:
        return rows
    start = header_start(rows)
    rows = [r[:] for r in rows[start:]]
    header = rows[0]

    motorista_idx = None
    for idx, name in enumerate(header):
        if is_driver_header(name):
            motorista_idx = idx
            header[idx] = "Motorista"
            break

    if motorista_idx is None:
        motorista_idx = len(header)
        header.append("Motorista")
        for row in rows[1:]:
            row.append("")

    current_header = hmap(header)
    original_idx = first_idx(current_header, "Motorista_Original")
    if original_idx is None:
        original_idx = len(header)
        header.append("Motorista_Original")
    ref_idx = first_idx(current_header, "Motorista_Referencial")
    if ref_idx is None:
        ref_idx = len(header)
        header.append("Motorista_Referencial")

    for row in rows[1:]:
        while len(row) < len(header):
            row.append("")
        for idx, column_name in enumerate(header):
            if norm(column_name) == "FORNECEDOR" and norm(row[idx]) == "POOL":
                row[idx] = "POOL"
        original = row[motorista_idx]
        resolved = resolve_name(original, official)
        row[original_idx] = row[original_idx] or original
        row[ref_idx] = resolved
        if resolved:
            row[motorista_idx] = resolved

    return rows


def excel_col(index: int) -> str:
    index += 1
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def sheet_xml(rows: list[list[str]]) -> str:
    body = []
    for r_idx, row in enumerate(rows, 1):
        cells = []
        for c_idx, value in enumerate(row, 1):
            text = "" if value is None else str(value)
            ref = f"{excel_col(c_idx - 1)}{r_idx}"
            cells.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{html.escape(text)}</t></is></c>'
            )
        body.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetData>{"".join(body)}</sheetData></worksheet>'
    )


def write_workbook(path: Path, sheets: dict[str, list[list[str]]]) -> None:
    now = dt.datetime.now(dt.UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    sheet_names = list(sheets)
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        overrides = [
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        ]
        for i in range(1, len(sheet_names) + 1):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + "".join(overrides)
            + "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
            + "".join(
                f'<sheet name="{html.escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                for i, name in enumerate(sheet_names, 1)
            )
            + "</sheets></workbook>",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, len(sheet_names) + 1)
            )
            + "</Relationships>",
        )
        zf.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
            "</cp:coreProperties>",
        )
        zf.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Codex XLSX Normalizer</Application></Properties>",
        )
        for i, name in enumerate(sheet_names, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(sheets[name]))


def main() -> None:
    root = Path.cwd()
    official = build_official_names(root)
    print(f"Nomes referenciais de motoristas: {len(official)}")
    for filename in FILES:
        source = root / filename
        target = root / f"{source.stem}_NORMALIZADO.xlsx"
        workbook = read_workbook(source)
        normalized = {
            name: normalize_sheet(rows, official)
            for name, rows in workbook.items()
        }
        write_workbook(target, normalized)
        print(f"Gerado: {target.name} ({len(normalized)} abas)")


if __name__ == "__main__":
    main()
