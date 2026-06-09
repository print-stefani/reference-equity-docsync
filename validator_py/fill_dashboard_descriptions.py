#!/usr/bin/env python3
import argparse
import csv
import html
import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path
from typing import Callable

PLACEHOLDER = "A preencher pela documentação funcional 5"
XLSX_PLACEHOLDER_PREFIX = "A preencher pela documentação funcional"
XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

TR_PAT = re.compile(r"<tr>(.*?)</tr>", re.I | re.S)
TD_PAT = re.compile(r"<td[^>]*>(.*?)</td>", re.I | re.S)
TAG_PAT = re.compile(r"<[^>]+>")


def strip_tags(s: str) -> str:
    return html.unescape(TAG_PAT.sub("", s)).strip()


def normalize_ws(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def excel_col_to_index(ref: str) -> int:
    col = "".join(ch for ch in (ref or "").upper() if "A" <= ch <= "Z")
    if not col:
        return 0
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def excel_index_to_col(idx: int) -> str:
    if idx < 0:
        return "A"
    n = idx + 1
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def _xlsx_load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    out: list[str] = []
    for si in root:
        if local_name(si.tag) != "si":
            continue
        chunks: list[str] = []
        for node in si.iter():
            if local_name(node.tag) == "t":
                chunks.append(node.text or "")
        out.append("".join(chunks))
    return out


def xlsx_cell_text(cell: ET.Element, shared: list[str]) -> str:
    ctype = cell.attrib.get("t", "")
    if ctype == "s":
        v_node = next((n for n in cell if local_name(n.tag) == "v"), None)
        if v_node is None:
            return ""
        raw = (v_node.text or "").strip()
        if raw.isdigit():
            i = int(raw)
            if 0 <= i < len(shared):
                return normalize_ws(shared[i])
        return ""
    if ctype == "inlineStr":
        chunks: list[str] = []
        for node in cell.iter():
            if local_name(node.tag) == "t":
                chunks.append(node.text or "")
        return normalize_ws("".join(chunks))
    v_node = next((n for n in cell if local_name(n.tag) == "v"), None)
    return normalize_ws(v_node.text or "") if v_node is not None else ""


def xlsx_row_values(row: ET.Element, shared: list[str]) -> dict[int, str]:
    values: dict[int, str] = {}
    for cell in row:
        if local_name(cell.tag) != "c":
            continue
        idx = excel_col_to_index(cell.attrib.get("r", ""))
        values[idx] = xlsx_cell_text(cell, shared)
    return values


def set_xlsx_inline_text(cell: ET.Element, value: str) -> None:
    cell.attrib["t"] = "inlineStr"
    for child in list(cell):
        cell.remove(child)
    is_node = ET.SubElement(cell, f"{{{XLSX_MAIN_NS}}}is")
    t_node = ET.SubElement(is_node, f"{{{XLSX_MAIN_NS}}}t")
    t_node.text = value


def find_or_create_cell(row: ET.Element, idx: int, row_num: int) -> ET.Element:
    wanted_ref = f"{excel_index_to_col(idx)}{row_num}"
    for cell in row:
        if local_name(cell.tag) != "c":
            continue
        if (cell.attrib.get("r", "") or "").upper() == wanted_ref:
            return cell

    new_cell = ET.Element(f"{{{XLSX_MAIN_NS}}}c")
    new_cell.attrib["r"] = wanted_ref
    row.append(new_cell)
    return new_cell


def enrich_single_xlsx(
    xlsx_path: Path,
    resolve_description: Callable[[str, str], str],
) -> tuple[int, int, list[str]]:
    ET.register_namespace("", XLSX_MAIN_NS)
    unresolved: list[str] = []
    candidates = 0
    updated = 0

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        names = zin.namelist()
        file_data = {name: zin.read(name) for name in names}
        shared = _xlsx_load_shared_strings(zin)

    changed = False
    sheet_names = [n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")]
    for sheet_name in sheet_names:
        raw = file_data.get(sheet_name, b"")
        if not raw:
            continue
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            continue
        sheet_data = next((x for x in root.iter() if local_name(x.tag) == "sheetData"), None)
        if sheet_data is None:
            continue

        rows = [r for r in sheet_data if local_name(r.tag) == "row"]
        if not rows:
            continue

        header_row = rows[0]
        header_vals = xlsx_row_values(header_row, shared)
        header_map = {normalize_ws(v).upper(): i for i, v in header_vals.items() if v}

        desc_idx = header_map.get("COLUMN DESCRIPTION")
        table_idx = header_map.get("ASSOCIATED TABLE TO EACH COLUMN")
        col_idx = header_map.get("COLUMN NAME")
        if desc_idx is None or table_idx is None or col_idx is None:
            continue

        for row in rows[1:]:
            row_num_text = row.attrib.get("r", "")
            row_num = int(row_num_text) if row_num_text.isdigit() else 0
            vals = xlsx_row_values(row, shared)

            desc_raw = vals.get(desc_idx, "")
            if not desc_raw.upper().startswith(XLSX_PLACEHOLDER_PREFIX.upper()):
                continue

            table_name = vals.get(table_idx, "")
            col_name = vals.get(col_idx, "")
            table_norm = normalize_table_name(table_name)
            col_norm = normalize_table_name(col_name)
            if not table_norm or table_norm in {"N/A", "MULTI_SOURCE"} or not col_norm:
                continue

            candidates += 1
            new_desc = resolve_description(table_norm, col_norm)
            if not new_desc:
                unresolved.append(f"{xlsx_path.name}:{table_norm}.{col_norm}")
                continue

            target_cell = find_or_create_cell(row, desc_idx, row_num if row_num > 0 else 1)
            set_xlsx_inline_text(target_cell, new_desc)
            updated += 1
            changed = True

        if changed:
            file_data[sheet_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    if changed:
        tmp_path = xlsx_path.with_suffix(f"{xlsx_path.suffix}.tmp")
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in file_data:
                    zout.writestr(name, file_data[name])
        tmp_path.replace(xlsx_path)

    return candidates, updated, sorted(set(unresolved))


def enrich_xlsx_directory(
    xlsx_dir: Path,
    resolve_description: Callable[[str, str], str],
) -> tuple[int, int, int, list[str]]:
    if not xlsx_dir.exists():
        return 0, 0, 0, []

    files = sorted(xlsx_dir.glob("*.xlsx"), key=lambda p: p.name.upper())
    file_count = 0
    total_candidate = 0
    total_updated = 0
    unresolved: list[str] = []

    for xlsx in files:
        file_count += 1
        cand, upd, unr = enrich_single_xlsx(xlsx, resolve_description)
        total_candidate += cand
        total_updated += upd
        unresolved.extend(unr)
        print(f"[XLSX] {xlsx.name}: candidatos={cand}, atualizados={upd}, não_resolvidos={len(unr)}")

    return file_count, total_candidate, total_updated, sorted(set(unresolved))


def read_xlsx_rows(xlsx_path: Path, sheet_name: str = "xl/worksheets/sheet1.xml") -> list[list[str]]:
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        shared = _xlsx_load_shared_strings(zf)
        raw = zf.read(sheet_name)
    root = ET.fromstring(raw)
    sheet_data = next((x for x in root.iter() if local_name(x.tag) == "sheetData"), None)
    if sheet_data is None:
        return []
    rows: list[list[str]] = []
    for row in sheet_data:
        if local_name(row.tag) != "row":
            continue
        values: dict[int, str] = {}
        max_idx = -1
        for cell in row:
            if local_name(cell.tag) != "c":
                continue
            ref = cell.attrib.get("r", "")
            idx = excel_col_to_index(ref)
            max_idx = max(max_idx, idx)
            ctype = cell.attrib.get("t", "")
            val = ""
            if ctype == "s":
                v_node = next((n for n in cell if local_name(n.tag) == "v"), None)
                if v_node is not None and (v_node.text or "").strip().isdigit():
                    s_idx = int((v_node.text or "0").strip())
                    if 0 <= s_idx < len(shared):
                        val = shared[s_idx]
            elif ctype == "inlineStr":
                chunks: list[str] = []
                for node in cell.iter():
                    if local_name(node.tag) == "t":
                        chunks.append(node.text or "")
                val = "".join(chunks)
            else:
                v_node = next((n for n in cell if local_name(n.tag) == "v"), None)
                if v_node is not None:
                    val = v_node.text or ""
            values[idx] = normalize_ws(val)
        if max_idx < 0:
            continue
        rows.append([values.get(i, "") for i in range(max_idx + 1)])
    return rows


def normalize_flex_setup_row(raw_row: list[str]) -> list[str]:
    row = [normalize_ws(x) for x in (raw_row or [])]
    filled = [x for x in row if x]
    if len(filled) == 1 and "," in filled[0]:
        try:
            parsed = next(csv.reader([filled[0]]))
            return [normalize_ws(x) for x in parsed]
        except Exception:
            return row
    return row


def load_flex_setup_reference(xlsx_path: Path) -> dict[str, dict[str, str]]:
    if not xlsx_path.exists():
        return {}
    rows = [normalize_flex_setup_row(r) for r in read_xlsx_rows(xlsx_path)]
    refs: dict[str, dict[str, str]] = {}
    for row in rows:
        if not row or not row[0]:
            continue
        tag = row[0].strip()
        if tag.upper().replace(" ", "") == "XMLTAGNAME":
            continue
        key = tag.upper()
        refs[key] = {
            "xml_tag_name": key,
            "column_description": row[1].strip() if len(row) > 1 else "",
            "associated_table_to_each_column": row[2].strip() if len(row) > 2 else "",
            "column_name": row[3].strip() if len(row) > 3 else "",
            "column_characteristic": row[4].strip() if len(row) > 4 else "",
        }
    return refs


def apply_local_flex_setup_updates(html_text: str, refs: dict[str, dict[str, str]]) -> tuple[str, int, int, list[str]]:
    replacements = {}
    candidate = 0
    updated = 0
    unresolved: list[str] = []
    current_dataset = ""

    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if len(tds) == 1:
            current_dataset = strip_tags(tds[0]).upper()
            continue
        if len(tds) < 5:
            continue
        if current_dataset != "G_FLEX_SETUP":
            continue
        if strip_tags(tds[1]) != PLACEHOLDER:
            continue

        xml_tag = strip_tags(tds[0]).upper()
        candidate += 1
        ref = refs.get(xml_tag)
        if not ref:
            # Fallback: allow stable suffix variants like ISV_INVOICE_TYPE_L.
            for key in sorted(refs.keys(), key=len, reverse=True):
                if xml_tag.startswith(f"{key}_"):
                    ref = refs[key]
                    break
        if not ref:
            unresolved.append(xml_tag)
            continue

        new_tr = tr
        desc = (ref.get("column_description") or "").strip()
        if desc:
            new_tr = new_tr.replace(f"<td>{tds[1]}</td>", f"<td>{html.escape(desc)}</td>", 1)

        assoc_in = strip_tags(tds[2]).upper()
        assoc_ref = (ref.get("associated_table_to_each_column") or "").strip()
        if assoc_ref and assoc_in in {"", "N/A", "MULTI_SOURCE"}:
            new_tr = new_tr.replace(f"<td>{tds[2]}</td>", f"<td>{html.escape(assoc_ref)}</td>", 1)

        col_in = strip_tags(tds[3]).upper()
        col_ref = (ref.get("column_name") or "").strip()
        if col_ref and col_in in {"", "N/A", "MULTI_SOURCE", "UNKNOWN_SOURCE_COLUMN"}:
            new_tr = new_tr.replace(f"<td>{tds[3]}</td>", f"<td>{html.escape(col_ref)}</td>", 1)

        typ_in = strip_tags(tds[4]).upper()
        typ_ref = (ref.get("column_characteristic") or "").strip()
        if typ_ref and typ_in in {"", "N/A", "VARCHAR2"}:
            new_tr = new_tr.replace(f"<td>{tds[4]}</td>", f"<td>{html.escape(typ_ref)}</td>", 1)

        replacements[(m.start(), m.end())] = new_tr
        updated += 1

    for (s, e), new_tr in sorted(replacements.items(), key=lambda x: x[0][0], reverse=True):
        html_text = html_text[:s] + new_tr + html_text[e:]
    return html_text, candidate, updated, sorted(set(unresolved))


def fetch(url: str, timeout: int = 30) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "ignore")


def normalize_table_name(name: str) -> str:
    n = (name or "").strip().strip('"').upper()
    if "." in n:
        n = n.split(".")[-1]
    return n


def strip_sql_comments(sql: str) -> str:
    sql = re.sub(r"/\*.*?\*/", " ", sql or "", flags=re.S)
    sql = re.sub(r"--[^\n\r]*", " ", sql)
    return sql


def split_top_level(text: str, delimiter: str = ",") -> list[str]:
    parts: list[str] = []
    buf: list[str] = []
    depth = 0
    quote = ""
    i = 0
    while i < len(text):
        ch = text[i]
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = ""
            i += 1
            continue
        if ch in ("'", '"'):
            quote = ch
            buf.append(ch)
            i += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            i += 1
            continue
        if ch == ")":
            depth = max(0, depth - 1)
            buf.append(ch)
            i += 1
            continue
        if ch == delimiter and depth == 0:
            item = "".join(buf).strip()
            if item:
                parts.append(item)
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)
    return parts


def find_keyword_top_level(sql: str, keyword: str, start: int = 0) -> int:
    kw = keyword.upper()
    depth = 0
    quote = ""
    i = start
    while i < len(sql):
        ch = sql[i]
        if quote:
            if ch == quote:
                quote = ""
            i += 1
            continue
        if ch in ("'", '"'):
            quote = ch
            i += 1
            continue
        if ch == "(":
            depth += 1
            i += 1
            continue
        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue
        if depth == 0 and sql[i : i + len(kw)].upper() == kw:
            prev_ok = i == 0 or not (sql[i - 1].isalnum() or sql[i - 1] == "_")
            nxt = i + len(kw)
            next_ok = nxt >= len(sql) or not (sql[nxt].isalnum() or sql[nxt] == "_")
            if prev_ok and next_ok:
                return i
        i += 1
    return -1


def extract_select_and_from(sql: str) -> tuple[str, str]:
    cleaned = re.sub(r"\s+", " ", strip_sql_comments(sql))
    sel = find_keyword_top_level(cleaned, "SELECT")
    if sel < 0:
        return "", ""
    frm = find_keyword_top_level(cleaned, "FROM", sel + 6)
    if frm < 0:
        return "", ""
    return cleaned[sel + 6 : frm].strip(), cleaned[frm + 4 :].strip()


def extract_query_text_from_oracle_doc(doc_html: str) -> str:
    blocks = re.findall(r"<pre[^>]*>(.*?)</pre>", doc_html, flags=re.I | re.S)
    best = ""
    for b in blocks:
        text = html.unescape(re.sub(r"<[^>]+>", " ", b))
        text = re.sub(r"\s+", " ", text).strip()
        if "SELECT " in text.upper() and " FROM " in text.upper() and len(text) > len(best):
            best = text
    if not best:
        p_blocks = re.findall(r"<p[^>]*class=[\"']p[\"'][^>]*>(.*?)</p>", doc_html, flags=re.I | re.S)
        if p_blocks:
            text = " ".join(strip_tags(x) for x in p_blocks)
            text = re.sub(r"\s+", " ", text).strip()
            if "SELECT " in text.upper() and " FROM " in text.upper():
                best = text
    return best


def parse_query_lineage(sql: str) -> tuple[dict[str, str], dict[str, str]]:
    select_part, from_part = extract_select_and_from(sql)
    if not select_part:
        return {}, {}

    alias_to_table: dict[str, str] = {}
    # Explicit FROM/JOIN form
    for m in re.finditer(r"\b(?:FROM|JOIN)\s+([A-Z0-9_$.\"#]+)(?:\s+(?:AS\s+)?([A-Z0-9_\"#]+))?", from_part, flags=re.I):
        tbl_raw = (m.group(1) or "").strip().rstrip(",")
        ali_raw = (m.group(2) or "").strip().rstrip(",")
        tbl = normalize_table_name(tbl_raw)
        if not tbl:
            continue
        if ali_raw:
            alias_to_table[normalize_table_name(ali_raw)] = tbl
        alias_to_table.setdefault(tbl, tbl)

    # Comma-separated FROM list form (legacy Oracle style).
    from_clause = from_part
    for kw in (" WHERE ", " GROUP BY ", " ORDER BY ", " UNION ", " CONNECT BY ", " START WITH "):
        i = from_clause.upper().find(kw)
        if i >= 0:
            from_clause = from_clause[:i]
            break
    for seg in split_top_level(from_clause, ","):
        s = seg.strip()
        if not s:
            continue
        # Remove leading JOIN/FROM keywords and trailing ON clause.
        s = re.sub(r"(?i)^(FROM|JOIN)\s+", "", s).strip()
        s = re.split(r"(?i)\bON\b", s, maxsplit=1)[0].strip()
        tokens = re.split(r"\s+", s)
        if not tokens:
            continue
        tbl = normalize_table_name(tokens[0])
        if not tbl:
            continue
        alias = ""
        if len(tokens) >= 2:
            if tokens[1].upper() == "AS" and len(tokens) >= 3:
                alias = normalize_table_name(tokens[2])
            elif tokens[1].upper() not in {"LEFT", "RIGHT", "FULL", "INNER", "OUTER", "CROSS", "JOIN"}:
                alias = normalize_table_name(tokens[1])
        if alias:
            alias_to_table[alias] = tbl
        alias_to_table.setdefault(tbl, tbl)

    output_to_expr: dict[str, str] = {}
    for item in split_top_level(select_part):
        it = item.strip()
        if not it:
            continue

        alias = ""
        expr = it

        m_as = re.match(r"(?is)^(.*)\s+AS\s+\"?([A-Z0-9_#$]+)\"?$", it)
        if m_as:
            expr = m_as.group(1).strip()
            alias = m_as.group(2).strip().upper()
        else:
            m_tail = re.match(r"(?is)^(.*?)(?:\s+)\"?([A-Z0-9_#$]+)\"?$", it)
            if m_tail:
                candidate_expr = m_tail.group(1).strip()
                candidate_alias = m_tail.group(2).strip().upper()
                if candidate_expr:
                    expr = candidate_expr
                    alias = candidate_alias

        if not alias:
            m_col = re.search(r"(?i)([A-Z0-9_#$]+)\.([A-Z0-9_#$]+)$", expr)
            if m_col:
                alias = m_col.group(2).upper()
            else:
                alias = normalize_table_name(expr)

        if alias:
            output_to_expr[alias] = expr

    return output_to_expr, alias_to_table


def alias_column_refs(expr: str) -> list[tuple[str, str]]:
    refs = re.findall(r"(?i)\b([A-Z0-9_#$]+)\.([A-Z0-9_#$]+)\b", expr or "")
    out: list[tuple[str, str]] = []
    for a, c in refs:
        out.append((normalize_table_name(a), c.upper()))
    return out


def extract_doc_info(doc_html: str) -> dict:
    col_desc = extract_col_desc_from_oracle_doc(doc_html)
    query_text = extract_query_text_from_oracle_doc(doc_html)
    select_map, alias_map = parse_query_lineage(query_text) if query_text else ({}, {})
    return {
        "col_desc": col_desc,
        "query_text": query_text,
        "select_map": select_map,
        "alias_map": alias_map,
    }


def resolve_table_url(table_name: str, table_to_url: dict[str, str]) -> str:
    return table_to_url.get(normalize_table_name(table_name), "")


def resolve_description_via_lineage(
    table_name: str,
    column_name: str,
    table_to_url: dict[str, str],
    fetch_cache: dict[str, str],
    parse_cache: dict[str, dict],
    visited: set[tuple[str, str]],
    depth: int = 0,
) -> str:
    t = normalize_table_name(table_name)
    c = normalize_table_name(column_name)
    if not t or not c or depth > 4:
        return ""
    key = (t, c)
    if key in visited:
        return ""
    visited.add(key)

    url = resolve_table_url(t, table_to_url)
    if not url:
        return ""

    if url not in fetch_cache:
        try:
            fetch_cache[url] = fetch(url)
        except Exception:
            return ""
    doc_html = fetch_cache[url]

    if url not in parse_cache:
        parse_cache[url] = extract_doc_info(doc_html)
    info = parse_cache[url]

    direct = info["col_desc"].get(c, "")
    if direct:
        return direct

    expr = info["select_map"].get(c, "")
    if not expr:
        return ""

    for alias, src_col in alias_column_refs(expr):
        src_table = info["alias_map"].get(alias, "")
        if not src_table:
            continue
        desc = resolve_description_via_lineage(
            table_name=src_table,
            column_name=src_col,
            table_to_url=table_to_url,
            fetch_cache=fetch_cache,
            parse_cache=parse_cache,
            visited=visited,
            depth=depth + 1,
        )
        if desc:
            return desc
    return ""


def extract_col_desc_from_oracle_doc(doc_html: str) -> dict:
    """Extract COLUMN_NAME -> description from Oracle OEDM-like pages.

    Expected row shape usually has at least 6 columns where:
    [Name, Type, Length, Precision, Not-null, Comments, ...]
    """
    col_desc = {}
    for m in re.finditer(r"<tr[^>]*>(.*?)</tr>", doc_html, re.I | re.S):
        cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", m.group(1), re.I | re.S)
        vals = [strip_tags(c) for c in cells]
        if len(vals) >= 6 and re.fullmatch(r"[A-Z0-9_#$]+", vals[0] or ""):
            if vals[1].upper() in ("TYPE", "DATATYPE"):
                continue
            desc = vals[5].strip()
            if desc and desc.upper() not in ("COMMENTS", "DESCRIPTION"):
                col_desc[vals[0].upper()] = desc
    return col_desc


def apply_table_updates(
    html_text: str,
    table_name: str,
    col_desc: dict,
    fallback_resolver: Callable[[str], str] | None = None,
) -> tuple[str, int, int, list]:
    replacements = {}
    candidate = 0
    updated = 0

    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if len(tds) < 5:
            continue

        desc_cell = strip_tags(tds[1])
        table_cell = strip_tags(tds[2]).upper()
        col_cell = strip_tags(tds[3]).upper()

        if desc_cell != PLACEHOLDER:
            continue
        if table_name.upper() not in table_cell:
            continue

        candidate += 1
        desc = col_desc.get(col_cell)
        if (not desc) and fallback_resolver is not None:
            desc = fallback_resolver(col_cell)
        if desc:
            new_tr = tr.replace(f"<td>{tds[1]}</td>", f"<td>{html.escape(desc)}</td>", 1)
            replacements[(m.start(), m.end())] = new_tr
            updated += 1

    for (s, e), new_tr in sorted(replacements.items(), key=lambda x: x[0][0], reverse=True):
        html_text = html_text[:s] + new_tr + html_text[e:]

    unresolved_cols = []
    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if len(tds) < 5:
            continue
        if strip_tags(tds[1]) == PLACEHOLDER and table_name.upper() in strip_tags(tds[2]).upper():
            unresolved_cols.append(strip_tags(tds[3]).upper())

    return html_text, candidate, updated, sorted(set(unresolved_cols))


def count_placeholder_candidates(html_text: str, table_name: str) -> int:
    """Count rows that are potential candidates for a table before replacement."""
    count = 0
    tname = table_name.upper()
    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if len(tds) < 5:
            continue
        if strip_tags(tds[1]) != PLACEHOLDER:
            continue
        if tname in strip_tags(tds[2]).upper():
            count += 1
    return count


def parse_mapping_file(path: Path) -> dict:
    """Parse TABLE=URL lines (ignore blank lines and comments starting with #)."""
    mapping = {}
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            raise ValueError(f"Invalid mapping line: {raw}")
        table, url = line.split("=", 1)
        mapping[table.strip().upper()] = url.strip()
    return mapping


def parse_lineage_file(path: Path) -> dict[tuple[str, str], tuple[str, str]]:
    """Parse lineage CSV with headers:
    source_table,source_column,target_table,target_column
    """
    lineage: dict[tuple[str, str], tuple[str, str]] = {}
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        required = {"source_table", "source_column", "target_table", "target_column"}
        if not required.issubset({(x or "").strip() for x in (reader.fieldnames or [])}):
            raise ValueError(f"Invalid lineage CSV headers in {path}. Required: {', '.join(sorted(required))}")
        for row in reader:
            src_t = (row.get("source_table") or "").strip().upper()
            src_c = (row.get("source_column") or "").strip().upper()
            dst_t = (row.get("target_table") or "").strip().upper()
            dst_c = (row.get("target_column") or "").strip().upper()
            if src_t and src_c and dst_t and dst_c:
                lineage[(src_t, src_c)] = (dst_t, dst_c)
    return lineage


def apply_lineage_updates(
    html_text: str,
    lineage_map: dict[tuple[str, str], tuple[str, str]],
    resolve_col_desc: Callable[[str], dict[str, str]],
) -> tuple[str, int, int, list[str]]:
    """Fallback updates using source->target table/column lineage mapping."""
    replacements = {}
    candidate = 0
    updated = 0
    unresolved: list[str] = []

    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if len(tds) < 5:
            continue
        if strip_tags(tds[1]) != PLACEHOLDER:
            continue

        src_table = strip_tags(tds[2]).upper()
        src_col = strip_tags(tds[3]).upper()
        key = (src_table, src_col)
        if key not in lineage_map:
            continue

        candidate += 1
        dst_table, dst_col = lineage_map[key]
        col_desc = resolve_col_desc(dst_table)
        desc = col_desc.get(dst_col)
        if not desc:
            unresolved.append(f"{src_table}.{src_col} -> {dst_table}.{dst_col}")
            continue

        new_tr = tr.replace(f"<td>{tds[1]}</td>", f"<td>{html.escape(desc)}</td>", 1)
        replacements[(m.start(), m.end())] = new_tr
        updated += 1

    for (s, e), new_tr in sorted(replacements.items(), key=lambda x: x[0][0], reverse=True):
        html_text = html_text[:s] + new_tr + html_text[e:]

    return html_text, candidate, updated, sorted(set(unresolved))


def collect_remaining_placeholders(html_text: str) -> list[dict[str, str]]:
    """Collect rows still containing PLACEHOLDER after processing.

    Expected dashboard row shape:
    [XML Tag Name, Column description, Associated table to each column, Column Name, Column characteristic]
    """
    rows: list[dict[str, str]] = []
    current_dataset = ""

    for m in TR_PAT.finditer(html_text):
        tr = m.group(0)
        tds = TD_PAT.findall(tr)
        if not tds:
            continue

        # Dataset separator row usually has a single cell with dataset name.
        if len(tds) == 1:
            ds = strip_tags(tds[0])
            if ds:
                current_dataset = ds
            continue

        if len(tds) < 5:
            continue
        if strip_tags(tds[1]) != PLACEHOLDER:
            continue

        rows.append(
            {
                "dataset": current_dataset,
                "xml_tag_name": strip_tags(tds[0]),
                "description": strip_tags(tds[1]),
                "associated_table": strip_tags(tds[2]),
                "column_name": strip_tags(tds[3]),
                "column_characteristic": strip_tags(tds[4]),
            }
        )

    return rows


def write_remaining_log(rows: list[dict[str, str]], log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "dataset",
                "xml_tag_name",
                "description",
                "associated_table",
                "column_name",
                "column_characteristic",
            ],
        )
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description="Preenche placeholders do dashboard a partir de documentação de referência mapeada por tabela/coluna.")
    ap.add_argument("--input", required=True, help="Input HTML path")
    ap.add_argument("--output", required=True, help="Output HTML path")
    ap.add_argument("--mapping-file", required=True, help="Text file with TABLE=URL per line")
    ap.add_argument(
        "--flex-setup-file",
        default="",
        help="Arquivo XLSX opcional com descrições locais de G_FLEX_SETUP.",
    )
    ap.add_argument(
        "--skip-web",
        action="store_true",
        help="Ignora consulta web e aplica somente regras locais/de fallback.",
    )
    ap.add_argument(
        "--lineage-file",
        default="",
        help="CSV opcional com regras fallback source_table/source_column/target_table/target_column",
    )
    ap.add_argument("--tables", default="", help="Filtro opcional de tabelas separado por vírgula (nomes em maiúsculas)")
    ap.add_argument("--alias", action="append", default=[], help="Alias opcional no formato ALIAS=REAL_TABLE")
    ap.add_argument(
        "--remaining-log",
        default="",
        help="Caminho CSV opcional para linhas que continuam com placeholder após o processamento",
    )
    ap.add_argument(
        "--diagnostic-log",
        default="",
        help="Caminho CSV opcional com diagnóstico por tabela (status de busca/parse/atualização)",
    )
    ap.add_argument(
        "--xlsx-dir",
        default="",
        help="Diretório opcional com arquivos XLSX de extractor para enriquecer (padrão: pasta irmã by_extractor).",
    )
    ap.add_argument(
        "--skip-xlsx",
        action="store_true",
        help="Ignora etapa de enriquecimento XLSX.",
    )
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    map_path = Path(args.mapping_file)
    flex_setup_path = Path(args.flex_setup_file) if args.flex_setup_file.strip() else map_path.with_name("G_FLEX_SETUP.xlsx")
    remaining_log_path = Path(args.remaining_log) if args.remaining_log.strip() else out_path.with_name(
        f"{out_path.stem}_remaining_tofill5.csv"
    )
    diagnostic_log_path = Path(args.diagnostic_log) if args.diagnostic_log.strip() else out_path.with_name(
        f"{out_path.stem}_table_diagnostics.csv"
    )
    xlsx_dir = Path(args.xlsx_dir) if args.xlsx_dir.strip() else out_path.parent / "by_extractor"

    html_text = in_path.read_text(encoding="utf-8", errors="ignore")
    table_to_url = parse_mapping_file(map_path)
    default_lineage_path = map_path.with_name("table_column_lineage.csv")
    lineage_path = Path(args.lineage_file) if args.lineage_file.strip() else default_lineage_path
    lineage_map: dict[tuple[str, str], tuple[str, str]] = {}
    if lineage_path.exists():
        lineage_map = parse_lineage_file(lineage_path)

    aliases = {}
    for a in args.alias:
        if "=" not in a:
            raise ValueError(f"Invalid alias: {a}")
        k, v = a.split("=", 1)
        aliases[k.strip().upper()] = v.strip().upper()

    selected = None
    if args.tables.strip():
        selected = {t.strip().upper() for t in args.tables.split(",") if t.strip()}

    # Build update order: tables present in mapping (filtered if requested)
    work_tables = list(table_to_url.keys())
    if selected is not None:
        work_tables = [t for t in work_tables if t in selected]

    total_candidate = 0
    total_updated = 0
    table_col_desc_cache: dict[str, dict[str, str]] = {}
    fetch_cache: dict[str, str] = {}
    parse_cache: dict[str, dict] = {}
    diagnostics: list[dict[str, str | int]] = []

    flex_refs = load_flex_setup_reference(flex_setup_path) if flex_setup_path.exists() else {}
    if flex_refs:
        html_text, flex_cand, flex_upd, flex_unresolved = apply_local_flex_setup_updates(html_text, flex_refs)
        total_candidate += flex_cand
        total_updated += flex_upd
        print(
            f"[LOCAL] G_FLEX_SETUP: candidatos={flex_cand}, atualizados={flex_upd}, "
            f"não_resolvidos={len(flex_unresolved)} (file: {flex_setup_path})"
        )
        if flex_unresolved:
            print(f"        unresolved xml tags: {', '.join(flex_unresolved[:20])}")
        diagnostics.append(
            {
                "table": "G_FLEX_SETUP",
                "real_table": "G_FLEX_SETUP",
                "url": str(flex_setup_path),
                "url_type": "local_xlsx",
                "status": "OK_LOCAL_PARTIAL" if flex_unresolved else "OK_LOCAL_FULL",
                "reason": "Filled from local G_FLEX_SETUP.xlsx",
                "pre_candidates": flex_cand,
                "parsed_columns": len(flex_refs),
                "updated": flex_upd,
                "unresolved_after": len(flex_unresolved),
                "unresolved_columns_sample": ", ".join(flex_unresolved[:20]),
            }
        )
    else:
        print(f"[LOCAL] G_FLEX_SETUP: local file not found or empty ({flex_setup_path})")

    def resolve_table_col_desc(table_name: str) -> dict[str, str]:
        t = table_name.upper()
        if t in table_col_desc_cache:
            return table_col_desc_cache[t]
        real_t = aliases.get(t, t)
        u = table_to_url.get(real_t) or table_to_url.get(t)
        if not u:
            table_col_desc_cache[t] = {}
            return {}
        try:
            parsed = extract_col_desc_from_oracle_doc(fetch(u))
        except Exception:
            parsed = {}
        table_col_desc_cache[t] = parsed
        return parsed

    def resolve_table_column_description(table_name: str, column_name: str) -> str:
        t = normalize_table_name(table_name)
        c = normalize_table_name(column_name)
        if not t or not c:
            return ""

        desc_map = resolve_table_col_desc(t)
        if c in desc_map:
            return desc_map[c]

        if args.skip_web:
            return ""

        return resolve_description_via_lineage(
            table_name=t,
            column_name=c,
            table_to_url=table_to_url,
            fetch_cache=fetch_cache,
            parse_cache=parse_cache,
            visited=set(),
        )

    if not args.skip_web:
        for table in work_tables:
            real_table = aliases.get(table, table)
            url = table_to_url.get(real_table) or table_to_url.get(table)
            pre_candidates = count_placeholder_candidates(html_text, table)

            if pre_candidates == 0:
                diagnostics.append(
                    {
                        "table": table,
                        "real_table": real_table,
                        "url": url or "",
                        "url_type": "search" if (url and "/search?" in url) else ("direct" if url else "none"),
                        "status": "SKIP_NO_CANDIDATES",
                        "reason": "No placeholder rows matched this table in HTML",
                        "pre_candidates": pre_candidates,
                        "parsed_columns": 0,
                        "updated": 0,
                        "unresolved_after": 0,
                        "unresolved_columns_sample": "",
                    }
                )
                continue

            if not url:
                print(f"[SKIP] {table}: no URL")
                diagnostics.append(
                    {
                        "table": table,
                        "real_table": real_table,
                        "url": "",
                        "url_type": "none",
                        "status": "SKIP_NO_URL",
                        "reason": "No mapping URL for table",
                        "pre_candidates": pre_candidates,
                        "parsed_columns": 0,
                        "updated": 0,
                        "unresolved_after": pre_candidates,
                        "unresolved_columns_sample": "",
                    }
                )
                continue

            try:
                doc_html = fetch(url)
            except Exception as e:
                print(f"[ERROR] {table}: failed to fetch doc URL: {e}")
                diagnostics.append(
                    {
                        "table": table,
                        "real_table": real_table,
                        "url": url,
                        "url_type": "search" if "/search?" in url else "direct",
                        "status": "ERROR_FETCH",
                        "reason": str(e),
                        "pre_candidates": pre_candidates,
                        "parsed_columns": 0,
                        "updated": 0,
                        "unresolved_after": pre_candidates,
                        "unresolved_columns_sample": "",
                    }
                )
                continue

            info = extract_doc_info(doc_html)
            col_desc = info["col_desc"]
            table_col_desc_cache[table.upper()] = col_desc
            parse_empty = False
            if not col_desc:
                parse_empty = True
                print(f"[WARN] {table}: no column descriptions parsed from {url}")

            def resolver(col_name: str) -> str:
                return resolve_description_via_lineage(
                    table_name=table,
                    column_name=col_name,
                    table_to_url=table_to_url,
                    fetch_cache=fetch_cache,
                    parse_cache=parse_cache,
                    visited=set(),
                )

            html_text, cand, upd, unresolved = apply_table_updates(html_text, table, col_desc, resolver)
            total_candidate += cand
            total_updated += upd

            print(f"[OK] {table}: candidatos={cand}, atualizados={upd}, não_resolvidos={len(unresolved)}")
            if unresolved:
                print(f"      unresolved columns: {', '.join(unresolved[:20])}")
            if cand == 0:
                status = "WARN_NO_CANDIDATES_IN_HTML"
                reason = "Table has mapping and parsed docs, but no placeholder rows matched this table in HTML"
            elif upd == 0:
                status = "WARN_NO_COLUMN_MATCH"
                reason = "Candidates exist but none of the column names matched parsed documentation"
            elif unresolved:
                status = "OK_PARTIAL"
                reason = "Some rows were updated, some remain unresolved"
            else:
                status = "OK_FULL"
                reason = "All candidate rows for this table were updated"

            if parse_empty and upd == 0:
                status = "WARN_PARSE_EMPTY"
                reason = "No parseable column-description rows found"
            elif parse_empty and upd > 0:
                status = "OK_LINEAGE_FALLBACK"
                reason = "Updated using query-lineage fallback from related base tables"

            diagnostics.append(
                {
                    "table": table,
                    "real_table": real_table,
                    "url": url,
                    "url_type": "search" if "/search?" in url else "direct",
                    "status": status,
                    "reason": reason,
                    "pre_candidates": pre_candidates,
                    "parsed_columns": len(col_desc),
                    "updated": upd,
                    "unresolved_after": len(unresolved),
                    "unresolved_columns_sample": ", ".join(unresolved[:20]),
                }
            )
    else:
        print("[SKIP] Consulta web desabilitada (--skip-web).")

    if lineage_map and not args.skip_web:
        html_text, lineage_candidate, lineage_updated, lineage_unresolved = apply_lineage_updates(
            html_text=html_text,
            lineage_map=lineage_map,
            resolve_col_desc=resolve_table_col_desc,
        )
        total_candidate += lineage_candidate
        total_updated += lineage_updated
        print(
            f"[LINEAGE] rules={len(lineage_map)} candidatos={lineage_candidate} "
            f"atualizados={lineage_updated} não_resolvidos={len(lineage_unresolved)}"
        )
        if lineage_unresolved:
            print(f"          unresolved lineage: {', '.join(lineage_unresolved[:10])}")

    if not args.skip_xlsx:
        xlsx_files, xlsx_candidate, xlsx_updated, xlsx_unresolved = enrich_xlsx_directory(
            xlsx_dir=xlsx_dir,
            resolve_description=resolve_table_column_description,
        )
        print(
            f"[XLSX] arquivos={xlsx_files} candidatos={xlsx_candidate} "
            f"atualizados={xlsx_updated} não_resolvidos={len(xlsx_unresolved)}"
        )
        if xlsx_unresolved:
            print(f"       unresolved sample: {', '.join(xlsx_unresolved[:10])}")
    else:
        print("[SKIP] Enriquecimento XLSX desabilitado (--skip-xlsx).")

    out_path.write_text(html_text, encoding="utf-8")

    remaining_rows = collect_remaining_placeholders(html_text)
    write_remaining_log(remaining_rows, remaining_log_path)
    with diagnostic_log_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "table",
                "real_table",
                "url",
                "url_type",
                "status",
                "reason",
                "pre_candidates",
                "parsed_columns",
                "updated",
                "unresolved_after",
                "unresolved_columns_sample",
            ],
        )
        w.writeheader()
        w.writerows(diagnostics)
    by_table = Counter(r["associated_table"].upper() for r in remaining_rows if r["associated_table"])

    print("\nConcluído")
    print(f"Saída: {out_path}")
    print(f"Total de candidatos: {total_candidate}")
    print(f"Total atualizado: {total_updated}")
    print(f"Restantes '{PLACEHOLDER}': {len(remaining_rows)}")
    print(f"CSV de pendências: {remaining_log_path}")
    print(f"CSV de diagnóstico: {diagnostic_log_path}")
    if by_table:
        print("Principais tabelas restantes:")
        for table, cnt in by_table.most_common(10):
            print(f"  - {table}: {cnt}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Fatal: {exc}", file=sys.stderr)
        sys.exit(1)

