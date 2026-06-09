#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import html
import json
from functools import lru_cache
import re
import unicodedata
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, quote_plus, unquote, urlparse
from urllib.request import Request, urlopen
import xml.etree.ElementTree as ET

NOISE_REGEXES: list[tuple[str, str]] = [
    ("TEST", r"(^|[_\-\s])TEST([_\-\s]|$)"),
    ("ORIGINAL", r"(^|[_\-\s])ORIGINAL([_\-\s]|$)"),
    ("BKP", r"(^|[_\-\s])BKP(?:\d+)?([_\-\s]|$)"),
    ("DEBUG", r"(^|[_\-\s])DEBUG([_\-\s]|$)"),
    ("COPY", r"(^|[_\-\s])COPY(?:\d+)?([_\-\s]|$)"),
    ("RC", r"(^|[_\-\s])RC(?:V?\d+)?([_\-\s]|$)"),
    ("PLUS", r"(^|[_\-\s])PLUS([_\-\s]|$)"),
    ("JAN20XXXX", r"(^|[_\-\s])JAN20[0-9]{2}([_\-\s]|$)"),
    ("AGUARDANDO", r"(^|[_\-\s])AGUARDANDO([_\-\s]|$)"),
    ("CLIENTE", r"(^|[_\-\s])CLIENTE([_\-\s]|$)"),
    ("ANALISE", r"(^|[_\-\s])ANALISE([_\-\s]|$)"),
    ("ANALYSE", r"(^|[_\-\s])ANALYSE([_\-\s]|$)"),
    ("PERFORMANCE", r"(^|[_\-\s])PERFORMANCE([_\-\s]|$)"),
    ("HYDRUS", r"(^|[_\-\s])HYDRUS([_\-\s]|$)"),
]
SQL_KEYWORDS = {
    "SELECT",
    "FROM",
    "JOIN",
    "WHERE",
    "GROUP",
    "ORDER",
    "BY",
    "AS",
    "ON",
    "AND",
    "OR",
    "WHEN",
    "THEN",
    "ELSE",
    "END",
    "UNION",
    "ALL",
    "DISTINCT",
    "WITH",
}

TYPE_RX = re.compile(r"\b(NUMBER|VARCHAR2|NVARCHAR2|DATE|TIMESTAMP|CHAR|CLOB|BLOB)\b", flags=re.IGNORECASE)
XSD_TO_ORACLE_TYPE: dict[str, str] = {
    "string": "VARCHAR2",
    "normalizedstring": "VARCHAR2",
    "token": "VARCHAR2",
    "anyuri": "VARCHAR2",
    "double": "NUMBER",
    "decimal": "NUMBER",
    "float": "NUMBER",
    "int": "NUMBER",
    "integer": "NUMBER",
    "long": "NUMBER",
    "short": "NUMBER",
    "nonnegativeinteger": "NUMBER",
    "positiveinteger": "NUMBER",
    "date": "DATE",
    "datetime": "TIMESTAMP",
    "time": "TIMESTAMP",
    "boolean": "VARCHAR2",
}

ANCHOR_CLUSTER_MAX_GAP = 6
RELEVANT_CLUSTER_MAX_GAP = 2
IGNORED_GROUP_PREFIXES = ("SYN_",)
IGNORED_GROUP_NAMES_NORMALIZED = {
    "SALDOEMTRANSITONOMESTRANSFERENCIAINTERORGANIZACAO",
}
G_FLEX_SETUP_ASSET_PATH = Path(__file__).resolve().parent.parent / "assets" / "G_FLEX_SETUP.xlsx"

def map_xsd_to_oracle_type(data_type: str) -> str:
    dt = (data_type or "").strip().lower()
    if not dt:
        return "N/A"
    if ":" in dt:
        dt = dt.split(":", 1)[1]
    return XSD_TO_ORACLE_TYPE.get(dt, "N/A")


def infer_oracle_type_from_sample(value: str) -> str:
    v = (value or "").strip()
    if not v:
        return "N/A"
    if re.match(r"^\d{4}-\d{2}-\d{2}[ tT]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$", v):
        return "TIMESTAMP"
    if re.match(r"^\d{4}-\d{2}-\d{2}$", v) or re.match(r"^\d{2}/\d{2}/\d{4}$", v):
        return "DATE"
    if re.match(r"^[+-]?\d+(?:[\.,]\d+)?$", v):
        return "NUMBER"
    return "VARCHAR2"


def sample_group_keys(group_name: str) -> list[str]:
    g = (group_name or "").upper()
    out: list[str] = []
    if g:
        out.append(g)
        if g.startswith("G_") or g.startswith("Q_"):
            out.append(g[2:])
    return list(dict.fromkeys(out))


def lookup_sample_value(
    sample_values: dict[str, dict[str, str]],
    group_name: str,
    source_name: str,
    field_name: str,
) -> str:
    fname = (field_name or "").upper()
    for key in sample_group_keys(group_name) + sample_group_keys(source_name):
        grp = sample_values.get(key, {})
        if fname in grp:
            return grp[fname]
    return ""


@dataclass
class Candidate:
    entry: str
    file_name: str
    group_key: str
    version: int
    score: float
    official: bool
    reasons: list[str]


@dataclass
class DatasetInfo:
    name: str
    source_dataset: str
    fields: list[str]
    sql_tables: list[str]
    sql_aliases: dict[str, str]
    sql_text: str
    field_details: dict[str, dict[str, str]]


@dataclass
class ExtractorAudit:
    extractor: str
    selected_entry: str
    selected_file: str
    status: str  # PASS / PASS_WITH_WARNINGS / FAIL / NOT_IN_DOC
    doc_search_names: list[str]
    doc_pages: list[int]
    field_count: int
    missing_fields_in_doc: list[str]
    datasets: list[DatasetInfo]
    rejected_candidates: list[dict]
    source_review_count: int
    enrich_warning_count: int
    warnings: list[str]
    scope_debug: dict[str, Any]


@dataclass
class OracleEnriched:
    description: str
    characteristic: str
    doc_url: str
    matched: bool
    provider: str


@dataclass
class CompareOptions:
    scope_mode: str = "anchor-window"  # anchor-window|name-match
    write_scope_debug_json: bool = False


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def normalize_token(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (value or "").upper())


def normalize_match_text(value: str) -> str:
    raw = (value or "").strip()
    no_accents = "".join(
        ch for ch in unicodedata.normalize("NFKD", raw)
        if not unicodedata.combining(ch)
    )
    return re.sub(r"[^A-Z0-9]", "", no_accents.upper())


def should_ignore_group(group_key: str) -> bool:
    g = (group_key or "").strip()
    if not g:
        return False
    g_upper = g.upper()
    if any(g_upper.startswith(prefix) for prefix in IGNORED_GROUP_PREFIXES):
        return True
    return normalize_match_text(g) in IGNORED_GROUP_NAMES_NORMALIZED


def normalize_ws(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def contains_field(text_raw: str, text_norm: str, field_name: str) -> bool:
    if re.search(rf"\b{re.escape(field_name)}\b", text_raw, flags=re.IGNORECASE):
        return True
    f_norm = normalize_token(field_name)
    return bool(f_norm) and f_norm in text_norm


def excel_col_to_index(ref: str) -> int:
    col = "".join(ch for ch in (ref or "").upper() if "A" <= ch <= "Z")
    if not col:
        return 0
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


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


def load_g_flex_setup_reference(xlsx_path: Optional[Path] = None) -> dict[str, dict[str, str]]:
    path = xlsx_path or G_FLEX_SETUP_ASSET_PATH
    if not path.exists():
        return {}
    rows = [normalize_flex_setup_row(r) for r in read_xlsx_rows(path)]
    refs: dict[str, dict[str, str]] = {}
    for row in rows:
        if not row or not row[0]:
            continue
        tag = row[0].strip()
        if normalize_token(tag) == "XMLTAGNAME":
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


@lru_cache(maxsize=1)
def load_g_flex_setup_reference_cached() -> dict[str, dict[str, str]]:
    return load_g_flex_setup_reference(G_FLEX_SETUP_ASSET_PATH)


def is_g_flex_setup_dataset(dataset_name: str) -> bool:
    return normalize_token(dataset_name) == "GFLEXSETUP"


def cluster_pages(page_numbers: list[int], max_gap: int) -> list[list[int]]:
    if not page_numbers:
        return []
    ordered = sorted(set(int(p) for p in page_numbers if int(p) > 0))
    if not ordered:
        return []
    clusters: list[list[int]] = [[ordered[0]]]
    for p in ordered[1:]:
        if p - clusters[-1][-1] <= max_gap:
            clusters[-1].append(p)
        else:
            clusters.append([p])
    return clusters


def pick_primary_cluster(clusters: list[list[int]]) -> list[int]:
    if not clusters:
        return []
    # Prefer the densest cluster; on tie keep the earliest one.
    return sorted(clusters, key=lambda c: (-len(c), c[0]))[0]


def extract_version(name: str) -> int:
    parts = re.findall(r"(?:_|-)v(\d+)\b|(?:^|[^a-z])v(\d+)\b", name, flags=re.IGNORECASE)
    nums = [int(a or b) for a, b in parts if (a or b)]
    return max(nums) if nums else 0


def clean_base(file_name: str) -> str:
    x = re.sub(r"\.xdmz$", "", file_name, flags=re.IGNORECASE)
    # Remove version tokens from the reference key (v1, v8, version12, RCv4, etc.).
    x = re.sub(r"[_-]?version\d+\b", "", x, flags=re.IGNORECASE)
    x = re.sub(r"[_-]?v\d+[a-z0-9]*\b", "", x, flags=re.IGNORECASE)
    x = re.sub(r"(?<=[A-Za-z])v\d+\b", "", x, flags=re.IGNORECASE)
    x = re.sub(r"__+", "_", x)
    return x.strip(" _-")


def classify(file_name: str, modified: datetime) -> Candidate:
    cleaned = clean_base(file_name)
    upper = cleaned.upper()
    reasons = [label for label, rx in NOISE_REGEXES if re.search(rx, upper)]
    official = len(reasons) == 0
    is_incremental = bool(re.search(r"(^|[_-])INCREMENTAL($|[_-])", upper))
    bucket = re.sub(r"(^|[_-])INCREMENTAL($|[_-])", "_", cleaned, flags=re.IGNORECASE).strip(" _-")
    group = f"{bucket}_Incremental" if is_incremental else bucket
    version = extract_version(file_name)
    # Prefer official files and, among them, prioritize baseline names (without explicit version suffix).
    base_bonus = 1_000_000 if version == 0 else 0
    score = (10_000_000 if official else 0) + base_bonus + modified.timestamp() / 1000.0 - version
    return Candidate(entry="", file_name=file_name, group_key=group, version=version, score=score, official=official, reasons=reasons)


def list_candidates(xdrz_path: Path) -> list[Candidate]:
    out: list[Candidate] = []
    with zipfile.ZipFile(xdrz_path, "r") as zf:
        for i in zf.infolist():
            if not i.filename.lower().endswith(".xdmz"):
                continue
            c = classify(Path(i.filename).name, datetime(*i.date_time))
            c.entry = i.filename
            out.append(c)
    return out


def pick_candidates(candidates: list[Candidate]) -> tuple[list[Candidate], dict[str, list[Candidate]]]:
    groups: dict[str, list[Candidate]] = {}
    for c in candidates:
        if should_ignore_group(c.group_key):
            continue
        groups.setdefault(c.group_key, []).append(c)
    selected: list[Candidate] = []
    rejected: dict[str, list[Candidate]] = {}
    for g, rows in groups.items():
        official_rows = [r for r in rows if r.official]
        # Ignore noise-only groups (TEST/BKP/ORIGINAL/DEBUG) from selection.
        if not official_rows:
            continue
        # Prefer canonical base file (without explicit version suffix) when present.
        base_rows = [
            r for r in official_rows
            if extract_version(r.file_name) == 0 and clean_base(r.file_name).upper() == g.upper()
        ]
        pool = base_rows if base_rows else official_rows
        s = sorted(pool, key=lambda x: x.score, reverse=True)
        selected.append(s[0])
        rejected[g] = [r for r in rows if r.entry != s[0].entry]
    selected.sort(key=lambda x: x.group_key.upper())
    return selected, rejected


def doc_names(group_key: str) -> list[str]:
    base = clean_base(group_key or "")
    variants = [base]
    dm_suffix = re.sub(r"Dm(?=(_|$))", "", base, flags=re.IGNORECASE)
    dm_any = re.sub(r"Dm", "", base, flags=re.IGNORECASE)
    no_extract = re.sub(r"(^|[_-])Extract(?=[_-]|$)", "_", base, flags=re.IGNORECASE)
    no_extract_dm = re.sub(r"Dm(?=(_|$))", "", no_extract, flags=re.IGNORECASE)
    no_extract_dm_any = re.sub(r"Dm", "", no_extract, flags=re.IGNORECASE)
    variants.extend([dm_suffix, dm_any, no_extract, no_extract_dm, no_extract_dm_any])
    out: list[str] = []
    for v in variants:
        vv = re.sub(r"__+", "_", v).strip(" _-")
        if vv and vv not in out:
            out.append(vv)
    return out


def extract_docx_paragraph_text(paragraph: ET.Element) -> str:
    chunks: list[str] = []
    for node in paragraph.iter():
        tag = local_name(node.tag)
        if tag == "t":
            chunks.append(node.text or "")
        elif tag == "tab":
            chunks.append(" ")
        elif tag in {"br", "cr"}:
            chunks.append("\n")
    return normalize_ws("".join(chunks))


def extract_docx_table_text(table: ET.Element) -> str:
    rows: list[str] = []
    for tr in table.iter():
        if local_name(tr.tag) != "tr":
            continue
        cells: list[str] = []
        for tc in tr:
            if local_name(tc.tag) != "tc":
                continue
            cell_parts: list[str] = []
            for node in tc.iter():
                if local_name(node.tag) != "p":
                    continue
                txt = extract_docx_paragraph_text(node)
                if txt:
                    cell_parts.append(txt)
            cell_text = normalize_ws(" ".join(cell_parts))
            if cell_text:
                cells.append(cell_text)
        if cells:
            rows.append(" | ".join(cells))
    return "\n".join(rows)


def extract_docx_blocks(doc_path: Path) -> tuple[list[str], list[str], dict[str, Any]]:
    pages: list[str] = []
    table_pages: list[str] = []
    paragraph_blocks = 0
    table_blocks = 0

    with zipfile.ZipFile(doc_path, "r") as zf:
        try:
            document_xml = zf.read("word/document.xml")
        except KeyError as exc:
            raise ValueError(f"DOCX invalido: word/document.xml nao encontrado ({doc_path})") from exc

    root = ET.fromstring(document_xml)
    body = next((child for child in root if local_name(child.tag) == "body"), None)
    if body is None:
        return [], [], {"block_count": 0, "paragraph_blocks": 0, "table_blocks": 0, "warning": "word/document.xml missing body"}

    for block in body:
        tag = local_name(block.tag)
        if tag == "p":
            text = extract_docx_paragraph_text(block)
            if text:
                pages.append(text)
                table_pages.append("")
                paragraph_blocks += 1
        elif tag == "tbl":
            table_text = extract_docx_table_text(block)
            if table_text:
                pages.append(table_text)
                table_pages.append(table_text)
                table_blocks += 1

    debug: dict[str, Any] = {
        "block_count": len(pages),
        "paragraph_blocks": paragraph_blocks,
        "table_blocks": table_blocks,
    }
    if not pages:
        debug["warning"] = "DOCX has no extractable text blocks."
    return pages, table_pages, debug


def find_pages(pages: list[str], names: list[str]) -> list[int]:
    idx: list[int] = []
    names_clean = [n for n in names if n]
    # 1) strict pass first, to avoid pulling unrelated extractor pages.
    for i, t in enumerate(pages, start=1):
        if any(re.search(rf"\b{re.escape(n)}\b", t, flags=re.IGNORECASE) for n in names_clean):
            idx.append(i)
    if idx:
        return idx

    # 2) fallback normalized search when strict match finds nothing.
    names_norm = [normalize_token(n) for n in names_clean]
    for i, t in enumerate(pages, start=1):
        t_norm = normalize_token(t)
        if any(nn and nn in t_norm for nn in names_norm):
            idx.append(i)
    return idx


def collect_scope_tokens(datasets: list[DatasetInfo]) -> list[str]:
    tokens: list[str] = []
    for ds in datasets:
        # Keep anti-noise behavior anchored in G_* outputs and their Q_* sources.
        for raw in [ds.name, ds.source_dataset]:
            tok = normalize_token(raw or "")
            if tok:
                tokens.append(tok)
            if raw and (raw.upper().startswith("G_") or raw.upper().startswith("Q_")):
                short_tok = normalize_token(raw[2:])
                if short_tok:
                    tokens.append(short_tok)
        for f in ds.fields[:24]:
            ft = normalize_token(f)
            if ft and len(ft) >= 6:
                tokens.append(ft)
    return list(dict.fromkeys(tokens))


def page_relevance_hits(page_text: str, tokens: list[str]) -> list[str]:
    norm = normalize_token(page_text)
    return [t for t in tokens if t and t in norm]


def find_anchor_windows(pages: list[str], selected: list[Candidate]) -> dict[str, dict[str, Any]]:
    # Window spans from first effective anchor of current extractor up to page before next one.
    anchor_data: list[dict[str, Any]] = []
    page_anchor_freq: dict[int, int] = {}
    for s in selected:
        names = doc_names(s.group_key)
        anchors = find_pages(pages, names)
        for p in anchors:
            page_anchor_freq[p] = page_anchor_freq.get(p, 0) + 1
        anchor_data.append(
            {
                "group_key": s.group_key,
                "doc_search_names": names,
                "anchor_pages": anchors,
            }
        )

    # Ignore TOC/front-matter anchors that appear in many extractors on early pages.
    # This prevents window collapse around pages like 2/3/5.
    frontmatter_threshold = max(3, len(selected) // 6)
    suppressed_frontmatter_pages = {
        p for p, cnt in page_anchor_freq.items()
        if p <= 20 and cnt >= frontmatter_threshold
    }

    for row in anchor_data:
        anchors = list(row["anchor_pages"])
        effective = [p for p in anchors if p not in suppressed_frontmatter_pages]
        clusters = cluster_pages(effective if effective else anchors, max_gap=ANCHOR_CLUSTER_MAX_GAP)
        primary_cluster = pick_primary_cluster(clusters)
        row["suppressed_frontmatter_pages"] = sorted([p for p in anchors if p in suppressed_frontmatter_pages])
        row["effective_anchor_pages"] = effective if effective else anchors
        row["anchor_clusters"] = clusters
        row["primary_anchor_cluster"] = primary_cluster
        row["first_anchor"] = primary_cluster[0] if primary_cluster else None

    anchored = [x for x in anchor_data if x["first_anchor"] is not None]
    anchored.sort(key=lambda x: (int(x["first_anchor"]), x["group_key"]))

    windows: dict[str, dict[str, Any]] = {}
    for i, row in enumerate(anchored):
        start = int(row["first_anchor"])
        next_start = int(anchored[i + 1]["first_anchor"]) if i + 1 < len(anchored) else len(pages) + 1
        end = max(start, next_start - 1)
        windows[row["group_key"]] = {
            "anchor_pages": row["anchor_pages"],
            "effective_anchor_pages": row["effective_anchor_pages"],
            "anchor_clusters": row["anchor_clusters"],
            "primary_anchor_cluster": row["primary_anchor_cluster"],
            "suppressed_frontmatter_pages": row["suppressed_frontmatter_pages"],
            "doc_search_names": row["doc_search_names"],
            "window_start": start,
            "window_end": end,
            "candidate_pages": list(range(start, end + 1)),
        }

    for row in anchor_data:
        if row["group_key"] in windows:
            continue
        windows[row["group_key"]] = {
            "anchor_pages": row["anchor_pages"],
            "effective_anchor_pages": row["effective_anchor_pages"],
            "anchor_clusters": row["anchor_clusters"],
            "primary_anchor_cluster": row["primary_anchor_cluster"],
            "suppressed_frontmatter_pages": row["suppressed_frontmatter_pages"],
            "doc_search_names": row["doc_search_names"],
            "window_start": None,
            "window_end": None,
            "candidate_pages": [],
        }
    return windows


def select_pages_for_extractor(
    pages: list[str],
    sel: Candidate,
    datasets: list[DatasetInfo],
    scope_mode: str,
    anchor_windows: dict[str, dict[str, Any]],
) -> tuple[list[int], dict[str, Any]]:
    names = doc_names(sel.group_key)
    debug: dict[str, Any] = {
        "scope_mode": scope_mode,
        "extractor": sel.group_key,
        "doc_search_names": names,
    }

    if scope_mode == "name-match":
        pgs = find_pages(pages, names)
        debug["anchor_pages"] = pgs
        debug["candidate_pages"] = pgs
        debug["included_pages"] = pgs
        debug["page_relevance_hits"] = {}
        return pgs, debug

    info = anchor_windows.get(sel.group_key, {})
    anchor_pages: list[int] = list(info.get("anchor_pages") or [])
    effective_anchor_pages: list[int] = list(info.get("effective_anchor_pages") or anchor_pages)
    anchor_clusters: list[list[int]] = [list(c) for c in (info.get("anchor_clusters") or [])]
    primary_anchor_cluster: list[int] = list(info.get("primary_anchor_cluster") or pick_primary_cluster(anchor_clusters))
    if not primary_anchor_cluster:
        primary_anchor_cluster = list(effective_anchor_pages)
    candidate_pages: list[int] = [p for p in (info.get("candidate_pages") or []) if 1 <= p <= len(pages)]
    debug["anchor_pages"] = anchor_pages
    debug["effective_anchor_pages"] = effective_anchor_pages
    debug["anchor_clusters"] = anchor_clusters
    debug["primary_anchor_cluster"] = primary_anchor_cluster
    debug["suppressed_frontmatter_pages"] = list(info.get("suppressed_frontmatter_pages") or [])
    debug["window_start"] = info.get("window_start")
    debug["window_end"] = info.get("window_end")
    debug["candidate_pages"] = candidate_pages

    anchor_neighbors: list[int] = []
    for a in primary_anchor_cluster:
        for p in [a - 1, a, a + 1]:
            if 1 <= p <= len(pages):
                anchor_neighbors.append(p)

    bridge_pages: list[int] = []
    sorted_eff = sorted(set(primary_anchor_cluster))
    for i in range(len(sorted_eff) - 1):
        a, b = sorted_eff[i], sorted_eff[i + 1]
        if 1 < (b - a) <= 3:
            bridge_pages.extend(range(a, b + 1))

    if not candidate_pages:
        fallback_pages = sorted(set(anchor_neighbors + bridge_pages + primary_anchor_cluster))
        debug["included_pages"] = fallback_pages
        debug["page_relevance_hits"] = {}
        debug["anchor_neighbor_pages"] = sorted(set(anchor_neighbors))
        debug["bridge_pages"] = sorted(set(bridge_pages))
        return fallback_pages, debug

    tokens = collect_scope_tokens(datasets)
    page_hits: dict[str, list[str]] = {}
    relevant_pages: list[int] = []
    for p in candidate_pages:
        hits = page_relevance_hits(pages[p - 1], tokens)
        if hits:
            relevant_pages.append(p)
            page_hits[str(p)] = hits[:20]

    # Keep contiguous range around detected relevant pages, preserving continuation pages.
    if relevant_pages:
        relevant_clusters = cluster_pages(relevant_pages, max_gap=RELEVANT_CLUSTER_MAX_GAP)
        if primary_anchor_cluster:
            primary_anchor = primary_anchor_cluster[0]
            best_relevant_cluster = sorted(
                relevant_clusters,
                key=lambda c: (abs(c[0] - primary_anchor), -len(c), c[0]),
            )[0]
        else:
            best_relevant_cluster = pick_primary_cluster(relevant_clusters)
        start = max(candidate_pages[0], min(best_relevant_cluster) - 1)
        end = min(candidate_pages[-1], max(best_relevant_cluster) + 1)
        included = list(range(start, end + 1))
        debug["relevant_clusters"] = relevant_clusters
        debug["best_relevant_cluster"] = best_relevant_cluster
    else:
        # Fallback to full anchor window if relevance scoring is inconclusive.
        included = candidate_pages

    included = sorted(set(included + primary_anchor_cluster + anchor_neighbors + bridge_pages))

    debug["page_relevance_hits"] = page_hits
    debug["relevant_pages"] = relevant_pages
    debug["anchor_neighbor_pages"] = sorted(set(anchor_neighbors))
    debug["bridge_pages"] = sorted(set(bridge_pages))
    debug["included_pages"] = included
    return included, debug


def strip_sql_comments(sql: str) -> str:
    sql = re.sub(r"/\*.*?\*/", " ", sql or "", flags=re.DOTALL)
    sql = re.sub(r"--[^\n\r]*", " ", sql)
    return sql


def split_top_level(text: str, delimiter: str = ",") -> list[str]:
    items: list[str] = []
    buf: list[str] = []
    depth = 0
    quote: Optional[str] = None
    i = 0
    while i < len(text):
        ch = text[i]
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = None
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
            if depth > 0:
                depth -= 1
            buf.append(ch)
            i += 1
            continue
        if ch == delimiter and depth == 0:
            item = "".join(buf).strip()
            if item:
                items.append(item)
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    tail = "".join(buf).strip()
    if tail:
        items.append(tail)
    return items


def find_keyword_top_level(sql: str, keyword: str, start: int = 0) -> int:
    kw = keyword.upper()
    depth = 0
    quote: Optional[str] = None
    i = start
    while i < len(sql):
        ch = sql[i]
        if quote:
            if ch == quote:
                quote = None
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
            if depth > 0:
                depth -= 1
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
    cleaned = normalize_ws(strip_sql_comments(sql))
    sel = find_keyword_top_level(cleaned, "SELECT")
    if sel < 0:
        return "", ""
    frm = find_keyword_top_level(cleaned, "FROM", sel + 6)
    if frm < 0:
        return "", ""
    select_part = cleaned[sel + len("SELECT") : frm].strip()
    from_part = cleaned[frm + len("FROM") :].strip()
    return select_part, from_part


def extract_from_clause_segment(sql: str) -> str:
    cleaned = normalize_ws(strip_sql_comments(sql or ""))
    _, main_sql = extract_with_ctes(cleaned)
    if not main_sql:
        main_sql = cleaned
    _, from_part = extract_select_and_from(main_sql)
    if not from_part:
        return ""

    cut_at = len(from_part)
    for kw in ("WHERE", "GROUP", "ORDER", "UNION", "CONNECT", "MINUS", "INTERSECT"):
        idx = find_keyword_top_level(from_part, kw)
        if idx >= 0:
            cut_at = min(cut_at, idx)
    return from_part[:cut_at].strip()


def is_table_name_candidate(name: str) -> bool:
    n = normalize_ws((name or "").strip('"').upper())
    if not n:
        return False
    if n in SQL_KEYWORDS:
        return False
    if n in {"N/A", "MULTI_SOURCE", "UNKNOWN_SUBQUERY", "WHH", "CNPJ"}:
        return False
    if any(x in n for x in (" ", "(", ")")):
        return False
    return bool(re.fullmatch(r"[A-Z][A-Z0-9_$#]*(?:\.[A-Z][A-Z0-9_$#]*)?", n))


def parse_tables(sql: str) -> list[str]:
    cleaned = extract_from_clause_segment(sql)
    tables: list[str] = []
    tables.extend(
        m.group(1).strip('"')
        for m in re.finditer(r"\b(?:FROM|JOIN)\s+([A-Z0-9_\.\"$]+)\b", cleaned, flags=re.IGNORECASE)
    )
    tables.extend(
        m.group(1).strip('"')
        for m in re.finditer(r"(?:^|,)\s*([A-Z0-9_\.\"$]+)\s+(?:AS\s+)?[A-Z0-9_]+\b", cleaned, flags=re.IGNORECASE)
    )
    out = [t for t in dict.fromkeys(tables) if is_table_name_candidate(t)]
    return out


def parse_aliases(sql: str) -> dict[str, str]:
    cleaned = extract_from_clause_segment(sql)
    d: dict[str, str] = {}
    for m in re.finditer(
        r"\b(?:FROM|JOIN)\s+([A-Z0-9_\.\"$]+)\s+(?:AS\s+)?([A-Z0-9_]+)\b",
        cleaned,
        flags=re.IGNORECASE,
    ):
        table = m.group(1).strip('"')
        alias = m.group(2).upper()
        if is_table_name_candidate(table) and alias not in SQL_KEYWORDS:
            d[alias] = table
    for m in re.finditer(
        r"(?:^|,)\s*([A-Z0-9_\.\"$]+)\s+(?:AS\s+)?([A-Z0-9_]+)\b",
        cleaned,
        flags=re.IGNORECASE,
    ):
        table = m.group(1).strip('"')
        alias = m.group(2).upper()
        if is_table_name_candidate(table) and alias not in SQL_KEYWORDS:
            d[alias] = table
    return d


def parse_projection_item(item: str) -> tuple[str, str]:
    s = normalize_ws(item.strip())
    m = re.match(r"(?is)^(.*?)\s+AS\s+([A-Z0-9_]+)$", s)
    if m:
        return m.group(1).strip(), m.group(2).upper()

    m2 = re.match(r"(?is)^(.*)\s+([A-Z0-9_]+)$", s)
    if m2:
        expr = m2.group(1).strip()
        alias = m2.group(2).upper()
        if alias not in SQL_KEYWORDS:
            return expr, alias

    m3 = re.match(r"(?is)^([A-Z0-9_]+)\.([A-Z0-9_]+)$", s)
    if m3:
        return s, m3.group(2).upper()

    m4 = re.match(r"(?is)^([A-Z0-9_]+)$", s)
    if m4:
        return s, m4.group(1).upper()

    norm_alias = normalize_token(s)[:64] or "EXPR"
    return s, norm_alias


def extract_with_ctes(sql: str) -> tuple[dict[str, dict[str, dict[str, str]]], str]:
    cte_maps: dict[str, dict[str, dict[str, str]]] = {}
    cleaned = normalize_ws(strip_sql_comments(sql or ""))
    if not re.match(r"(?is)^WITH\b", cleaned):
        return cte_maps, cleaned

    i = len("WITH")
    n = len(cleaned)
    while i < n:
        while i < n and cleaned[i].isspace():
            i += 1
        m_name = re.match(r"([A-Z0-9_]+)", cleaned[i:], flags=re.IGNORECASE)
        if not m_name:
            break
        cte_name = m_name.group(1).upper()
        i += len(m_name.group(0))

        while i < n and cleaned[i].isspace():
            i += 1
        if i < n and cleaned[i] == "(":
            d = 1
            i += 1
            while i < n and d > 0:
                if cleaned[i] == "(":
                    d += 1
                elif cleaned[i] == ")":
                    d -= 1
                i += 1

        while i < n and cleaned[i].isspace():
            i += 1
        if not re.match(r"(?is)^AS\b", cleaned[i:]):
            break
        i += 2
        while i < n and cleaned[i].isspace():
            i += 1
        if i >= n or cleaned[i] != "(":
            break

        start = i
        d = 1
        i += 1
        while i < n and d > 0:
            if cleaned[i] == "(":
                d += 1
            elif cleaned[i] == ")":
                d -= 1
            i += 1
        if d != 0:
            break
        inner_sql = cleaned[start + 1 : i - 1].strip()
        cte_maps[cte_name] = parse_select_map(inner_sql, depth_limit=2, cte_maps=cte_maps)

        while i < n and cleaned[i].isspace():
            i += 1
        if i < n and cleaned[i] == ",":
            i += 1
            continue
        break

    return cte_maps, cleaned[i:].strip()


def parse_subquery_aliases(
    from_part: str,
    depth_limit: int = 2,
    cte_maps: Optional[dict[str, dict[str, dict[str, str]]]] = None,
) -> dict[str, dict[str, dict[str, str]]]:
    out: dict[str, dict[str, dict[str, str]]] = {}
    if depth_limit <= 0:
        return out
    base_ctes = cte_maps or {}
    txt = from_part or ""
    i = 0
    while i < len(txt):
        if txt[i] != "(":
            i += 1
            continue
        start = i
        d = 1
        i += 1
        while i < len(txt) and d > 0:
            if txt[i] == "(":
                d += 1
            elif txt[i] == ")":
                d -= 1
            i += 1
        if d != 0:
            break
        inner = txt[start + 1 : i - 1].strip()
        if not re.match(r"(?is)^(WITH\b|SELECT\b)", inner):
            continue
        m_alias = re.match(r"\s*([A-Z0-9_]+)", txt[i:], flags=re.IGNORECASE)
        if not m_alias:
            continue
        alias = m_alias.group(1).upper()
        out[alias] = parse_select_map(inner, depth_limit=depth_limit - 1, cte_maps=base_ctes)
        i += len(m_alias.group(0))
    return out


def resolve_column_reference(
    ref_alias: str,
    ref_col: str,
    table_aliases: dict[str, str],
    subquery_aliases: dict[str, dict[str, dict[str, str]]],
    cte_maps: Optional[dict[str, dict[str, dict[str, str]]]] = None,
) -> list[tuple[str, str]]:
    alias = ref_alias.upper()
    col = ref_col.upper()
    base_ctes = cte_maps or {}
    if alias in subquery_aliases:
        sq = subquery_aliases[alias]
        hit = sq.get(col)
        if not hit:
            return [("UNKNOWN_SUBQUERY", col)]
        return [(hit.get("source_table", "UNKNOWN_SUBQUERY"), hit.get("source_column", col))]
    if alias in table_aliases:
        table = table_aliases[alias].upper()
        cte_hit = base_ctes.get(table, {}).get(col)
        if cte_hit:
            return [(cte_hit.get("source_table", table), cte_hit.get("source_column", col))]
        return [(table, col)]
    if alias in base_ctes:
        cte_hit = base_ctes[alias].get(col)
        if cte_hit:
            return [(cte_hit.get("source_table", alias), cte_hit.get("source_column", col))]
    return [(alias, col)]


def resolve_refs_in_expression(
    expr: str,
    table_aliases: dict[str, str],
    subquery_aliases: dict[str, dict[str, dict[str, str]]],
    cte_maps: Optional[dict[str, dict[str, dict[str, str]]]] = None,
) -> list[tuple[str, str]]:
    refs = re.findall(r"\b([A-Z_][A-Z0-9_$#]*)\.([A-Z_][A-Z0-9_$#]*)\b", expr or "", flags=re.IGNORECASE)
    resolved_refs: list[tuple[str, str]] = []
    for ra, rc in refs:
        resolved_refs.extend(resolve_column_reference(ra, rc, table_aliases, subquery_aliases, cte_maps=cte_maps))
    return list(dict.fromkeys((t.upper(), c.upper()) for t, c in resolved_refs))


def pick_dominant_expression_source(
    expr: str,
    table_aliases: dict[str, str],
    subquery_aliases: dict[str, dict[str, dict[str, str]]],
    cte_maps: Optional[dict[str, dict[str, dict[str, str]]]] = None,
) -> Optional[tuple[str, str]]:
    s = normalize_ws(expr)

    m_decode = re.match(r"(?is)^DECODE\s*\((.*)\)$", s)
    if m_decode:
        args = split_top_level(m_decode.group(1), delimiter=",")
        result_exprs: list[str] = []
        if len(args) >= 3:
            for idx in range(2, len(args), 2):
                result_exprs.append(args[idx])
            if len(args) % 2 == 0:
                result_exprs.append(args[-1])

        result_refs: list[tuple[str, str]] = []
        for part in result_exprs:
            result_refs.extend(resolve_refs_in_expression(part, table_aliases, subquery_aliases, cte_maps=cte_maps))
        result_refs = list(dict.fromkeys(result_refs))
        if len(result_refs) == 1:
            return result_refs[0]

    m_func = re.match(r"(?is)^(NVL|COALESCE)\s*\((.*)\)$", s)
    if m_func:
        args = split_top_level(m_func.group(2), delimiter=",")
        for part in args:
            refs = resolve_refs_in_expression(part, table_aliases, subquery_aliases, cte_maps=cte_maps)
            if len(refs) == 1:
                return refs[0]

    return None


def parse_select_map(
    sql: str,
    depth_limit: int = 2,
    cte_maps: Optional[dict[str, dict[str, dict[str, str]]]] = None,
) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    base_ctes = dict(cte_maps or {})
    local_ctes, main_sql = extract_with_ctes(sql or "")
    base_ctes.update(local_ctes)

    select_part, from_part = extract_select_and_from(main_sql or "")
    if not select_part:
        return out
    table_aliases = parse_aliases(main_sql or "")
    subquery_aliases = parse_subquery_aliases(from_part, depth_limit=depth_limit, cte_maps=base_ctes)
    items = split_top_level(select_part, delimiter=",")

    for raw_item in items:
        expr, out_alias = parse_projection_item(raw_item)
        resolved_refs = resolve_refs_in_expression(expr, table_aliases, subquery_aliases, cte_maps=base_ctes)

        source_mode = "UNKNOWN"
        source_confidence = "LOW"
        source_table = "N/A"
        source_column = "UNKNOWN_SOURCE_COLUMN"

        direct_rx = re.match(r"(?is)^([A-Z0-9_]+)\.([A-Z0-9_]+)$", normalize_ws(expr))
        if direct_rx and resolved_refs:
            source_table, source_column = resolved_refs[0]
            source_mode = "ALIAS" if out_alias != source_column else "DIRECT"
            source_confidence = "HIGH"
        elif len(resolved_refs) == 1:
            source_table, source_column = resolved_refs[0]
            source_mode = "EXPRESSION"
            source_confidence = "MEDIUM"
        elif len(resolved_refs) > 1:
            dominant = pick_dominant_expression_source(expr, table_aliases, subquery_aliases, cte_maps=base_ctes)
            if dominant:
                source_table, source_column = dominant
                source_mode = "EXPRESSION"
                source_confidence = "MEDIUM"
            else:
                source_table = "MULTI_SOURCE"
                source_column = "MULTI_SOURCE"
                source_mode = "MULTI_SOURCE"
                source_confidence = "LOW"
        else:
            lone = re.match(r"(?is)^([A-Z0-9_]+)$", normalize_ws(expr))
            if lone:
                source_column = lone.group(1).upper()
                source_mode = "DIRECT"
                source_confidence = "MEDIUM"

        out[out_alias.upper()] = {
            "source_table": source_table,
            "source_column": source_column,
            "source_mode": source_mode,
            "source_confidence": source_confidence,
            "source_sql_snippet": normalize_ws(raw_item)[:220],
        }
    return out

class OracleDocsEnricher:
    def __init__(self, enabled: bool, cache_path: Path, timeout_sec: int) -> None:
        self.enabled = enabled
        self.cache_path = cache_path
        self.timeout_sec = timeout_sec
        self.cache: dict[str, dict] = {}
        if self.enabled:
            self._load_cache()

    def _load_cache(self) -> None:
        if self.cache_path.exists():
            try:
                self.cache = json.loads(self.cache_path.read_text(encoding="utf-8"))
            except Exception:
                self.cache = {}

    def flush_cache(self) -> None:
        if not self.enabled:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(json.dumps(self.cache, indent=2, ensure_ascii=False), encoding="utf-8")

    def _default(self) -> OracleEnriched:
        return OracleEnriched(
            description="A preencher pela documentação funcional 1",
            characteristic="N/A",
            doc_url="",
            matched=False,
            provider="none",
        )

    def enrich(self, table_name: str, column_name: str) -> OracleEnriched:
        if not self.enabled:
            return self._default()
        key = f"{(table_name or '').upper()}|{(column_name or '').upper()}"
        if key in self.cache:
            try:
                return OracleEnriched(**self.cache[key])
            except Exception:
                pass

        result = self._try_http(table_name, column_name)
        if result is None:
            result = self._try_botcity(table_name, column_name)
        if result is None:
            result = self._default()
            result.provider = "not_found"

        self.cache[key] = asdict(result)
        return result

    def _request_text(self, url: str) -> str:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 ReferenceEquityDocSync/1.0"})
        with urlopen(req, timeout=self.timeout_sec) as resp:
            return resp.read().decode("utf-8", errors="replace")

    def _is_oracle_url(self, url: str) -> bool:
        try:
            host = (urlparse(url).netloc or "").lower()
            return host.endswith("docs.oracle.com")
        except Exception:
            return False

    def _find_oracle_doc_link(self, html_text: str, table_name: str) -> str:
        links = self._extract_oracle_doc_links(html_text)
        if not links:
            return ""
        table_lower = (table_name or "").lower()
        for link in links:
            if table_lower and table_lower in link.lower():
                return link
        return links[0]

    def _extract_oracle_doc_links(self, html_text: str) -> list[str]:
        raw_links: list[str] = []
        raw_links.extend(re.findall(r'href="(https?://docs\.oracle\.com[^"]+)"', html_text or "", flags=re.IGNORECASE))
        raw_links.extend(re.findall(r'href="(/en/cloud/[^"]+)"', html_text or "", flags=re.IGNORECASE))
        raw_links.extend(re.findall(r'"(https?://docs\.oracle\.com[^"]+/oedmf/[^"]+)"', html_text or "", flags=re.IGNORECASE))

        out: list[str] = []
        for link in raw_links:
            url = link.strip()
            if url.startswith("/"):
                url = f"https://docs.oracle.com{url}"
            url = url.replace("&amp;", "&")
            low = url.lower()
            if any(x in low for x in ["sign-into-cloud", "/privacy/", "/legal/", "lookup?ctx=", "/cpyr"]):
                continue
            if self._is_oracle_url(url) and "/oedmf/" in url.lower():
                out.append(url)
            elif self._is_oracle_url(url):
                out.append(url)
        return list(dict.fromkeys(out))

    def _pick_best_doc_link(self, links: list[str], table_name: str) -> str:
        if not links:
            return ""
        table_token = normalize_token(table_name or "")
        table_lower = (table_name or "").lower()

        scored = []
        for link in links:
            lu = link.lower()
            score = 0
            if "/oedmf/" in lu:
                score += 5
            if table_lower and table_lower in lu:
                score += 4
            if table_token and normalize_token(lu).find(table_token) >= 0:
                score += 2
            scored.append((score, link))
        scored.sort(key=lambda x: x[0], reverse=True)
        if not scored or scored[0][0] < 4:
            return ""
        return scored[0][1]

    def _extract_oracle_doc_link_from_google(self, html_text: str, table_name: str) -> str:
        links = re.findall(r'href="(/url\?q=https?://[^"]+)"', html_text, flags=re.IGNORECASE)
        if not links:
            links = re.findall(r'href="(https?://docs\.oracle\.com[^"]+)"', html_text, flags=re.IGNORECASE)
        candidates: list[str] = []
        for raw in links:
            link = raw
            if link.startswith("/url?"):
                q = parse_qs(urlparse(link).query).get("q", [""])[0]
                link = unquote(q)
            elif "google.com/url?" in link.lower():
                parsed = parse_qs(urlparse(link).query)
                q = (parsed.get("url") or parsed.get("q") or [""])[0]
                link = unquote(q) if q else link
            if self._is_oracle_url(link):
                candidates.append(link)
        if not candidates:
            return ""
        return self._pick_best_doc_link(candidates, table_name)

    def _extract_oracle_links_from_driver(self, driver: object, table_name: str) -> list[str]:
        try:
            from selenium.webdriver.common.by import By  # type: ignore
        except Exception:
            return []
        candidates: list[str] = []
        try:
            anchors = driver.find_elements(By.TAG_NAME, "a")
        except Exception:
            return []
        for a in anchors:
            try:
                href = a.get_attribute("href") or ""
            except Exception:
                href = ""
            if not href:
                continue
            link = href
            if "google.com/url?" in href.lower():
                parsed = parse_qs(urlparse(href).query)
                q = (parsed.get("url") or parsed.get("q") or [""])[0]
                if q:
                    link = unquote(q)
            if self._is_oracle_url(link):
                candidates.append(link)
        picked = self._pick_best_doc_link(list(dict.fromkeys(candidates)), table_name)
        return [picked] if picked else []

    def _clean_html_text(self, value: str) -> str:
        if not value:
            return ""
        x = re.sub(r"(?is)<[^>]+>", " ", value)
        x = html.unescape(x)
        return normalize_ws(x)

    def _as_html_text(self, page_obj: object) -> str:
        if page_obj is None:
            return ""
        if isinstance(page_obj, str):
            return page_obj
        try:
            return str(page_obj)
        except Exception:
            return ""

    def _extract_doc_metadata_from_html(self, page_html: str, column_name: str) -> tuple[str, str]:
        col_token = normalize_token(column_name or "")
        if not page_html or not col_token:
            return "", ""

        rows = re.findall(r"(?is)<tr\b[^>]*>(.*?)</tr>", page_html)
        parsed_rows: list[list[str]] = []
        for row in rows:
            cells = re.findall(r"(?is)<t[hd]\b[^>]*>(.*?)</t[hd]>", row)
            clean_cells = [self._clean_html_text(c) for c in cells]
            if cells:
                parsed_rows.append(clean_cells)

        header_idx = -1
        col_idx = -1
        type_idx = -1
        comments_idx = -1

        for i, row in enumerate(parsed_rows):
            norms = [normalize_token(c) for c in row]
            for j, n in enumerate(norms):
                if n in {"COLUMNNAME", "COLUMN", "NAME"}:
                    col_idx = j
                elif n in {"DATATYPE", "TYPE"}:
                    type_idx = j
                elif n in {"COMMENTS", "COMMENT", "DESCRIPTION"}:
                    comments_idx = j
            if col_idx >= 0 and (type_idx >= 0 or comments_idx >= 0):
                header_idx = i
                break

        if header_idx >= 0 and col_idx >= 0:
            for row in parsed_rows[header_idx + 1 :]:
                if col_idx >= len(row):
                    continue
                if normalize_token(row[col_idx]) != col_token:
                    continue
                dtype = row[type_idx] if 0 <= type_idx < len(row) else "N/A"
                desc = row[comments_idx] if 0 <= comments_idx < len(row) else ""
                return desc, (dtype or "N/A")

        # Fallback for Oracle OEDMF pages where headers may be absent or non-standard.
        for row in parsed_rows:
            norm_cells = [normalize_token(c) for c in row]
            if col_token not in norm_cells:
                continue
            raw_cells = list(row)
            dtype = "N/A"
            desc = ""

            for c in raw_cells:
                m = TYPE_RX.search(c or "")
                if m:
                    dtype = normalize_ws(c)
                    break
            for c in raw_cells:
                nc = normalize_token(c)
                if nc in {col_token, normalize_token(dtype)}:
                    continue
                if len(c or "") >= 8:
                    desc = normalize_ws(c)
                    break
            return desc, dtype

        return "", ""

    def _extract_doc_metadata(self, page_text: str, column_name: str) -> tuple[str, str]:
        desc_html, char_html = self._extract_doc_metadata_from_html(page_text, column_name)
        if desc_html or (char_html and char_html != "N/A"):
            return desc_html, char_html
        up = page_text.upper()
        col = (column_name or "").upper()
        idx = up.find(col)
        if idx < 0:
            return "", ""
        start = max(0, idx - 220)
        end = min(len(page_text), idx + 320)
        snippet = normalize_ws(page_text[start:end])
        type_m = TYPE_RX.search(snippet)
        characteristic = type_m.group(1).upper() if type_m else "N/A"
        description = snippet[:200] if snippet else ""
        return description, characteristic

    def _try_http(self, table_name: str, column_name: str) -> Optional[OracleEnriched]:
        try:
            queries = [
                f"{table_name or ''} {column_name or ''}",
                f"{table_name or ''}",
                f"{table_name or ''} oedmf",
            ]
            candidates: list[str] = []
            for q in queries:
                search_url = f"https://docs.oracle.com/search/?q={quote_plus(q)}"
                search_html = self._request_text(search_url)
                candidates.extend(self._extract_oracle_doc_links(search_html))
            candidates = list(dict.fromkeys(candidates))
            if not candidates:
                return None

            ordered: list[str] = []
            best = self._pick_best_doc_link(candidates, table_name)
            if best:
                ordered.append(best)
            ordered.extend([c for c in candidates if c != best])

            for doc_url in ordered[:6]:
                if not doc_url or not self._is_oracle_url(doc_url):
                    continue
                page = self._request_text(doc_url)
                desc, charac = self._extract_doc_metadata(page, column_name)
                if desc or (charac and charac != "N/A"):
                    return OracleEnriched(
                        description=desc or "A preencher pela documentação funcional 2",
                        characteristic=charac or "N/A",
                        doc_url=doc_url,
                        matched=bool(desc),
                        provider="http",
                    )
            return None
        except Exception:
            return None

    def _try_botcity(self, table_name: str, column_name: str) -> Optional[OracleEnriched]:
        try:
            from botcity.web import Browser, WebBot  # type: ignore
            from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
        except Exception:
            return None
        try:
            query = f"{table_name or ''} {column_name or ''} oracle docs"
            google_url = f"https://www.google.com/search?q={quote_plus(query)}"
            docs_search_url = f"https://docs.oracle.com/search/?q={quote_plus((table_name or '') + ' ' + (column_name or ''))}"
            docs_table_url = f"https://docs.oracle.com/search/?q={quote_plus(table_name or '')}"
            bot = WebBot()
            bot.headless = True
            bot.browser = Browser.CHROME
            try:
                bot.driver_path = ChromeDriverManager().install()
            except Exception:
                # Fallback to PATH if webdriver-manager fails.
                pass

            bot.browse(google_url)
            google_html = self._as_html_text(bot.page_source())
            driver_links = self._extract_oracle_links_from_driver(bot.driver, table_name)
            doc_url = driver_links[0] if driver_links else self._extract_oracle_doc_link_from_google(google_html, table_name)
            if not doc_url:
                bot.browse(docs_search_url)
                docs_search_html = self._as_html_text(bot.page_source())
                doc_links = self._extract_oracle_doc_links(docs_search_html)
                driver_links = self._extract_oracle_links_from_driver(bot.driver, table_name)
                if driver_links:
                    doc_links.extend(driver_links)
                doc_url = self._pick_best_doc_link(doc_links, table_name)
            if not doc_url:
                bot.browse(docs_table_url)
                docs_table_html = self._as_html_text(bot.page_source())
                doc_links = self._extract_oracle_doc_links(docs_table_html)
                driver_links = self._extract_oracle_links_from_driver(bot.driver, table_name)
                if driver_links:
                    doc_links.extend(driver_links)
                doc_url = self._pick_best_doc_link(doc_links, table_name)
            if not doc_url or not self._is_oracle_url(doc_url):
                bot.stop_browser()
                return None

            bot.browse(doc_url)
            detail_page = self._as_html_text(bot.page_source())
            bot.stop_browser()

            desc, charac = self._extract_doc_metadata(detail_page, column_name)
            return OracleEnriched(
                description=desc or "A preencher pela documentação funcional 3",
                characteristic=charac or "N/A",
                doc_url=doc_url,
                matched=bool(desc),
                provider="botcity",
            )
        except Exception:
            return None


def normalize_source_reference(ds: DatasetInfo, meta: dict[str, str]) -> None:
    source_table = normalize_ws(str(meta.get("source_table", "N/A") or "N/A")).upper()
    source_column = normalize_ws(str(meta.get("source_column", "") or "")).upper()
    source_mode = normalize_ws(str(meta.get("source_mode", "UNKNOWN") or "UNKNOWN")).upper()
    snippet = normalize_ws(str(meta.get("source_sql_snippet", "") or ""))

    if source_mode == "MULTI_SOURCE":
        meta["source_table"] = "MULTI_SOURCE"
        meta["source_column"] = "MULTI_SOURCE"
        meta["source_validation_reason"] = "multi_source_expression"
        return

    alias_map = {k.upper(): v.upper() for k, v in ds.sql_aliases.items() if is_table_name_candidate(v)}
    known_tables = {t.upper() for t in ds.sql_tables if is_table_name_candidate(t)}
    known_tables.update(alias_map.values())
    reason = "ok"

    if source_table in alias_map:
        source_table = alias_map[source_table]
        reason = "alias_to_table"

    if not source_column:
        source_column = "UNKNOWN_SOURCE_COLUMN"

    if source_table in {"UNKNOWN_SUBQUERY", "WHH", "CNPJ"} or source_table not in known_tables:
        m = re.search(r"\b([A-Z_][A-Z0-9_$#]*)\.([A-Z_][A-Z0-9_$#]*)\b", snippet, flags=re.IGNORECASE)
        if m:
            ref_alias = m.group(1).upper()
            ref_col = m.group(2).upper()
            if ref_alias in alias_map:
                source_table = alias_map[ref_alias]
                source_column = ref_col
                reason = "resolved_from_sql_snippet_alias"
            elif ref_alias in known_tables:
                source_table = ref_alias
                source_column = ref_col
                reason = "resolved_from_sql_snippet_table"

    if source_table in {"UNKNOWN_SUBQUERY", "WHH", "CNPJ"} or (source_table not in known_tables and source_table not in {"N/A", "MULTI_SOURCE"}):
        source_table = "N/A"
        if source_mode in {"DIRECT", "ALIAS", "EXPRESSION", "UNKNOWN"}:
            source_mode = "UNRESOLVED"
        meta["source_confidence"] = "LOW"
        reason = "unresolved_or_derived_alias"

    meta["source_table"] = source_table or "N/A"
    meta["source_column"] = source_column or "UNKNOWN_SOURCE_COLUMN"
    meta["source_mode"] = source_mode or "UNKNOWN"
    meta["source_validation_reason"] = reason


def parse_datamodel(xdmz_bytes: bytes, enricher: OracleDocsEnricher) -> list[DatasetInfo]:
    with zipfile.ZipFile(BytesIO(xdmz_bytes), "r") as zf:
        xdm = [n for n in zf.namelist() if n.lower().endswith("_datamodel.xdm") or n.lower().endswith(".xdm")]
        if not xdm:
            return []
        root = ET.fromstring(zf.read(xdm[0]))

        sample_values: dict[str, dict[str, str]] = {}
        sample_names = [n for n in zf.namelist() if n.lower().endswith("sample.xml")]
        if sample_names:
            try:
                sample_root = ET.fromstring(zf.read(sample_names[0]))
                for node in sample_root.iter():
                    children = list(node)
                    if not children:
                        continue
                    gname = local_name(node.tag).upper()
                    for ch in children:
                        if list(ch):
                            continue
                        fname = local_name(ch.tag).upper()
                        sval = (ch.text or "").strip()
                        if not fname or not sval:
                            continue
                        sample_values.setdefault(gname, {})
                        sample_values[gname].setdefault(fname, sval)
            except Exception:
                sample_values = {}

    datasets: dict[str, DatasetInfo] = {}
    dataset_sql: dict[str, str] = {}
    dataset_tables: dict[str, list[str]] = {}
    dataset_aliases: dict[str, dict[str, str]] = {}

    for n in root.iter():
        if local_name(n.tag).lower() not in {"dataset", "sqldataset", "xmldataset"}:
            continue
        name = (n.attrib.get("name") or n.attrib.get("NAME") or "").strip()
        if not name:
            continue
        sql = ""
        for ch in n.iter():
            if local_name(ch.tag).lower() in {"sql", "query"} and (ch.text or "").strip():
                sql = (ch.text or "").strip()
        key = name.upper()
        dataset_sql[key] = sql
        dataset_tables[key] = parse_tables(sql)
        dataset_aliases[key] = parse_aliases(sql)

    group_nodes = [g for g in root.iter() if local_name(g.tag).lower() == "group"]

    def load_groups(strict_q_only: bool) -> None:
        for g in group_nodes:
            src = (g.attrib.get("source") or "").strip().upper()
            group_name = (g.attrib.get("name") or g.attrib.get("NAME") or "").strip()
            if not src or src not in dataset_sql:
                continue
            if strict_q_only and not src.startswith("Q_"):
                continue

            view_name = group_name if group_name else src
            ds_key = view_name.upper()
            src_key = src

            if ds_key not in datasets:
                datasets[ds_key] = DatasetInfo(
                    name=view_name,
                    source_dataset=src_key,
                    fields=[],
                    sql_tables=dataset_tables.get(src_key, []),
                    sql_aliases=dataset_aliases.get(src_key, {}),
                    sql_text=dataset_sql.get(src_key, ""),
                    field_details={},
                )

            for ch in list(g):
                if local_name(ch.tag).lower() not in {"field", "column", "element"}:
                    continue
                fname = ch.attrib.get("name") or ch.attrib.get("columnName") or ch.attrib.get("fieldName")
                if not fname:
                    continue
                if fname not in datasets[ds_key].fields:
                    datasets[ds_key].fields.append(fname)

                meta = datasets[ds_key].field_details.setdefault(fname.upper(), {})
                raw_dt = (ch.attrib.get("dataType") or ch.attrib.get("datatype") or "").strip()
                if raw_dt:
                    meta["data_type_raw"] = raw_dt
                    mapped = map_xsd_to_oracle_type(raw_dt)
                    if mapped != "N/A":
                        meta["data_type"] = mapped

                sample_val = lookup_sample_value(sample_values, view_name, src_key, fname)
                if sample_val and "sample_value" not in meta:
                    meta["sample_value"] = sample_val
                if (not meta.get("data_type") or meta.get("data_type") == "N/A") and sample_val:
                    inferred = infer_oracle_type_from_sample(sample_val)
                    if inferred != "N/A":
                        meta["data_type"] = inferred

    # Preferred path: business output groups sourced by Q_* datasets.
    load_groups(strict_q_only=True)
    # Fallback for exceptional extractors that do not use Q_* as source.
    if not datasets:
        load_groups(strict_q_only=False)
    for key, ds in datasets.items():
        smap = parse_select_map(ds.sql_text or "") if ds.sql_text else {}
        for f in ds.fields:
            meta = ds.field_details.setdefault(f.upper(), {})
            if f.upper() in smap:
                meta.update(smap[f.upper()])
            meta.setdefault("source_table", ds.sql_tables[0] if ds.sql_tables else "N/A")
            meta.setdefault("source_column", f.upper())
            meta.setdefault("source_mode", "UNKNOWN")
            meta.setdefault("source_confidence", "LOW")
            meta.setdefault("source_sql_snippet", "")
            meta.setdefault("data_type", "N/A")
            normalize_source_reference(ds, meta)
            if meta.get("source_mode") in {"UNKNOWN", "UNRESOLVED"} and ds.sql_text:
                xml_tag = f.upper()
                if xml_tag.endswith("_LINK"):
                    base_col = re.sub(r"_LINK$", "", xml_tag)
                    if re.search(rf"\b{re.escape(base_col)}\b", strip_sql_comments(ds.sql_text), flags=re.IGNORECASE):
                        meta["source_column"] = base_col
                        meta["source_mode"] = "ALIAS"
                        meta["source_confidence"] = "MEDIUM"
                        if not meta.get("source_sql_snippet"):
                            meta["source_sql_snippet"] = f"{base_col} ... {xml_tag}"

            meta.setdefault("oracle_doc_url", "")
            meta.setdefault("oracle_description", "A preencher pela documentação funcional 4")
            meta.setdefault("oracle_characteristic", meta.get("data_type", "N/A"))
            meta.setdefault("oracle_provider", "none")
            meta.setdefault("oracle_matched", "false")
            meta.setdefault("source_validation_reason", "ok")

    return list(datasets.values())


def compute_status(
    doc_pages: list[int],
    field_count: int,
    missing_count: int,
) -> str:
    if not doc_pages:
        return "NOT_IN_DOC"
    ratio = (missing_count / field_count) if field_count else 0.0
    if missing_count >= 500 or ratio >= 0.50:
        return "FAIL"
    if missing_count > 0:
        return "PASS_WITH_WARNINGS"
    return "PASS"


def find_meta_for_field(datasets: list[DatasetInfo], field_name: str) -> tuple[str, dict[str, str]]:
    for ds in datasets:
        meta = ds.field_details.get(field_name.upper())
        if meta:
            return ds.name, meta
    return "N/A", {}


def suppress_flex_setup_false_positives(
    missing_fields: list[str],
    datasets: list[DatasetInfo],
    scope_debug: dict[str, Any],
) -> list[str]:
    refs = load_g_flex_setup_reference_cached()
    if not refs:
        scope_debug["flex_setup_reference_loaded"] = False
        return missing_fields

    flex_fields_in_extractor = {
        f.upper()
        for ds in datasets
        if is_g_flex_setup_dataset(ds.name)
        for f in ds.fields
    }
    suppressed: list[str] = []
    out: list[str] = []
    for fld in missing_fields:
        f_up = fld.upper()
        if f_up in refs and f_up in flex_fields_in_extractor:
            suppressed.append(fld)
            continue
        out.append(fld)
    scope_debug["flex_setup_reference_loaded"] = True
    scope_debug["flex_setup_reference_count"] = len(refs)
    scope_debug["flex_setup_fields_in_extractor_count"] = len(flex_fields_in_extractor)
    scope_debug["flex_setup_suppressed_missing_fields"] = suppressed
    return out


def apply_flex_setup_reference_metadata(datasets: list[DatasetInfo]) -> None:
    refs = load_g_flex_setup_reference_cached()
    if not refs:
        return
    for ds in datasets:
        if not is_g_flex_setup_dataset(ds.name):
            continue
        for fld in ds.fields:
            ref = refs.get(fld.upper())
            if not ref:
                continue
            meta = ds.field_details.setdefault(fld.upper(), {})
            if ref.get("column_description"):
                meta["oracle_description"] = ref["column_description"]
            if ref.get("column_characteristic"):
                meta["oracle_characteristic"] = ref["column_characteristic"]
            if ref.get("associated_table_to_each_column") and meta.get("source_table") in {"", "N/A"}:
                meta["source_table"] = ref["associated_table_to_each_column"]
            if ref.get("column_name") and meta.get("source_column") in {"", "N/A"}:
                meta["source_column"] = ref["column_name"]


def audit_one(
    sel: Candidate,
    rej: list[Candidate],
    pages: list[str],
    table_pages: list[str],
    xdrz: Path,
    enricher: OracleDocsEnricher,
    options: CompareOptions,
    anchor_windows: dict[str, dict[str, Any]],
) -> ExtractorAudit:
    with zipfile.ZipFile(xdrz, "r") as zf:
        bytes_xdmz = zf.read(sel.entry)
    datasets = parse_datamodel(bytes_xdmz, enricher)
    apply_flex_setup_reference_metadata(datasets)
    pgs, scope_debug = select_pages_for_extractor(
        pages=pages,
        sel=sel,
        datasets=datasets,
        scope_mode=options.scope_mode,
        anchor_windows=anchor_windows,
    )
    names = list(scope_debug.get("doc_search_names") or doc_names(sel.group_key))

    scoped_text_raw = "\n".join(pages[p - 1] for p in pgs if 1 <= p <= len(pages)) if pgs else ""
    scoped_table_raw = "\n".join(table_pages[p - 1] for p in pgs if 1 <= p <= len(table_pages)) if pgs else ""
    global_text_raw = "\n".join(pages)
    global_table_raw = "\n".join(table_pages)
    if not scoped_text_raw:
        scoped_text_raw = global_text_raw
    if not scoped_table_raw:
        scoped_table_raw = global_table_raw
    scoped_text_norm = normalize_token(scoped_text_raw)
    scoped_table_norm = normalize_token(scoped_table_raw)

    fields = list(dict.fromkeys(f for ds in datasets for f in ds.fields))
    # Compare fields only within the extractor's matched pages.
    # Global documentation fallback can create false positives across unrelated extractors.
    flex_fields = {
        f.upper()
        for ds in datasets
        if is_g_flex_setup_dataset(ds.name)
        for f in ds.fields
    }
    global_text_norm = normalize_token(global_text_raw)
    global_table_norm = normalize_token(global_table_raw)
    missing: list[str] = []
    for f in fields:
        in_scope = contains_field(scoped_text_raw, scoped_text_norm, f) or contains_field(scoped_table_raw, scoped_table_norm, f)
        if in_scope:
            continue
        if f.upper() in flex_fields:
            # G_FLEX_SETUP can be documented in shared/global sections; avoid scoped-window false positives.
            in_global = contains_field(global_text_raw, global_text_norm, f) or contains_field(global_table_raw, global_table_norm, f)
            if in_global:
                continue
        missing.append(f)
    missing = suppress_flex_setup_false_positives(missing, datasets, scope_debug)
    scope_debug["field_presence"] = {
        f: ("present" if f not in missing else "missing")
        for f in fields
    }

    source_review_count = 0
    enrich_warning_count = 0
    warnings: list[str] = []
    for fld in missing:
        _, meta = find_meta_for_field(datasets, fld)
        enriched = enricher.enrich(meta.get("source_table", "N/A"), meta.get("source_column", fld.upper()))
        meta["oracle_doc_url"] = enriched.doc_url
        if enriched.matched and enriched.description:
            meta["oracle_description"] = enriched.description
        else:
            meta["oracle_description"] = "A preencher pela documentação funcional 5"
        if enriched.characteristic and enriched.characteristic != "N/A":
            meta["oracle_characteristic"] = enriched.characteristic
        else:
            meta["oracle_characteristic"] = meta.get("data_type", "N/A")
        meta["oracle_provider"] = enriched.provider
        meta["oracle_matched"] = "true" if enriched.matched else "false"
    st = compute_status(pgs, len(fields), len(missing))

    return ExtractorAudit(
        extractor=sel.group_key,
        selected_entry=sel.entry,
        selected_file=sel.file_name,
        status=st,
        doc_search_names=names,
        doc_pages=pgs,
        field_count=len(fields),
        missing_fields_in_doc=missing,
        datasets=datasets,
        rejected_candidates=[{"file_name": r.file_name, "entry": r.entry, "official": r.official, "reasons": r.reasons, "version": r.version, "score": round(r.score, 2)} for r in rej],
        source_review_count=source_review_count,
        enrich_warning_count=enrich_warning_count,
        warnings=warnings,
        scope_debug=scope_debug,
    )


def collect_tables(a: ExtractorAudit) -> list[str]:
    t: list[str] = []
    for ds in a.datasets:
        for x in ds.sql_tables:
            if x not in t:
                t.append(x)
    return t


def split_doc_scope(audits: list[ExtractorAudit]) -> tuple[list[ExtractorAudit], list[ExtractorAudit]]:
    in_doc = [a for a in audits if a.status != "NOT_IN_DOC"]
    not_doc = [a for a in audits if a.status == "NOT_IN_DOC"]
    return in_doc, not_doc


def missing_rows_grouped_by_dataset(a: ExtractorAudit) -> dict[str, list[dict[str, str]]]:
    grouped: dict[str, list[dict[str, str]]] = {}
    for fld in a.missing_fields_in_doc:
        ds_name, meta = find_meta_for_field(a.datasets, fld)
        src_table = str(meta.get("source_table", "N/A") or "N/A")
        src_col = meta.get("source_column", fld)
        dtype = meta.get("oracle_characteristic") or meta.get("data_type", "N/A")
        desc = meta.get("oracle_description", "A preencher pela documentação funcional 6")
        doc_url = meta.get("oracle_doc_url", "")
        grouped.setdefault(ds_name, []).append(
            {
                "dataset": ds_name,
                "xml_tag_name": fld,
                "column_name": src_col,
                "column_characteristic": dtype,
                "associated_table_to_each_column": src_table,
                "column_description": desc,
                "oracle_doc_url": doc_url,
            }
        )
    return grouped


def build_markdown(audits: list[ExtractorAudit], out_md: Path) -> None:
    in_doc, not_doc = split_doc_scope(audits)
    fail = sum(1 for a in in_doc if a.status == "FAIL")
    warn = sum(1 for a in in_doc if a.status == "PASS_WITH_WARNINGS")
    pas = sum(1 for a in in_doc if a.status == "PASS")
    top = sorted([a for a in in_doc if a.status != "PASS"], key=lambda a: len(a.missing_fields_in_doc), reverse=True)[:10]

    lines = [
        "# Relatório de Comparação: Reference Equity DocSync",
        "",
        f"- extractors no escopo da documentação: **{len(in_doc)}**",
        f"- extractors fora do escopo da documentação: **{len(not_doc)}**",
        f"- FAIL: **{fail}** | PASS_WITH_WARNINGS: **{warn}** | PASS: **{pas}**",
        "",
        "## Principais prioridades de revisão",
        "| Prioridade | Extractor | Status | Campos ausentes |",
        "|---:|---|---|---:|",
    ]
    for i, a in enumerate(top, start=1):
        lines.append(f"| {i} | {a.extractor} | {a.status} | {len(a.missing_fields_in_doc)} |")

    out_md.write_text("\n".join(lines), encoding="utf-8")

def safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "extractor"


def build_extractor_table_diagnostics_rows(a: ExtractorAudit) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for ds in a.datasets:
        for fld in ds.fields:
            meta = ds.field_details.get(fld.upper(), {})
            src_table = str(meta.get("source_table", "N/A") or "N/A")
            src_col = str(meta.get("source_column", "UNKNOWN_SOURCE_COLUMN") or "UNKNOWN_SOURCE_COLUMN")
            rows.append(
                {
                    "dataset": ds.name,
                    "xml_tag_name": fld,
                    "source_table": src_table,
                    "source_column": src_col,
                    "source_mode": str(meta.get("source_mode", "UNKNOWN") or "UNKNOWN"),
                    "source_confidence": str(meta.get("source_confidence", "LOW") or "LOW"),
                    "is_valid_source_table": "true" if is_table_name_candidate(src_table) else "false",
                    "source_validation_reason": str(meta.get("source_validation_reason", "") or ""),
                    "source_sql_snippet": str(meta.get("source_sql_snippet", "") or ""),
                }
            )
    return rows


def write_per_extractor_csvs(audits: list[ExtractorAudit], out_dir: Path) -> None:
    by_dir = out_dir / "by_extractor"
    by_dir.mkdir(parents=True, exist_ok=True)
    headers = [
        "dataset",
        "xml_tag_name",
        "column_name",
        "column_characteristic",
        "associated_table_to_each_column",
    ]

    in_doc, _ = split_doc_scope(audits)
    for a in in_doc:
        grouped = missing_rows_grouped_by_dataset(a)
        rows = [row for rows_ds in grouped.values() for row in rows_ds]

        base_name = safe_name(a.extractor)
        if rows:
            out_rows = [{h: r.get(h, "") for h in headers} for r in rows]
            out = by_dir / f"{base_name}.csv"
            try:
                with out.open("w", encoding="utf-8-sig", newline="") as f:
                    w = csv.DictWriter(f, fieldnames=headers)
                    w.writeheader()
                    w.writerows(out_rows)
            except PermissionError:
                fallback = by_dir / f"{base_name}_new.csv"
                with fallback.open("w", encoding="utf-8-sig", newline="") as f:
                    w = csv.DictWriter(f, fieldnames=headers)
                    w.writeheader()
                    w.writerows(out_rows)

        diag_headers = [
            "dataset",
            "xml_tag_name",
            "source_table",
            "source_column",
            "source_mode",
            "source_confidence",
            "is_valid_source_table",
            "source_validation_reason",
            "source_sql_snippet",
        ]
        diag_rows = build_extractor_table_diagnostics_rows(a)
        diag_out = by_dir / f"{base_name}_table_diagnostics.csv"
        with diag_out.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=diag_headers)
            w.writeheader()
            w.writerows(diag_rows)


def format_associated_table_cell(row: dict[str, str]) -> str:
    table = row.get("associated_table_to_each_column", "N/A")
    url = row.get("oracle_doc_url", "")
    if url:
        return f"<a href='{html.escape(url)}' target='_blank' rel='noreferrer noopener'>{html.escape(table)}</a>"
    return html.escape(table)


def html_missing_rows_main(rows: list[dict[str, str]]) -> str:
    chunks: list[str] = []
    for r in rows:
        chunks.append(
            "<tr>"
            f"<td>{html.escape(r['xml_tag_name'])}</td>"
            f"<td>{html.escape(r['column_description'])}</td>"
            f"<td>{format_associated_table_cell(r)}</td>"
            f"<td>{html.escape(r['column_name'])}</td>"
            f"<td>{html.escape(r.get('column_characteristic', 'N/A'))}</td>"
            "</tr>"
        )
    return "".join(chunks) if chunks else "<tr><td colspan='5'>None</td></tr>"


def build_dashboard(audits: list[ExtractorAudit], out_html: Path, include_pass_cards: bool = False) -> None:
    in_doc, not_doc = split_doc_scope(audits)
    pass_only = [a for a in in_doc if a.status == "PASS"] if include_pass_cards else []
    actionable = [a for a in in_doc if a.status != "PASS"]

    fail = sum(1 for a in actionable if a.status == "FAIL")
    warn = sum(1 for a in actionable if a.status == "PASS_WITH_WARNINGS")
    miss_total = sum(len(a.missing_fields_in_doc) for a in actionable)
    top = sorted(actionable, key=lambda a: len(a.missing_fields_in_doc), reverse=True)[:10]

    highlights = "".join(
        f"<li><strong>{html.escape(a.extractor)}</strong></li>"
        for a in top
    ) or "<li>Sem pend\u00eancias.</li>"

    notdoc_items = "".join(
        f"<li><strong>{html.escape(a.extractor)}</strong></li>"
        for a in sorted(not_doc, key=lambda x: x.extractor.upper())
    ) or "<li>Nenhum.</li>"

    actionable_sorted = sorted(
        actionable,
        key=lambda a: (
            0 if a.status == "FAIL" else 1,
            -len(a.missing_fields_in_doc),
            a.extractor.upper(),
        ),
    )

    actionable_cards: list[str] = []
    for a in actionable_sorted:
        grouped = missing_rows_grouped_by_dataset(a)

        missing_main_sections: list[str] = []
        for ds_name, rows in grouped.items():
            missing_main_sections.append(
                f"<tr><td colspan='5'><strong>{html.escape(ds_name)}</strong></td></tr>{html_missing_rows_main(rows)}"
            )
        main_rows_html = "".join(missing_main_sections) if missing_main_sections else "<tr><td colspan='5'>None</td></tr>"

        rejected = "".join(
            f"<tr><td>{html.escape(r['file_name'])}</td><td>{html.escape(','.join(r['reasons']) if r['reasons'] else 'lower_score')}</td></tr>"
            for r in a.rejected_candidates
        ) or "<tr><td colspan='2'>None</td></tr>"

        actionable_cards.append(
            f"<details class='card actionable-card' data-extractor='{html.escape(a.extractor)}' data-status='{a.status}' data-missing='{len(a.missing_fields_in_doc)}'>"
            f"<summary><span class='title'>{html.escape(a.extractor)}</span> <span class='status {a.status}'>{a.status}</span> "
            f"<span class='meta'>missing: {len(a.missing_fields_in_doc)}</span></summary>"
            "<div class='grid'>"
            "<div class='box missing-box'><h3>Campos ausentes by dataset (document scope)</h3>"
            "<table class='inner'><thead><tr><th>XML Tag Name</th><th>Descrição da coluna</th><th>Tabela associada a cada coluna</th><th>Nome da coluna</th><th>Característica da coluna</th></tr></thead><tbody>"
            f"{main_rows_html}</tbody></table></div>"
            "<div class='box'><h3>Auditoria de seleção (rejeitados)</h3><p>Arquivos rejeitados durante a seleção da referência (TEST/BKP/ORIGINAL etc.).</p>"
            f"<table class='inner'><thead><tr><th>Arquivo rejeitado</th><th>Motivo</th></tr></thead><tbody>{rejected}</tbody></table></div>"
            "</div></details>"
        )

    pass_cards: list[str] = []
    if include_pass_cards:
        for a in sorted(pass_only, key=lambda x: x.extractor.upper()):
            pass_cards.append(
                f"<div class='card pass-card' data-extractor='{html.escape(a.extractor)}' data-status='PASS' data-missing='0'>"
                f"<div><span class='title'>{html.escape(a.extractor)}</span> <span class='status PASS'>PASS</span>"
                f" <span class='meta'>missing: 0</span></div>"
                "<div><strong>Nenhuma ação necessária.</strong></div>"
                "</div>"
            )

    actionable_cards_html = "".join(actionable_cards) if actionable_cards else (
        "<div class='card'><strong>Sem pend\u00eancias de ajuste.</strong></div>"
    )
    pass_cards_html = "".join(pass_cards)
    pass_section_html = (
        "<h2 class='section-title'>Extractors PASS</h2>"
        f"<div class='cards' id='passCards'>{pass_cards_html}</div>"
        if include_pass_cards
        else "<div class='cards' id='passCards' style='display:none'></div>"
    )

    doc = f"""<!doctype html><html lang='pt-br'><head><meta charset='utf-8'/><meta name='viewport' content='width=device-width, initial-scale=1'/><title>Reference Equity DocSync Dashboard</title>
<style>
body{{margin:0;background:#eef2f7;font-family:Segoe UI,Tahoma,sans-serif;color:#0f172a}}
.wrap{{width:min(96vw,1820px);margin:16px auto}}
.hero,.kpi,.card,.box{{background:#fff;border-radius:12px;box-shadow:0 3px 12px rgba(15,23,42,.08)}}
.hero{{padding:14px}} .kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px;margin-top:10px}}
.kpi{{padding:10px}} .n{{font-size:1.35rem;font-weight:700}} .l{{color:#475569;font-size:.9rem}}
.panel{{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}}
.filters{{display:grid;grid-template-columns:1fr 220px;gap:8px;margin-top:10px}}
.search-input,.filter-select{{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:10px;font-size:.95rem;box-sizing:border-box}}
.cards{{display:grid;gap:8px;margin-top:12px}} .card{{padding:8px}}
.section-title{{margin:14px 0 6px 0;font-size:1.02rem}}
.pass-card{{border-left:4px solid #0f766e}}
.actionable-card{{border-left:4px solid #b45309}}
.card summary{{cursor:pointer;display:flex;gap:8px;flex-wrap:wrap;align-items:center;list-style:none}} .card summary::-webkit-details-marker{{display:none}}
.title{{font-weight:700}} .meta{{font-size:.85rem;color:#475569}}
.grid{{display:grid;grid-template-columns:repeat(2,minmax(320px,1fr));gap:8px;margin-top:8px}} .box{{padding:10px;overflow:auto}}
.missing-box{{grid-column:1 / -1}}
.inner{{width:100%;border-collapse:collapse;min-width:980px}} .inner th,.inner td{{padding:7px;border-bottom:1px solid #e2e8f0;text-align:left;font-size:.9rem;vertical-align:top}}
.status{{font-weight:700;padding:2px 8px;border-radius:999px}} .PASS{{color:#0f766e;background:#ccfbf1}} .PASS_WITH_WARNINGS{{color:#b45309;background:#fef3c7}} .FAIL{{color:#b91c1c;background:#fee2e2}} .NOT_IN_DOC{{color:#334155;background:#e2e8f0}}
.kpi-green .n{{color:#0f766e}} .kpi-orange .n{{color:#b45309}} .kpi-red .n{{color:#b91c1c}} .kpi-blue .n{{color:#1d4ed8}}
@media (max-width:1100px){{.grid,.panel,.filters{{grid-template-columns:1fr}} .inner{{min-width:760px}}}}
</style></head><body><div class='wrap'>
<div class='hero'><h1>Reference Equity DocSync</h1><p>Compara referências de implementação contra o escopo da documentação. Lacunas acionáveis aparecem primeiro; itens PASS podem ser incluídos para contexto de auditoria.</p></div>
<div class='kpis'>
<div class='kpi kpi-blue'><div class='n'>{len(in_doc)}</div><div class='l'>No escopo da documentação</div></div>
<div class='kpi'><div class='n'>{len(not_doc)}</div><div class='l'>Fora do escopo da documentação</div></div>
<div class='kpi kpi-red'><div class='n'>{fail}</div><div class='l'>FAIL</div></div>
<div class='kpi kpi-orange'><div class='n'>{warn}</div><div class='l'>PASS_WITH_WARNINGS</div></div>
<div class='kpi kpi-green'><div class='n'>{len(pass_only)}</div><div class='l'>PASS (sem ação)</div></div>
<div class='kpi kpi-red'><div class='n'>{miss_total}</div><div class='l'>Campos ausentes</div></div>
</div>
<div class='panel'>
<div class='box'><strong>Principais prioridades de revisão</strong><ul>{highlights}</ul></div>
<div class='box'><strong>Extractors out of documentation scope</strong><ul>{notdoc_items}</ul></div>
</div>
<div class='filters'>
<input id='searchInput' class='search-input' type='text' placeholder='Buscar extractor, tabela, campo ou status...'/>
<select id='statusFilter' class='filter-select'>
<option value='ALL'>Todos os status</option>
<option value='FAIL'>Somente FAIL</option>
<option value='PASS_WITH_WARNINGS'>Somente PASS_WITH_WARNINGS</option>
<option value='PASS'>Somente PASS</option>
</select>
</div>
<h2 class='section-title'>Revisar primeiro</h2>
<div class='cards' id='actionableCards'>{actionable_cards_html}</div>
{pass_section_html}
</div>
<script>
(function() {{
  const searchInput = document.getElementById('searchInput');
  const statusFilter = document.getElementById('statusFilter');
  const cards = Array.from(document.querySelectorAll('#actionableCards .card, #passCards .card'));

  function applyFilters() {{
    const q = (searchInput.value || '').toLowerCase().trim();
    const status = statusFilter.value;

    cards.forEach((card) => {{
      const text = card.textContent.toLowerCase();
      const cardStatus = card.getAttribute('data-status') || '';
      const matchesText = !q || text.includes(q);
      const matchesStatus = (status === 'ALL') || (cardStatus === status);
      card.style.display = (matchesText && matchesStatus) ? '' : 'none';
    }});
  }}

  searchInput.addEventListener('input', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
}})();
</script>
</body></html>"""
    out_html.write_text(doc, encoding="utf-8")

def main() -> None:
    ap = argparse.ArgumentParser(description="Comparação Reference Equity DocSync")
    ap.add_argument("--xdrz", required=True)
    ap.add_argument("--doc", default="", help="Caminho do arquivo DOCX de documentação.")
    ap.add_argument("--pdf", default="", help="Alias legado de --doc. Aceita somente arquivos .docx.")
    ap.add_argument("--out", required=True)
    ap.add_argument("--enrich-web", action="store_true", help="Habilita enriquecimento opcional de referência Oracle (HTTP primeiro, BotCity como fallback).")
    ap.add_argument("--enrich-cache", default="", help="Path to enrichment cache json.")
    ap.add_argument("--enrich-timeout-sec", type=int, default=12)
    ap.add_argument("--scope-mode", choices=["anchor-window", "name-match"], default="anchor-window")
    ap.add_argument("--write-scope-debug-json", action="store_true")
    ap.add_argument(
        "--extractor",
        default="",
        help="Opcional: processa somente extractor(es) informado(s), separado por virgula.",
    )
    ap.add_argument(
        "--include-pass-in-dashboard",
        action="store_true",
        help="Inclui cards PASS no dashboard. Por padrao, mostra apenas alteracoes/novos datasets.",
    )
    args = ap.parse_args()

    doc_arg = (args.doc or "").strip()
    pdf_alias = (args.pdf or "").strip()
    if not doc_arg and pdf_alias:
        print("WARNING: --pdf esta deprecated; use --doc.", flush=True)
        doc_arg = pdf_alias
    if not doc_arg:
        ap.error("Parâmetro obrigatório ausente: informe --doc <arquivo.docx>.")

    xdrz, doc, out = Path(args.xdrz), Path(doc_arg), Path(args.out)
    if doc.suffix.lower() != ".docx":
        ap.error(f"Formato de documentação inválido: {doc}. Use .docx.")
    out.mkdir(parents=True, exist_ok=True)
    cache_path = Path(args.enrich_cache) if args.enrich_cache else (out / "cache" / "reference_enrich.json")
    options = CompareOptions(
        scope_mode=args.scope_mode,
        write_scope_debug_json=bool(args.write_scope_debug_json),
    )

    enricher = OracleDocsEnricher(enabled=args.enrich_web, cache_path=cache_path, timeout_sec=max(args.enrich_timeout_sec, 3))

    pages, table_pages, doc_extraction_debug = extract_docx_blocks(doc)
    selected, rejected = pick_candidates(list_candidates(xdrz))
    if args.extractor.strip():
        requested = [normalize_match_text(x) for x in args.extractor.split(",") if normalize_match_text(x)]
        selected = [s for s in selected if any(normalize_match_text(s.group_key).find(r) >= 0 for r in requested)]
        if not selected:
            ap.error(f"Nenhum extractor corresponde a --extractor={args.extractor}")
    anchor_windows = find_anchor_windows(pages, selected) if options.scope_mode == "anchor-window" else {}
    audits = [
        audit_one(s, rejected.get(s.group_key, []), pages, table_pages, xdrz, enricher, options, anchor_windows)
        for s in selected
    ]
    enricher.flush_cache()

    (out / "comparison_report.json").write_text(json.dumps([asdict(a) for a in audits], indent=2, ensure_ascii=False), encoding="utf-8")
    build_markdown(audits, out / "comparison_report.md")
    write_per_extractor_csvs(audits, out)
    build_dashboard(audits, out / "comparison_dashboard.html", include_pass_cards=bool(args.include_pass_in_dashboard))

    if options.write_scope_debug_json:
        scope_debug_payload = {
            "doc_extraction": doc_extraction_debug,
            "scope_mode": options.scope_mode,
            "extractors": [
                {
                    "extractor": a.extractor,
                    "selected_file": a.selected_file,
                    "doc_pages": a.doc_pages,
                    "missing_fields_count": len(a.missing_fields_in_doc),
                    "scope_debug": a.scope_debug,
                    "missing_fields_sample": a.missing_fields_in_doc[:40],
                }
                for a in audits
            ],
        }
        (out / "comparison_scope_debug.json").write_text(
            json.dumps(scope_debug_payload, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    in_doc, _ = split_doc_scope(audits)
    fail = sum(1 for a in in_doc if a.status == "FAIL")
    warn = sum(1 for a in in_doc if a.status == "PASS_WITH_WARNINGS")
    print("FAIL" if fail > 0 else "PASS_WITH_WARNINGS" if warn > 0 else "PASS")


if __name__ == "__main__":
    main()












