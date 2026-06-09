import sys
import tempfile
import unittest
from pathlib import Path
import zipfile
from unittest.mock import patch

VALIDATOR_DIR = Path(__file__).resolve().parents[1] / 'validator_py'
if str(VALIDATOR_DIR) not in sys.path:
    sys.path.insert(0, str(VALIDATOR_DIR))

import docsync_compare as mod


class TestSqlOriginParser(unittest.TestCase):
    @staticmethod
    def _write_docx(document_xml: str) -> mod.Path:
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.close()
        with zipfile.ZipFile(tmp.name, "w") as zf:
            zf.writestr("word/document.xml", document_xml)
        return mod.Path(tmp.name)

    def _cleanup_docx(self, path: mod.Path) -> None:
        self.addCleanup(lambda: path.exists() and path.unlink())

    @staticmethod
    def _write_xlsx_with_sheet_rows(rows: list[list[str]]) -> mod.Path:
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        sheet_rows: list[str] = []
        for r_idx, row in enumerate(rows, start=1):
            cells: list[str] = []
            for c_idx, value in enumerate(row, start=1):
                col = chr(ord("A") + c_idx - 1)
                cell_ref = f"{col}{r_idx}"
                safe = (value or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                cells.append(
                    f'<c r="{cell_ref}" t="inlineStr"><is><t>{safe}</t></is></c>'
                )
            sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
        sheet_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{ns}"><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'
        )
        with zipfile.ZipFile(tmp.name, "w") as zf:
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        return mod.Path(tmp.name)

    def test_alias_org_id_link(self):
        sql = """
        SELECT RCTL.ORG_ID ORG_ID_LINK, RCTL.CUSTOMER_TRX_ID
        FROM RA_CUSTOMER_TRX_LINES_ALL RCTL
        """
        smap = mod.parse_select_map(sql)
        hit = smap["ORG_ID_LINK"]
        self.assertEqual(hit["source_column"], "ORG_ID")
        self.assertEqual(hit["source_table"], "RA_CUSTOMER_TRX_LINES_ALL")
        self.assertEqual(hit["source_mode"], "ALIAS")

    def test_alias_as_org_id_link(self):
        sql = """
        SELECT RCTA.ORG_ID AS ORG_ID_LINK
        FROM RA_CUSTOMER_TRX_ALL RCTA
        """
        smap = mod.parse_select_map(sql)
        hit = smap["ORG_ID_LINK"]
        self.assertEqual(hit["source_column"], "ORG_ID")
        self.assertEqual(hit["source_table"], "RA_CUSTOMER_TRX_ALL")
        self.assertEqual(hit["source_mode"], "ALIAS")

    def test_case_multi_source(self):
        sql = """
        SELECT CASE WHEN 1=1 THEN A.COL1 ELSE B.COL2 END AS X
        FROM T_A A
        JOIN T_B B ON B.ID = A.ID
        """
        smap = mod.parse_select_map(sql)
        hit = smap["X"]
        self.assertEqual(hit["source_mode"], "MULTI_SOURCE")
        self.assertEqual(hit["source_column"], "MULTI_SOURCE")

    def test_function_single_source(self):
        sql = """
        SELECT NVL(A.COL1, 0) AS X
        FROM T_A A
        """
        smap = mod.parse_select_map(sql)
        hit = smap["X"]
        self.assertEqual(hit["source_mode"], "EXPRESSION")
        self.assertEqual(hit["source_column"], "COL1")

    def test_subquery_alias_trace(self):
        sql = """
        SELECT linhas.ORG_ID_LINK
        FROM (
            SELECT RCTL.ORG_ID AS ORG_ID_LINK
            FROM RA_CUSTOMER_TRX_LINES_ALL RCTL
        ) linhas
        """
        smap = mod.parse_select_map(sql)
        hit = smap["ORG_ID_LINK"]
        self.assertEqual(hit["source_column"], "ORG_ID")
        self.assertEqual(hit["source_table"], "RA_CUSTOMER_TRX_LINES_ALL")

    def test_oracle_html_metadata_table(self):
        html_page = """
        <html><body>
        <table>
          <tr><th>Column Name</th><th>Datatype</th><th>Comments</th></tr>
          <tr><td>LOCATION_ID</td><td>NUMBER</td><td>Location identifier</td></tr>
          <tr><td>LOCATION_NAME</td><td>VARCHAR2(240)</td><td>Location display name</td></tr>
        </table>
        </body></html>
        """
        enricher = mod.OracleDocsEnricher(enabled=False, cache_path=mod.Path("."), timeout_sec=2)
        desc, dtype = enricher._extract_doc_metadata_from_html(html_page, "LOCATION_NAME")
        self.assertEqual(desc, "Location display name")
        self.assertEqual(dtype, "VARCHAR2(240)")

    def test_oracle_html_metadata_not_found(self):
        html_page = "<html><body><p>No table here</p></body></html>"
        enricher = mod.OracleDocsEnricher(enabled=False, cache_path=mod.Path("."), timeout_sec=2)
        desc, dtype = enricher._extract_doc_metadata_from_html(html_page, "LOCATION_NAME")
        self.assertEqual(desc, "")
        self.assertEqual(dtype, "")

    def test_decode_with_lookup_result_source(self):
        sql = """
        SELECT DECODE(ARX.JG_EXT_GL_ATTR_CAT, 'JLxBRFsclHdrAtrbExtGoodsExport', LOCN.LOCATION_NAME, '') AS EXPORT_SHIP_LOCATION_DESC
        FROM (
            SELECT F.GLOBAL_ATTRIBUTE_CATEGORY AS JG_EXT_GL_ATTR_CAT,
                   F.GLOBAL_ATTRIBUTE_NUMBER1 AS JG_EXT_GL_ATTR_NUMBER1
            FROM JG_FSCL_HDRS_ATRB_EXT_ALL F
        ) ARX
        LEFT JOIN LOC LOCN ON ARX.JG_EXT_GL_ATTR_NUMBER1 = LOCN.LOCATION_ID
        """
        smap = mod.parse_select_map(sql)
        hit = smap["EXPORT_SHIP_LOCATION_DESC"]
        self.assertEqual(hit["source_table"], "LOC")
        self.assertEqual(hit["source_column"], "LOCATION_NAME")
        self.assertEqual(hit["source_mode"], "EXPRESSION")

    def test_with_cte_lookup_resolves_physical_table(self):
        sql = """
        WITH LOC AS (
            SELECT LOT.LOCATION_NAME, LOC.LOCATION_ID
            FROM PER_LOCATION_DETAILS_F LOC
            JOIN PER_LOCATION_DETAILS_F_TL LOT ON LOT.LOCATION_DETAILS_ID = LOC.LOCATION_DETAILS_ID
        )
        SELECT DECODE(A.CAT, 'X', LOCN.LOCATION_NAME, '') AS EXPORT_SHIP_LOCATION_DESC
        FROM (SELECT 'X' CAT, 1 AS ID FROM DUAL) A
        LEFT JOIN LOC LOCN ON LOCN.LOCATION_ID = A.ID
        """
        smap = mod.parse_select_map(sql)
        hit = smap["EXPORT_SHIP_LOCATION_DESC"]
        self.assertEqual(hit["source_table"], "PER_LOCATION_DETAILS_F_TL")
        self.assertEqual(hit["source_column"], "LOCATION_NAME")
        self.assertEqual(hit["source_mode"], "EXPRESSION")

    def test_anchor_window_scope_keeps_continuation_pages(self):
        pages = [
            "Header CustomerReference G_HEADER",
            "continuation page with G_FLEX_SETUP ISV_PROTOCOL_NUM details",
            "OrderReference G_OTHER",
        ]
        selected = [
            mod.Candidate(entry="a", file_name="a.xdmz", group_key="CustomerReference", version=0, score=1.0, official=True, reasons=[]),
            mod.Candidate(entry="b", file_name="b.xdmz", group_key="OrderReference", version=0, score=1.0, official=True, reasons=[]),
        ]
        windows = mod.find_anchor_windows(pages, selected)
        ds = mod.DatasetInfo(
            name="G_FLEX_SETUP",
            source_dataset="Q_FLEX_SETUP",
            fields=["ISV_PROTOCOL_NUM"],
            sql_tables=[],
            sql_aliases={},
            sql_text="",
            field_details={},
        )
        included, dbg = mod.select_pages_for_extractor(
            pages=pages,
            sel=selected[0],
            datasets=[ds],
            scope_mode="anchor-window",
            anchor_windows=windows,
        )
        self.assertEqual(included, [1, 2])
        self.assertEqual(dbg["window_start"], 1)
        self.assertEqual(dbg["window_end"], 2)

    def test_docx_extracts_simple_paragraph(self):
        doc_xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Extractor One</w:t></w:r></w:p>
          </w:body>
        </w:document>
        """
        path = self._write_docx(doc_xml)
        self._cleanup_docx(path)
        pages, table_pages, debug = mod.extract_docx_blocks(path)
        self.assertEqual(pages, ["Extractor One"])
        self.assertEqual(table_pages, [""])
        self.assertEqual(debug["paragraph_blocks"], 1)
        self.assertEqual(debug["table_blocks"], 0)

    def test_docx_extracts_table_block(self):
        doc_xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>COL_A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>DESC_A</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
          </w:body>
        </w:document>
        """
        path = self._write_docx(doc_xml)
        self._cleanup_docx(path)
        pages, table_pages, debug = mod.extract_docx_blocks(path)
        self.assertEqual(pages, ["COL_A | DESC_A"])
        self.assertEqual(table_pages, ["COL_A | DESC_A"])
        self.assertEqual(debug["table_blocks"], 1)

    def test_docx_preserves_paragraph_table_order(self):
        doc_xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>FIRST</w:t></w:r></w:p>
            <w:tbl>
              <w:tr>
                <w:tc><w:p><w:r><w:t>T1</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>T2</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p><w:r><w:t>LAST</w:t></w:r></w:p>
          </w:body>
        </w:document>
        """
        path = self._write_docx(doc_xml)
        self._cleanup_docx(path)
        pages, table_pages, _ = mod.extract_docx_blocks(path)
        self.assertEqual(pages, ["FIRST", "T1 | T2", "LAST"])
        self.assertEqual(table_pages, ["", "T1 | T2", ""])

    def test_docx_empty_document(self):
        doc_xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body></w:body>
        </w:document>
        """
        path = self._write_docx(doc_xml)
        self._cleanup_docx(path)
        pages, table_pages, debug = mod.extract_docx_blocks(path)
        self.assertEqual(pages, [])
        self.assertEqual(table_pages, [])
        self.assertIn("warning", debug)

    def test_name_match_scope_works_with_docx_blocks(self):
        doc_xml = """
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>CustomerReference details</w:t></w:r></w:p>
            <w:p><w:r><w:t>Some continuation text</w:t></w:r></w:p>
            <w:p><w:r><w:t>OrderReference details</w:t></w:r></w:p>
          </w:body>
        </w:document>
        """
        path = self._write_docx(doc_xml)
        self._cleanup_docx(path)
        pages, _, _ = mod.extract_docx_blocks(path)
        selected = mod.Candidate(
            entry="a",
            file_name="a.xdmz",
            group_key="CustomerReference",
            version=0,
            score=1.0,
            official=True,
            reasons=[],
        )
        included, debug = mod.select_pages_for_extractor(
            pages=pages,
            sel=selected,
            datasets=[],
            scope_mode="name-match",
            anchor_windows={},
        )
        self.assertEqual(included, [1])
        self.assertEqual(debug["scope_mode"], "name-match")

    def test_load_g_flex_setup_reference_from_csv_like_sheet(self):
        xlsx = self._write_xlsx_with_sheet_rows([
            ['XML Tag Name,"Column description","Associated table to each column","Column Name","Column characteristic"'],
            ['ISV_MODEL,"Desc model","FND_DESCR_FLEX_COL_USAGE_VL","APPLICATION_COLUMN_NAME","VARCHAR2(30)"'],
            ['ISV_AGENCY,"Desc agency","FND_DESCR_FLEX_COL_USAGE_VL","APPLICATION_COLUMN_NAME","VARCHAR2(150)"'],
        ])
        self.addCleanup(lambda: xlsx.exists() and xlsx.unlink())
        refs = mod.load_g_flex_setup_reference(xlsx)
        self.assertIn("ISV_MODEL", refs)
        self.assertEqual(refs["ISV_MODEL"]["column_description"], "Desc model")
        self.assertEqual(refs["ISV_AGENCY"]["column_characteristic"], "VARCHAR2(150)")

    def test_suppress_flex_setup_false_positives(self):
        ds = mod.DatasetInfo(
            name="G_FLEX_SETUP",
            source_dataset="Q_FLEX_SETUP",
            fields=["ISV_MODEL", "ISV_AGENCY"],
            sql_tables=[],
            sql_aliases={},
            sql_text="",
            field_details={},
        )
        missing = ["ISV_MODEL", "OTHER_FIELD"]
        scope_debug = {}
        refs = {
            "ISV_MODEL": {
                "xml_tag_name": "ISV_MODEL",
                "column_description": "Desc",
                "associated_table_to_each_column": "TAB",
                "column_name": "COL",
                "column_characteristic": "VARCHAR2(30)",
            }
        }
        with patch("docsync_compare.load_g_flex_setup_reference_cached", return_value=refs):
            filtered = mod.suppress_flex_setup_false_positives(missing, [ds], scope_debug)
        self.assertEqual(filtered, ["OTHER_FIELD"])
        self.assertIn("ISV_MODEL", scope_debug["flex_setup_suppressed_missing_fields"])

    def test_apply_flex_setup_reference_metadata(self):
        ds = mod.DatasetInfo(
            name="G_FLEX_SETUP",
            source_dataset="Q_FLEX_SETUP",
            fields=["ISV_MODEL"],
            sql_tables=[],
            sql_aliases={},
            sql_text="",
            field_details={"ISV_MODEL": {"source_table": "N/A", "source_column": "N/A"}},
        )
        refs = {
            "ISV_MODEL": {
                "xml_tag_name": "ISV_MODEL",
                "column_description": "Desc model",
                "associated_table_to_each_column": "FND_DESCR_FLEX_COL_USAGE_VL",
                "column_name": "APPLICATION_COLUMN_NAME",
                "column_characteristic": "VARCHAR2(30)",
            }
        }
        with patch("docsync_compare.load_g_flex_setup_reference_cached", return_value=refs):
            mod.apply_flex_setup_reference_metadata([ds])
        meta = ds.field_details["ISV_MODEL"]
        self.assertEqual(meta["oracle_description"], "Desc model")
        self.assertEqual(meta["oracle_characteristic"], "VARCHAR2(30)")
        self.assertEqual(meta["source_table"], "FND_DESCR_FLEX_COL_USAGE_VL")
        self.assertEqual(meta["source_column"], "APPLICATION_COLUMN_NAME")


if __name__ == "__main__":
    unittest.main()
