import io
import os
import zipfile
import tempfile
import unittest

# Import from the sibling module
from xlsx_to_csv import convert_xlsx_to_csvs

WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml"/>
</Relationships>
"""

# One row with three inline strings: plain, contains comma, contains quote
SHEET1_XML = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
      <c r="B1" t="inlineStr"><is><t>a,b</t></is></c>
      <c r="C1" t="inlineStr"><is><t>He said "hi"</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

def write_minimal_xlsx(path: str) -> None:
    """Create a minimal XLSX containing one sheet with inline strings."""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""")
        z.writestr("xl/workbook.xml", WORKBOOK_XML)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/worksheets/sheet1.xml", SHEET1_XML)
        # styles.xml not necessary for these tests; sharedStrings.xml not needed due to inlineStr


class TestQuotingOptions(unittest.TestCase):
    def setUp(self) -> None:
        self.tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmpdir.cleanup)
        self.xlsx_path = os.path.join(self.tmpdir.name, "input.xlsx")
        write_minimal_xlsx(self.xlsx_path)

    def read_output_csv(self) -> str:
        # The converter names output "input - Sheet1.csv"
        out_csv = os.path.join(self.tmpdir.name, "input - Sheet1.csv")
        with open(out_csv, "r", encoding="utf-8") as f:
            return f.read()

    def test_quoting_minimal_default(self):
        outputs = convert_xlsx_to_csvs(
            xlsx_path=self.xlsx_path,
            outdir=self.tmpdir.name,
            include_hidden=False,
            delimiter=",",
            quotechar='"',
            encoding="utf-8",
            quoting_name="minimal",
            escapechar=None,
        )
        self.assertEqual(len(outputs), 1)
        content = self.read_output_csv()
        # Minimal quoting: Hello,"a,b","He said ""hi"""
        self.assertEqual(content, 'Hello,"a,b","He said ""hi"""\n')

    def test_quoting_all(self):
        outputs = convert_xlsx_to_csvs(
            xlsx_path=self.xlsx_path,
            outdir=self.tmpdir.name,
            include_hidden=False,
            delimiter=",",
            quotechar='"',
            encoding="utf-8",
            quoting_name="all",
            escapechar=None,
        )
        self.assertEqual(len(outputs), 1)
        content = self.read_output_csv()
        self.assertEqual(content, '"Hello","a,b","He said ""hi"""\n')

    def test_quoting_none_with_escape(self):
        outputs = convert_xlsx_to_csvs(
            xlsx_path=self.xlsx_path,
            outdir=self.tmpdir.name,
            include_hidden=False,
            delimiter=",",
            quotechar='"',
            encoding="utf-8",
            quoting_name="none",
            escapechar="\\",
        )
        self.assertEqual(len(outputs), 1)
        content = self.read_output_csv()
        # QUOTE_NONE with escapechar: delimiters and quotes are escaped
        self.assertEqual(content, 'Hello,a\\,b,He said \\"hi\\"\n')


if __name__ == "__main__":
    unittest.main(verbosity=2)