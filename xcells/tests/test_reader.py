import pytest
import zipfile
from unittest.mock import patch, MagicMock
from lxml import etree
from xcells.core.reader import Reader
from xcells.core.workbook import Workbook
from xcells.core.worksheet import Worksheet
from xcells.core.cell import Cell


@pytest.fixture
def reader():
    return Reader()


def test_read_invalid_file_extension(reader):
    with pytest.raises(ValueError, match="File must be in .xlsx format"):
        reader.read("invalid_file.txt")


@patch("xcells.core.reader.zipfile.ZipFile")
def test_read_valid_file(mock_zipfile, reader):
    mock_zip = MagicMock()
    mock_zip.read.side_effect = [
            b"""
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1" />
                </sheets>
            </workbook>
            """,  # workbook.xml
            b"""
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <si><t>String1</t></si>
            </sst>
            """,  # sharedStrings.xml
            b"""
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
                    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" />
                    <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" />
                </cellXfs>
            </styleSheet>
            """,  # styles.xml
            b"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" s="1" ><v>0</v></c>
                        <c r="B1" s="2" ><v>0.75</v></c>
                    </row>
                </sheetData>
            </worksheet>
            """  # worksheet.xml
        ]
    mock_zipfile.return_value.__enter__.return_value = mock_zip

    workbook = reader.read("valid_file.xlsx")

    assert isinstance(workbook, Workbook)
    assert len(workbook.worksheets) == 1


def test_parse_workbook(reader):
    xml_data = b"""
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheets>
            <sheet name="Sheet1" sheetId="1"/>
            <sheet name="Sheet2" sheetId="2"/>
        </sheets>
    </workbook>
    """
    sheets = reader._parse_workbook(xml_data)
    assert len(sheets) == 2
    assert sheets[0].name == "Sheet1"
    assert sheets[1].name == "Sheet2"



def test_extract_shared_strings(reader):
    xml_data = b"""
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
        <si><t>String1</t></si>
        <si><t>String2</t></si>
    </sst>
    """
    shared_strings = reader._extract_shared_strings(xml_data)
    assert len(shared_strings) == 2
    assert shared_strings[0] == "String1"
    assert shared_strings[1] == "String2"
