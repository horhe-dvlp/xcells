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
            <si><t>String2</t></si>
        </sst>
        """,  # sharedStrings.xml
        b"""
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1" t="s"><v>1</v></c>
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


def test_parse_worksheet(reader):
    xml_data = b"""
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row>
                <c r="A1"><v>1</v></c>
                <c r="B1"><v>2</v></c>
            </row>
            <row>
                <c r="A2"><v>3</v></c>
                <c r="B2"><v>4</v></c>
            </row>
        </sheetData>
    </worksheet>
    """
    sheet_data = reader._parse_worksheet(xml_data)
    assert len(sheet_data) == 2
    assert len(sheet_data[0]) == 2
    assert sheet_data[0][0].value_id == 1
    assert sheet_data[0][1].value_id == 2
    assert sheet_data[1][0].value_id == 3
    assert sheet_data[1][1].value_id == 4


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
