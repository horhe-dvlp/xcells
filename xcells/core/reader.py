import zipfile
from lxml import etree
from .workbook import Workbook
from .worksheet import Worksheet
from .cell import CellFabric
from .logger_config import logger
from .namespaces import NAMESPACES as ns



class Reader:
    """
    A singleton class to read Excel files in .xlsx format.

    Attributes:
        _instance (Reader): The singleton instance of the Reader class.
    """

    _instance = None

    def __new__(cls, *args, **kwargs) -> 'Reader':
        """
        Create or return the singleton instance of the Reader class.

        Args:
            *args: Variable length argument list.
            **kwargs: Arbitrary keyword arguments.

        Returns:
            Reader: The singleton instance of the Reader class.
        """
        if not cls._instance:
            cls._instance = super(Reader, cls).__new__(cls, *args, **kwargs)
        return cls._instance

    def read(self, filename: str) -> 'Workbook':
        """
        Read an Excel file in .xlsx format.

        Args:
            filename (str): The path to the Excel file.

        Returns:
            Workbook: The parsed workbook object.

        Raises:
            ValueError: If the file is not in .xlsx format.
        """

        if not filename.endswith('.xlsx'):
            raise ValueError("File must be in .xlsx format")

        workbook = Workbook()

        with zipfile.ZipFile(filename, 'r') as zipf:
            workbook_xml = zipf.read('xl/workbook.xml')
            workbook._shared_strings = self._extract_shared_strings(
                zipf.read('xl/sharedStrings.xml'))
            
            sheets = self._parse_workbook(workbook_xml)
            
            styles_list = self._extract_cell_styles(
                zipf.read('xl/styles.xml'))
            CellFabric(styles_list)

            for sheet in sheets:
                sheet.cells = self._parse_worksheet(
                    zipf.read(f'xl/worksheets/sheet{sheet.sheet_id}.xml'))

                workbook.add_worksheet(sheet)

        return workbook

    def _parse_workbook(self, xml_data: bytes) -> list:
        """
        Parse XML data for the workbook.

        Args:
            xml_data (bytes): The XML data of the workbook.

        Returns:
            list: A list of Worksheet objects.
        """
        root = etree.fromstring(xml_data)

        sheets = []
        for sheet in root.findall('.//xl:sheets/xl:sheet', ns):
            sheets.append(
                Worksheet(sheet.get("sheetId"), sheet.get("name")))

        logger.debug(f"Found {len(sheets)} worksheets, with names: " +
                     f"{', '.join([sheet.name for sheet in sheets])}")

        return sheets

    def _parse_worksheet(self, xml_data: bytes) -> list:
        """
        Parse XML data for the worksheet.

        Args:
            xml_data (bytes): The XML data of the worksheet.

        Returns:
            list: A list of rows, where each row is a list of Cell objects.
        """
        root = etree.fromstring(xml_data)
        sheet_data = []
        
        cell_fabric = CellFabric()

        for row in root.findall('.//xl:row', ns):
            row_data = []
            for cell_pos in row.findall('.//xl:c', ns):
                pos = cell_pos.get('r')
                cell_style = int(cell_pos.get('s'))
                cell_type = cell_pos.get('t')
                
                cell = cell_fabric.create_cell(pos[0], pos[1], 
                    cell_pos.find('.//xl:v', ns).text, cell_style)
                row_data.append(cell)
            sheet_data.append(row_data)

        logger.debug(f"Found {len(sheet_data)} rows in worksheet " +
                     f"with first row data id: {sheet_data[0]}")
        return sheet_data
    
    def _extract_cell_styles(self, xml_data: bytes) -> list:
        """
        Extract cell styles from the XML data.

        Args:
            xml_data (bytes): The XML data of the cell styles.

        Returns:
            list: A list of cell styles.
        """
        root = etree.fromstring(xml_data)

        cell_styles = []
        for elem in root.findall('.//xl:cellXfs/xl:xf', ns):
            num_fmt_id = elem.get('numFmtId')
            cell_styles.append(num_fmt_id)

        logger.debug(f"Found {len(cell_styles)} cell styles " +
                     f"with first style numFmtId: {cell_styles[0]}")

        return cell_styles

    def _extract_shared_strings(self, xml_data: bytes) -> list:
        """
        Extract shared strings from the XML data.

        Args:
            xml_data (bytes): The XML data of the shared strings.

        Returns:
            list: A list of shared strings.
        """
        root = etree.fromstring(xml_data)

        shared_strings = [elem.text for elem in root.findall('.//xl:t', ns)]

        logger.debug(f"Found {len(shared_strings)} shared strings " +
                     f"with first string: {shared_strings[0]}")

        return shared_strings
