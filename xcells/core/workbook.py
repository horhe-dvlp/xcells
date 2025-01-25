from .worksheet import Worksheet
from .cell import Cell
from typing import List, Optional

class Workbook:
    def __init__(self):
        self.worksheets: List[Worksheet] = []
        self._shared_strings: List[str] = []
        
    def add_worksheet(self, worksheet: Worksheet) -> None:
        """Add a worksheet to the workbook"""
        self._fill_cell_values(worksheet)
        self.worksheets.append(worksheet)

    def get_worksheets(self) -> List[Worksheet]:
        """Return all worksheets"""
        return self.worksheets

    def get_sheet(self, index: int) -> Optional[Worksheet]:
        """Get a specific worksheet by index"""
        if 0 <= index < len(self.worksheets):
            return self.worksheets[index]
        return None
    
    def _fill_cell_values(self, worksheet: Worksheet) -> None:
        """Fill cell values"""
        for row in worksheet.cells:
            for cell in row:
                cell.value = self._shared_strings[cell.value_id]