from typing import Optional
from .cell import Cell

class Worksheet:
    """
    A class to represent a worksheet in a workbook.
    
    Attributes:
    -----------
    name : str
        The name of the worksheet.
    sheet_id : int
        The identifier of the worksheet.
    cells : list
        A list to store cells in the worksheet.
    
    Methods:
    --------
    __str__():
        Returns a string representation of the worksheet.
    get_cell(row: int, col: int) -> Optional[Cell]:
        Get the value of a cell.
    """
    def __init__(self, sheet_id: int, name: str):
        """
        Initialize a Worksheet object.
        
        Parameters:
        -----------
        sheet_id : int
            The identifier of the worksheet.
        name : str
            The name of the worksheet.
        """
        self.name = name
        self.sheet_id = sheet_id
        self.cells = []
        
    def __str__(self) -> str:
        """
        Returns a string representation of the worksheet.
        
        Returns:
        --------
        str
            A string representation of the worksheet.
        """
        return f"Worksheet: {self.name}, id: {self.sheet_id}"

    def get_cell(self, row: int, col: int) -> Optional[Cell]:
        """
        Get the value of a cell.
        
        Parameters:
        -----------
        row : int
            The row index of the cell.
        col : int
            The column index of the cell.
        
        Returns:
        --------
        Optional[Cell]
            The cell at the specified row and column, or None if not found.
        """
        return self.cells.get((row, col), None)