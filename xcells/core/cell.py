from decimal import Decimal
from fractions import Fraction
from datetime import datetime, timedelta, time
from .logger_config import logger

class Cell:
    """
    Base Cell class.
    """
    def __init__(self, col: str, row: str, value: str):
        self.row = row
        self.col = col
        self.value = value
        
    def __str__(self) -> str:
        """
        Returns a string representation of the cell.
        
        Returns:
        --------
        str
            A string representation of the cell.
        """
        return f"Cell({self.row}, {self.col}, {self.value})"

    def __repr__(self) -> str:
        """
        Returns a string representation of the cell for debugging.
        
        Returns:
        --------
        str
            A string representation of the cell for debugging.
        """
        return str(self)
    
class NumberCell(Cell):
    """
    A class to represent number cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = float(value)
        
class CurrencyCell(Cell):
    """
    A class to represent currency cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = Decimal(value)

class DateCell(Cell):
    """
    A class to represent date cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = self.excel_to_date(int(value))
    
    @staticmethod
    def excel_to_date(excel_number):
        excel_epoch = datetime(1900, 1, 1)
        days_offset = excel_number - 2
        return excel_epoch + timedelta(days=days_offset)
    
class LongDateCell(DateCell):
    """
    A class to represent long date cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = self.excel_to_date(int(value)).strftime("%A, %B %d, %Y")
    
class TimeCell(Cell):
    
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = self.excel_decimal_to_time(Decimal(value))
    
    @staticmethod
    def excel_decimal_to_time(excel_decimal):
        full_days = int(excel_decimal)
        fractional_day = excel_decimal - full_days
        

        hours = int(fractional_day * 24)

        minutes = int((fractional_day * 24 - hours) * 60)

        return time(hours, minutes)


class StringCell(Cell):
    """
    A class to represent a cell.
    """
    def __init__(self, col: str, row: str, value_id: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = ""
        self.value_id = int(value_id)
        
class PercentCell(Cell):
    """
    A class to represent a cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = f"{float(value) * 100}%"

class FractionCell(Cell):
    """
    A class to represent a cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a Cell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        """
        self.row = row
        self.col = col
        self.value = self.get_shortest_fraction(Decimal(value))
        
    @staticmethod
    def get_shortest_fraction(decimal_number):
        fraction = Fraction(decimal_number).limit_denominator()

        return f"{fraction.numerator} / {fraction.denominator}"
    

class ExponentialCell(Cell):
    def __init__(self, col, row, value):
        self.row = row
        self.col = col
        self.value = "{:.3e}".format(float(value))
        


class CellFabric:
    """
    A class to fabricate cells.
    """
    
    _instance = None
    
    CELL_STYLES = {
        "0": Cell,
        "1": Cell,
        "2": NumberCell,
        "165": CurrencyCell,
        "44": CurrencyCell,
        "14": DateCell,
        "166": LongDateCell,
        "167": TimeCell,
        "10": PercentCell,
        "12": FractionCell,
        "13": FractionCell,
        "11": ExponentialCell,
        "49": StringCell
    }
    
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(CellFabric, cls).__new__(cls)
        return cls._instance
    
    def __init__(self, styles_list: list|None = None):
        if styles_list:
            self.styles_list = styles_list
    

    def create_cell(self, col: str, row: str, value_or_value_id: str, cell_style_id: str) -> Cell:
        """
        Create a cell.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : int
            An identifier for the value in the cell.
        cell_type : str
            The type of the cell.
        
        Returns:
        --------
        Cell
            A cell object.
        """
        
        cell_style = self.styles_list[cell_style_id]
        
        cell_class = CellFabric.CELL_STYLES.get(cell_style, None)
        
        if not cell_class:
            logger.warning(f"Unknown cell style: {cell_style} using default cell style.")
            cell_class = Cell
            

        
        return cell_class(col, row, value_or_value_id)
    
    @classmethod
    def delete_instance(cls):
        """
        Delete the singleton instance.
        """
        if cls._instance:
            logger.info("Deleting singleton instance.")
            cls._instance = None  # Remove the instance
        else:
            logger.warning("No singleton instance to delete.")
    
    
