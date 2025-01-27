from decimal import Decimal
from fractions import Fraction
from datetime import datetime, timedelta, time
from typing import Type, Union, Optional, List
from .logger_config import logger

class Cell:
    """
    Base Cell class.
    """
    def __init__(self, col: str, row: str, value: str):
        self.row: str = row
        self.col: str = col
        self.value: str = value
        self.raw_value: str = value
        
    def __str__(self) -> str:
        """
        Returns a string representation of the cell.
        
        Returns:
        --------
        str
            A string representation of the cell.
        """
        return f"{self.__class__.__name__}(row={self.row}, col={self.col}, value={self.value}, raw_value={self.raw_value})"

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
        Initialize a NumberCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: float = float(value)
        
class CurrencyCell(Cell):
    """
    A class to represent currency cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a CurrencyCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: Decimal = Decimal(value)

class DateCell(Cell):
    """
    A class to represent date cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a DateCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: datetime = self.excel_to_date(int(value))
    
    @staticmethod
    def excel_to_date(excel_number: int) -> datetime:
        """
        Convert an Excel date number to a datetime object.
        
        Parameters:
        -----------
        excel_number : int
            The Excel date number.
        
        Returns:
        --------
        datetime
            The corresponding datetime object.
        """
        excel_epoch = datetime(1900, 1, 1)
        days_offset = excel_number - 2
        return excel_epoch + timedelta(days=days_offset)
    
class LongDateCell(DateCell):
    """
    A class to represent long date cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a LongDateCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: str = self.excel_to_date(int(value)).strftime("%A, %B %d, %Y")
    
class TimeCell(Cell):
    """
    A class to represent time cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a TimeCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: time = self.excel_decimal_to_time(Decimal(value))
    
    @staticmethod
    def excel_decimal_to_time(excel_decimal: Decimal) -> time:
        """
        Convert an Excel time decimal to a time object.
        
        Parameters:
        -----------
        excel_decimal : Decimal
            The Excel time decimal.
        
        Returns:
        --------
        time
            The corresponding time object.
        """
        full_days = int(excel_decimal)
        fractional_day = excel_decimal - full_days
        hours = int(fractional_day * 24)
        minutes = int((fractional_day * 24 - hours) * 60)
        return time(hours, minutes)

class StringCell(Cell):
    """
    A class to represent a string cell.
    """
    def __init__(self, col: str, row: str, value_id: str):
        """
        Initialize a StringCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_id : str
            An identifier for the value in the cell.
        """
        super().__init__(col, row, value_id)
        self.value: str = ""
        self.raw_value: str = value_id
        self.value_id: int = int(value_id)
        
class PercentCell(Cell):
    """
    A class to represent a percent cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a PercentCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: str = f"{float(value) * 100}%"

class FractionCell(Cell):
    """
    A class to represent a fraction cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize a FractionCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: str = self.get_shortest_fraction(Decimal(value))
        
    @staticmethod
    def get_shortest_fraction(decimal_number: Decimal) -> str:
        """
        Get the shortest fraction representation of a decimal number.
        
        Parameters:
        -----------
        decimal_number : Decimal
            The decimal number.
        
        Returns:
        --------
        str
            The shortest fraction representation.
        """
        fraction = Fraction(decimal_number).limit_denominator()
        return f"{fraction.numerator} / {fraction.denominator}"
    
class ExponentialCell(Cell):
    """
    A class to represent an exponential cell.
    """
    def __init__(self, col: str, row: str, value: str):
        """
        Initialize an ExponentialCell object.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value : str
            The value in the cell.
        """
        super().__init__(col, row, value)
        self.value: str = "{:.3e}".format(float(value))

class CellFabric:
    """
    A class to fabricate cells.
    """
    
    _instance: Optional['CellFabric'] = None
    
    CELL_STYLES: dict[str, Type[Cell]] = {
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
    
    def __new__(cls, *args, **kwargs) -> 'CellFabric':
        if not cls._instance:
            cls._instance = super(CellFabric, cls).__new__(cls)
        return cls._instance
    
    def __init__(self, styles_list: Optional[List[str]] = None):
        if styles_list:
            self.styles_list: List[str] = styles_list
    
    def create_cell(self, col: str, row: str, value_or_value_id: str, cell_style_id: str) -> Cell:
        """
        Create a cell.
        
        Parameters:
        -----------
        col : str
            The column index of the cell.
        row : str
            The row index of the cell.
        value_or_value_id : str
            The value or value identifier in the cell.
        cell_style_id : str
            The style identifier of the cell.
        
        Returns:
        --------
        Cell
            A cell object.
        """
        cell_style = self.styles_list[int(cell_style_id)]
        cell_class = CellFabric.CELL_STYLES.get(cell_style, None)
        
        if not cell_class:
            logger.warning(f"Unknown cell style: {cell_style} using default cell style.")
            cell_class = Cell
        
        return cell_class(col, row, value_or_value_id)
    
    @classmethod
    def delete_instance(cls) -> None:
        """
        Delete the singleton instance.
        """
        if cls._instance:
            logger.info("Deleting singleton instance.")
            cls._instance = None  # Remove the instance
        else:
            logger.warning("No singleton instance to delete.")
