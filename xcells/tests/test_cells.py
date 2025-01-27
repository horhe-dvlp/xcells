import unittest
from decimal import Decimal
from datetime import datetime, time
from xcells.core.cell import Cell, NumberCell, CurrencyCell, DateCell, LongDateCell, TimeCell, StringCell, PercentCell, FractionCell, ExponentialCell, CellFabric

class TestCell(unittest.TestCase):
    def test_cell_initialization(self):
        cell = Cell('A', '1', 'test')
        self.assertEqual(cell.col, 'A')
        self.assertEqual(cell.row, '1')
        self.assertEqual(cell.value, 'test')
        self.assertEqual(cell.raw_value, 'test')
        
    def test_cell_str(self):
        cell = Cell('A', '1', 'test')
        self.assertEqual(str(cell), "Cell(row=1, col=A, value=test, raw_value=test)")
        
    def test_number_cell(self):
        cell = NumberCell('A', '1', '123.45')
        self.assertEqual(cell.value, 123.45)
        
    def test_currency_cell(self):
        cell = CurrencyCell('A', '1', '123.45')
        self.assertEqual(cell.value, Decimal('123.45'))
        
    def test_date_cell(self):
        cell = DateCell('A', '1', '44197')
        self.assertEqual(cell.value, datetime(2021, 1, 1))
        
    def test_long_date_cell(self):
        cell = LongDateCell('A', '1', '44197')
        self.assertEqual(cell.value, "Friday, January 01, 2021")
        
    def test_time_cell(self):
        cell = TimeCell('A', '1', '0.5')
        self.assertEqual(cell.value, time(12, 0))
        
    def test_string_cell(self):
        cell = StringCell('A', '1', '123')
        self.assertEqual(cell.value_id, 123)
        
    def test_percent_cell(self):
        cell = PercentCell('A', '1', '0.75')
        self.assertEqual(cell.value, '75.0%')
        
    def test_fraction_cell(self):
        cell = FractionCell('A', '1', '0.75')
        self.assertEqual(cell.value, '3 / 4')
        
    def test_exponential_cell(self):
        cell = ExponentialCell('A', '1', '12345')
        self.assertEqual(cell.value, '1.234e+04')
        
    def test_cell_fabric_singleton(self):
        fabric1 = CellFabric()
        fabric2 = CellFabric()
        self.assertIs(fabric1, fabric2)
        
    def test_cell_fabric_create_cell(self):
        fabric = CellFabric(styles_list=["0", "1", "2", "165", "44", "14", "166", "167", "10", "12", "13", "11", "49"])
        cell = fabric.create_cell('A', '1', '123.45', '2')
        self.assertIsInstance(cell, NumberCell)
        self.assertEqual(cell.value, 123.45)
        
    def test_cell_fabric_delete_instance(self):
        fabric = CellFabric()
        CellFabric.delete_instance()
        self.assertIsNone(CellFabric._instance)

if __name__ == '__main__':
    unittest.main()