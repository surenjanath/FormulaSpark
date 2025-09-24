"""
FormulaSpark Excel Integration Tools
Handles Excel connection, data reading, and formula insertion
"""

import xlwings as xw
from typing import List, Dict, Optional, Tuple

class ExcelHandler:
    """Handles Excel operations and data extraction"""
    
    def __init__(self):
        self.active_workbook = None
    
    def connect_to_active_workbook(self) -> Tuple[bool, str, Optional[object]]:
        """
        Connect to the active Excel workbook
        
        Returns:
            Tuple of (success, message, workbook_object)
        """
        try:
            # Add timeout to prevent hanging
            import time
            start_time = time.time()
            timeout = 5  # 5 second timeout
            
            # First check if Excel is running
            try:
                xw.apps.active
            except Exception:
                return False, "Excel is not running. Please start Excel and open a workbook first.", None
            
            # Try to get active workbook with timeout
            wb = None
            while wb is None and (time.time() - start_time) < timeout:
                try:
                    wb = xw.books.active
                    if wb is None:
                        time.sleep(0.1)  # Small delay before retry
                except Exception:
                    time.sleep(0.1)
                    continue
            
            if wb is None:
                return False, "No active workbook found. Please ensure Excel is running and has an open workbook.", None
            
            self.active_workbook = wb
            return True, f"Successfully connected to {wb.name}", wb
            
        except ImportError:
            return False, "xlwings library not found. Please install it via: pip install xlwings", None
        except Exception as e:
            error_msg = str(e)
            if "Call was rejected by callee" in error_msg:
                return False, "Excel is not responding. Please ensure Excel is running and try again.", None
            elif "COM" in error_msg:
                return False, "Cannot connect to Excel. Please ensure Excel is running and try again.", None
            else:
                return False, f"Connection failed: {error_msg}", None
    
    def get_sheet_names(self) -> List[str]:
        """Get list of sheet names from active workbook"""
        if not self.active_workbook:
            return []
        
        try:
            return [sheet.name for sheet in self.active_workbook.sheets]
        except Exception:
            return []
    
    def get_headers(self, sheet_name: str) -> List[str]:
        """
        Get column headers from the first row of a sheet
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            List of header strings
        """
        if not self.active_workbook:
            return []
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            headers = sheet.range('A1').expand('right').value
            
            if isinstance(headers, str):
                headers = [headers]
            
            return [h for h in headers if h]
        except Exception:
            return []
    
    def get_headers_with_column_info(self, sheet_name: str) -> Dict[str, Dict]:
        """
        Get headers with their column information
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Dictionary mapping headers to column info
        """
        headers = self.get_headers(sheet_name)
        result = {}
        
        for i, header in enumerate(headers):
            column_letter = chr(65 + i)  # A, B, C, etc.
            result[header] = {
                'column': column_letter,
                'range': f"{column_letter}:{column_letter}",
                'index': i
            }
        
        return result
    
    def insert_formula(self, sheet_name: str, cell_address: str, formula: str) -> Tuple[bool, str]:
        """
        Insert formula into a specific cell
        
        Args:
            sheet_name: Name of the sheet
            cell_address: Cell address (e.g., 'A1')
            formula: Formula to insert
            
        Returns:
            Tuple of (success, message)
        """
        if not self.active_workbook:
            return False, "Not connected to Excel"
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            cell = sheet.range(cell_address)
            cell.formula = formula
            return True, f"Formula inserted into {cell_address}"
        except Exception as e:
            return False, f"Failed to insert formula: {e}"
    
    def insert_formula_to_active_cell(self, formula: str) -> Tuple[bool, str]:
        """
        Insert formula into the currently active cell
        
        Args:
            formula: Formula to insert
            
        Returns:
            Tuple of (success, message)
        """
        if not self.active_workbook:
            return False, "Not connected to Excel"
        
        try:
            active_cell = self.active_workbook.selection
            active_cell.formula = formula
            return True, f"Formula inserted into cell {active_cell.address}"
        except Exception as e:
            return False, f"Failed to insert formula: {e}"
    
    def test_formula_in_cell(self, sheet_name: str, cell_address: str, formula: str) -> Tuple[bool, str]:
        """
        Test a formula in a temporary cell
        
        Args:
            sheet_name: Name of the sheet
            cell_address: Cell address to test in
            formula: Formula to test
            
        Returns:
            Tuple of (success, message)
        """
        if not self.active_workbook:
            return False, "Not connected to Excel"
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            test_cell = sheet.range(cell_address)
            
            # Store original value
            original_value = test_cell.value
            
            # Test the formula
            test_cell.formula = formula
            
            # Restore original value
            test_cell.value = original_value
            
            return True, "Formula test successful"
        except Exception as e:
            return False, f"Formula test failed: {e}"
    
    def get_cell_value(self, sheet_name: str, cell_address: str) -> Tuple[bool, any]:
        """
        Get value from a specific cell
        
        Args:
            sheet_name: Name of the sheet
            cell_address: Cell address
            
        Returns:
            Tuple of (success, value)
        """
        if not self.active_workbook:
            return False, None
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            cell = sheet.range(cell_address)
            return True, cell.value
        except Exception:
            return False, None
    
    def get_range_values(self, sheet_name: str, range_address: str) -> Tuple[bool, List]:
        """
        Get values from a range of cells
        
        Args:
            sheet_name: Name of the sheet
            range_address: Range address (e.g., 'A1:C10')
            
        Returns:
            Tuple of (success, values_list)
        """
        if not self.active_workbook:
            return False, []
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            range_obj = sheet.range(range_address)
            return True, range_obj.value
        except Exception:
            return False, []
    
    def is_connected(self) -> bool:
        """Check if connected to Excel"""
        return self.active_workbook is not None
    
    def disconnect(self):
        """Disconnect from Excel"""
        self.active_workbook = None
    
    def detect_date_columns(self, sheet_name: str, sample_size: int = 10) -> Dict[str, str]:
        """
        Detect which columns contain date data and their format
        
        Args:
            sheet_name: Name of the sheet
            sample_size: Number of rows to sample for detection
            
        Returns:
            Dictionary with column names and their detected date format
        """
        if not self.active_workbook:
            return {}
        
        try:
            sheet = self.active_workbook.sheets[sheet_name]
            headers = self.get_headers(sheet_name)
            date_columns = {}
            
            for i, header in enumerate(headers):
                column_letter = chr(65 + i)
                # Sample a few rows to detect date format
                sample_range = f"{column_letter}2:{column_letter}{min(sample_size + 1, sheet.used_range.last_cell.row)}"
                sample_data = sheet.range(sample_range).value
                
                if isinstance(sample_data, list):
                    sample_data = [d for d in sample_data if d is not None]
                else:
                    sample_data = [sample_data] if sample_data is not None else []
                
                # Check if any values look like dates
                date_like_count = 0
                for value in sample_data:
                    if isinstance(value, (int, float)) and 1 <= value <= 50000:  # Excel date range
                        date_like_count += 1
                    elif isinstance(value, str) and any(char in value for char in ['/', '-', '.']) and len(value) > 4:
                        date_like_count += 1
                
                if date_like_count > len(sample_data) * 0.5:  # More than 50% look like dates
                    date_columns[header] = "DATE"
            
            return date_columns
        except Exception as e:
            print(f"Error detecting date columns: {e}")
            return {}
    
    def create_formula_sheet(self, formula: str, sheet_name: str = None, source_sheet_name: str = None) -> Tuple[bool, str, str]:
        """
        Create a new sheet and insert formula with headers
        
        Args:
            formula: Formula to insert
            sheet_name: Optional name for the new sheet (default: auto-generated)
            source_sheet_name: Name of the source sheet that the formula references
            
        Returns:
            Tuple of (success, message, new_sheet_name)
        """
        if not self.active_workbook:
            return False, "Not connected to Excel", ""
        
        try:
            import datetime
            
            # Generate sheet name if not provided
            if not sheet_name:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                sheet_name = f"Formula_{timestamp}"
            
            # Create new sheet
            new_sheet = self.active_workbook.sheets.add(sheet_name)
            
            # Add headers
            new_sheet.range('A1').value = "Generated Formula"
            new_sheet.range('A2').value = "Generated on:"
            new_sheet.range('B2').value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_sheet.range('A3').value = "Source Sheet:"
            if source_sheet_name:
                new_sheet.range('B3').value = f"'{source_sheet_name}'"
            else:
                new_sheet.range('B3').value = f"'{self.active_workbook.sheets[0].name}'"  # Reference to first sheet
            new_sheet.range('A4').value = "Formula:"
            new_sheet.range('A5').value = formula
            
            # Insert formula in A6 - use formula2 for better compatibility
            try:
                new_sheet.range('A6').formula2 = formula
                print(f"DEBUG: Formula inserted using formula2: {formula[:100]}...")
            except Exception as e1:
                print(f"DEBUG: formula2 failed: {e1}")
                try:
                    # Fallback to regular formula if formula2 fails
                    new_sheet.range('A6').formula = formula
                    print(f"DEBUG: Formula inserted using formula: {formula[:100]}...")
                except Exception as e2:
                    print(f"DEBUG: Both formula methods failed: {e2}")
                    # Last resort - insert as text and let user copy
                    new_sheet.range('A6').value = f"= {formula}"
                    new_sheet.range('A7').value = "Note: Formula inserted as text. Please copy and paste manually."
            
            # Add some formatting
            new_sheet.range('A1').font.bold = True
            new_sheet.range('A3').font.bold = True
            new_sheet.range('A4').font.name = 'Courier New'
            new_sheet.range('A4').font.size = 10
            
            # Auto-fit columns
            new_sheet.autofit()
            
            return True, f"Formula inserted into new sheet '{sheet_name}'", sheet_name
            
        except Exception as e:
            return False, f"Failed to create formula sheet: {e}", ""
