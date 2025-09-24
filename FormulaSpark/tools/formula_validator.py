"""
FormulaSpark Formula Validation Tools
Handles formula validation, testing, and error checking
"""

from typing import Tuple, Optional

class FormulaValidator:
    """Validates Excel formulas before insertion"""
    
    @staticmethod
    def validate_formula(formula: str) -> Tuple[bool, str]:
        """
        Validate Excel formula syntax
        
        Args:
            formula: The formula string to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not formula.strip():
            return False, "Formula cannot be empty"
            
        if not formula.startswith('='):
            return False, "Formula must start with '='"
        
        # Check for mismatched parentheses
        if formula.count('(') != formula.count(')'):
            return False, "Mismatched parentheses"
        
        # Check for common syntax issues
        if '==' in formula:
            return False, "Use single '=' for comparison, not '=='"
        
        # Check for invalid characters
        invalid_chars = ['<', '>', '&', '|']
        for char in invalid_chars:
            if char in formula:
                return False, f"Invalid character '{char}' in formula"
        
        # Check for common function syntax issues
        if formula.count('"') % 2 != 0:
            return False, "Mismatched quotes in formula"
        
        return True, ""
    
    @staticmethod
    def test_formula_in_excel(formula: str, sheet) -> Tuple[bool, str]:
        """
        Test formula in a temporary Excel cell
        
        Args:
            formula: The formula to test
            sheet: Excel sheet object
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        try:
            # Use a cell that's unlikely to be used
            test_cell = sheet.range('ZZ999')
            test_cell.formula = formula
            # Check if formula evaluates without error
            test_cell.clear()
            return True, ""
        except Exception as e:
            return False, str(e)
    
    @staticmethod
    def get_formula_suggestions(formula: str) -> list:
        """
        Get suggestions for improving a formula
        
        Args:
            formula: The formula to analyze
            
        Returns:
            List of suggestion strings
        """
        suggestions = []
        
        # Check for common issues and suggest improvements
        if 'VLOOKUP' in formula and 'FALSE' not in formula:
            suggestions.append("Consider using FALSE for exact match in VLOOKUP")
        
        if 'SUMIF' in formula and 'SUMIFS' not in formula:
            suggestions.append("Consider SUMIFS for multiple criteria")
        
        if 'COUNTIF' in formula and 'COUNTIFS' not in formula:
            suggestions.append("Consider COUNTIFS for multiple criteria")
        
        if formula.count('IF') > 3:
            suggestions.append("Consider using IFS or SWITCH for multiple conditions")
        
        return suggestions
