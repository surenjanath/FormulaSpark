"""
FormulaSpark Helper Utilities
Contains utility functions and helper classes
"""

import re
from typing import List, Dict, Any

def clean_formula(formula: str) -> str:
    """
    Clean and normalize a formula string
    
    Args:
        formula: Raw formula string
        
    Returns:
        Cleaned formula string
    """
    if not formula:
        return ""
    
    # Remove markdown code blocks
    formula = formula.strip().replace("```excel", "").replace("```", "")
    
    # Remove leading/trailing whitespace
    formula = formula.strip()
    
    # Ensure it starts with =
    if not formula.startswith('='):
        formula = '=' + formula
    
    return formula

def extract_column_letter(column_index: int) -> str:
    """
    Convert column index to Excel column letter
    
    Args:
        column_index: 0-based column index
        
    Returns:
        Excel column letter (A, B, C, etc.)
    """
    result = ""
    while column_index >= 0:
        result = chr(65 + (column_index % 26)) + result
        column_index = column_index // 26 - 1
    return result

def parse_cell_reference(cell_ref: str) -> tuple:
    """
    Parse Excel cell reference to row and column
    
    Args:
        cell_ref: Cell reference like "A1", "B2", etc.
        
    Returns:
        Tuple of (row, column) as integers
    """
    match = re.match(r'([A-Z]+)(\d+)', cell_ref.upper())
    if not match:
        return None, None
    
    col_str, row_str = match.groups()
    
    # Convert column letters to number
    col_num = 0
    for char in col_str:
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    
    return int(row_str), col_num - 1

def generate_smart_tag(header: str) -> str:
    """
    Generate a smart tag from header name
    
    Args:
        header: Header name
        
    Returns:
        Generated tag
    """
    # Remove special characters and normalize
    tag = re.sub(r'[^\w\s]', '', header)
    tag = re.sub(r'\s+', '_', tag.strip())
    
    # Convert to camelCase
    words = tag.split('_')
    if len(words) == 1:
        return f"@{words[0].capitalize()}"
    
    # Handle common prefixes
    if words[0].lower() in ['beginning', 'start']:
        return f"@Begin{''.join(word.capitalize() for word in words[1:])}"
    elif words[0].lower() in ['ending', 'end']:
        return f"@End{''.join(word.capitalize() for word in words[1:])}"
    elif words[0].lower() in ['total', 'sum']:
        return f"@Total{''.join(word.capitalize() for word in words[1:])}"
    else:
        return f"@{''.join(word.capitalize() for word in words)}"

def validate_excel_range(range_str: str) -> bool:
    """
    Validate Excel range string
    
    Args:
        range_str: Range string like "A1:B10"
        
    Returns:
        True if valid range
    """
    # Simple range validation
    if ':' not in range_str:
        return False
    
    start, end = range_str.split(':', 1)
    start_row, start_col = parse_cell_reference(start)
    end_row, end_col = parse_cell_reference(end)
    
    return (start_row is not None and start_col is not None and 
            end_row is not None and end_col is not None)

def format_formula_for_display(formula: str, max_length: int = 50) -> str:
    """
    Format formula for display in UI
    
    Args:
        formula: Formula string
        max_length: Maximum length before truncation
        
    Returns:
        Formatted formula string
    """
    if len(formula) <= max_length:
        return formula
    
    return formula[:max_length-3] + "..."

class FormulaAnalyzer:
    """Analyzes formulas for patterns and suggestions"""
    
    @staticmethod
    def get_function_usage(formula: str) -> Dict[str, int]:
        """
        Get count of Excel functions used in formula
        
        Args:
            formula: Formula string
            
        Returns:
            Dictionary of function names and their counts
        """
        # Common Excel functions
        functions = [
            'SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP',
            'INDEX', 'MATCH', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS',
            'AVERAGEIF', 'AVERAGEIFS', 'CONCATENATE', 'LEFT', 'RIGHT', 'MID',
            'LEN', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'TRIM',
            'UPPER', 'LOWER', 'PROPER', 'TEXT', 'VALUE', 'DATE', 'TIME',
            'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'NOW', 'TODAY'
        ]
        
        usage = {}
        formula_upper = formula.upper()
        
        for func in functions:
            count = formula_upper.count(func + '(')
            if count > 0:
                usage[func] = count
        
        return usage
    
    @staticmethod
    def get_complexity_score(formula: str) -> int:
        """
        Calculate complexity score for formula
        
        Args:
            formula: Formula string
            
        Returns:
            Complexity score (higher = more complex)
        """
        score = 0
        
        # Base score for length
        score += len(formula) // 10
        
        # Add points for nested functions
        score += formula.count('(') * 2
        
        # Add points for complex functions
        complex_functions = ['VLOOKUP', 'INDEX', 'MATCH', 'IF', 'SUMIFS', 'COUNTIFS']
        for func in complex_functions:
            score += formula.upper().count(func + '(') * 3
        
        # Add points for array formulas
        if formula.startswith('{') and formula.endswith('}'):
            score += 10
        
        return score
    
    @staticmethod
    def suggest_improvements(formula: str) -> List[str]:
        """
        Suggest improvements for a formula
        
        Args:
            formula: Formula string
            
        Returns:
            List of improvement suggestions
        """
        suggestions = []
        
        # Check for common issues
        if 'VLOOKUP' in formula and 'FALSE' not in formula:
            suggestions.append("Consider using FALSE for exact match in VLOOKUP")
        
        if 'SUMIF' in formula and 'SUMIFS' not in formula:
            suggestions.append("Consider SUMIFS for multiple criteria")
        
        if 'COUNTIF' in formula and 'COUNTIFS' not in formula:
            suggestions.append("Consider COUNTIFS for multiple criteria")
        
        if formula.count('IF') > 3:
            suggestions.append("Consider using IFS or SWITCH for multiple conditions")
        
        if 'A:A' in formula or 'B:B' in formula:
            suggestions.append("Consider using specific ranges instead of entire columns for better performance")
        
        return suggestions
