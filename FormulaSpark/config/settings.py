"""
FormulaSpark Configuration Management
Handles all application settings, constants, and configuration loading/saving
"""

import json
import os
from typing import Dict, Any, Optional

# Application Constants
APP_NAME = "FormulaSpark"
APP_VERSION = "1.0.0"
AUTHOR = "A Custom Tool for an Excel Innovator"
CONFIG_FILE = "formulaspark_config.json"

# Default Configuration
DEFAULT_CONFIG = {
    "ollama_base_url": "http://localhost:11434",
    "history": [],
    "temperature": 0.2,
    "top_p": 0.9,
    "max_retries": 3,
    "use_context": True,
    "auto_validate": True,
    "cache_enabled": True,
    "history_limit": 1000,
    "timeout": 90,
    "selected_headers": {}
}

# Formula Templates
FORMULA_TEMPLATES = {
    "Sum with Condition": "=SUMIF({range}, {criteria}, {sum_range})",
    "Count with Multiple Conditions": "=COUNTIFS({range1}, {criteria1}, {range2}, {criteria2})",
    "Lookup Value": "=VLOOKUP({lookup_value}, {table_array}, {col_index}, {range_lookup})",
    "Date Difference": "=DATEDIF({start_date}, {end_date}, \"D\")",
    "Text Concatenation": "=CONCATENATE({text1}, {text2})",
    "Average with Condition": "=AVERAGEIF({range}, {criteria}, {average_range})",
    "Find Maximum": "=MAX({range})",
    "Find Minimum": "=MIN({range})",
    "Count Non-Empty": "=COUNTA({range})",
    "Count Empty": "=COUNTBLANK({range})"
}

# Prompt Templates
PROMPT_TEMPLATE_SIMPLE = """
You are a world-class expert in Microsoft Excel formulas. Your sole purpose is to generate a single, valid Excel formula based on a user's request.
- **Analyze the Request:** The user wants a formula for the sheet named '{sheet_name}'. The request is: "{user_prompt}".
- **Constraint:** You MUST provide ONLY the Excel formula itself.
- **Do Not:** Do not include any explanations, introductory text, code blocks (like ```excel), or notes.
- **Example:** If the user asks to "sum A1 and B1", your response must be exactly `=SUM(A1,B1)`.
The final, complete Excel formula is:
"""

PROMPT_TEMPLATE_WITH_CONTEXT = """
You are a world-class expert in Microsoft Excel formulas. Your sole purpose is to generate a single, valid Excel formula based on a user's request and the provided sheet context.

- **Sheet Context:** The user is working on the sheet named '{sheet_name}'. The first row contains the following headers: {column_headers}.
- **User Request:** The user's request is: "{user_prompt}".
- **Analyze and Infer:** Use the column headers to infer the correct ranges and criteria. For example, if the user asks to "sum sales for 'Product A'", and the headers are "Product Name" in column B and "Sales" in column D, you should use `SUMIF(B:B, "Product A", D:D)`.
- **Constraint:** You MUST provide ONLY the Excel formula itself.
- **Do Not:** Do not include any explanations, introductory text, code blocks (like ```excel), or notes.

The final, complete Excel formula is:
"""

PROMPT_TEMPLATE_WITH_TAGS = """
You are a world-class expert in Microsoft Excel formulas. Your sole purpose is to generate valid Excel formulas based on a user's request and the provided sheet context with tagged headers.

- **Sheet Context:** The user is working on the sheet named '{sheet_name}'.
- **Tagged Headers:** {tagged_headers}
- **User Request:** The user's request is: "{user_prompt}".
- **Tag Usage:** The user may reference headers using tags (e.g., @PaymentDate, @BeginBalance). When you see tags in the request, use the corresponding column ranges.
- **Analyze and Infer:** Use the tagged headers to infer the correct ranges and criteria. For example, if the user asks to "sum @Sales where @Region is 'North'", use the corresponding column ranges for @Sales and @Region.
- **Dynamic Row Detection:** Always use dynamic row detection instead of hardcoded ranges. Use MAX(IF(column:column<>"",ROW(column:column))) to find the last row with data, then use INDIRECT("column"&lastRow) to create dynamic ranges.
- **Sheet References:** ALWAYS use sheet references in formulas. Use 'SheetName'!column:column format for all ranges to ensure formulas work across sheets. The sheet name is provided in the context as '{sheet_name}'.

- **Pivot Table Requests:** When the user asks for a "pivot table" that shows "years against payable" or similar structure:
  1. **PRIMARY - Advanced Dynamic Array Formula**: =LET(lastRow,MAX(IF('{sheet_name}'!K:K<>"",ROW('{sheet_name}'!K:K))),years,UNIQUE(YEAR('{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow))),totals,MAP(years,LAMBDA(y,SUMIFS('{sheet_name}'!AG2:INDIRECT("'{sheet_name}'!AG"&lastRow),'{sheet_name}'!E2:INDIRECT("'{sheet_name}'!E"&lastRow),"Outstanding Claims",'{sheet_name}'!D2:INDIRECT("'{sheet_name}'!D"&lastRow),"OFFLINE",'{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow),">="&DATE(y,1,1),'{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow),"<="&DATE(y,12,31)))),HSTACK(years,totals))
  2. **Alternative - Multiple SUMPRODUCT formulas**: Create separate formulas for each year
  3. **For 2022**: =SUMPRODUCT((YEAR(K2:K1000)=2022)*(E2:E1000="Outstanding Claims")*(D2:D1000="OFFLINE")*AG2:AG1000)
  4. **For 2023**: =SUMPRODUCT((YEAR(K2:K1000)=2023)*(E2:E1000="Outstanding Claims")*(D2:D1000="OFFLINE")*AG2:AG1000)
  5. **For 2024**: =SUMPRODUCT((YEAR(K2:K1000)=2024)*(E2:E1000="Outstanding Claims")*(D2:D1000="OFFLINE")*AG2:AG1000)
  6. **For 2025**: =SUMPRODUCT((YEAR(K2:K1000)=2025)*(E2:E1000="Outstanding Claims")*(D2:D1000="OFFLINE")*AG2:AG1000)

**IMPORTANT DATE HANDLING:**
- For date ranges, use Excel's DATE function for proper date formatting: DATE(year, month, day)
- For year-only queries (e.g., "2024", "2025"), use: >=DATE(2024,1,1) and <=DATE(2024,12,31)
- For date comparisons, use: >=DATE(2024,1,1) instead of ">=1/1/2024"
- For "in 2024 and 2025", use: (>=DATE(2024,1,1) AND <=DATE(2025,12,31))
- For "this year", use: >=DATE(YEAR(TODAY()),1,1) AND <=DATE(YEAR(TODAY()),12,31)
- For "last year", use: >=DATE(YEAR(TODAY())-1,1,1) AND <=DATE(YEAR(TODAY())-1,12,31)
- For "this month", use: >=DATE(YEAR(TODAY()),MONTH(TODAY()),1) AND <=EOMONTH(TODAY(),0)
- For "last month", use: >=DATE(YEAR(TODAY()),MONTH(TODAY())-1,1) AND <=EOMONTH(TODAY(),-1)

**ADVANCED FEATURES:**
- For "pivot table" that shows "years against payable" or similar structures:
  * **PRIMARY METHOD - Advanced Dynamic Array with Dynamic Rows and Sheet References**: =LET(lastRow,MAX(IF('{sheet_name}'!K:K<>"",ROW('{sheet_name}'!K:K))),years,UNIQUE(YEAR('{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow))),totals,MAP(years,LAMBDA(y,SUMIFS('{sheet_name}'!AG2:INDIRECT("'{sheet_name}'!AG"&lastRow),'{sheet_name}'!E2:INDIRECT("'{sheet_name}'!E"&lastRow),"Outstanding Claims",'{sheet_name}'!D2:INDIRECT("'{sheet_name}'!D"&lastRow),"OFFLINE",'{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow),">="&DATE(y,1,1),'{sheet_name}'!K2:INDIRECT("'{sheet_name}'!K"&lastRow),"<="&DATE(y,12,31)))),HSTACK(years,totals))
  * **Alternative - Multiple SUMPRODUCT formulas with Dynamic Rows** for each year:
  * **For 2022**: =SUMPRODUCT((YEAR(K2:INDIRECT("K"&MAX(IF(K:K<>"",ROW(K:K)))))=2022)*(E2:INDIRECT("E"&MAX(IF(E:E<>"",ROW(E:E))))="Outstanding Claims")*(D2:INDIRECT("D"&MAX(IF(D:D<>"",ROW(D:D))))="OFFLINE")*AG2:INDIRECT("AG"&MAX(IF(AG:AG<>"",ROW(AG:AG)))))
  * **For 2023**: =SUMPRODUCT((YEAR(K2:INDIRECT("K"&MAX(IF(K:K<>"",ROW(K:K)))))=2023)*(E2:INDIRECT("E"&MAX(IF(E:E<>"",ROW(E:E))))="Outstanding Claims")*(D2:INDIRECT("D"&MAX(IF(D:D<>"",ROW(D:D))))="OFFLINE")*AG2:INDIRECT("AG"&MAX(IF(AG:AG<>"",ROW(AG:AG)))))
  * **For 2024**: =SUMPRODUCT((YEAR(K2:INDIRECT("K"&MAX(IF(K:K<>"",ROW(K:K)))))=2024)*(E2:INDIRECT("E"&MAX(IF(E:E<>"",ROW(E:E))))="Outstanding Claims")*(D2:INDIRECT("D"&MAX(IF(D:D<>"",ROW(D:D))))="OFFLINE")*AG2:INDIRECT("AG"&MAX(IF(AG:AG<>"",ROW(AG:AG)))))
  * **For 2025**: =SUMPRODUCT((YEAR(K2:INDIRECT("K"&MAX(IF(K:K<>"",ROW(K:K)))))=2025)*(E2:INDIRECT("E"&MAX(IF(E:E<>"",ROW(E:E))))="Outstanding Claims")*(D2:INDIRECT("D"&MAX(IF(D:D<>"",ROW(D:D))))="OFFLINE")*AG2:INDIRECT("AG"&MAX(IF(AG:AG<>"",ROW(AG:AG)))))
  * **Use LET, UNIQUE, HSTACK** for modern Excel versions - these create dynamic pivot-like structures
  * **Always use dynamic row detection** - Use MAX(IF(column:column<>"",ROW(column:column))) to find last row
- For simple aggregations, use SUMIFS, COUNTIFS, or AVERAGEIFS with specific ranges
- For text searches, use wildcards: "*text*" for contains, "text*" for starts with, "*text" for ends with
- For case-insensitive searches, use UPPER() or LOWER() functions
- ALWAYS use specific ranges (e.g., K2:K1000) instead of full columns (K:K) for better performance

- **Constraint:** You MUST provide ONLY the Excel formula itself.
- **Do Not:** Do not include any explanations, introductory text, code blocks (like ```excel), or notes.

The final, complete Excel formula is:
"""

class ConfigManager:
    """Manages application configuration"""
    
    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from file or create default"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    merged_config = DEFAULT_CONFIG.copy()
                    merged_config.update(config)
                    return merged_config
            except Exception as e:
                print(f"Error loading config: {e}")
                return DEFAULT_CONFIG.copy()
        return DEFAULT_CONFIG.copy()
    
    def save_config(self) -> bool:
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value"""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set configuration value"""
        self.config[key] = value
    
    def update(self, updates: Dict[str, Any]) -> None:
        """Update multiple configuration values"""
        self.config.update(updates)
    
    def reset_to_defaults(self) -> None:
        """Reset configuration to defaults"""
        self.config = DEFAULT_CONFIG.copy()
    
    def get_ollama_url(self) -> str:
        """Get Ollama base URL"""
        return self.get("ollama_base_url", "http://localhost:11434")
    
    def get_model_settings(self) -> Dict[str, Any]:
        """Get model settings"""
        return {
            "temperature": self.get("temperature", 0.2),
            "top_p": self.get("top_p", 0.9),
            "max_retries": self.get("max_retries", 3),
            "timeout": self.get("timeout", 90)
        }
    
    def get_ui_settings(self) -> Dict[str, Any]:
        """Get UI settings"""
        return {
            "use_context": self.get("use_context", True),
            "auto_validate": self.get("auto_validate", True),
            "cache_enabled": self.get("cache_enabled", True),
            "history_limit": self.get("history_limit", 1000)
        }
    
    def add_history_entry(self, prompt: str, formula: str) -> None:
        """Add entry to history"""
        history = self.get("history", [])
        history.insert(0, (prompt, formula))
        
        # Limit history size
        history_limit = self.get("history_limit", 1000)
        if len(history) > history_limit:
            history = history[:history_limit]
        
        self.set("history", history)
    
    def clear_history(self) -> None:
        """Clear history"""
        self.set("history", [])
    
    def get_selected_headers(self, sheet_name: str) -> Dict[str, str]:
        """Get selected headers for a sheet"""
        selected_headers = self.get("selected_headers", {})
        return selected_headers.get(sheet_name, {})
    
    def set_selected_headers(self, sheet_name: str, headers: Dict[str, str]) -> None:
        """Set selected headers for a sheet"""
        selected_headers = self.get("selected_headers", {})
        selected_headers[sheet_name] = headers
        self.set("selected_headers", selected_headers)
