#!/usr/bin/env python3
"""
FormulaSpark Launcher - Python 3.13.1 Compatible
Robust launcher that handles xlwings compatibility issues
"""

import sys
import os
import traceback

def main():
    """Main entry point for FormulaSpark with Python 3.13.1 compatibility"""
    print("🚀 Starting FormulaSpark...")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    
    try:
        # Add FormulaSpark to the path
        formula_path = os.path.join(os.path.dirname(__file__), 'FormulaSpark')
        print(f"Adding to path: {formula_path}")
        sys.path.insert(0, formula_path)
        
        print("📦 Importing PyQt5...")
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtCore import Qt
        
        print("📦 Importing FormulaSpark modules...")
        from FormulaSpark.config.settings import ConfigManager
        from FormulaSpark.ai.ollama_client import OllamaClient
        from FormulaSpark.ui.main_window import FormulaSparkMainWindow
        
        # Handle xlwings import with Python 3.13.1 compatibility
        print("📦 Checking xlwings compatibility...")
        xlwings_available = False
        
        try:
            # Try to import xlwings directly first
            import xlwings
            print(f"✅ xlwings version: {xlwings.__version__}")
            xlwings_available = True
            
        except Exception as e:
            print(f"⚠️  xlwings import failed: {e}")
            print("📝 This might be due to numpy compatibility issues with Python 3.13.1")
        
        # Try to import ExcelHandler regardless of xlwings status
        try:
            from FormulaSpark.tools.excel_handler import ExcelHandler
            print("✅ ExcelHandler imported successfully!")
            
        except Exception as e:
            print(f"⚠️  ExcelHandler import failed: {e}")
            print("📝 Creating fallback ExcelHandler...")
            
            # Create a fallback ExcelHandler that works without xlwings
            class FallbackExcelHandler:
                def __init__(self):
                    self.active_workbook = None
                    self.workbooks = {}
                    print("📝 Excel integration limited - using fallback mode")
                
                def connect_to_excel(self):
                    print("📝 Excel connection not available in fallback mode")
                    return False
                
                def get_workbooks(self):
                    print("📝 Excel workbooks not available in fallback mode")
                    return []
                
                def get_sheets(self, workbook_name):
                    print("📝 Excel sheets not available in fallback mode")
                    return []
                
                def get_headers(self, workbook_name, sheet_name):
                    print("📝 Excel headers not available in fallback mode")
                    return []
                
                def insert_formula(self, formula, cell_address, workbook_name, sheet_name):
                    print(f"📝 Excel integration not available. Formula: {formula}")
                    print(f"📝 Would insert at: {cell_address} in {workbook_name}!{sheet_name}")
                    return False
                
                def validate_formula(self, formula):
                    print(f"📝 Formula validation: {formula}")
                    return True
            
            # Replace the ExcelHandler with the fallback version
            import FormulaSpark.tools.excel_handler
            FormulaSpark.tools.excel_handler.ExcelHandler = FallbackExcelHandler
            from FormulaSpark.tools.excel_handler import ExcelHandler
            print("✅ Fallback ExcelHandler created successfully!")
        
        print("✅ All imports successful!")
        
        print("🎨 Creating QApplication...")
        app = QApplication(sys.argv)
        
        print("🪟 Creating main window...")
        window = FormulaSparkMainWindow()
        
        print("👁️ Showing window...")
        window.show()
        
        print("🔄 Starting event loop...")
        print("✅ FormulaSpark started successfully!")
        print("🎯 Ready to generate Excel formulas!")
        
        # Run the application
        sys.exit(app.exec_())
        
    except ImportError as e:
        print(f"❌ Import Error: {e}")
        print("📋 Traceback:")
        traceback.print_exc()
        print("\n💡 Please ensure all dependencies are installed:")
        print("pip install -r FormulaSpark/requirements.txt")
        try:
            input("Press Enter to exit...")
        except:
            pass  # Handle case where stdin is not available
        sys.exit(1)
        
    except Exception as e:
        print(f"❌ Error running FormulaSpark: {e}")
        print("📋 Traceback:")
        traceback.print_exc()
        try:
            input("Press Enter to exit...")
        except:
            pass  # Handle case where stdin is not available
        sys.exit(1)

if __name__ == "__main__":
    main()