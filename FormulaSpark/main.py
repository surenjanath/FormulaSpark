#!/usr/bin/env python3
"""
FormulaSpark - Main Entry Point
An intelligent Excel formula generator powered by Ollama AI
"""

import sys
import os
from PyQt5.QtWidgets import QApplication

# Add the parent directory to the path so we can import FormulaSpark modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from FormulaSpark.ui.main_window import FormulaSparkMainWindow

def main():
    """Main entry point for FormulaSpark"""
    app = QApplication(sys.argv)
    
    # Create main window (methods are automatically integrated)
    window = FormulaSparkMainWindow()
    
    # Show the window
    window.show()
    
    # Run the application
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
