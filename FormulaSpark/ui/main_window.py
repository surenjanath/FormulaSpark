"""
FormulaSpark Main Window
The main application window with all UI components
"""

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit, QListWidget,
    QListWidgetItem, QCheckBox, QProgressBar, QSplitter, QGroupBox,
    QMenuBar, QAction, QStatusBar, QStyle, qApp
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QTextCharFormat, QColor, QIcon, QPixmap, QPainter

class InlineAutocompleteTextEdit(QTextEdit):
    """Custom QTextEdit with inline autocomplete functionality"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.autocomplete_suggestion = ""
        self.autocomplete_start = 0
        self.autocomplete_end = 0
        self.is_autocomplete_visible = False
        
    def show_autocomplete(self, suggestion, start_pos, end_pos):
        """Show inline autocomplete suggestion"""
        self.autocomplete_suggestion = suggestion
        self.autocomplete_start = start_pos
        self.autocomplete_end = end_pos
        self.is_autocomplete_visible = True
        self.update()
    
    def hide_autocomplete(self):
        """Hide inline autocomplete"""
        self.is_autocomplete_visible = False
        self.update()
    
    def accept_autocomplete(self):
        """Accept the current autocomplete suggestion"""
        if not self.is_autocomplete_visible:
            return
        
        cursor = self.textCursor()
        cursor.setPosition(self.autocomplete_start)
        cursor.setPosition(self.autocomplete_end, cursor.KeepAnchor)
        cursor.insertText(self.autocomplete_suggestion)
        self.hide_autocomplete()
    
    def keyPressEvent(self, event):
        """Handle key press events"""
        if event.key() == Qt.Key_Tab and self.is_autocomplete_visible:
            self.accept_autocomplete()
            return
        elif event.key() == Qt.Key_Escape and self.is_autocomplete_visible:
            self.hide_autocomplete()
            return
        
        super().keyPressEvent(event)
    
    def paintEvent(self, event):
        """Custom paint event to show autocomplete suggestion"""
        super().paintEvent(event)
        
        if self.is_autocomplete_visible and self.autocomplete_suggestion:
            # This is a simplified approach - in a real implementation you'd
            # draw the grayed-out text at the cursor position
            pass

from .dialogs import HeaderPickerDialog, TemplateDialog, SettingsDialog, AboutDialog
from ..ai.ollama_client import OllamaClient
from ..tools.excel_handler import ExcelHandler
from ..tools.formula_validator import FormulaValidator
from ..config.settings import ConfigManager

class FormulaSparkMainWindow(QMainWindow):
    """Main application window for FormulaSpark"""
    
    def __init__(self):
        super().__init__()
        
        # Initialize components
        self.config_manager = ConfigManager()
        self.excel_handler = ExcelHandler()
        self.ollama_client = OllamaClient(self.config_manager)
        self.validator = FormulaValidator()
        
        # UI state
        self.generation_thread = None
        self.current_worker = None
        self.selected_headers_with_tags = {}
        self.header_picker_data = None
        
        self._integrate_methods()  # Integrate methods before UI so signals bind to real methods
        self.init_ui()
        self.check_ollama_connection()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle(f'FormulaSpark v{self.config_manager.get("APP_VERSION", "1.0.0")}')
        self.setGeometry(200, 200, 800, 900)
        self.setStyleSheet(self.get_stylesheet())
        
        # Set window icon
        self.setWindowIcon(self.create_lightning_icon())
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)
        
        # Create menu bar and status bar
        self.create_menu_bar()
        self.setStatusBar(QStatusBar())
        
        # Create splitter for better layout
        splitter = QSplitter(Qt.Vertical)
        
        # Top panel - Main controls
        top_panel = self.create_top_panel()
        splitter.addWidget(top_panel)
        
        # Bottom panel - History and templates
        bottom_panel = self.create_bottom_panel()
        splitter.addWidget(bottom_panel)
        
        # Set splitter proportions
        splitter.setSizes([500, 300])
        
        main_layout.addWidget(splitter)
        self.update_ui_state()
    
    def create_top_panel(self):
        """Create the top panel with main controls"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # Connection Section
        layout.addWidget(self.create_section_label("1. Connect to Excel"))
        self.file_path_display = QLineEdit("Not connected to any Excel file")
        self.file_path_display.setReadOnly(True)
        layout.addWidget(self.file_path_display)
        
        connect_layout = QHBoxLayout()
        self.connect_button = QPushButton("Connect to Active Workbook")
        self.connect_button.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.connect_button.clicked.connect(self.connect_to_excel)
        connect_layout.addWidget(self.connect_button)
        
        self.refresh_button = QPushButton("Refresh Connection")
        self.refresh_button.clicked.connect(self.check_ollama_connection)
        connect_layout.addWidget(self.refresh_button)
        layout.addLayout(connect_layout)
        
        layout.addSpacing(15)
        
        # Configuration Section
        layout.addWidget(self.create_section_label("2. Configure Generator"))
        config_layout = QFormLayout()
        config_layout.setSpacing(10)
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        
        self.model_combo = QComboBox()
        
        model_layout = QHBoxLayout()
        model_layout.addWidget(self.model_combo)
        self.status_indicator = QLabel("OFFLINE")
        self.status_indicator.setObjectName("StatusIndicator")
        self.status_indicator.setStyleSheet("color: red;")
        model_layout.addWidget(self.status_indicator)
        
        config_layout.addRow("Active Sheet:", self.sheet_combo)
        config_layout.addRow("Ollama Model:", model_layout)
        layout.addLayout(config_layout)
        
        # Context and validation options
        options_layout = QHBoxLayout()
        self.context_checkbox = QCheckBox("Analyze column headers for context")
        self.context_checkbox.setChecked(self.config_manager.get("use_context", True))
        options_layout.addWidget(self.context_checkbox)
        
        self.auto_validate_checkbox = QCheckBox("Auto-validate formulas")
        self.auto_validate_checkbox.setChecked(self.config_manager.get("auto_validate", True))
        options_layout.addWidget(self.auto_validate_checkbox)
        layout.addLayout(options_layout)
        
        # Header picker section
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("Selected Headers:"))
        self.header_picker_button = QPushButton("Pick Headers & Tags")
        self.header_picker_button.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        self.header_picker_button.clicked.connect(self.show_header_picker)
        header_layout.addWidget(self.header_picker_button)
        
        self.selected_headers_label = QLabel("No headers selected")
        self.selected_headers_label.setStyleSheet("color: #666; font-style: italic;")
        header_layout.addWidget(self.selected_headers_label)
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        layout.addSpacing(15)
        
        # Prompt and Generate Section
        layout.addWidget(self.create_section_label("3. Generate Formula"))
        
        # Template button
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("Quick Templates:"))
        self.template_button = QPushButton("Browse Templates")
        self.template_button.clicked.connect(self.show_templates)
        template_layout.addWidget(self.template_button)
        template_layout.addStretch()
        layout.addLayout(template_layout)
        
        # Create custom text edit with inline autocomplete
        self.prompt_input = InlineAutocompleteTextEdit()
        self.prompt_input.setPlaceholderText("e.g., Sum @BeginBalance where @PaymentDate is after 2020")
        self.prompt_input.setMinimumHeight(100)
        
        # Setup autocomplete for this text edit
        self.setup_text_edit_autocomplete()
        
        layout.addWidget(self.prompt_input, 1)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Generate button
        self.generate_button = QPushButton("Generate Formula")
        self.generate_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.generate_button.clicked.connect(self.generate_formula)
        layout.addWidget(self.generate_button)
        
        # Result display
        result_layout = QHBoxLayout()
        self.result_display = QLineEdit()
        self.result_display.setReadOnly(True)
        self.result_display.setPlaceholderText("Generated formula will appear here...")
        
        copy_button = QPushButton("Copy")
        copy_button.setIcon(self.style().standardIcon(QStyle.SP_FileDialogToParent))
        copy_button.clicked.connect(self.copy_to_clipboard)
        
        insert_button = QPushButton("Insert to New Sheet")
        insert_button.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        insert_button.clicked.connect(self.insert_into_cell)
        
        validate_button = QPushButton("Validate")
        validate_button.clicked.connect(self.validate_current_formula)
        
        result_layout.addWidget(self.result_display)
        result_layout.addWidget(copy_button)
        result_layout.addWidget(insert_button)
        result_layout.addWidget(validate_button)
        layout.addLayout(result_layout)
        
        return panel
    
    def create_bottom_panel(self):
        """Create the bottom panel with history"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # History Section
        history_header_layout = QHBoxLayout()
        history_header_layout.addWidget(self.create_section_label("Generation History"))
        history_header_layout.addStretch()
        
        clear_history_button = QPushButton("Clear History")
        clear_history_button.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        clear_history_button.setObjectName("ClearButton")
        clear_history_button.clicked.connect(self.clear_history)
        history_header_layout.addWidget(clear_history_button)
        layout.addLayout(history_header_layout)
        
        self.history_list = QListWidget()
        self.history_list.itemDoubleClicked.connect(self.reuse_history_item)
        self.populate_history()
        layout.addWidget(self.history_list, 1)
        
        return panel
    
    def create_section_label(self, text):
        """Create a section label"""
        label = QLabel(text)
        label.setObjectName("SectionLabel")
        return label
    
    def create_menu_bar(self):
        """Create the menu bar"""
        menu_bar = self.menuBar()
        
        # File menu
        file_menu = menu_bar.addMenu("&File")
        
        settings_action = QAction("&Settings", self)
        settings_action.triggered.connect(self.open_settings)
        file_menu.addAction(settings_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("&Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Tools menu
        tools_menu = menu_bar.addMenu("&Tools")
        
        clear_cache_action = QAction("Clear &Cache", self)
        clear_cache_action.triggered.connect(self.clear_cache)
        tools_menu.addAction(clear_cache_action)
        
        # Help menu
        help_menu = menu_bar.addMenu("&Help")
        about_action = QAction("&About", self)
        about_action.triggered.connect(self.open_about)
        help_menu.addAction(about_action)
    
    def get_stylesheet(self):
        """Get the application stylesheet"""
        return """
        QMainWindow, QDialog {
            background-color: #f7f7f7;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        QLabel#SectionLabel {
            font-size: 11pt;
            font-weight: bold;
            color: #333;
            margin-top: 10px;
            margin-bottom: 5px;
        }
        QPushButton {
            background-color: #0078d4;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            font-size: 10pt;
            min-height: 20px;
        }
        QPushButton#StopButton {
            background-color: #c42b1c;
        }
        QPushButton#StopButton:hover {
            background-color: #a32417;
        }
        QPushButton#ClearButton {
            background-color: #e81123;
            max-width: 100px;
        }
        QPushButton#ClearButton:hover {
            background-color: #a20b17;
        }
        QPushButton:hover {
            background-color: #005a9e;
        }
        QPushButton:disabled {
            background-color: #cccccc;
            color: #666666;
        }
        QLineEdit, QTextEdit, QComboBox, QListWidget, QDoubleSpinBox, QSpinBox {
            border: 1px solid #dcdcdc;
            border-radius: 4px;
            padding: 5px;
            background-color: #ffffff;
            font-size: 10pt;
        }
        QComboBox::drop-down {
            border: 0px;
        }
        QStatusBar {
            background-color: #0078d4;
            color: white;
        }
        QListWidget::item:hover {
            background-color: #e6f2fa;
        }
        QLabel#StatusIndicator {
            font-weight: bold;
        }
        QCheckBox {
            font-size: 10pt;
            spacing: 5px;
        }
        QProgressBar {
            border: 1px solid #dcdcdc;
            border-radius: 4px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #0078d4;
            border-radius: 3px;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #dcdcdc;
            border-radius: 4px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        """
    
    # Import and integrate methods
    def _integrate_methods(self):
        """Integrate methods from FormulaSparkMainWindowMethods"""
        from .main_window_methods import FormulaSparkMainWindowMethods
        methods = FormulaSparkMainWindowMethods(self)
        
        # Copy all methods from the methods class to this instance
        for method_name in dir(methods):
            if not method_name.startswith('_') and callable(getattr(methods, method_name)):
                setattr(self, method_name, getattr(methods, method_name))
    
    def setup_text_edit_autocomplete(self):
        """Setup autocomplete for the custom text edit"""
        # Connect signals
        self.prompt_input.textChanged.connect(self.on_text_changed)
        
        # Common Excel functions and keywords
        self.excel_functions = [
            "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN",
            "IF", "IFS", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
            "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP",
            "CONCATENATE", "TEXT", "LEFT", "RIGHT", "MID", "LEN",
            "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "TRIM",
            "UPPER", "LOWER", "PROPER", "VALUE", "DATE", "TIME",
            "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
            "NOW", "TODAY", "DATEDIF", "NETWORKDAYS", "WORKDAY",
            "ROUND", "ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR",
            "ABS", "SQRT", "POWER", "LOG", "EXP", "RAND", "RANDBETWEEN"
        ]
        
        # Common formula keywords
        self.formula_keywords = [
            "where", "and", "or", "not", "greater than", "less than",
            "equal to", "not equal to", "contains", "starts with", "ends with",
            "between", "in", "not in", "is empty", "is not empty",
            "sum", "count", "average", "maximum", "minimum", "total"
        ]
    
    def on_text_changed(self):
        """Handle text changes for inline autocomplete"""
        cursor = self.prompt_input.textCursor()
        cursor_pos = cursor.position()
        text = self.prompt_input.toPlainText()
        
        # Find the current word being typed
        word_start = text.rfind(' ', 0, cursor_pos)
        if word_start == -1:
            word_start = 0
        else:
            word_start += 1
        
        current_word = text[word_start:cursor_pos].strip()
        
        # Get suggestions
        suggestions = self.get_suggestions(current_word)
        
        if suggestions and len(current_word) >= 1:
            # Show inline autocomplete
            self.prompt_input.show_autocomplete(suggestions[0], word_start, cursor_pos)
        else:
            # Hide autocomplete
            self.prompt_input.hide_autocomplete()
    
    def get_suggestions(self, current_word):
        """Get suggestions for the current word"""
        suggestions = []
        current_word_lower = current_word.lower()
        
        # Get selected headers with tags
        headers_with_tags = self.get_headers_with_tags()
        
        # Add header tags
        for tag in headers_with_tags.keys():
            if current_word_lower in tag.lower():
                suggestions.append(tag)
        
        # Add Excel functions
        for func in self.excel_functions:
            if current_word_lower in func.lower():
                suggestions.append(func)
        
        # Add formula keywords
        for keyword in self.formula_keywords:
            if current_word_lower in keyword.lower():
                suggestions.append(keyword)
        
        return suggestions[:5]  # Limit to 5 suggestions
    
    def show_inline_autocomplete(self, suggestion, word_start, cursor_pos):
        """Show inline autocomplete suggestion"""
        if self.is_autocomplete_active:
            return  # Already showing autocomplete
        
        self.autocomplete_text = suggestion
        self.autocomplete_start_pos = word_start
        self.autocomplete_end_pos = cursor_pos
        self.is_autocomplete_active = True
        
        # Apply gray text formatting to show the suggestion
        self.apply_autocomplete_formatting()
    
    def hide_inline_autocomplete(self):
        """Hide inline autocomplete"""
        if not self.is_autocomplete_active:
            return
        
        self.is_autocomplete_active = False
        self.clear_autocomplete_formatting()
    
    def apply_autocomplete_formatting(self):
        """Apply gray text formatting to show autocomplete suggestion"""
        # This is a simplified approach - in a real implementation you'd use
        # QTextCharFormat to show grayed-out text
        pass
    
    def clear_autocomplete_formatting(self):
        """Clear autocomplete formatting"""
        pass
    
    def show_autocomplete(self, current_word, word_start, cursor_pos):
        """Show autocomplete suggestions"""
        print(f"DEBUG: show_autocomplete called with word: '{current_word}'")
        suggestions = []
        current_word_lower = current_word.lower()
        
        # Get selected headers with tags
        headers_with_tags = self.get_headers_with_tags()
        print(f"DEBUG: Headers with tags: {headers_with_tags}")
        
        # Add header tags
        for tag in headers_with_tags.keys():
            if current_word_lower in tag.lower():
                suggestions.append(("Header", tag, f"Use header: {tag}"))
                print(f"DEBUG: Added header suggestion: {tag}")
        
        # Add Excel functions
        for func in self.excel_functions:
            if current_word_lower in func.lower():
                suggestions.append(("Function", func, f"Excel function: {func}()"))
                print(f"DEBUG: Added function suggestion: {func}")
        
        # Add formula keywords
        for keyword in self.formula_keywords:
            if current_word_lower in keyword.lower():
                suggestions.append(("Keyword", keyword, f"Formula keyword: {keyword}"))
                print(f"DEBUG: Added keyword suggestion: {keyword}")
        
        # Limit suggestions
        suggestions = suggestions[:10]
        print(f"DEBUG: Total suggestions: {len(suggestions)}")
        
        if suggestions:
            print("DEBUG: Clearing and populating autocomplete popup")
            self.autocomplete_popup.clear()
            
            # Store suggestions for keyboard navigation
            self.current_suggestions = []
            self.selected_suggestion_index = 0
            
            for category, text, description in suggestions:
                item = QListWidgetItem(f"{text} - {description}")
                item.setData(Qt.UserRole, (text, word_start, cursor_pos))
                self.autocomplete_popup.addItem(item)
                self.current_suggestions.append((text, word_start, cursor_pos))
                print(f"DEBUG: Added item: {text} - {description}")
            
            # Position the popup
            cursor_rect = self.prompt_input.cursorRect()
            popup_pos = self.prompt_input.mapToGlobal(cursor_rect.bottomLeft())
            print(f"DEBUG: Moving popup to position: {popup_pos}")
            self.autocomplete_popup.move(popup_pos)
            
            # Select first item
            self.autocomplete_popup.setCurrentRow(0)
            
            print("DEBUG: Showing autocomplete popup")
            self.autocomplete_popup.show()
            print(f"DEBUG: Popup visible: {self.autocomplete_popup.isVisible()}")
        else:
            print("DEBUG: No suggestions, hiding popup")
            self.autocomplete_popup.hide()
            self.current_suggestions = []
    
    def eventFilter(self, obj, event):
        """Handle keyboard events for inline autocomplete"""
        if obj == self.prompt_input:
            if event.type() == event.KeyPress:
                if event.key() == Qt.Key_Tab and self.is_autocomplete_active:
                    # Accept current suggestion
                    self.accept_inline_suggestion()
                    return True
                elif event.key() == Qt.Key_Escape and self.is_autocomplete_active:
                    # Hide autocomplete
                    self.hide_inline_autocomplete()
                    return True
        return super().eventFilter(obj, event)
    
    def accept_inline_suggestion(self):
        """Accept the current inline suggestion"""
        if not self.is_autocomplete_active:
            return
        
        # Replace the current word with the suggestion
        cursor = self.prompt_input.textCursor()
        cursor.setPosition(self.autocomplete_start_pos)
        cursor.setPosition(self.autocomplete_end_pos, cursor.KeepAnchor)
        cursor.insertText(self.autocomplete_text)
        
        # Hide autocomplete
        self.hide_inline_autocomplete()
    
    def navigate_suggestions(self, direction):
        """Navigate through suggestions with arrow keys"""
        if not self.current_suggestions:
            return
        
        self.selected_suggestion_index += direction
        
        # Wrap around
        if self.selected_suggestion_index < 0:
            self.selected_suggestion_index = len(self.current_suggestions) - 1
        elif self.selected_suggestion_index >= len(self.current_suggestions):
            self.selected_suggestion_index = 0
        
        # Update selection
        self.autocomplete_popup.setCurrentRow(self.selected_suggestion_index)
        self.autocomplete_popup.scrollToItem(self.autocomplete_popup.currentItem())
    
    def accept_current_suggestion(self):
        """Accept the currently selected suggestion"""
        if not self.current_suggestions or self.selected_suggestion_index >= len(self.current_suggestions):
            return
        
        selected_suggestion = self.current_suggestions[self.selected_suggestion_index]
        text, word_start, cursor_pos = selected_suggestion
        
        # Replace the current word with the selected text
        cursor = self.prompt_input.textCursor()
        cursor.setPosition(word_start)
        cursor.setPosition(cursor_pos, cursor.KeepAnchor)
        cursor.insertText(text)
        
        # Hide the popup
        self.autocomplete_popup.hide()
        
        # Set focus back to the text input
        self.prompt_input.setFocus()
    
    def on_autocomplete_selected(self, item):
        """Handle autocomplete selection"""
        text, word_start, cursor_pos = item.data(Qt.UserRole)
        
        # Replace the current word with the selected text
        cursor = self.prompt_input.textCursor()
        cursor.setPosition(word_start)
        cursor.setPosition(cursor_pos, cursor.KeepAnchor)
        cursor.insertText(text)
        
        # Hide the popup
        self.autocomplete_popup.hide()
        
        # Set focus back to the text input
        self.prompt_input.setFocus()

    # Placeholder methods that will be replaced by _integrate_methods
    def check_ollama_connection(self): pass
    def connect_to_excel(self): pass
    def on_sheet_changed(self): pass
    def show_header_picker(self): pass
    def update_selected_headers_display(self): pass
    def get_headers_with_tags(self): return {}
    def show_templates(self): pass
    def validate_current_formula(self): pass
    def clear_cache(self): pass
    def update_ui_state(self, is_generating=False): pass
    def generate_formula(self): pass
    def on_generation_progress(self, message): pass
    def stop_generation(self): pass
    def on_generation_finished(self, formula): pass
    def on_generation_error(self, error_message): pass
    def copy_to_clipboard(self): pass
    def insert_into_cell(self): pass
    def populate_history(self): pass
    def clear_history(self): pass
    def reuse_history_item(self, item): pass
    def open_settings(self): pass
    def open_about(self): pass
    
    def create_lightning_icon(self):
        """Create a lightning bolt icon for the application"""
        # Create a 32x32 pixmap
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Set lightning bolt color (blue gradient)
        painter.setPen(QColor(102, 126, 234))  # #667eea
        painter.setBrush(QColor(102, 126, 234))
        
        # Draw lightning bolt shape
        # This is a simplified lightning bolt using lines
        points = [
            (16, 4),   # Top point
            (10, 16),  # Left middle
            (14, 16),  # Right middle
            (8, 28),   # Bottom left
            (22, 12),  # Right point
            (18, 12),  # Left point
            (24, 4)    # Top right
        ]
        
        # Draw the lightning bolt as a polygon
        from PyQt5.QtCore import QPoint
        polygon_points = [QPoint(x, y) for x, y in points]
        painter.drawPolygon(polygon_points)
        
        painter.end()
        
        return QIcon(pixmap)
