"""
FormulaSpark UI Dialogs
Contains all dialog classes for settings, templates, header picker, etc.
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
    QPushButton, QDialogButtonBox, QComboBox, QCheckBox, QDoubleSpinBox,
    QSpinBox, QTextEdit, QListWidget, QListWidgetItem, QScrollArea,
    QWidget, QGridLayout, QGroupBox, QTabWidget, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter, QColor

from ..config.settings import FORMULA_TEMPLATES

class HeaderPickerDialog(QDialog):
    """Dialog for selecting and tagging Excel headers"""
    
    def __init__(self, headers, parent=None, excel_handler=None, sheet_name=None):
        super().__init__(parent)
        self.setWindowTitle("Select Headers for Formula Generation")
        self.setMinimumSize(700, 600)
        self.headers = headers
        self.selected_headers = []
        self.header_tags = {}
        self.excel_handler = excel_handler
        self.sheet_name = sheet_name
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Instructions
        instructions = QLabel(
            "Select the headers you want to use in your formula. " 
            "You can assign custom tags to make referencing easier."
        )
        instructions.setWordWrap(True)
        instructions.setStyleSheet("font-weight: bold; color: #333; margin-bottom: 10px;")
        layout.addWidget(instructions)
        
        # Excel selection section
        if self.excel_handler and self.sheet_name:
            excel_group = QGroupBox("Excel Selection")
            excel_layout = QVBoxLayout(excel_group)
            
            excel_instructions = QLabel(
                "1. In Excel, select the row that contains your headers (click and drag across the row)\n"
                "2. Click 'Use Selected Row as Headers' to set those as your header row\n"
                "3. The dialog will update with your selected headers"
            )
            excel_instructions.setWordWrap(True)
            excel_instructions.setStyleSheet("color: #666; margin-bottom: 10px;")
            excel_layout.addWidget(excel_instructions)
            
            excel_buttons = QHBoxLayout()
            self.use_selected_btn = QPushButton("Use Selected Row as Headers")
            self.use_selected_btn.clicked.connect(self.use_selected_row_as_headers)
            excel_buttons.addWidget(self.use_selected_btn)
            
            self.refresh_btn = QPushButton("Refresh from Excel Selection")
            self.refresh_btn.clicked.connect(self.refresh_from_excel_selection)
            excel_buttons.addWidget(self.refresh_btn)
            
            excel_buttons.addStretch()
            excel_layout.addLayout(excel_buttons)
            
            layout.addWidget(excel_group)
        
        # Headers selection area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMaximumHeight(300)
        
        headers_widget = QWidget()
        self.headers_layout = QGridLayout(headers_widget)
        
        self.header_checkboxes = {}
        self.tag_inputs = {}
        
        for i, header in enumerate(self.headers):
            # Create checkbox for header selection
            checkbox = QCheckBox(header)
            checkbox.stateChanged.connect(self.on_header_selection_changed)
            self.header_checkboxes[header] = checkbox
            
            # Create tag input
            tag_input = QLineEdit()
            tag_input.setPlaceholderText(f"Tag for {header}")
            tag_input.setText(self.generate_default_tag(header))
            tag_input.setMaximumWidth(150)
            self.tag_inputs[header] = tag_input
            
            # Add to layout
            self.headers_layout.addWidget(checkbox, i, 0)
            self.headers_layout.addWidget(QLabel("Tag:"), i, 1)
            self.headers_layout.addWidget(tag_input, i, 2)
        
        scroll_area.setWidget(headers_widget)
        layout.addWidget(scroll_area)
        
        # Preview section
        preview_group = QGroupBox("Preview - How to use in prompts:")
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_text = QTextEdit()
        self.preview_text.setMaximumHeight(100)
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText("Select headers to see how to reference them in your prompts...")
        preview_layout.addWidget(self.preview_text)
        
        layout.addWidget(preview_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all_headers)
        button_layout.addWidget(select_all_btn)
        
        clear_all_btn = QPushButton("Clear All")
        clear_all_btn.clicked.connect(self.clear_all_headers)
        button_layout.addWidget(clear_all_btn)
        
        button_layout.addStretch()
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_layout.addWidget(button_box)
        
        layout.addLayout(button_layout)
        
        # Connect signals
        for tag_input in self.tag_inputs.values():
            tag_input.textChanged.connect(self.update_preview)
    
    def generate_default_tag(self, header):
        """Generate a default tag from header name"""
        # Remove special characters and spaces, convert to camelCase
        tag = ''.join(word.capitalize() for word in header.replace(' ', '_').split('_'))
        # Remove common prefixes and make it shorter
        if tag.startswith('Beginning'):
            return f"@Begin{tag[9:]}"
        elif tag.startswith('Ending'):
            return f"@End{tag[6:]}"
        elif tag.startswith('Total'):
            return f"@Total{tag[5:]}"
        else:
            return f"@{tag[:10]}"  # Limit length
    
    def on_header_selection_changed(self):
        """Handle header selection changes"""
        self.update_preview()
    
    def select_all_headers(self):
        """Select all headers"""
        for checkbox in self.header_checkboxes.values():
            checkbox.setChecked(True)
        self.update_preview()
    
    def clear_all_headers(self):
        """Clear all header selections"""
        for checkbox in self.header_checkboxes.values():
            checkbox.setChecked(False)
        self.update_preview()
    
    def update_preview(self):
        """Update the preview text"""
        selected_headers = []
        for header, checkbox in self.header_checkboxes.items():
            if checkbox.isChecked():
                tag = self.tag_inputs[header].text().strip()
                if tag:
                    selected_headers.append(f"{tag} = {header}")
        
        if selected_headers:
            preview_text = "Selected headers:\n"
            preview_text += "\n".join(selected_headers)
            preview_text += "\n\nExample usage in prompts:\n"
            if len(selected_headers) >= 2:
                preview_text += f"• Sum {selected_headers[0].split(' = ')[0]} where {selected_headers[1].split(' = ')[0]} is greater than 0\n"
                preview_text += f"• Count rows where {selected_headers[0].split(' = ')[0]} contains 'Active'"
        else:
            preview_text = "No headers selected. Select headers to see usage examples."
        
        self.preview_text.setText(preview_text)
    
    def get_selected_headers_with_tags(self):
        """Get selected headers with their tags"""
        result = {}
        for header, checkbox in self.header_checkboxes.items():
            if checkbox.isChecked():
                tag = self.tag_inputs[header].text().strip()
                if tag:
                    # Use the header text from the checkbox (which includes column info)
                    checkbox_text = checkbox.text()
                    # Extract just the header name (before the column info)
                    if ' (' in checkbox_text:
                        header_name = checkbox_text.split(' (')[0]
                    else:
                        header_name = checkbox_text
                    result[header_name] = tag
        return result
    
    def use_selected_row_as_headers(self):
        """Use the currently selected row in Excel as headers"""
        print("DEBUG: Starting use_selected_row_as_headers")
        
        if not self.excel_handler or not self.sheet_name:
            print("DEBUG: Missing excel_handler or sheet_name")
            return
        
        try:
            print(f"DEBUG: excel_handler exists: {self.excel_handler is not None}")
            print(f"DEBUG: sheet_name: {self.sheet_name}")
            print(f"DEBUG: active_workbook: {self.excel_handler.active_workbook}")
            
            # Get the active sheet
            sheet = self.excel_handler.active_workbook.sheets[self.sheet_name]
            print(f"DEBUG: Got sheet: {sheet}")
            
            # Get the current selection
            print("DEBUG: Getting selection address...")
            selection_address = sheet.api.Application.Selection.Address
            print(f"DEBUG: Selection address: {selection_address}")
            
            selection = sheet.range(selection_address)
            print(f"DEBUG: Selection range: {selection}")
            print(f"DEBUG: Selection rows: {selection.rows.count}")
            print(f"DEBUG: Selection columns: {selection.columns.count}")
            
            # Check if it's a single row selection
            if selection.rows.count != 1:
                from PyQt5.QtWidgets import QMessageBox
                QMessageBox.warning(self, "Invalid Selection", 
                                  f"Please select a single row in Excel (click and drag across one row).\n"
                                  f"Current selection: {selection.rows.count} rows, {selection.columns.count} columns")
                return
            
            # Get the values from the selected row
            print("DEBUG: Getting row values...")
            row_values = selection.value
            print(f"DEBUG: Row values type: {type(row_values)}")
            print(f"DEBUG: Row values: {row_values}")
            
            # Get the actual column positions
            print("DEBUG: Getting column positions...")
            start_column = selection.column
            print(f"DEBUG: Start column: {start_column}")
            
            if isinstance(row_values, list):
                # Convert to strings and clean up, storing actual column positions
                new_headers = []
                for i, val in enumerate(row_values):
                    if val is not None:
                        # Calculate actual Excel column letter
                        actual_column = start_column + i
                        column_letter = self.get_column_letter(actual_column)
                        header_text = str(val)
                        new_headers.append({
                            'text': header_text,
                            'column': column_letter,
                            'column_number': actual_column
                        })
            else:
                # Single cell selected
                column_letter = self.get_column_letter(start_column)
                new_headers = [{
                    'text': str(row_values) if row_values is not None else "Column_1",
                    'column': column_letter,
                    'column_number': start_column
                }]
            
            print(f"DEBUG: New headers with column info: {new_headers}")
            
            # Update the headers list
            self.headers = new_headers
            
            # Store header data in main window for column mapping
            if hasattr(self, 'main_window') and self.main_window:
                self.main_window.header_picker_data = new_headers
                print(f"DEBUG: Stored header picker data: {new_headers}")
            
            # Clear existing UI elements
            for checkbox in self.header_checkboxes.values():
                checkbox.deleteLater()
            for tag_input in self.tag_inputs.values():
                tag_input.deleteLater()
            
            self.header_checkboxes = {}
            self.tag_inputs = {}
            
            # Rebuild the headers section
            print("DEBUG: Rebuilding headers section...")
            self.rebuild_headers_section()
            
            from PyQt5.QtWidgets import QMessageBox
            # Extract header texts for display
            header_texts = [h['text'] if isinstance(h, dict) else h for h in new_headers]
            QMessageBox.information(self, "Headers Updated", 
                                  f"Successfully set {len(new_headers)} headers from your selection:\n" + 
                                  ", ".join(header_texts[:5]) + ("..." if len(header_texts) > 5 else ""))
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Full error details: {error_details}")
            
            # Create a custom error dialog with selectable text
            self.show_detailed_error_dialog(e, error_details)
    
    def rebuild_headers_section(self):
        """Rebuild the headers section with new headers"""
        # Clear the existing layout
        for i in reversed(range(self.headers_layout.count())):
            self.headers_layout.itemAt(i).widget().setParent(None)
        
        # Add new headers
        for i, header_info in enumerate(self.headers):
            # Handle both old format (string) and new format (dict)
            if isinstance(header_info, dict):
                header_text = header_info['text']
                column_info = f" ({header_info['column']})"
            else:
                header_text = header_info
                column_info = ""
            
            # Create checkbox for header selection
            checkbox = QCheckBox(f"{header_text}{column_info}")
            checkbox.stateChanged.connect(self.on_header_selection_changed)
            self.header_checkboxes[header_text] = checkbox
            
            # Create tag input
            tag_input = QLineEdit()
            tag_input.setPlaceholderText(f"Tag for {header_text}")
            tag_input.setText(self.generate_default_tag(header_text))
            tag_input.setMaximumWidth(150)
            self.tag_inputs[header_text] = tag_input
            
            # Add to layout
            self.headers_layout.addWidget(checkbox, i, 0)
            self.headers_layout.addWidget(QLabel("Tag:"), i, 1)
            self.headers_layout.addWidget(tag_input, i, 2)
        
        # Connect signals for new tag inputs
        for tag_input in self.tag_inputs.values():
            tag_input.textChanged.connect(self.update_preview)
    
    def get_column_letter(self, column_number):
        """Convert column number to Excel column letter (1=A, 2=B, 27=AA, etc.)"""
        print(f"DEBUG: Converting column number {column_number} to letter")
        result = ""
        original_number = column_number
        while column_number > 0:
            column_number -= 1
            result = chr(65 + (column_number % 26)) + result
            column_number //= 26
        print(f"DEBUG: Column {original_number} -> '{result}'")
        return result
    
    def show_detailed_error_dialog(self, error, traceback_details):
        """Show a detailed error dialog with selectable text"""
        from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                                   QTextEdit, QPushButton, QDialogButtonBox)
        
        error_dialog = QDialog(self)
        error_dialog.setWindowTitle("Detailed Error Information")
        error_dialog.setMinimumSize(600, 400)
        
        layout = QVBoxLayout(error_dialog)
        
        # Error title
        title = QLabel("Error occurred while using selected row as headers:")
        title.setStyleSheet("font-weight: bold; color: #d32f2f; margin-bottom: 10px;")
        layout.addWidget(title)
        
        # Error message
        error_label = QLabel("Error:")
        error_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(error_label)
        
        error_text = QTextEdit()
        error_text.setPlainText(str(error))
        error_text.setMaximumHeight(60)
        error_text.setStyleSheet("background-color: #f5f5f5; border: 1px solid #ccc; padding: 5px;")
        error_text.setReadOnly(True)
        layout.addWidget(error_text)
        
        # Traceback section
        traceback_label = QLabel("Full Traceback (selectable):")
        traceback_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(traceback_label)
        
        traceback_text = QTextEdit()
        traceback_text.setPlainText(traceback_details)
        traceback_text.setStyleSheet("background-color: #f5f5f5; border: 1px solid #ccc; padding: 5px; font-family: monospace; font-size: 9pt;")
        traceback_text.setReadOnly(True)
        layout.addWidget(traceback_text)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(error_dialog.accept)
        layout.addWidget(button_box)
        
        error_dialog.exec_()
    
    def refresh_from_excel_selection(self):
        """Refresh the dialog based on Excel selection"""
        if not self.excel_handler or not self.sheet_name:
            return
        
        try:
            # Get the active sheet
            sheet = self.excel_handler.active_workbook.sheets[self.sheet_name]
            
            # Get the current selection using xlwings
            selection = sheet.range(sheet.api.Selection.Address)
            
            # Get the column indices of the selection
            selected_columns = []
            if selection.rows.count == 1:  # Single row selection
                for col in range(selection.columns.count):
                    col_index = selection.column + col - 1  # Convert to 0-based
                    if col_index < len(self.headers):
                        selected_columns.append(col_index)
            
            # Update checkboxes based on selection
            if selected_columns:
                # Clear all selections first
                for checkbox in self.header_checkboxes.values():
                    checkbox.setChecked(False)
                
                # Check the selected columns
                for col_index in selected_columns:
                    if col_index < len(self.headers):
                        header = self.headers[col_index]
                        if header in self.header_checkboxes:
                            self.header_checkboxes[header].setChecked(True)
                
                self.update_preview()
                
                from PyQt5.QtWidgets import QMessageBox
                QMessageBox.information(self, "Selection Updated", 
                                      f"Updated selection based on Excel selection: {len(selected_columns)} columns selected.")
            else:
                from PyQt5.QtWidgets import QMessageBox
                QMessageBox.information(self, "No Selection", 
                                      "Please select a single row in Excel (like the header row) and try again.")
            
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Error", f"Could not refresh from Excel selection: {e}")
    
    def update_header_selection_from_excel(self):
        """Update header selection based on Excel interaction"""
        # This is a placeholder - in a full implementation, you'd:
        # 1. Detect which cells are selected in Excel
        # 2. Map those cells to header names
        # 3. Update the checkboxes accordingly
        
        # For now, we'll just show a message
        from PyQt5.QtWidgets import QMessageBox
        QMessageBox.information(self, "Excel Integration", 
                              "Excel integration is working! The header row is highlighted. "
                              "You can now manually select the headers you want using the checkboxes below.")

class TemplateDialog(QDialog):
    """Dialog for selecting formula templates"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Formula Templates")
        self.setMinimumSize(500, 400)
        self.selected_template = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Template list
        self.template_list = QListWidget()
        for name, template in FORMULA_TEMPLATES.items():
            item = QListWidgetItem(f"{name}: {template}")
            item.setData(Qt.UserRole, (name, template))
            self.template_list.addItem(item)
        
        layout.addWidget(QLabel("Select a formula template:"))
        layout.addWidget(self.template_list)
        
        # Preview
        self.preview_label = QLabel("Preview: ")
        layout.addWidget(self.preview_label)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept_template)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        # Connect selection
        self.template_list.itemSelectionChanged.connect(self.update_preview)
    
    def update_preview(self):
        current_item = self.template_list.currentItem()
        if current_item:
            name, template = current_item.data(Qt.UserRole)
            self.preview_label.setText(f"Preview: {template}")
    
    def accept_template(self):
        current_item = self.template_list.currentItem()
        if current_item:
            self.selected_template = current_item.data(Qt.UserRole)
            self.accept()

class SettingsDialog(QDialog):
    """Enhanced settings dialog with tabs"""
    
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("FormulaSpark Settings")
        self.setMinimumWidth(500)
        self.config_manager = config_manager
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Create tab widget
        tab_widget = QTabWidget()
        
        # General Settings Tab
        general_tab = self.create_general_tab()
        tab_widget.addTab(general_tab, "General")
        
        # Ollama Settings Tab
        ollama_tab = self.create_ollama_tab()
        tab_widget.addTab(ollama_tab, "Ollama")
        
        # Advanced Settings Tab
        advanced_tab = self.create_advanced_tab()
        tab_widget.addTab(advanced_tab, "Advanced")
        
        layout.addWidget(tab_widget)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def create_general_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        
        self.context_checkbox = QCheckBox("Analyze column headers for context")
        self.context_checkbox.setChecked(self.config_manager.get("use_context", True))
        layout.addRow("Context Analysis:", self.context_checkbox)
        
        self.auto_validate_checkbox = QCheckBox("Validate formulas before insertion")
        self.auto_validate_checkbox.setChecked(self.config_manager.get("auto_validate", True))
        layout.addRow("Auto-validate:", self.auto_validate_checkbox)
        
        self.cache_enabled_checkbox = QCheckBox("Enable formula caching")
        self.cache_enabled_checkbox.setChecked(self.config_manager.get("cache_enabled", True))
        layout.addRow("Enable Cache:", self.cache_enabled_checkbox)
        
        return tab
    
    def create_ollama_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        
        self.ollama_url_input = QLineEdit(self.config_manager.get("ollama_base_url", "http://localhost:11434"))
        layout.addRow("Ollama Base URL:", self.ollama_url_input)
        
        self.temperature_input = QDoubleSpinBox()
        self.temperature_input.setRange(0.0, 2.0)
        self.temperature_input.setSingleStep(0.1)
        self.temperature_input.setValue(self.config_manager.get("temperature", 0.2))
        layout.addRow("Temperature:", self.temperature_input)
        
        self.top_p_input = QDoubleSpinBox()
        self.top_p_input.setRange(0.0, 1.0)
        self.top_p_input.setSingleStep(0.1)
        self.top_p_input.setValue(self.config_manager.get("top_p", 0.9))
        layout.addRow("Top P:", self.top_p_input)
        
        self.max_retries_input = QSpinBox()
        self.max_retries_input.setRange(1, 10)
        self.max_retries_input.setValue(self.config_manager.get("max_retries", 3))
        layout.addRow("Max Retries:", self.max_retries_input)
        
        return tab
    
    def create_advanced_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        
        self.history_limit_input = QSpinBox()
        self.history_limit_input.setRange(10, 10000)
        self.history_limit_input.setValue(self.config_manager.get("history_limit", 1000))
        layout.addRow("History Limit:", self.history_limit_input)
        
        self.timeout_input = QSpinBox()
        self.timeout_input.setRange(10, 300)
        self.timeout_input.setValue(self.config_manager.get("timeout", 90))
        layout.addRow("Request Timeout (s):", self.timeout_input)
        
        return tab
    
    def get_settings(self):
        return {
            "ollama_base_url": self.ollama_url_input.text().strip(),
            "temperature": self.temperature_input.value(),
            "top_p": self.top_p_input.value(),
            "max_retries": self.max_retries_input.value(),
            "use_context": self.context_checkbox.isChecked(),
            "auto_validate": self.auto_validate_checkbox.isChecked(),
            "cache_enabled": self.cache_enabled_checkbox.isChecked(),
            "history_limit": self.history_limit_input.value(),
            "timeout": self.timeout_input.value()
        }

class AboutDialog(QDialog):
    """Clean About dialog for FormulaSpark"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About FormulaSpark")
        self.setFixedSize(480, 380)
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                border-radius: 8px;
            }
        """)
        self.init_ui()
    
    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 10)
        main_layout.setSpacing(10)
        
        # Header with app info
        header_layout = QVBoxLayout()
        header_layout.setSpacing(4)
        
        # App icon and name
        title_layout = QHBoxLayout()
        title_layout.setSpacing(8)
        
        # Icon
        icon_label = QLabel()
        icon_label.setPixmap(self.create_lightning_icon().pixmap(24, 24))
        title_layout.addWidget(icon_label)
        
        # Title and version
        title_info = QVBoxLayout()
        title_info.setSpacing(2)
        
        title = QLabel("FormulaSpark")
        title.setFont(QFont('Segoe UI', 16, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        title_info.addWidget(title)
        
        version = QLabel("v1.0.0")
        version.setFont(QFont('Segoe UI', 9))
        version.setStyleSheet("color: #7f8c8d;")
        title_info.addWidget(version)
        
        title_layout.addLayout(title_info)
        title_layout.addStretch()
        
        header_layout.addLayout(title_layout)
        
        # Tagline
        tagline = QLabel("AI-Powered Excel Formula Generator")
        tagline.setFont(QFont('Segoe UI', 10))
        tagline.setStyleSheet("color: #5d6d7e;")
        header_layout.addWidget(tagline)
        
        main_layout.addLayout(header_layout)
        
        # Creator section
        creator_frame = QFrame()
        creator_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 6px;
                border: 1px solid #e9ecef;
            }
        """)
        creator_layout = QHBoxLayout(creator_frame)
        creator_layout.setContentsMargins(8, 6, 8, 6)
        creator_layout.setSpacing(8)
        
        # Avatar
        avatar = QLabel("👨‍💻")
        avatar.setFont(QFont('Segoe UI', 16))
        avatar.setFixedSize(36, 36)
        avatar.setStyleSheet("""
            QLabel {
                background-color: #e3f2fd;
                border-radius: 18px;
                border: 1px solid #bbdefb;
            }
        """)
        avatar.setAlignment(Qt.AlignCenter)
        creator_layout.addWidget(avatar)
        
        # Creator info
        creator_info = QVBoxLayout()
        creator_info.setSpacing(2)
        
        name = QLabel("Surenjanath Singh")
        name.setFont(QFont('Segoe UI', 10, QFont.Bold))
        name.setStyleSheet("color: #2c3e50;")
        creator_info.addWidget(name)
        
        title_text = QLabel("Data Solutions Engineer & Systems Architect")
        title_text.setFont(QFont('Segoe UI', 8))
        title_text.setStyleSheet("color: #7f8c8d;")
        creator_info.addWidget(title_text)
        
        # Email contact
        email_text = QLabel("surenjanath.singh@gmail.com")
        email_text.setFont(QFont('Segoe UI', 7))
        email_text.setStyleSheet("color: #5d6d7e; text-decoration: underline;")
        email_text.setCursor(Qt.PointingHandCursor)
        email_text.mousePressEvent = lambda e: self.open_email()
        creator_info.addWidget(email_text)
        
        creator_layout.addLayout(creator_info)
        creator_layout.addStretch()
        
        # Social media buttons
        social_layout = QVBoxLayout()
        social_layout.setSpacing(2)
        
        # First row of buttons
        social_row1 = QHBoxLayout()
        social_row1.setSpacing(6)
        
        # LinkedIn button
        linkedin_btn = QPushButton("Link")
        linkedin_btn.setFont(QFont('Segoe UI', 4, QFont.Bold))
        linkedin_btn.setFixedSize(60, 22)
        linkedin_btn.setStyleSheet("""
            QPushButton {
                background-color: #0077b5;
                color: white;
                border: none;
                border-radius: 11px;
                font-weight: bold;
                padding: 2px 6px;
            }
            QPushButton:hover {
                background-color: #005885;
            }
        """)
        linkedin_btn.clicked.connect(self.open_linkedin)
        social_row1.addWidget(linkedin_btn)
        
        # GitHub button
        github_btn = QPushButton("Git")
        github_btn.setFont(QFont('Segoe UI', 4, QFont.Bold))
        github_btn.setFixedSize(55, 22)
        github_btn.setStyleSheet("""
            QPushButton {
                background-color: #333;
                color: white;
                border: none;
                border-radius: 11px;
                font-weight: bold;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #555;
            }
        """)
        github_btn.clicked.connect(self.open_github)
        social_row1.addWidget(github_btn)
        
        # Second row of buttons
        social_row2 = QHBoxLayout()
        social_row2.setSpacing(6)
        
        # Medium button
        medium_btn = QPushButton("Med")
        medium_btn.setFont(QFont('Segoe UI', 4, QFont.Bold))
        medium_btn.setFixedSize(60, 22)
        medium_btn.setStyleSheet("""
            QPushButton {
                background-color: #00ab6c;
                color: white;
                border: none;
                border-radius: 11px;
                font-weight: bold;
                padding: 2px 6px;
            }
            QPushButton:hover {
                background-color: #008f5a;
            }
        """)
        medium_btn.clicked.connect(self.open_medium)
        social_row2.addWidget(medium_btn)
        
        # Fiverr button
        fiverr_btn = QPushButton("Fiv")
        fiverr_btn.setFont(QFont('Segoe UI', 4, QFont.Bold))
        fiverr_btn.setFixedSize(55, 22)
        fiverr_btn.setStyleSheet("""
            QPushButton {
                background-color: #1dbf73;
                color: white;
                border: none;
                border-radius: 11px;
                font-weight: bold;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #19a463;
            }
        """)
        fiverr_btn.clicked.connect(self.open_fiverr)
        social_row2.addWidget(fiverr_btn)
        
        social_layout.addLayout(social_row1)
        social_layout.addLayout(social_row2)
        creator_layout.addLayout(social_layout)
        
        main_layout.addWidget(creator_frame)
        
        # Features section
        features_label = QLabel("Key Features")
        features_label.setFont(QFont('Segoe UI', 10, QFont.Bold))
        features_label.setStyleSheet("color: #34495e; margin-bottom: 5px;")
        main_layout.addWidget(features_label)
        
        features_text = QLabel(
            "• Natural language to Excel formulas\n"
            "• Smart tag system for intuitive references\n"
            "• Context-aware AI with header analysis\n"
            "• Intelligent validation & caching\n"
            "• Complete privacy with local AI"
        )
        features_text.setFont(QFont('Segoe UI', 8))
        features_text.setStyleSheet("color: #5d6d7e; line-height: 1.4; margin-bottom: 8px;")
        features_text.setWordWrap(True)
        main_layout.addWidget(features_text)
        
        # Tech stack
        tech_label = QLabel("Built with Python • PyQt5 • Ollama • xlwings")
        tech_label.setFont(QFont('Segoe UI', 7))
        tech_label.setStyleSheet("color: #95a5a6; font-style: italic; margin-top: 3px; margin-bottom: 3px;")
        tech_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(tech_label)
        
        # Footer
        footer = QLabel("© 2025 FormulaSpark")
        footer.setFont(QFont('Segoe UI', 7))
        footer.setStyleSheet("color: #bdc3c7; margin-top: 2px; margin-bottom: 6px;")
        footer.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(footer)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.setFont(QFont('Segoe UI', 8, QFont.Bold))
        close_btn.setFixedSize(75, 26)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #667eea;
                color: white;
                border: none;
                border-radius: 13px;
                font-weight: bold;
                padding: 3px 12px;
            }
            QPushButton:hover {
                background-color: #5a6fd8;
            }
        """)
        close_btn.clicked.connect(self.accept)
        main_layout.addWidget(close_btn, alignment=Qt.AlignCenter)
    
    def open_linkedin(self):
        """Open LinkedIn profile in default browser"""
        import webbrowser
        linkedin_url = "https://www.linkedin.com/in/surenjanath"
        webbrowser.open(linkedin_url)
    
    def open_github(self):
        """Open GitHub profile in default browser"""
        import webbrowser
        github_url = "https://github.com/surenjanath"
        webbrowser.open(github_url)
    
    def open_medium(self):
        """Open Medium profile in default browser"""
        import webbrowser
        medium_url = "https://medium.com/@surenjanath"
        webbrowser.open(medium_url)
    
    def open_fiverr(self):
        """Open Fiverr profile in default browser"""
        import webbrowser
        fiverr_url = "https://www.fiverr.com/surenjanath"
        webbrowser.open(fiverr_url)
    
    def open_email(self):
        """Open default email client with pre-filled email, subject, and body"""
        import webbrowser
        import urllib.parse
        
        subject = "FormulaSpark Inquiry"
        body = """Hello Surenjanath,

I'm interested in learning more about FormulaSpark and your Excel automation services.

Please let me know more about:
- FormulaSpark features and capabilities
- Your Excel automation services
- Pricing and availability
- Any other relevant information

Thank you for your time!

Best regards,
[Your Name]"""
        
        # URL encode the subject and body
        subject_encoded = urllib.parse.quote(subject)
        body_encoded = urllib.parse.quote(body)
        
        # Create mailto URL with subject and body
        email_url = f"mailto:surenjanath.singh@gmail.com?subject={subject_encoded}&body={body_encoded}"
        webbrowser.open(email_url)
    
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
