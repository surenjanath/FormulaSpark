"""
FormulaSpark Main Window Methods
Contains all the event handlers and business logic methods
"""

from PyQt5.QtWidgets import QMessageBox, QApplication, QStyle, qApp, QListWidgetItem
from PyQt5.QtCore import QThread, Qt
from .dialogs import HeaderPickerDialog, TemplateDialog, SettingsDialog, AboutDialog

class FormulaSparkMainWindowMethods:
    """Contains all the methods for the main window"""
    
    def __init__(self, main_window):
        self.main_window = main_window
    
    def check_ollama_connection(self):
        """Check Ollama connection and populate models"""
        try:
            is_online, status = self.main_window.ollama_client.check_connection()
            if is_online:
                self.main_window.status_indicator.setText("ONLINE")
                self.main_window.status_indicator.setStyleSheet("color: green;")
                models = self.main_window.ollama_client.get_available_models()
                self.main_window.model_combo.clear()
                self.main_window.model_combo.addItems(models)
            else:
                self.main_window.status_indicator.setText("OFFLINE")
                self.main_window.status_indicator.setStyleSheet("color: red;")
                self.main_window.model_combo.clear()
        except Exception as e:
            self.main_window.status_indicator.setText("ERROR")
            self.main_window.status_indicator.setStyleSheet("color: red;")
    
    def connect_to_excel(self):
        """Connect to active Excel workbook"""
        # Disable the connect button to prevent multiple attempts
        self.main_window.connect_button.setEnabled(False)
        self.main_window.connect_button.setText("Connecting...")
        self.main_window.statusBar().showMessage("Attempting to connect to Excel...")
        QApplication.processEvents()
        
        try:
            success, message, workbook = self.main_window.excel_handler.connect_to_active_workbook()
            
            if success:
                self.main_window.file_path_display.setText(workbook.name)
                sheet_names = self.main_window.excel_handler.get_sheet_names()
                self.main_window.sheet_combo.clear()
                self.main_window.sheet_combo.addItems(sheet_names)
                # Clear selected headers when connecting to new workbook
                self.main_window.selected_headers_with_tags = {}
                self.update_selected_headers_display()
                self.main_window.statusBar().showMessage(f"Successfully connected to {workbook.name}", 5000)
            else:
                self.main_window.file_path_display.setText("Not connected")
                self.main_window.sheet_combo.clear()
                self.main_window.statusBar().showMessage(message, 5000)
                QMessageBox.warning(self.main_window, "Connection Failed", 
                                  f"{message}\n\nPlease ensure Excel is running and a workbook is open.")
        except Exception as e:
            self.main_window.file_path_display.setText("Not connected")
            self.main_window.statusBar().showMessage(f"Connection error: {e}", 5000)
            QMessageBox.warning(self.main_window, "Connection Error", 
                              f"An error occurred while connecting to Excel: {e}")
        finally:
            # Re-enable the connect button
            self.main_window.connect_button.setEnabled(True)
            self.main_window.connect_button.setText("Connect to Active Workbook")
            self.update_ui_state()
    
    def on_sheet_changed(self):
        """Handle sheet selection change"""
        # Clear selected headers when sheet changes
        self.main_window.selected_headers_with_tags = {}
        self.update_selected_headers_display()
        self.main_window.statusBar().showMessage("Sheet changed - headers cleared", 2000)
    
    def show_header_picker(self):
        """Show the header picker dialog"""
        if not self.main_window.excel_handler.is_connected() or not self.main_window.sheet_combo.currentText():
            QMessageBox.warning(self.main_window, "No Sheet Selected", "Please connect to Excel and select a sheet first.")
            return
        
        try:
            headers = self.main_window.excel_handler.get_headers(self.main_window.sheet_combo.currentText())
            if not headers:
                QMessageBox.warning(self.main_window, "No Headers Found", "Could not find headers in the selected sheet.")
                return
            
            dialog = HeaderPickerDialog(headers, self.main_window, self.main_window.excel_handler, self.main_window.sheet_combo.currentText())
            if dialog.exec_():
                self.main_window.selected_headers_with_tags = dialog.get_selected_headers_with_tags()
                self.update_selected_headers_display()
                self.main_window.statusBar().showMessage(f"Selected {len(self.main_window.selected_headers_with_tags)} headers with tags", 3000)
        except Exception as e:
            QMessageBox.critical(self.main_window, "Error", f"Failed to read headers: {e}")
    
    def update_selected_headers_display(self):
        """Update the display of selected headers"""
        if not self.main_window.selected_headers_with_tags:
            self.main_window.selected_headers_label.setText("No headers selected")
            self.main_window.selected_headers_label.setStyleSheet("color: #666; font-style: italic;")
        else:
            tags = list(self.main_window.selected_headers_with_tags.values())
            if len(tags) <= 3:
                display_text = f"Selected: {', '.join(tags)}"
            else:
                display_text = f"Selected: {', '.join(tags[:3])}... (+{len(tags)-3} more)"
            
            self.main_window.selected_headers_label.setText(display_text)
            self.main_window.selected_headers_label.setStyleSheet("color: #0078d4; font-weight: bold;")
    
    def get_column_letter(self, column_number):
        """Convert column number to Excel column letter (1=A, 2=B, 27=AA, etc.)"""
        result = ""
        while column_number > 0:
            column_number -= 1
            result = chr(65 + (column_number % 26)) + result
            column_number //= 26
        return result
    
    def get_headers_with_tags(self):
        """Get headers with their corresponding column letters and tags"""
        if not self.main_window.selected_headers_with_tags:
            return {}
        
        try:
            # Check if we have header picker data with column info
            print(f"DEBUG: Checking header_picker_data: {hasattr(self.main_window, 'header_picker_data')}")
            print("=" * 80)
            print("DEBUG: ABOUT TO CHECK HEADER PICKER DATA")
            print("=" * 80)
            if hasattr(self.main_window, 'header_picker_data') and self.main_window.header_picker_data:
                print("=" * 80)
                print("DEBUG: HEADER PICKER DATA FOUND!")
                print(f"DEBUG: header_picker_data: {self.main_window.header_picker_data}")
                print(f"DEBUG: selected_headers_with_tags: {self.main_window.selected_headers_with_tags}")
                print(f"DEBUG: Number of picker headers: {len(self.main_window.header_picker_data)}")
                print(f"DEBUG: Number of selected headers: {len(self.main_window.selected_headers_with_tags)}")
                print("=" * 80)
                
                # Check if we have any matches
                for header_text, tag in self.main_window.selected_headers_with_tags.items():
                    print(f"DEBUG: Looking for '{header_text}' in picker data...")
                    found = False
                    for picker_header in self.main_window.header_picker_data:
                        if isinstance(picker_header, dict) and picker_header['text'] == header_text:
                            print(f"DEBUG: FOUND MATCH: '{header_text}' -> {picker_header['column']}")
                            found = True
                            break
                    if not found:
                        print(f"DEBUG: NO MATCH for '{header_text}'")
                
                result = {}
                for header_text, tag in self.main_window.selected_headers_with_tags.items():
                    print(f"DEBUG: Looking for header '{header_text}' in picker data")
                    # Find the header info from the picker data
                    for header_info in self.main_window.header_picker_data:
                        if isinstance(header_info, dict) and header_info['text'] == header_text:
                            print(f"DEBUG: Found match for '{header_text}' -> column {header_info['column']}")
                            result[tag] = {
                                'header': header_text,
                                'column': header_info['column'],
                                'range': f"{header_info['column']}:{header_info['column']}"
                            }
                            break
                    else:
                        print(f"DEBUG: No match found for '{header_text}'")
                
                print(f"DEBUG: Result from picker data: {result}")
                if result:
                    return result
                else:
                    print("DEBUG: No matches found, falling back to old method")
            
            # Fallback to old method if no picker data
            print("DEBUG: Using fallback method - fixing column mapping")
            print(f"DEBUG: selected_headers_with_tags: {self.main_window.selected_headers_with_tags}")
            sheet_name = self.main_window.sheet_combo.currentText()
            headers = self.main_window.excel_handler.get_headers(sheet_name)
            
            result = {}
            for i, header in enumerate(headers):
                if header in self.main_window.selected_headers_with_tags:
                    column_letter = self.get_column_letter(i + 1)  # Use proper Excel column mapping
                    tag = self.main_window.selected_headers_with_tags[header]
                    result[tag] = {
                        'header': header,
                        'column': column_letter,
                        'range': f"{column_letter}:{column_letter}"
                    }
            print(f"DEBUG: Result from fallback method: {result}")
            return result
        except Exception as e:
            print(f"Error getting headers with tags: {e}")
            return {}
    
    def show_templates(self):
        """Show formula templates dialog"""
        dialog = TemplateDialog(self.main_window)
        if dialog.exec_():
            name, template = dialog.selected_template
            self.main_window.prompt_input.setText(f"Use template: {name}")
            self.main_window.result_display.setText(template)
            self.main_window.statusBar().showMessage(f"Template '{name}' loaded", 3000)
    
    def validate_current_formula(self):
        """Validate the current formula"""
        formula = self.main_window.result_display.text()
        if not formula:
            QMessageBox.warning(self.main_window, "No Formula", "No formula to validate")
            return
        
        is_valid, error_msg = self.main_window.validator.validate_formula(formula)
        
        if is_valid:
            QMessageBox.information(self.main_window, "Validation", "Formula syntax is valid!")
        else:
            QMessageBox.warning(self.main_window, "Validation Error", f"Formula validation failed:\n{error_msg}")
    
    def clear_cache(self):
        """Clear the formula cache"""
        reply = QMessageBox.question(self.main_window, "Clear Cache", "Are you sure you want to clear the formula cache?", 
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.main_window.ollama_client.cache.cache.clear()
            self.main_window.ollama_client.cache.save_cache()
            self.main_window.statusBar().showMessage("Cache cleared", 3000)
    
    def update_ui_state(self, is_generating=False):
        """Update UI state based on current status"""
        connected = self.main_window.excel_handler.is_connected()
        self.main_window.sheet_combo.setEnabled(connected and not is_generating)
        self.main_window.model_combo.setEnabled(connected and not is_generating)
        self.main_window.prompt_input.setEnabled(connected and not is_generating)
        self.main_window.connect_button.setEnabled(not is_generating)
        self.main_window.refresh_button.setEnabled(not is_generating)
        self.main_window.context_checkbox.setEnabled(connected and not is_generating)
        self.main_window.auto_validate_checkbox.setEnabled(not is_generating)
        
        self.main_window.generate_button.setEnabled(connected)
        
        if is_generating:
            self.main_window.generate_button.setText("Stop Generation")
            self.main_window.generate_button.setObjectName("StopButton")
            self.main_window.generate_button.setIcon(self.main_window.style().standardIcon(QStyle.SP_MediaStop))
            try:
                self.main_window.generate_button.clicked.disconnect()
            except TypeError:
                pass
            self.main_window.generate_button.clicked.connect(self.stop_generation)
            self.main_window.progress_bar.setVisible(True)
        else:
            self.main_window.generate_button.setText("Generate Formula")
            self.main_window.generate_button.setObjectName("")
            self.main_window.generate_button.setIcon(self.main_window.style().standardIcon(QStyle.SP_MediaPlay))
            try:
                self.main_window.generate_button.clicked.disconnect()
            except TypeError:
                pass
            self.main_window.generate_button.clicked.connect(self.generate_formula)
            self.main_window.progress_bar.setVisible(False)
        
        # Re-apply stylesheet
        self.main_window.generate_button.style().unpolish(self.main_window.generate_button)
        self.main_window.generate_button.style().polish(self.main_window.generate_button)
    
    def generate_formula(self):
        """Generate formula using AI"""
        print("DEBUG: Starting formula generation")
        
        if not self.main_window.excel_handler.is_connected() or not self.main_window.sheet_combo.currentText() or not self.main_window.model_combo.currentText() or not self.main_window.prompt_input.toPlainText().strip():
            QMessageBox.warning(self.main_window, "Missing Information", "Please ensure you are connected to Excel and all fields are filled out.")
            return
        
        print("DEBUG: All checks passed, updating UI state")
        self.update_ui_state(is_generating=True)
        self.main_window.progress_bar.setValue(0)
        
        user_prompt = self.main_window.prompt_input.toPlainText().strip()
        sheet_name = self.main_window.sheet_combo.currentText()
        model = self.main_window.model_combo.currentText()
        
        print(f"DEBUG: User prompt: {user_prompt}")
        print(f"DEBUG: Sheet name: {sheet_name}")
        print(f"DEBUG: Model: {model}")
        
        # Check cache first
        if self.main_window.config_manager.get("cache_enabled", True):
            print("DEBUG: Checking cache")
            headers = self.main_window.excel_handler.get_headers(sheet_name)
            header_context = ", ".join([f"'{h}'" for h in headers])
            cached_formula = self.main_window.ollama_client.cache.get_cached_formula(user_prompt, header_context)
            if cached_formula:
                print("DEBUG: Found cached formula")
                self.main_window.result_display.setText(cached_formula)
                self.main_window.statusBar().showMessage("Formula loaded from cache!", 3000)
                self.update_ui_state(is_generating=False)
                return
        
        print("DEBUG: Creating worker for async generation")
        # Create worker for async generation
        headers = self.main_window.excel_handler.get_headers(sheet_name)
        tagged_headers = self.get_headers_with_tags()
        
        print(f"DEBUG: Headers: {headers}")
        print(f"DEBUG: Tagged headers: {tagged_headers}")
        
        try:
            print(f"DEBUG: About to create worker with:")
            print(f"DEBUG: - user_prompt: {user_prompt}")
            print(f"DEBUG: - sheet_name: {sheet_name}")
            print(f"DEBUG: - headers: {headers}")
            print(f"DEBUG: - tagged_headers: {tagged_headers}")
            print(f"DEBUG: - model: {model}")
            
            worker = self.main_window.ollama_client.create_worker(
                user_prompt, sheet_name, headers, tagged_headers, model
            )
            print("DEBUG: Worker created successfully")
        except Exception as e:
            print(f"DEBUG: Error creating worker: {e}")
            print(f"DEBUG: Error type: {type(e)}")
            import traceback
            print(f"DEBUG: Full traceback: {traceback.format_exc()}")
            QMessageBox.critical(self.main_window, "Error", f"Failed to create generation worker: {e}")
            self.update_ui_state(is_generating=False)
            return
        
        print("DEBUG: Setting up thread and connections")
        self.main_window.generation_thread = QThread()
        print(f"DEBUG: Thread created: {self.main_window.generation_thread}")
        
        # Store worker reference to prevent garbage collection
        self.main_window.current_worker = worker
        print("DEBUG: Worker reference stored")
        
        worker.moveToThread(self.main_window.generation_thread)
        print("DEBUG: Worker moved to thread")
        
        # Connect signals BEFORE starting the thread
        self.main_window.generation_thread.started.connect(worker.run)
        print("DEBUG: Connected started signal to worker.run")
        
        worker.finished.connect(self.on_generation_finished)
        worker.error.connect(self.on_generation_error)
        worker.progress.connect(self.on_generation_progress)
        print("DEBUG: Connected worker signals")
        
        worker.finished.connect(self.main_window.generation_thread.quit)
        worker.finished.connect(worker.deleteLater)
        self.main_window.generation_thread.finished.connect(self.main_window.generation_thread.deleteLater)
        print("DEBUG: Connected cleanup signals")
        
        print("DEBUG: Starting generation thread")
        self.main_window.generation_thread.start()
        print(f"DEBUG: Thread started, isRunning: {self.main_window.generation_thread.isRunning()}")
        
        # Force the event loop to process the started signal
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        print("DEBUG: Processed events to trigger started signal")
    
    def on_generation_progress(self, message):
        """Handle generation progress updates"""
        self.main_window.statusBar().showMessage(message, 1000)
        self.main_window.progress_bar.setValue(self.main_window.progress_bar.value() + 10)
    
    def stop_generation(self):
        """Stop formula generation"""
        if self.main_window.generation_thread and self.main_window.generation_thread.isRunning():
            self.main_window.generation_thread.requestInterruption()
            self.main_window.generation_thread.quit()
            self.main_window.generation_thread.wait()
            self.main_window.statusBar().showMessage("Generation stopped by user.", 3000)
        self.update_ui_state(is_generating=False)
    
    def on_generation_finished(self, formula):
        """Handle successful formula generation"""
        self.main_window.result_display.setText(formula)
        self.main_window.statusBar().showMessage("Formula generated successfully!", 3000)
        
        # Add to history
        prompt_text = self.main_window.prompt_input.toPlainText().strip()
        self.main_window.config_manager.add_history_entry(prompt_text, formula)
        self.main_window.config_manager.save_config()
        
        # Update history display
        history_item = QListWidgetItem(f"'{prompt_text}' -> {formula}")
        history_item.setData(Qt.UserRole, (prompt_text, formula))
        self.main_window.history_list.insertItem(0, history_item)
        
        # Auto-validate if enabled
        if self.main_window.config_manager.get("auto_validate", True):
            is_valid, error_msg = self.main_window.validator.validate_formula(formula)
            if not is_valid:
                QMessageBox.warning(self.main_window, "Formula Validation", f"Generated formula may have issues:\n{error_msg}")
        
        self.update_ui_state(is_generating=False)
    
    def on_generation_error(self, error_message):
        """Handle formula generation errors"""
        self.main_window.result_display.clear()
        QMessageBox.critical(self.main_window, "Generation Error", error_message)
        self.update_ui_state(is_generating=False)
    
    def copy_to_clipboard(self):
        """Copy formula to clipboard"""
        qApp.clipboard().setText(self.main_window.result_display.text())
        self.main_window.statusBar().showMessage("Formula copied to clipboard!", 2000)
    
    def insert_into_cell(self):
        """Insert formula into a new Excel sheet with headers"""
        formula = self.main_window.result_display.text()
        if not formula:
            self.main_window.statusBar().showMessage("No formula to insert.", 3000)
            return
        
        # Validate before insertion if enabled
        if self.main_window.config_manager.get("auto_validate", True):
            is_valid, error_msg = self.main_window.validator.validate_formula(formula)
            if not is_valid:
                reply = QMessageBox.question(self.main_window, "Validation Warning", 
                                           f"Formula validation failed:\n{error_msg}\n\nDo you want to insert anyway?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No:
                    return
        
        # Get current sheet name for reference
        current_sheet_name = None
        try:
            if hasattr(self.main_window, 'sheet_combo') and self.main_window.sheet_combo.currentText():
                current_sheet_name = self.main_window.sheet_combo.currentText()
        except:
            pass
        
        # Create new sheet with formula
        success, message, sheet_name = self.main_window.excel_handler.create_formula_sheet(formula, source_sheet_name=current_sheet_name)
        if success:
            self.main_window.statusBar().showMessage(f"Formula inserted into new sheet '{sheet_name}'", 5000)
        else:
            QMessageBox.critical(self.main_window, "Excel Error", f"Could not create formula sheet:\n{message}")
    
    def populate_history(self):
        """Populate the history list"""
        self.main_window.history_list.clear()
        history = self.main_window.config_manager.get("history", [])
        for prompt_text, formula in history:
            history_item = QListWidgetItem(f"'{prompt_text}' -> {formula}")
            history_item.setData(Qt.UserRole, (prompt_text, formula))
            self.main_window.history_list.addItem(history_item)
    
    def clear_history(self):
        """Clear the history"""
        reply = QMessageBox.question(self.main_window, "Clear History", "Are you sure you want to clear all history?", 
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.main_window.history_list.clear()
            self.main_window.config_manager.clear_history()
            self.main_window.config_manager.save_config()
            self.main_window.statusBar().showMessage("History cleared.", 3000)
    
    def reuse_history_item(self, item):
        """Reuse a history item"""
        prompt, formula = item.data(Qt.UserRole)
        self.main_window.prompt_input.setText(prompt)
        self.main_window.result_display.setText(formula)
    
    def open_settings(self):
        """Open settings dialog"""
        dialog = SettingsDialog(self.main_window.config_manager, self.main_window)
        if dialog.exec_():
            self.main_window.config_manager.update(dialog.get_settings())
            self.main_window.config_manager.save_config()
            self.main_window.statusBar().showMessage("Settings updated", 5000)
            self.check_ollama_connection()
    
    def open_about(self):
        """Open about dialog"""
        dialog = AboutDialog(self.main_window)
        dialog.exec_()
