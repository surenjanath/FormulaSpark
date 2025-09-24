"""
FormulaSpark Ollama AI Client
Handles all AI-related operations including API calls, retry logic, and caching
"""

import requests
import time
import hashlib
import json
from datetime import datetime
from typing import Optional, Dict, Any, Tuple
from PyQt5.QtCore import QObject, QThread, pyqtSignal

from ..config.settings import ConfigManager, PROMPT_TEMPLATE_SIMPLE, PROMPT_TEMPLATE_WITH_CONTEXT, PROMPT_TEMPLATE_WITH_TAGS

class FormulaCache:
    """Caches similar formula requests for performance"""
    
    def __init__(self, cache_file: str = "formula_cache.json"):
        self.cache_file = cache_file
        self.cache = self.load_cache()
    
    def load_cache(self) -> Dict:
        """Load cache from file"""
        try:
            with open(self.cache_file, 'r') as f:
                return json.load(f)
        except:
            return {}
    
    def save_cache(self):
        """Save cache to file"""
        try:
            with open(self.cache_file, 'w') as f:
                json.dump(self.cache, f, indent=2)
        except Exception as e:
            print(f"Failed to save cache: {e}")
    
    def get_cache_key(self, prompt: str, headers: str) -> str:
        """Generate cache key from prompt and context"""
        content = f"{prompt}:{headers}".lower().strip()
        return hashlib.md5(content.encode()).hexdigest()
    
    def get_cached_formula(self, prompt: str, headers: str) -> Optional[str]:
        """Get cached formula if available"""
        key = self.get_cache_key(prompt, headers)
        cached = self.cache.get(key)
        if cached and (datetime.now() - datetime.fromisoformat(cached['timestamp'])).days < 7:
            return cached['formula']
        return None
    
    def cache_formula(self, prompt: str, headers: str, formula: str):
        """Cache successful formula generation"""
        key = self.get_cache_key(prompt, headers)
        self.cache[key] = {
            'formula': formula,
            'timestamp': datetime.now().isoformat(),
            'usage_count': self.cache.get(key, {}).get('usage_count', 0) + 1
        }
        self.save_cache()

class RetryableOllamaWorker(QObject):
    """Enhanced worker with retry logic and better error handling"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    
    def __init__(self, url: str, payload: Dict[str, Any], max_retries: int = 3):
        super().__init__()
        self.url = url
        self.payload = payload
        self.max_retries = max_retries
        
    def run(self):
        """Execute the API call with retry logic"""
        print("=" * 50)
        print("DEBUG: WORKER RUN METHOD CALLED!")
        print("=" * 50)
        print(f"DEBUG: Worker run() started with URL: {self.url}")
        print(f"DEBUG: Payload: {self.payload}")
        
        for attempt in range(self.max_retries):
            try:
                if QThread.currentThread().isInterruptionRequested():
                    print("DEBUG: Thread interruption requested, stopping")
                    return
                    
                print(f"DEBUG: Attempt {attempt + 1}/{self.max_retries}")
                self.progress.emit(f"Attempt {attempt + 1}/{self.max_retries}")
                
                print("DEBUG: Making API request...")
                response = requests.post(self.url, json=self.payload, timeout=90)
                print(f"DEBUG: Response status: {response.status_code}")
                
                response.raise_for_status()
                
                print("DEBUG: Parsing response...")
                response_json = response.json()
                print(f"DEBUG: Response JSON: {response_json}")
                
                formula = response_json.get("response", "").strip().replace("`", "")
                print(f"DEBUG: Extracted formula: {formula}")
                
                if formula.lower().startswith("excel"):
                    formula = formula.splitlines()[0][5:].strip()
                    print(f"DEBUG: Cleaned formula: {formula}")
                
                if not QThread.currentThread().isInterruptionRequested():
                    print("DEBUG: Emitting finished signal")
                    self.finished.emit(formula)
                return
                
            except requests.exceptions.Timeout:
                print(f"DEBUG: Timeout on attempt {attempt + 1}")
                if attempt < self.max_retries - 1:
                    wait_time = 2 ** attempt
                    print(f"DEBUG: Retrying in {wait_time} seconds...")
                    self.progress.emit(f"Timeout, retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue
                else:
                    print("DEBUG: Max retries reached for timeout")
                    if not QThread.currentThread().isInterruptionRequested():
                        self.error.emit("Request timed out after multiple attempts")
            except requests.exceptions.RequestException as e:
                print(f"DEBUG: Request exception on attempt {attempt + 1}: {e}")
                if attempt < self.max_retries - 1:
                    wait_time = 2 ** attempt
                    print(f"DEBUG: Retrying in {wait_time} seconds...")
                    self.progress.emit(f"Connection error, retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue
                else:
                    print("DEBUG: Max retries reached for request exception")
                    if not QThread.currentThread().isInterruptionRequested():
                        self.error.emit(f"API Error: Could not connect after {self.max_retries} attempts. Details: {e}")
            except Exception as e:
                print(f"DEBUG: Unexpected error: {e}")
                if not QThread.currentThread().isInterruptionRequested():
                    self.error.emit(f"An unexpected error occurred: {e}")
                return

class OllamaClient:
    """Main Ollama AI client for formula generation"""
    
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        self.cache = FormulaCache()
    
    def check_connection(self) -> Tuple[bool, str]:
        """Check if Ollama is accessible"""
        try:
            response = requests.get(f"{self.config.get_ollama_url()}/api/tags", timeout=5)
            response.raise_for_status()
            return True, "ONLINE"
        except requests.exceptions.RequestException as e:
            return False, f"OFFLINE: {e}"
    
    def get_available_models(self) -> list:
        """Get list of available Ollama models"""
        try:
            response = requests.get(f"{self.config.get_ollama_url()}/api/tags", timeout=5)
            response.raise_for_status()
            models = [m['name'] for m in response.json().get('models', [])]
            return models
        except:
            return []
    
    def generate_formula_simple(self, prompt: str, sheet_name: str, model: str) -> str:
        """Generate formula using simple prompt template"""
        full_prompt = PROMPT_TEMPLATE_SIMPLE.format(
            sheet_name=sheet_name,
            user_prompt=prompt
        )
        return self._call_ollama_api(full_prompt, model)
    
    def generate_formula_with_context(self, prompt: str, sheet_name: str, headers: list, model: str) -> str:
        """Generate formula using context-aware prompt template"""
        header_context = ", ".join([f"'{h}'" for h in headers])
        full_prompt = PROMPT_TEMPLATE_WITH_CONTEXT.format(
            sheet_name=sheet_name,
            user_prompt=prompt,
            column_headers=header_context
        )
        return self._call_ollama_api(full_prompt, model)
    
    def generate_formula_with_tags(self, prompt: str, sheet_name: str, tagged_headers: Dict[str, Dict], model: str) -> str:
        """Generate formula using tagged headers prompt template"""
        tagged_headers_str = "\n".join([
            f"- {tag} ({info['header']}) = Column {info['column']} ({info['range']})"
            for tag, info in tagged_headers.items()
        ])
        
        full_prompt = PROMPT_TEMPLATE_WITH_TAGS.format(
            sheet_name=sheet_name,
            user_prompt=prompt,
            tagged_headers=tagged_headers_str
        )
        return self._call_ollama_api(full_prompt, model)
    
    def _call_ollama_api(self, full_prompt: str, model: str) -> str:
        """Make API call to Ollama"""
        # Check cache first
        if self.config.get("cache_enabled", True):
            cached_formula = self.cache.get_cached_formula(full_prompt, "")
            if cached_formula:
                return cached_formula
        
        # Make API call
        model_settings = self.config.get_model_settings()
        payload = {
            "model": model,
            "prompt": full_prompt,
            "stream": False,
            "options": {
                "temperature": model_settings["temperature"],
                "top_p": model_settings["top_p"]
            }
        }
        
        url = f"{self.config.get_ollama_url()}/api/generate"
        
        try:
            response = requests.post(url, json=payload, timeout=model_settings["timeout"])
            response.raise_for_status()
            
            response_json = response.json()
            formula = response_json.get("response", "").strip().replace("`", "")
            
            if formula.lower().startswith("excel"):
                formula = formula.splitlines()[0][5:].strip()
            
            # Cache the result
            if self.config.get("cache_enabled", True):
                self.cache.cache_formula(full_prompt, "", formula)
            
            return formula
            
        except Exception as e:
            raise Exception(f"Failed to generate formula: {e}")
    
    def create_worker(self, prompt: str, sheet_name: str, headers: list, tagged_headers: Dict[str, Dict], model: str) -> RetryableOllamaWorker:
        """Create a worker thread for async formula generation"""
        # Detect date columns for better date handling
        date_columns = {}
        try:
            from ..tools.excel_handler import ExcelHandler
            excel_handler = ExcelHandler()
            if excel_handler.is_connected():
                date_columns = excel_handler.detect_date_columns(sheet_name)
        except Exception as e:
            print(f"Error detecting date columns: {e}")
        
        # Determine which prompt template to use
        if tagged_headers:
            print(f"DEBUG: Processing tagged headers: {tagged_headers}")
            try:
                tagged_headers_str = "\n".join([
                    f"- {tag} ({info['header']}) = Column {info['column']} ({info['range']})"
                    for tag, info in tagged_headers.items()
                ])
            except (KeyError, TypeError) as e:
                print(f"DEBUG: Error processing tagged headers: {e}")
                print(f"DEBUG: Problematic entry type: {type(e)}")
                # Check each entry individually
                for tag, info in tagged_headers.items():
                    try:
                        print(f"DEBUG: Processing {tag}: {info}")
                        test_str = f"- {tag} ({info['header']}) = Column {info['column']} ({info['range']})"
                    except Exception as entry_error:
                        print(f"DEBUG: Error with entry {tag}: {entry_error}")
                        print(f"DEBUG: Entry data: {info}")
                        print(f"DEBUG: Entry keys: {list(info.keys()) if isinstance(info, dict) else 'Not a dict'}")
                
                # Fallback to safe access
                tagged_headers_str = "\n".join([
                    f"- {tag} ({info.get('header', 'Unknown')}) = Column {info.get('column', '?')} ({info.get('range', '?:?')})"
                    for tag, info in tagged_headers.items()
                ])
            
            # Add date column information if available
            if date_columns:
                date_info = "\n- **Date Columns Detected:** " + ", ".join([f"{col} ({date_columns[col]})" for col in date_columns.keys()])
                tagged_headers_str += date_info
            
            full_prompt = PROMPT_TEMPLATE_WITH_TAGS.format(
                sheet_name=sheet_name,
                user_prompt=prompt,
                tagged_headers=tagged_headers_str
            )
        elif headers:
            header_context = ", ".join([f"'{h}'" for h in headers])
            full_prompt = PROMPT_TEMPLATE_WITH_CONTEXT.format(
                sheet_name=sheet_name,
                user_prompt=prompt,
                column_headers=header_context
            )
        else:
            full_prompt = PROMPT_TEMPLATE_SIMPLE.format(
                sheet_name=sheet_name,
                user_prompt=prompt
            )
        
        model_settings = self.config.get_model_settings()
        payload = {
            "model": model,
            "prompt": full_prompt,
            "stream": False,
            "options": {
                "temperature": model_settings["temperature"],
                "top_p": model_settings["top_p"]
            }
        }
        
        url = f"{self.config.get_ollama_url()}/api/generate"
        
        return RetryableOllamaWorker(url, payload, model_settings["max_retries"])
