import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.relative_locator import locate_with
import pyperclip
from urllib.parse import urlparse
import uuid
import keyboard
import json
import os
import logging

# User Configuration Section
SELENIUM_DRIVER_PATH = r'C:\Peter\msedgedriver.exe'  # Update this path
DEFAULT_URL = "https://www.google.com"
OUTPUT_FILE_NAME = "web_automation_script.py"

class CombinedAutomationApp:
    def __init__(self, master):
        self.master = master
        master.title("Web Automation Tool v7.15")
        master.geometry("800x760")

        self.selenium_driver_path = SELENIUM_DRIVER_PATH  # Add this line
        
        self.setup_ui()
        self.selector_check.state(['selected'])
        self.current_window = None
        
        self.driver = None
        self.is_detecting = False
        self.current_selector = ""
        self.current_alternate_selector = ""
        self.actions = []

    def setup_ui(self):
        self.create_url_frame()
        self.create_webdriver_frame()  # New frame for webdriver path
        self.create_start_button()
        self.create_selector_frame()
        self.create_action_frame()
        self.create_action_list()
        self.create_control_buttons()
        self.create_window_selection_frame()

    def create_url_frame(self):
        self.url_frame = tk.Frame(self.master)
        self.url_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(self.url_frame, text="URL:").pack(side=tk.LEFT)

        self.url_entry = tk.Entry(self.url_frame, width=50)
        self.url_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))
        self.url_entry.insert(0, DEFAULT_URL)

    def create_window_selection_frame(self):
        self.window_frame = tk.Frame(self.master)
        self.window_frame.pack(pady=10, padx=10, fill=tk.X)
    
        tk.Label(self.window_frame, text="Active Window:").pack(side=tk.LEFT)
        self.window_dropdown = ttk.Combobox(self.window_frame, width=40)
        self.window_dropdown.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))
        self.window_dropdown.bind("<<ComboboxSelected>>", self.on_window_selected)
    
        self.refresh_button = tk.Button(self.window_frame, text="Refresh", command=self.refresh_windows)
        self.refresh_button.pack(side=tk.RIGHT)
    
    def refresh_windows(self):
        if self.driver:
            window_handles = self.driver.window_handles
            window_titles = []
            for handle in window_handles:
                self.driver.switch_to.window(handle)
                window_titles.append(self.driver.title)
            self.window_dropdown['values'] = window_titles
            if len(window_titles) == 1:
                self.window_dropdown.set(window_titles[0])
                self.current_window = window_handles[0]
            elif len(window_titles) > 1:
                self.window_dropdown.set("Select a window")
            self.master.focus_force()

    def on_window_selected(self, event):
        selected_title = self.window_dropdown.get()
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if self.driver.title == selected_title:
                self.current_window = handle
                break
        self.detect_css()

    def update_copy_target(self):
        # This method will be called when either checkbox is clicked
        # It ensures that only one checkbox is checked at a time
        if self.selector_check.instate(['selected']):
            self.alternate_selector_check.state(['!alternate'])
        elif self.alternate_selector_check.instate(['selected']):
            self.selector_check.state(['!alternate'])

    def create_webdriver_frame(self):
        self.webdriver_frame = tk.Frame(self.master)
        self.webdriver_frame.pack(pady=5, padx=10, fill=tk.X)

        tk.Label(self.webdriver_frame, text="WebDriver Path:").pack(side=tk.LEFT)

        self.webdriver_entry = tk.Entry(self.webdriver_frame, width=50)
        self.webdriver_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))
        self.webdriver_entry.insert(0, self.selenium_driver_path)

        self.browse_button = tk.Button(self.webdriver_frame, text="Browse", command=self.browse_webdriver)
        self.browse_button.pack(side=tk.RIGHT)

    def browse_webdriver(self):
        filename = filedialog.askopenfilename(
            title="Select WebDriver",
            filetypes=(("Executable files", "*.exe"), ("All files", "*.*"))
        )
        if filename:
            self.webdriver_entry.delete(0, tk.END)
            self.webdriver_entry.insert(0, filename)
            self.selenium_driver_path = filename

    def create_start_button(self):
        self.start_button = tk.Button(self.master, text="Start Detection", command=self.start_detection)
        self.start_button.pack(pady=10)

    def create_selector_frame(self):
        self.selector_frame = tk.Frame(self.master, bd=2)
        self.selector_frame.pack(fill=tk.X, padx=5, pady=5)

        # Primary Selector Row
        primary_frame = tk.Frame(self.selector_frame)
        primary_frame.pack(fill=tk.X, expand=True)

        tk.Label(primary_frame, text="Current CSS Selector:", anchor='w').pack(side=tk.LEFT, padx=5)

        self.selector_display = tk.Entry(primary_frame, width=50, bg='black', fg='white')
        self.selector_display.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.selector_check = ttk.Checkbutton(primary_frame, command=self.update_copy_target)
        self.selector_check.pack(side=tk.LEFT, padx=5)
        self.selector_check.state(['selected']) 

        # Alternate Selector Row
        alternate_frame = tk.Frame(self.selector_frame)
        alternate_frame.pack(fill=tk.X, expand=True, pady=(5, 0))

        tk.Label(alternate_frame, text="Alternate Selector:", anchor='w').pack(side=tk.LEFT, padx=5)

        self.alternate_selector_display = tk.Entry(alternate_frame, width=50, bg='black', fg='white')
        self.alternate_selector_display.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.alternate_selector_check = ttk.Checkbutton(alternate_frame, command=self.update_copy_target)
        self.alternate_selector_check.pack(side=tk.LEFT, padx=5)
        self.alternate_selector_check.state(['!alternate'])  # Unchecked by default

        # Copy Status Label
        self.copy_status_label = tk.Label(self.master, text="", fg="green")
        self.copy_status_label.pack(pady=5)

    def create_action_frame(self):
        self.action_frame = tk.Frame(self.master)
        self.action_frame.pack(pady=10, padx=10, fill=tk.X)
    
        self.action_types = ["Click", "Dropdown", "Input", "URL", "Sleep", "Keypress", "Relative Click", "Windows Selector"]
        self.special_keys = ["", "Enter", "Tab", "Shift", "Ctrl", "Alt", "Esc", "Backspace", "Delete", "PageUp", "PageDown", "Home", "End", "Insert", "Up", "Down", "Left", "Right", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]
        self.relative_directions = ["above", "below", "toLeftOf", "toRightOf", "near"]
    
        ttk.Label(self.action_frame, text="Action Type:").grid(row=0, column=0, padx=5, pady=5)
        self.action_type = ttk.Combobox(self.action_frame, values=self.action_types)
        self.action_type.grid(row=0, column=1, padx=5, pady=5)
        self.action_type.set(self.action_types[0])
        self.action_type.bind("<<ComboboxSelected>>", self.on_action_type_change)
    
        self.selector_label = ttk.Label(self.action_frame, text="Selector:")
        self.selector_label.grid(row=1, column=0, padx=5, pady=5)
        self.selector = ttk.Entry(self.action_frame, width=50)
        self.selector.grid(row=1, column=1, padx=5, pady=5)
    
        self.text_label = ttk.Label(self.action_frame, text="Text:")
        self.text_label.grid(row=2, column=0, padx=5, pady=5)
        self.text = ttk.Entry(self.action_frame, width=50)
        self.text.grid(row=2, column=1, padx=5, pady=5)
    
        self.special_key_label = ttk.Label(self.action_frame, text="Special Key:")
        self.special_key_label.grid(row=3, column=0, padx=5, pady=5)
        self.special_key = ttk.Combobox(self.action_frame, values=self.special_keys)
        self.special_key.grid(row=3, column=1, padx=5, pady=5)
        self.special_key.bind("<<ComboboxSelected>>", self.on_special_key_change)
    
        self.special_key_input = ttk.Entry(self.action_frame, width=5, state='disabled')
        self.special_key_input.grid(row=3, column=2, padx=5, pady=5)
    
        # New fields for Relative Click
        self.target_label = ttk.Label(self.action_frame, text="Target Selector:")
        self.target_label.grid(row=4, column=0, padx=5, pady=5)
        self.target = ttk.Entry(self.action_frame, width=50)
        self.target.grid(row=4, column=1, padx=5, pady=5)
    
        self.direction_label = ttk.Label(self.action_frame, text="Relative Direction:")
        self.direction_label.grid(row=5, column=0, padx=5, pady=5)
        self.direction = ttk.Combobox(self.action_frame, values=self.relative_directions)
        self.direction.grid(row=5, column=1, padx=5, pady=5)
    
        # Initially hide the new fields
        self.target_label.grid_remove()
        self.target.grid_remove()
        self.direction_label.grid_remove()
        self.direction.grid_remove()
    
        self.add_button = ttk.Button(self.action_frame, text="Add", command=self.add_action)
        self.add_button.grid(row=6, column=0, columnspan=2, pady=10)

    def create_action_list(self):
        self.action_list = tk.Listbox(self.master, width=70, height=10, selectmode=tk.SINGLE)
        self.action_list.pack(padx=5, pady=5)

    def create_control_buttons(self):
        self.remove_button = ttk.Button(self.master, text="Remove", command=self.remove_action)
        self.remove_button.pack(pady=10)

        ttk.Label(self.master, text="Insert Position:").pack()
        self.insert_position = ttk.Entry(self.master, width=10)
        self.insert_position.pack()

        self.complete_button = ttk.Button(self.master, text="Complete", command=self.complete)
        self.complete_button.pack(pady=10)

    def start_detection(self):
        if self.is_detecting:
            self.stop_detection()
            return
    
        url = self.url_entry.get().strip()
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
    
        try:
            result = urlparse(url)
            if all([result.scheme, result.netloc]):
                options = Options()
                options.add_argument("inprivate")
                options.add_argument("--log-level=3")
                options.add_experimental_option("detach", True)  # Keep browser open after script finishes
                service = Service(self.selenium_driver_path)
                self.driver = webdriver.Edge(service=service, options=options)
                self.driver.get(url)
                self.is_detecting = True
                self.start_button.config(text="Stop Detection")
                self.inject_mouse_move_script()
                self.detect_css_wrapper()  # Use the wrapper to keep detecting CSS
                keyboard.on_press_key('ctrl', self.copy_to_clipboard)
                # Automatically add URL action
                self.add_url_action(url)
                # Force focus back to the Tkinter window
                self.master.after(100, self.master.focus_force)
            else:
                tk.messagebox.showerror("Invalid URL", "Please enter a valid URL.")
        except WebDriverException as e:
            tk.messagebox.showerror("WebDriver Error", f"Error initializing WebDriver: {str(e)}\n\nPlease check the WebDriver path and try again.")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def add_url_action(self, url):
        action = {"type": "url", "url": url}
        self.actions.append(action)
        self.action_list.insert(tk.END, f"URL: {url}")

    def stop_detection(self):
        self.is_detecting = False
        self.start_button.config(text="Start Detection")
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.driver = None
        keyboard.unhook_all()
        self.selector_display.delete(0, tk.END)
        self.alternate_selector_display.delete(0, tk.END)
        self.copy_status_label.config(text="")

    def inject_mouse_move_script(self):
        script = """
        window.lastElement = null;
        window.currentSelector = '';
        window.currentAlternateSelector = '';
        
        function getFullSelector(element) {
            if (!(element instanceof Element)) return '';
            var path = [];
            while (element.nodeType === Node.ELEMENT_NODE) {
                var selector = element.nodeName.toLowerCase();
                if (element.id) {
                    selector += '#' + element.id;
                    path.unshift(selector);
                    break;
                } else {
                    var sibling = element;
                    var nth = 1;
                    while (sibling.previousElementSibling) {
                        sibling = sibling.previousElementSibling;
                        if (sibling.nodeName.toLowerCase() === selector)
                            nth++;
                    }
                    if (nth !== 1)
                        selector += ":nth-of-type("+nth+")";
                }
                path.unshift(selector);
                element = element.parentNode;
            }
            return path.join(' > ');
        }
    
        function getAlternateSelector(element) {
            if (!(element instanceof Element)) return '';
            if (element.id) return '#' + element.id;
            if (element.name) return element.tagName.toLowerCase() + '[name="' + element.name + '"]';
            
            var classes = Array.from(element.classList).join('.');
            if (classes) return element.tagName.toLowerCase() + '.' + classes;
            
            return getFullSelector(element);
        }
    
        document.addEventListener('mousemove', function(e) {
            var element = document.elementFromPoint(e.clientX, e.clientY);
            if (element !== window.lastElement) {
                window.lastElement = element;
                window.currentSelector = getFullSelector(element);
                window.currentAlternateSelector = getAlternateSelector(element);
            }
        });
        """
        self.driver.execute_script(script)

    def detect_css_wrapper(self):
        if self.is_detecting:
            self.detect_css()
            self.master.after(1000, self.detect_css_wrapper)

    def detect_css(self):
        if not self.is_detecting:
            return
    
        try:
            if self.current_window:
                self.driver.switch_to.window(self.current_window)   
            current_url = self.driver.current_url
            if current_url != self.url_entry.get():
                self.url_entry.delete(0, tk.END)
                self.url_entry.insert(0, current_url)
                self.inject_mouse_move_script()
    
            selector = self.driver.execute_script("return window.currentSelector || '';")
            alternate_selector = self.driver.execute_script("return window.currentAlternateSelector || '';")
            
            if selector:
                self.selector_display.delete(0, tk.END)
                self.selector_display.insert(0, selector)
                self.current_selector = selector
    
            if alternate_selector:
                self.alternate_selector_display.delete(0, tk.END)
                self.alternate_selector_display.insert(0, alternate_selector)
                self.current_alternate_selector = alternate_selector
    
        except NoSuchWindowException:
            self.stop_detection()
        except Exception as e:
            print(f"Error in detect_css: {e}")
    
        self.master.update_idletasks()

    def copy_to_clipboard(self, e):
        if self.selector_check.instate(['selected']):
            pyperclip.copy(self.current_selector)
            self.selector.delete(0, tk.END)
            self.selector.insert(0, self.current_selector)
            self.copy_status_label.config(text="Primary CSS selector copied to clipboard!")
        elif self.alternate_selector_check.instate(['selected']):
            pyperclip.copy(self.current_alternate_selector)
            self.selector.delete(0, tk.END)
            self.selector.insert(0, self.current_alternate_selector)
            self.copy_status_label.config(text="Alternate selector copied to clipboard!")
        else:
            self.copy_status_label.config(text="Please select a selector to copy!")

        self.master.after(2000, lambda: self.copy_status_label.config(text=""))

    def on_action_type_change(self, event):
        action_type = self.action_type.get()
        
        # Reset all fields and labels
        self.selector_label.config(text="Selector:")
        self.text_label.config(text="Text:")
        self.selector_label.grid()
        self.selector.grid()
        self.text_label.grid()
        self.text.grid()
        self.special_key_label.grid_remove()
        self.special_key.grid_remove()
        self.special_key_input.grid_remove()
        self.target_label.grid_remove()
        self.target.grid_remove()
        self.direction_label.grid_remove()
        self.direction.grid_remove()
        self.text.config(state='normal')
    
        if action_type == "URL":
            self.selector_label.grid_remove()
            self.selector.grid_remove()
            self.text_label.config(text="URL:")
        elif action_type == "Click":
            self.text_label.grid_remove()
            self.text.grid_remove()
        elif action_type == "Dropdown":
            self.text_label.config(text="Reference word:")
        elif action_type == "Sleep":
            self.selector_label.grid_remove()
            self.selector.grid_remove()
            self.text_label.config(text="Duration (ms):")
        elif action_type == "Keypress":
            self.selector_label.grid_remove()
            self.selector.grid_remove()
            self.text_label.grid_remove()
            self.text.grid_remove()
            self.special_key_label.grid()
            self.special_key.grid()
            self.special_key.set("")
        elif action_type == "Relative Click":
            self.selector_label.config(text="Anchor Selector (CSS):")
            self.text_label.grid_remove()
            self.text.grid_remove()
            self.target_label.grid()
            self.target.grid()
            self.direction_label.grid()
            self.direction.grid()
            self.direction.set("")  # Reset direction dropdown
        elif action_type == "Windows Selector":
            self.selector_label.grid_remove()
            self.selector.grid_remove()
            self.text_label.config(text="Window Title:")
        
        # Clear all input fields
        self.selector.delete(0, tk.END)
        self.text.delete(0, tk.END)
        self.special_key.set("")
        self.special_key_input.delete(0, tk.END)
        self.target.delete(0, tk.END)
        self.direction.set("")

    def on_special_key_change(self, event):
        special_key = self.special_key.get()
        if special_key in ["Shift", "Ctrl"]:
            self.special_key_input.config(state='normal')
            self.special_key_input.grid(row=3, column=2, padx=5, pady=5)
            self.text.config(state='disabled')
        elif special_key:
            self.special_key_input.grid_remove()
            self.text.config(state='disabled')
        else:
            self.special_key_input.grid_remove()
            self.text.config(state='normal')

    def add_action(self):
        action_type = self.action_type.get().lower()
        selector = self.selector.get()
        text = self.text.get()
        special_key = self.special_key.get()
        special_key_input = self.special_key_input.get()
        target = self.target.get()
        direction = self.direction.get()
    
        if action_type == "url":
            action = {"type": action_type, "url": text}
            display_text = f"URL: {text}"
        elif action_type == "click":
            action = {"type": action_type, "selector": selector}
            display_text = f"Click: {selector}"
        elif action_type == "dropdown":
            action = {"type": action_type, "selector": selector, "text": text}
            display_text = f"Dropdown: {selector} (Reference: {text})"
        elif action_type == "input":
            action = {"type": action_type, "selector": selector, "text": text}
            display_text = f"Input: {selector} (Text: {text})"
        elif action_type == "sleep":
            try:
                duration = int(text)
                action = {"type": action_type, "duration": duration}
                display_text = f"Sleep: {duration} ms"
            except ValueError:
                messagebox.showwarning("Invalid Input", "Please enter a valid integer for sleep duration.")
                return
        elif action_type == "keypress":
            key = special_key if special_key else text
            if special_key in ["Shift", "Ctrl"] and special_key_input:
                key = f"{special_key}+{special_key_input}"
            action = {"type": action_type, "key": key}
            display_text = f"Keypress: {key}"
        elif action_type == "relative click":
            if not all([selector, target, direction]):
                messagebox.showwarning("Missing Information", "Please fill in all fields for Relative Click.")
                return
            action = {
                "type": action_type,
                "anchor": selector,
                "target": target,
                "direction": direction
            }
            display_text = f"Relative Click: Anchor {selector}, Target {target}, Direction: {direction}"

        elif action_type == "windows selector":
            action = {"type": "switch_window", "window_name": text}
            display_text = f"Switch Window: {text}"
        else:
            messagebox.showwarning("Invalid Action", f"Unknown action type: {action_type}")
            return
    
        insert_position = self.insert_position.get()
        if insert_position:
            try:
                position = int(insert_position) - 1
                self.actions.insert(position, action)
                self.action_list.insert(position, f"{position + 1}. {display_text}")
                self.renumber_actions()
            except ValueError:
                messagebox.showwarning("Invalid Position", "Please enter a valid integer for the insert position.")
                return
        else:
            self.actions.append(action)
            self.action_list.insert(tk.END, f"{self.action_list.size() + 1}. {display_text}")
        
        self.clear_input_fields()

    def clear_input_fields(self):
        self.selector.delete(0, tk.END)
        self.text.delete(0, tk.END)
        self.special_key.set("")
        self.special_key_input.delete(0, tk.END)
        self.insert_position.delete(0, tk.END)

    def remove_action(self):
        try:
            index = self.action_list.curselection()[0]
            self.action_list.delete(index)
            self.actions.pop(index)
            self.renumber_actions()
        except IndexError:
            messagebox.showwarning("Warning", "No action selected!")

    def renumber_actions(self):
        for i in range(self.action_list.size()):
            item = self.action_list.get(i)
            new_item = f"{i + 1}. {item.split('. ', 1)[1]}"
            self.action_list.delete(i)
            self.action_list.insert(i, new_item)

    def complete(self):
        if not self.actions:
            messagebox.showwarning("Warning", "No actions added!")
            return
        
        code = self.generate_code()
        if code is None:
            return
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", OUTPUT_FILE_NAME)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(code)
        
        messagebox.showinfo("Success", f"Script saved to {file_path}")


    def generate_code(self):
        try:
            actions_json = json.dumps(self.actions, indent=4)
            script_id = uuid.uuid4().hex[:8]
            webdriver_path = self.webdriver_entry.get()
            code = f"""
import logging
import time
import traceback
import keyboard
import json
import os
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, WebDriverException, NoSuchElementException
from selenium.webdriver.support.relative_locator import locate_with
from typing import Callable, Dict, Any

# Configuration
SCRIPT_ID = "250527d9"
WEBDRIVER_PATH = r"{webdriver_path}"
WAIT_TIME = 5
RETRY_ATTEMPTS = 3
DEBUG = True

# Setup logging
DESKTOP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop')
LOG_FOLDER = os.path.join(DESKTOP_PATH, 'automation_log')
if not os.path.exists(LOG_FOLDER):
    os.makedirs(LOG_FOLDER)
LOG_FILE = os.path.join(LOG_FOLDER, f"web_automation_{{SCRIPT_ID}}.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=LOG_FILE,
    filemode='w'
)
logger = logging.getLogger(__name__)

class WebAutomation:
    def __init__(self, driver_path: str):
        self.driver_path = driver_path
        self.driver = self.setup_webdriver()

    def setup_webdriver(self):
        service = Service(executable_path=self.driver_path)
        edge_options = Options()
        edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        edge_options.add_experimental_option('useAutomationExtension', False)
        edge_options.add_argument("inprivate")
        edge_options.add_experimental_option("prefs", {{"profile.default_content_setting_values.notifications": 2}})
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {{get: () => undefined}})")
        return driver

    def wait_for_element(self, selector: str, condition: Callable, timeout: int = WAIT_TIME) -> Any:
        logger.debug(f"Waiting for element: {{selector}}")
        return WebDriverWait(self.driver, timeout).until(condition((By.CSS_SELECTOR, selector)))

    def scroll_into_view(self, selector: str, by: By = By.CSS_SELECTOR) -> None:
        logger.debug(f"Scrolling element into view: {{selector}}")
        element = self.driver.find_element(by, selector)
        self.driver.execute_script("arguments[0].scrollIntoView({{block: 'center'}});", element)
        time.sleep(0.5)  # Allow time for scrolling animation

    def retry_on_stale(func: Callable) -> Callable:
        def wrapper(*args, **kwargs):
            for attempt in range(RETRY_ATTEMPTS):
                try:
                    return func(*args, **kwargs)
                except StaleElementReferenceException:
                    logger.warning(f"Stale element reference exception. Retrying (attempt {{attempt + 1}}/{{RETRY_ATTEMPTS}})")
                    time.sleep(0.2)  # Add a small delay before retrying
                    if attempt == RETRY_ATTEMPTS - 1:
                        raise
        return wrapper

    @retry_on_stale
    def wait_and_click(self, action: Dict[str, Any]) -> None:
        selector = action["selector"]
        by = action.get("by", By.CSS_SELECTOR)
        timeout = action.get("timeout", WAIT_TIME)
        scroll = action.get("scroll", True)

        if scroll:
            self.scroll_into_view(selector, by)

        try:
            element = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, selector)))
            logger.info(f"Clicking element: {{selector}}")
            element.click()
        except TimeoutException:
            logger.error(f"Element not clickable: {{selector}}")
            raise

    @retry_on_stale
    def wait_and_input(self, selector: str, text: str) -> None:
        self.scroll_into_view(selector)
        element = self.wait_for_element(selector, EC.element_to_be_clickable)
        logger.info(f"Inputting text: {{text}}")
        element.clear()
        element.send_keys(text)

    @retry_on_stale
    def wait_and_select(self, selector: str, text: str) -> None:
        self.scroll_into_view(selector)
        dropdown = self.wait_for_element(selector, EC.element_to_be_clickable)
        
        try:
            select = Select(dropdown)
            logger.info(f"Selecting option: {{text}}")
            select.select_by_visible_text(text)
        except WebDriverException:
            # If it's not a <select> element, try to handle it as a custom dropdown
            logger.info(f"Handling as custom dropdown: {{selector}}")
            dropdown.click()
            time.sleep(0.3)  # Wait for dropdown to expand
            option = self.driver.find_element(By.XPATH, f"//*[contains(text(), '{{text}}')]")
            option.click()

    def wait_and_relative_click(self, anchor_selector: str, target_selector: str, direction: str) -> None:
        logger.debug(f"Waiting for anchor element: {{anchor_selector}}")
        anchor = self.wait_for_element(anchor_selector, EC.presence_of_element_located)
        logger.info(f"Clicking element relative to {{anchor_selector}}: {{direction}}")
        relative_locator = getattr(locate_with(By.CSS_SELECTOR, target_selector), direction)(anchor)
        relative_element = self.driver.find_element(relative_locator)
        relative_element.click()

    def switch_to_window(self, window_name: str) -> None:
        logger.info(f"Switching to window: {{window_name}}")
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if window_name in self.driver.title or window_name == handle:
                logger.info(f"Switched to window: {{self.driver.title}}")
                return
        logger.error(f"Window not found: {{window_name}}")
        raise NoSuchElementException(f"Window not found: {{window_name}}")

    def send_key(self, key: str) -> str:
        special_keys = {{
            'Enter': Keys.ENTER, 'Tab': Keys.TAB, 'Shift': Keys.SHIFT,
            'Ctrl': Keys.CONTROL, 'Alt': Keys.ALT, 'Esc': Keys.ESCAPE,
            'Backspace': Keys.BACKSPACE, 'Delete': Keys.DELETE,
            'PageUp': Keys.PAGE_UP, 'PageDown': Keys.PAGE_DOWN,
            'Up': Keys.ARROW_UP, 'Down': Keys.ARROW_DOWN,
            'Left': Keys.ARROW_LEFT, 'Right': Keys.ARROW_RIGHT,
            'Home': Keys.HOME, 'End': Keys.END, 'Insert': Keys.INSERT,
            **{{f'F{{i}}': getattr(Keys, f'F{{i}}') for i in range(1, 13)}}
        }}
        if '+' in key:
            modifier, letter = key.split('+')
            return getattr(Keys, modifier.upper()) + letter.lower()
        return special_keys.get(key, key)

    def perform_action(self, action: Dict[str, Any]) -> None:
            action_type = action["type"]
            logger.info(f"Performing action: {{action_type}}")
            try:
                if action_type == "url":
                    self.driver.get(action["url"])
                    WebDriverWait(self.driver, WAIT_TIME).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                elif action_type == "dropdown":
                    self.wait_and_select(action["selector"], action["text"])
                elif action_type == "click":
                    self.wait_and_click(action)
                    # Add the code to print window handles and titles here
                    print("Available window handles:")
                    for handle in self.driver.window_handles:
                        self.driver.switch_to.window(handle)
                        print(f"Handle: {{handle}}, Title: {{self.driver.title}}")
                elif action_type == "input":
                    self.wait_and_input(action["selector"], action["text"])
                elif action_type == "sleep":
                    duration = action["duration"] / 1000
                    logger.info(f"Sleeping for {{duration}} seconds")
                    time.sleep(duration)
                elif action_type == "keypress":
                    key = self.send_key(action["key"])
                    logger.info(f"Pressing key: {{action['key']}}")
                    active_element = self.driver.switch_to.active_element
                    active_element.send_keys(key)
                elif action_type == "relative click":
                    self.wait_and_relative_click(action["anchor"], action["target"], action["direction"])
                elif action_type == "switch_window":
                    window_name = action["window_name"]
                    self.switch_to_window(window_name)
                else:
                    logger.warning(f"Unknown action type: {{action_type}}")
                time.sleep(0.1)  # Small delay between actions
            except Exception as e:
                logger.error(f"Error performing action {{action_type}}: {{e}}")
                raise
    
def show_error_popup(error_message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    with open(LOG_FILE, 'r') as log_file:
        lines = log_file.readlines()
    error_text = f"Error: {{error_message}}"
    messagebox.showerror("Error", error_text)
    root.destroy()

def main():
    automation = None
    try:
        automation = WebAutomation(WEBDRIVER_PATH)

        actions = {json.dumps(self.actions, indent=4)}

        for action in actions:
            automation.perform_action(action)
        logger.info("Script execution completed successfully.")
    except Exception as e:
        logger.error(f"An error occurred: {{{{e}}}}")
        if DEBUG:
            logger.error(traceback.format_exc())
        show_error_popup(str(e))
    finally:
        if automation and automation.driver:
            print("Press 'F10' to close the browser...")
            while True:
                if keyboard.is_pressed('F10'):
                    break
                time.sleep(0.2)
            automation.driver.quit()
        logger.info(f"Log file saved as {{{{LOG_FILE}}}}")

if __name__ == "__main__":
    main()
    """
            return code
        except Exception as e:
            messagebox.showerror("Error Generating Code", f"An error occurred while generating the code:\n\n{str(e)}")
            return None

    def on_closing(self):
        self.stop_detection()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedAutomationApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
