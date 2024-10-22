import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.relative_locator import locate_with
import pygetwindow as gw
import pyperclip
from urllib.parse import urlparse
import uuid
import keyboard
import json
import os
import logging
import ast
import win32gui
import win32con

# User Configuration Section
SELENIUM_DRIVER_PATH = r'D:\Edge Extension\Seleium WebDriver\msedgedriver.exe'
DEFAULT_URL = "https://www.google.com"
OUTPUT_FILE_NAME = "web_automation_script.txt"

class CombinedAutomationApp:
    def __init__(self, master):
        self.master = master
        master.title("Web Automation Tool v1.7.86")
        master.geometry("500x1050")

        self.canvas = tk.Canvas(master)
        self.scrollbar = tk.Scrollbar(master, orient="vertical", command=self.canvas.yview)
        self.master = tk.Frame(self.canvas)
        self.master.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
    
        self.canvas.create_window((0, 0), window=self.master, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.selenium_driver_path = SELENIUM_DRIVER_PATH
        self.enable_visualization = tk.BooleanVar()  # Initialize here
        self.insert_position = None
        self.setup_ui()
        self.selector_check.state(['selected'])
        self.driver = None
        self.is_detecting = False
        self.current_selector = ""
        self.current_alternate_selector = ""
        self.loop_start = None
        self.loop_end = None
        self.actions = []
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.style = ttk.Style()
        self.style.configure("TFrame", background="white")
        self.style.configure("TLabel", background="white")
        self.style.configure("TButton", background="white")
        self.master.configure(background="white")
        self.canvas.configure(background="white")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def setup_ui(self):
        self.create_url_frame()
        self.create_webdriver_frame()
        self.create_start_button()
        self.create_selector_frame()
        self.create_action_frame()
        self.create_loop_frame()
        self.create_action_list()
        self.create_control_buttons()
        self.create_window_selection_frame()
        self.create_visualization_checkbox()
        self.on_action_type_change(None)
        self.create_import_button()

    def create_url_frame(self):
        self.url_frame = tk.Frame(self.master)
        self.url_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Label(self.url_frame, text="URL:").pack(side=tk.LEFT)
        self.url_entry = tk.Entry(self.url_frame, width=50)
        self.url_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))
        self.url_entry.insert(0, DEFAULT_URL)

    def create_import_button(self):
        self.import_button = ttk.Button(self.master, text="Import", command=self.import_actions)
        self.import_button.pack(pady=10)

    def import_actions(self):
        file_path = filedialog.askopenfilename(
            title="Select Actions File",
            filetypes=(("All supported files", "*.txt *.py *.json"), ("All files", "*.*"))
        )
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    file_content = file.read()
                    imported_actions = self.extract_actions(file_content)
                
                if not imported_actions:
                    raise ValueError("No valid actions found in the file.")

                # Clear existing actions
                self.actions.clear()
                self.action_list.delete(0, tk.END)
                
                # Add imported actions
                for action in imported_actions:
                    self.actions.append(action)
                    display_text = self.get_display_text(action)
                    self.action_list.insert(tk.END, display_text)
                
                self.renumber_actions()
                messagebox.showinfo("Import Successful", f"Imported {len(imported_actions)} actions.")
            except Exception as e:
                messagebox.showerror("Import Error", f"An error occurred while importing actions:\n\n{str(e)}")

    def extract_actions(self, content):
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            return self.extract_actions_from_python(content)

    def extract_actions_from_python(self, content):
        actions = []
        in_actions_block = False
        action_block = []
        
        for line in content.split('\n'):
            line = line.strip()
            if line.startswith('actions = ['):
                in_actions_block = True
                continue
            elif in_actions_block and line.startswith(']'):
                in_actions_block = False
                break
            elif in_actions_block:
                action_block.append(line)
    
        if action_block:
            actions_str = '[' + ''.join(action_block) + ']'
            try:
                actions = ast.literal_eval(actions_str)
            except (ValueError, SyntaxError):
                raise ValueError("Unable to parse actions from the Python file.")
    
        return actions

    def get_display_text(self, action):
        action_type = action.get('type', '').lower()
        if action_type == "url":
            return f"URL: {action.get('url', '')}"
        elif action_type == "click":
            return f"Click: {action.get('name', '')}"
        elif action_type == "dropdown":
            return f"Dropdown: {action.get('name', '')} (Reference: {action.get('text', '')})"
        elif action_type == "input":
            return f"Input: {action.get('name', '')} (Text: {action.get('text', '')})"
        elif action_type == "sleep":
            return f"Sleep: {action.get('duration', '')} ms"
        elif action_type == "keypress":
            return f"Keypress: {action.get('key', '')}"
        elif action_type == "relative click":
            return f"Relative Click: Anchor {action.get('anchor', '')}, Target {action.get('target', '')}, Direction: {action.get('direction', '')}"
        elif action_type == "switch_window":
            return f"Switch Window: {action.get('window_name', '')}"
        elif action_type == "ask and input":
            return f"Ask and input: {action.get('name', '')}"
        elif action_type == "mouseover":
            return f"MouseOver: {action.get('name', '')}"
        elif action_type == "file dialog":
            return f"File Dialog: {action.get('dialog_title', '')} (File: {action.get('file_path', '')}, Keys: {action.get('key_sequence', '')}, Text: {action.get('additional_text', '')})"
        else:
            return f"Unknown action: {action_type}"

    def create_webdriver_frame(self):
        self.webdriver_frame = tk.Frame(self.master)
        self.webdriver_frame.pack(pady=5, padx=10, fill=tk.X)
        tk.Label(self.webdriver_frame, text="WebDriver Path:").pack(side=tk.LEFT)
        self.webdriver_entry = tk.Entry(self.webdriver_frame, width=50)
        self.webdriver_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))
        self.webdriver_entry.insert(0, self.selenium_driver_path)
        self.browse_button = tk.Button(self.webdriver_frame, text="Browse", command=self.browse_webdriver)
        self.browse_button.pack(side=tk.RIGHT)

    def create_loop_frame(self):
        self.loop_frame = tk.Frame(self.master)
        self.loop_frame.pack(pady=10, padx=10, fill=tk.X)
        ttk.Label(self.loop_frame, text="Loop Control:").grid(row=0, column=0, padx=5, pady=5)    
        self.loop_start_button = ttk.Button(self.loop_frame, text="Start Loop", command=self.add_loop_start)
        self.loop_start_button.grid(row=0, column=1, padx=5, pady=5)
        
        self.loop_end_button = ttk.Button(self.loop_frame, text="End Loop", command=self.add_loop_end)
        self.loop_end_button.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(self.loop_frame, text="Repeat:").grid(row=0, column=3, padx=5, pady=5)
        self.loop_repeat = ttk.Entry(self.loop_frame, width=5)
        self.loop_repeat.grid(row=0, column=4, padx=5, pady=5)

    def add_loop_start(self):
        insert_position = self.get_insert_position()
        self.loop_start = insert_position
        self.action_list.insert(insert_position, f"{insert_position + 1}. --- Loop Start ---")
        self.renumber_actions()

    def add_loop_end(self):
        if self.loop_start is None:
            messagebox.showwarning("Warning", "Please add a loop start first!")
            return
        insert_position = self.get_insert_position()
        if insert_position <= self.loop_start:
            messagebox.showwarning("Warning", "Loop end must be after loop start!")
            return
        self.loop_end = insert_position
        repeat = self.loop_repeat.get() or "1"
        self.action_list.insert(insert_position, f"{insert_position + 1}. --- Loop End (Repeat: {repeat}) ---")
        self.renumber_actions()
        self.actions.insert(self.loop_start, {"type": "loop_start"})
        self.actions.insert(self.loop_end + 1, {"type": "loop_end", "repeat": int(repeat)})
        self.loop_start = None
        self.loop_end = None
        self.loop_repeat.delete(0, tk.END)


    def get_insert_position(self):
        insert_position = self.insert_position.get()
        if insert_position:
            try:
                return int(insert_position) - 1
            except ValueError:
                messagebox.showwarning("Invalid Position", "Please enter a valid integer for the insert position.")
                return self.action_list.size()
        return self.action_list.size()

    def create_window_selection_frame(self):
        self.window_frame = tk.Frame(self.master)
        self.window_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Label(self.window_frame, text="Active Window:").pack(side=tk.LEFT)
        self.window_dropdown = ttk.Combobox(self.window_frame, width=40)
        self.window_dropdown.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))
        self.window_dropdown.bind("<<ComboboxSelected>>", self.on_window_selected)
        self.refresh_button = tk.Button(self.window_frame, text="Refresh", command=self.refresh_windows)
        self.refresh_button.pack(side=tk.RIGHT)

    def create_visualization_checkbox(self):
        self.visualization_check = ttk.Checkbutton(
            self.master, 
            text="Enable Action Visualization", 
            variable=self.enable_visualization
        )
        self.visualization_check.pack(pady=5)
    
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
        if self.selector_var.get():
            self.alternate_selector_var.set(False)
        else:
            self.alternate_selector_var.set(True)
        if not self.selector_var.get() and not self.alternate_selector_var.get():
            self.selector_var.set(True)
        self.selector_check.state(['selected' if self.selector_var.get() else '!selected'])
        self.alternate_selector_check.state(['selected' if self.alternate_selector_var.get() else '!selected'])
        self.update_selector_backgrounds()

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
        self.button_frame = tk.Frame(self.master)
        self.button_frame.pack(pady=10)
        self.start_button = tk.Button(self.button_frame, text="Start Detection", command=self.start_detection)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.import_button = tk.Button(self.button_frame, text="Import", command=self.import_actions)
        self.import_button.pack(side=tk.LEFT, padx=5)

    def create_selector_frame(self):
        self.selector_frame = tk.Frame(self.master, bd=2)
        self.selector_frame.pack(fill=tk.X, padx=5, pady=5)
    
        # Primary Selector Row
        primary_frame = tk.Frame(self.selector_frame)
        primary_frame.pack(fill=tk.X, expand=True)
    
        tk.Label(primary_frame, text="Current CSS Selector:", anchor='w').pack(side=tk.LEFT, padx=5)
        self.selector_display = tk.Entry(primary_frame, width=50, bg='black', fg='white')
        self.selector_display.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
    
        self.selector_var = tk.BooleanVar(value=True)
        self.selector_check = ttk.Checkbutton(primary_frame, variable=self.selector_var, command=self.update_copy_target)
        self.selector_check.pack(side=tk.LEFT, padx=5)
        self.selector_check.state(['!alternate'])
    
        # Alternate Selector Row
        alternate_frame = tk.Frame(self.selector_frame)
        alternate_frame.pack(fill=tk.X, expand=True, pady=(5, 0))
    
        tk.Label(alternate_frame, text="Alternate Selector:", anchor='w').pack(side=tk.LEFT, padx=5)
        self.alternate_selector_display = tk.Entry(alternate_frame, width=50, bg='black', fg='white')
        self.alternate_selector_display.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
    
        self.alternate_selector_var = tk.BooleanVar(value=False)
        self.alternate_selector_check = ttk.Checkbutton(alternate_frame, variable=self.alternate_selector_var, command=self.update_copy_target)
        self.alternate_selector_check.pack(side=tk.LEFT, padx=5)
        self.alternate_selector_check.state(['!alternate']) 
    
        # Copy Status Label
        self.copy_status_label = tk.Label(self.master, text="", fg="green")
        self.copy_status_label.pack(pady=5)
    
        self.update_selector_backgrounds()
    
    def update_selector_backgrounds(self):
        active_color = 'dark green'
        inactive_color = 'black'
    
        self.selector_display.config(bg=active_color if self.selector_var.get() else inactive_color)
        self.alternate_selector_display.config(bg=active_color if self.alternate_selector_var.get() else inactive_color)
    
    def create_action_frame(self):
        self.action_frame = tk.Frame(self.master)
        self.action_frame.pack(pady=10, padx=10, fill=tk.X)
    
        self.action_types = [
            "Click", "Dropdown", "Input", "ask and input", "Windows Selector",
            "Keypress", "Sleep", "Relative Click", "URL", "File Dialog", "mouseover"
        ]
        self.special_keys = [
            "", "Enter", "Tab", "Shift", "Ctrl", "Alt", "Esc", "Backspace", "Delete",
            "PageUp", "PageDown", "Home", "End", "Insert", "Up", "Down", "Left", "Right",
            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"
        ]
        self.relative_directions = ["above", "below", "toLeftOf", "toRightOf", "near"]
    
        ttk.Label(self.action_frame, text="Action Type:").grid(row=0, column=0, padx=5, pady=5)
        self.action_type = ttk.Combobox(self.action_frame, values=self.action_types)
        self.action_type.grid(row=0, column=1, padx=5, pady=5)
        self.action_type.set("Click")
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
    
        self.target_label = ttk.Label(self.action_frame, text="Target Selector:")
        self.target_label.grid(row=4, column=0, padx=5, pady=5)
        self.target = ttk.Entry(self.action_frame, width=50)
        self.target.grid(row=4, column=1, padx=5, pady=5)
    
        self.direction_label = ttk.Label(self.action_frame, text="Relative Direction:")
        self.direction_label.grid(row=5, column=0, padx=5, pady=5)
        self.direction = ttk.Combobox(self.action_frame, values=self.relative_directions)
        self.direction.grid(row=5, column=1, padx=5, pady=5)
    
        self.file_path_entry = ttk.Entry(self.action_frame, width=40)
        self.file_path_entry.grid(row=2, column=1, padx=5, pady=5)
        self.file_path_entry.grid_remove()
    
        self.browse_button = ttk.Button(self.action_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=2, column=2, padx=5, pady=5)
        self.browse_button.grid_remove()
    
        # Add new fields for file dialog
        self.key_sequence_label = ttk.Label(self.action_frame, text="Key Sequence:")
        self.key_sequence_label.grid(row=6, column=0, padx=5, pady=5)
        self.key_sequence_entry = ttk.Entry(self.action_frame, width=50)
        self.key_sequence_entry.grid(row=6, column=1, padx=5, pady=5)
    
        self.wait_time_label = ttk.Label(self.action_frame, text="Wait Time (s):")
        self.wait_time_label.grid(row=7, column=0, padx=5, pady=5)
        self.wait_time_entry = ttk.Entry(self.action_frame, width=10)
        self.wait_time_entry.grid(row=7, column=1, padx=5, pady=5)
        self.wait_time_entry.insert(0, "1")
    
        self.file_dialog_key_label = ttk.Label(self.action_frame, text="Special Key:")
        self.file_dialog_key_label.grid(row=9, column=0, padx=5, pady=5)
        self.file_dialog_key = ttk.Combobox(self.action_frame, values=self.special_keys)
        self.file_dialog_key.grid(row=9, column=1, padx=5, pady=5)
        self.file_dialog_key.bind("<<ComboboxSelected>>", self.on_file_dialog_key_change)
        self.file_dialog_key_input = ttk.Entry(self.action_frame, width=5)
        self.file_dialog_key_input.grid(row=9, column=2, padx=5, pady=5)
    
        self.file_dialog_text_label = ttk.Label(self.action_frame, text="Additional Text:")
        self.file_dialog_text_label.grid(row=10, column=0, padx=5, pady=5)
        self.file_dialog_text = ttk.Entry(self.action_frame, width=50)
        self.file_dialog_text.grid(row=10, column=1, columnspan=2, padx=5, pady=5)

        self.add_button = ttk.Button(self.action_frame, text="Add", command=self.add_action)
        self.add_button.grid(row=11, column=0, columnspan=2, pady=10)
    
        # Initially hide the file dialog specific fields
        self.file_dialog_key_label.grid_remove()
        self.file_dialog_key.grid_remove()
        self.file_dialog_key_input.grid_remove()
        self.file_dialog_text_label.grid_remove()
        self.file_dialog_text.grid_remove()
        self.key_sequence_label.grid_remove()
        self.key_sequence_entry.grid_remove()
        self.wait_time_label.grid_remove()
        self.wait_time_entry.grid_remove()
    
        self.on_action_type_change(None)

    def on_file_dialog_key_change(self, event):
        special_key = self.file_dialog_key.get()
        if special_key in ["Shift", "Ctrl", "Alt"]:
            self.file_dialog_key_input.config(state='normal')
        else:
            self.file_dialog_key_input.config(state='disabled')
            self.file_dialog_key_input.delete(0, tk.END)

    def create_action_list(self):
        self.action_list = tk.Listbox(self.master, width=70, height=10, selectmode=tk.SINGLE)
        self.action_list.pack(padx=5, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)

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
                keyboard.on_press_key('`', self.copy_to_clipboard)
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
        # Clear all fields except for File Dialog's dialog title
        self.clear_input_fields()
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
        self.file_path_entry.grid_remove()
        self.browse_button.grid_remove()
        self.key_sequence_label.grid_remove()
        self.key_sequence_entry.grid_remove()
        self.wait_time_label.grid_remove()
        self.wait_time_entry.grid_remove()
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
        elif action_type == "ask and input":
            self.text_label.grid_remove()
            self.text.grid_remove()
        elif action_type == "mouseover":
            self.text_label.grid_remove()
            self.text.grid_remove()
        if action_type == "File Dialog":
            self.selector_label.config(text="Dialog Title:")
            self.text_label.config(text="File Path:")
            self.file_path_entry.grid()
            self.browse_button.grid()
            self.wait_time_label.grid()
            self.wait_time_entry.grid()
            self.file_dialog_key_label.grid()
            self.file_dialog_key.grid()
            self.file_dialog_key_input.grid()
            self.file_dialog_text_label.grid()
            self.file_dialog_text.grid()
        else:
            self.selector.delete(0, tk.END)
            self.file_dialog_key_label.grid_remove()
            self.file_dialog_key.grid_remove()
            self.file_dialog_key_input.grid_remove()
            self.file_dialog_text_label.grid_remove()
            self.file_dialog_text.grid_remove()
        
        # Clear all input fields
        self.selector.delete(0, tk.END)
        self.text.delete(0, tk.END)
        self.special_key.set("")
        self.special_key_input.delete(0, tk.END)
        self.target.delete(0, tk.END)
        self.direction.set("")
        self.file_path_entry.delete(0, tk.END)
        self.key_sequence_entry.delete(0, tk.END)
        self.wait_time_entry.delete(0, tk.END)
        self.wait_time_entry.insert(0, "0.1")  # Reset to default wait time

    def on_special_key_change(self, event):
        special_key = self.special_key.get()
        if special_key in ["Shift", "Ctrl", "Alt"]:
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
        special_key = self.special_key.get() if hasattr(self, 'special_key') else ""
        special_key_input = self.special_key_input.get() if hasattr(self, 'special_key_input') else ""
    
        action = {}
        display_text = ""
    
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
            if special_key in ["Shift", "Ctrl", "Alt"] and special_key_input:
                key = f"{special_key}+{special_key_input}"
            action = {"type": action_type, "key": key}
            display_text = f"Keypress: {key}"
        elif action_type == "relative click":
            if not all([selector, text, special_key]):  # Assuming 'text' is used for target and 'special_key' for direction
                messagebox.showwarning("Missing Information", "Please fill in all fields for Relative Click.")
                return
            action = {
                "type": action_type,
                "anchor": selector,
                "target": text,
                "direction": special_key
            }
            display_text = f"Relative Click: Anchor {selector}, Target {text}, Direction: {special_key}"
        elif action_type == "windows selector":
            action = {"type": "switch_window", "window_name": text}
            display_text = f"Switch Window: {text}"
        elif action_type == "ask and input":
            action = {"type": action_type, "selector": selector, "prompt": "Enter value:"}
            display_text = f"Ask and input: {selector}"
        elif action_type == "mouseover":
            action = {"type": action_type, "selector": selector}
            display_text = f"MouseOver: {selector}"
        elif action_type == "file dialog":
            special_key = self.file_dialog_key.get()
            special_key_input = self.file_dialog_key_input.get()
            additional_text = self.file_dialog_text.get()
    
            if special_key in ["Shift", "Ctrl", "Alt"] and special_key_input:
                key_sequence = f"{special_key}+{special_key_input}"
            elif special_key:
                key_sequence = special_key
            else:
                key_sequence = ""
    
            action = {
                "type": action_type,
                "dialog_title": selector,
                "file_path": self.file_path_entry.get(),
                "key_sequence": key_sequence,
                "additional_text": additional_text,
                "wait_time": float(self.wait_time_entry.get())
            }
            display_text = f"File Dialog: {selector} (File: {self.file_path_entry.get()}, Keys: {key_sequence}, Text: {additional_text})"
        
        if action and display_text:
            insert_position = self.get_insert_position()
            if insert_position is not None:
                self.actions.insert(insert_position, action)
                self.action_list.insert(insert_position, display_text)
            else:
                self.actions.append(action)
                self.action_list.insert(tk.END, display_text)
            self.renumber_actions()
            self.clear_input_fields()

    def clear_input_fields(self):
        current_action_type = self.action_type.get()
        
        if current_action_type != "File Dialog":
            self.selector.delete(0, tk.END)
        
        self.text.delete(0, tk.END)
        self.special_key.set("")
        self.special_key_input.delete(0, tk.END)      
        self.file_path_entry.delete(0, tk.END)
        self.file_dialog_key.set("")
        self.file_dialog_key_input.delete(0, tk.END)
        self.file_dialog_text.delete(0, tk.END)
        self.wait_time_entry.delete(0, tk.END)
        self.wait_time_entry.insert(0, "0.1")  # Reset to default wait time

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
            if item[0].isdigit() and '. ' in item:
                item_text = item.split('. ', 1)[-1]
            else:
                item_text = item
            new_item = f"{i + 1}. {item_text}"
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
            renamed_actions = []
            for action in self.actions:
                renamed_action = action.copy()
                if 'selector' in renamed_action:
                    renamed_action['name'] = renamed_action.pop('selector')
                renamed_actions.append(renamed_action)
            actions_json = json.dumps(renamed_actions, indent=4)
            script_id = uuid.uuid4().hex[:8]
            webdriver_path = self.webdriver_entry.get()
            enable_visualization = self.enable_visualization.get()
############################################
            imports = """
import logging
import time
import traceback
import keyboard
import json
import os
import tkinter as tk
from tkinter import messagebox, simpledialog
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
from typing import Callable, Dict, Any, Union
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import pyautogui
import win32gui
import win32com.client
"""
############################################
            config = f"""
# Configuration
SCRIPT_ID = "{script_id}"
WEBDRIVER_PATH = r"{webdriver_path}"
WAIT_TIME = 1
RETRY_ATTEMPTS = 3
DEBUG = True
ENABLE_VISUALIZATION = {enable_visualization}
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
"""
############################################
            web_automation_class = """
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
        edge_options.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 2})
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver

    def wait_for_element(self, name: str, condition: Callable, timeout: int = WAIT_TIME) -> Any:
        logger.debug(f"Waiting for element: {name}")
        return WebDriverWait(self.driver, timeout).until(condition((By.CSS_SELECTOR, name)))

    def scroll_into_view(self, name: str, by: By = By.CSS_SELECTOR) -> None:
        logger.debug(f"Scrolling element into view: {name}")
        element = self.driver.find_element(by, name)
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.05)  # Allow time for scrolling animation

    def ask_and_input(self, name: str, prompt: str) -> None:
        root = tk.Tk()
        root.withdraw()
        user_input = simpledialog.askstring("Input", prompt)
        root.destroy()
        if user_input is not None:
            self.wait_and_input(name, user_input)
        else:
            logger.warning(f"User cancelled input for prompt: {prompt}")

    def retry_on_stale(func):
        def wrapper(*args, **kwargs):
            for attempt in range(RETRY_ATTEMPTS):
                try:
                    return func(*args, **kwargs)
                except StaleElementReferenceException:
                    logger.warning(f"Stale element reference exception. Retrying (attempt {attempt + 1}/{RETRY_ATTEMPTS})")
                    time.sleep(0.05)  # Add a small delay before retrying
                    if attempt == RETRY_ATTEMPTS - 1:
                        raise
        return wrapper

    @retry_on_stale
    def wait_and_click(self, action: Dict[str, Any]) -> None:
        name = action["name"]
        by = action.get("by", By.CSS_SELECTOR)
        timeout = action.get("timeout", WAIT_TIME)
        scroll = action.get("scroll", True)

        if scroll:
            self.scroll_into_view(name, by)

        try:
            element = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, name)))
            logger.info(f"Clicking element: {name}")
            if ENABLE_VISUALIZATION:
                self.highlight_element(element)
            element.click()
        except TimeoutException:
            logger.error(f"Element not clickable: {name}")
            raise

    @retry_on_stale
    def wait_and_input(self, name: str, text: str) -> None:
        if name:
            self.scroll_into_view(name)
            element = self.wait_for_element(name, EC.element_to_be_clickable)
            logger.info(f"Inputting text: {text}")
            if ENABLE_VISUALIZATION:
                self.highlight_element(element)
            element.clear()
            element.send_keys(text)
        else:
            logger.info(f"Pasting text: {text}")
            ActionChains(self.driver).send_keys(text).perform()

    @retry_on_stale
    def wait_and_select(self, name: str, text: str) -> None:
        self.scroll_into_view(name)
        dropdown = self.wait_for_element(name, EC.element_to_be_clickable)
        
        try:
            select = Select(dropdown)
            logger.info(f"Selecting option: {text}")
            if ENABLE_VISUALIZATION:
                self.highlight_element(dropdown)
            select.select_by_visible_text(text)
        except WebDriverException:
            # If it's not a <select> element, try to handle it as a custom dropdown
            logger.info(f"Handling as custom dropdown: {name}")
            dropdown.click()
            time.sleep(0.05)  # Wait for dropdown to expand
            option = self.driver.find_element(By.XPATH, f"//*[contains(text(), '{text}')]")
            if ENABLE_VISUALIZATION:
                self.highlight_element(option)
            option.click()

    @retry_on_stale
    def wait_and_mouseover(self, name: str) -> None:
        self.scroll_into_view(name)
        element = self.wait_for_element(name, EC.presence_of_element_located)
        logger.info(f"Moving mouse over element: {name}")
        if ENABLE_VISUALIZATION:
            self.highlight_element(element)
        ActionChains(self.driver).move_to_element(element).perform()

    def wait_and_relative_click(self, anchor_name: str, target_name: str, direction: str) -> None:
        logger.debug(f"Waiting for anchor element: {anchor_name}")
        anchor = self.wait_for_element(anchor_name, EC.presence_of_element_located)
        logger.info(f"Clicking element relative to {anchor_name}: {direction}")
        relative_locator = getattr(locate_with(By.CSS_SELECTOR, target_name), direction)(anchor)
        relative_element = self.driver.find_element(relative_locator)
        if ENABLE_VISUALIZATION:
            self.highlight_element(relative_element)
        relative_element.click()

    def switch_to_window(self, window_name: str) -> None:
        logger.info(f"Switching to window: {{window_name}}")
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if window_name in self.driver.title or window_name == handle:
                logger.info(f"Switched to window: {{self.driver.title}}")
                return
        logger.error(f"Window not found: {window_name}")
        raise NoSuchElementException(f"Window not found: {window_name}")

    def send_key(self, key: str) -> Union[str, Keys]:
        special_keys = {
            'Enter': Keys.ENTER, 'Tab': Keys.TAB, 'Shift': Keys.SHIFT,
            'Ctrl': Keys.CONTROL, 'Alt': Keys.ALT, 'Esc': Keys.ESCAPE,
            'Backspace': Keys.BACKSPACE, 'Delete': Keys.DELETE,
            'PageUp': Keys.PAGE_UP, 'PageDown': Keys.PAGE_DOWN,
            'Up': Keys.ARROW_UP, 'Down': Keys.ARROW_DOWN,
            'Left': Keys.ARROW_LEFT, 'Right': Keys.ARROW_RIGHT,
            'Home': Keys.HOME, 'End': Keys.END, 'Insert': Keys.INSERT,
            **{f'F{i}': getattr(Keys, f'F{i}') for i in range(1, 13)}
        }
        if '+' in key:
            modifier, letter = key.split('+')
            modifier_key = special_keys.get(modifier.capitalize(), modifier)
            return f"{modifier_key}+{letter.lower()}"
        return special_keys.get(key, key)

    def highlight_element(self, element):
        original_style = element.get_attribute('style')
        self.driver.execute_script(
            "arguments[0].setAttribute('style', arguments[1]);",
            element,
            "border: 2px solid red; background: yellow;"
        )
        time.sleep(0.5)  # Highlight for 0.5 seconds
        self.driver.execute_script(
            "arguments[0].setAttribute('style', arguments[1]);",
            element,
            original_style
        )

    def handle_file_dialog(self, action: Dict[str, Any]) -> None:
        dialog_title = action.get("dialog_title", "Open")
        file_path = action["file_path"]
        key_sequence = action["key_sequence"]
        additional_text = action["additional_text"]
        wait_time = action["wait_time"]
    
        logger.info(f"Handling file dialog: {dialog_title}")
        logger.info(f"File path: {file_path}")
        logger.info(f"Key sequence: {key_sequence}")
        logger.info(f"Additional text: {additional_text}")
    
        try:
            # Wait for the dialog to appear
            time.sleep(wait_time)
    
            # Connect to the dialog using the win32 backend
            app = Application(backend="win32").connect(title=dialog_title, visible_only=False)
            dlg = app.window(title=dialog_title)
            dlg.set_focus()
    
            # Define a mapping for special key sequences
            special_keys = {
                'Alt+D': ('alt', 'd'),
                'Ctrl+A': ('ctrl', 'a'),
                'Ctrl+C': ('ctrl', 'c'),
                'Ctrl+V': ('ctrl', 'v'),
                'Shift+Tab': ('shift', 'tab'),
                'Tab': 'tab',
                'Enter': 'enter',
                'Down': 'down',
                'Up': 'up',
                'Left': 'left',
                'Right': 'right',
                'Esc': 'esc',
                'Backspace': 'backspace',
                'Delete': 'delete',
                'PageUp': 'pageup',
                'PageDown': 'pagedown',
                'Home': 'home',
                'End': 'end',
                'Insert': 'insert',
                'F1': 'f1',
                'F2': 'f2',
                'F3': 'f3',
                'F4': 'f4',
                'F5': 'f5',
                'F6': 'f6',
                'F7': 'f7',
                'F8': 'f8',
                'F9': 'f9',
                'F10': 'f10',
                'F11': 'f11',
                'F12': 'f12',
                # Add other combinations as needed
            }
    
            # Send the key sequence if it's in the mapping
            if key_sequence in special_keys:
                keys = special_keys[key_sequence]
                if isinstance(keys, tuple):
                    for key in keys:
                        pyautogui.press(key)
                else:
                    pyautogui.press(keys)
            else:
                pyautogui.press(key_sequence)
    
            # Send additional text if any
            if additional_text:
                time.sleep(0.1)
                # Split the additional text into individual characters and send them one by one
                for char in additional_text:
                    pyautogui.press(char)
    
            time.sleep(0.2)
            pyautogui.press('enter')
    
            logger.info("File dialog interaction completed")
    
        except Exception as e:
            logger.error(f"Error handling file dialog: {e}")
            logger.error(traceback.format_exc())
            # Optionally, you can add a user prompt to decide whether to continue or exit
            # For now, we'll just log the error and continue
    
        finally:
            # Switch back to the browser window
            self.driver.switch_to.window(self.driver.window_handles[0])
            logger.info("Switched back to the browser window")

    def perform_action(self, action: Dict[str, Any]) -> str:
        action_type = action["type"]
        logger.info(f"Performing action: {action_type}")
        try:
            if action_type == "loop_start":
                logger.info("Starting loop")
            elif action_type == "loop_end":
                logger.info(f"Ending loop (Repeat: {action['repeat']})")
            elif action_type == "url":
                self.driver.get(action["url"])
                WebDriverWait(self.driver, WAIT_TIME).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            elif action_type == "dropdown":
                self.wait_and_select(action["name"], action["text"])
            elif action_type == "click":
                self.wait_and_click(action)
            elif action_type == "input":
                self.wait_and_input(action.get("name", ""), action["text"])
            elif action_type == "paste":
                self.wait_and_input("", action["text"])
            elif action_type == "sleep":
                duration = action["duration"] / 1000
                logger.info(f"Sleeping for {duration} seconds")
                time.sleep(duration)
            elif action_type == "keypress":
                key = self.send_key(action["key"])
                logger.info(f"Pressing key: {action['key']}")
                if isinstance(key, str) and '+' in key:
                    modifier, letter = key.split('+')
                    ActionChains(self.driver).key_down(modifier).send_keys(letter).key_up(modifier).perform()
                else:
                    active_element = self.driver.switch_to.active_element
                    if ENABLE_VISUALIZATION:
                        self.highlight_element(active_element)
                    active_element.send_keys(key)
            elif action_type == "relative click":
                self.wait_and_relative_click(action["anchor"], action["target"], action["direction"])
            elif action_type == "switch_window":
                window_name = action["window_name"]
                self.switch_to_window(window_name)
            elif action_type == "ask and input":
                self.ask_and_input(action["name"], action["prompt"])
            elif action_type == "mouseover":
                self.wait_and_mouseover(action["name"])
            elif action_type == "file dialog":
                self.handle_file_dialog(action)
            else:
                logger.warning(f"Unknown action type: {action_type}")
            time.sleep(0.05)  # Small delay between actions
        except NoSuchElementException as e:
            logger.error(f"Element not found: {e}")
            choice = self.handle_element_not_found(action)
            if choice == 'skip':
                logger.info("Skipping to next action")
                return
            elif choice == 'retry':
                logger.info("Retrying action")
                self.perform_action(action)
            elif choice == 'exit':
                logger.info("Exiting script")
                raise
            return "success"
        except NoSuchElementException as e:
            logger.error(f"Element not found: {e}")
            return self.handle_element_not_found(action)
        except Exception as e:
            logger.error(f"Error performing action {action_type}: {e}")
            raise

    def handle_element_not_found(self, action: Dict[str, Any]) -> str:
        root = tk.Tk()
        root.withdraw()
        element_name = action['name']
        message = f'''ELEMENT NOT FOUND: <b>{element_name}</b>
    What do you want to do?
    
    YES will skip this element and Continue
    NO will exit this script'''
        choice = messagebox.askquestion("Element Not Found", message, icon="warning")
        root.destroy()
        if choice == 'yes':
            return 'skip'
        elif choice == 'no':
            return 'exit'
        else:
            return 'exit'

def show_error_popup(error_message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    with open(LOG_FILE, 'r') as log_file:
        lines = log_file.readlines()
    error_text = f"Error: {error_message}"
    messagebox.showerror("Error", error_text)
    root.destroy()

def main():
    automation = None
    try:
        automation = WebAutomation(WEBDRIVER_PATH)

        actions = {actions_json}

        def execute_actions(start_index, end_index, repeat=1):
            i = start_index
            while i < end_index:
                for _ in range(repeat):
                    action = actions[i]
                    if action['type'] == 'loop_start':
                        loop_end = next(j for j in range(i + 1, len(actions)) if actions[j]['type'] == 'loop_end')
                        execute_actions(i + 1, loop_end, actions[loop_end]['repeat'])
                        i = loop_end
                    elif action['type'] == 'loop_end':
                        break
                    else:
                        result = automation.perform_action(action)
                        if result == "retry":
                            continue  # Retry the same action
                        elif result == "exit":
                            logger.info("User chose to exit the script")
                            return  # Exit the script
                        # If result is "skip" or "success", move to the next action
                i += 1

        execute_actions(0, len(actions))
        logger.info("Script execution completed successfully.")
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        if DEBUG:
            logger.error(traceback.format_exc())
        show_error_popup(str(e))
    finally:
        if automation and automation.driver:
            print("Press 'F10' to close the browser...")
            while True:
                if keyboard.is_pressed('F10'):
                    break
                time.sleep(5)
            automation.driver.quit()
        logger.info(f"Log file saved as {LOG_FILE}")

if __name__ == "__main__":
    main()
"""
    
            code = imports + config + web_automation_class
    
            # Replace the actions JSON in the main function
            code = code.replace('{actions_json}', actions_json)
    
            return code
        except Exception as e:
            messagebox.showerror("Error Generating Code", f"An error occurred while generating the code:\n\n{str(e)}")
            return None

    def on_closing(self):
        try:
            self.stop_detection()
        except Exception as e:
            print(f"Error during stop_detection: {e}")
        
        try:
            self.master.destroy()
        except Exception as e:
            print(f"Error during master.destroy: {e}")
        
        # Force exit if needed
        import sys
        sys.exit(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedAutomationApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
