import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
import pyperclip
from urllib.parse import urlparse
import keyboard
import json
import os

# User Configuration Section
SELENIUM_DRIVER_PATH = r'C:\Users\Desktop\msedgedriver.exe'  # Update your Selenium Driver
DEFAULT_URL = "https://www.google.com" # Your Default opening URL for detection
OUTPUT_FILE_NAME = "web_automation_script.py" # Your name of the script (default save location is in desktop)

class CombinedAutomationApp:
    def __init__(self, master):
        self.master = master
        master.title("Web Automation Tool")
        master.geometry("800x600")

        self.setup_ui()
        
        self.driver = None
        self.is_detecting = False
        self.current_selector = ""
        self.actions = []

    def setup_ui(self):
        self.create_url_frame()
        self.create_start_button()
        self.create_selector_frame()
        self.create_action_frame()
        self.create_action_list()
        self.create_control_buttons()

    def create_url_frame(self):
        self.url_frame = tk.Frame(self.master)
        self.url_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(self.url_frame, text="URL:").pack(side=tk.LEFT)

        self.url_entry = tk.Entry(self.url_frame, width=50)
        self.url_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))
        self.url_entry.insert(0, DEFAULT_URL)

    def create_start_button(self):
        self.start_button = tk.Button(self.master, text="Start Detection", command=self.start_detection)
        self.start_button.pack(pady=10)

    def create_selector_frame(self):
        self.selector_frame = tk.Frame(self.master, bd=2)
        self.selector_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(self.selector_frame, text="Current CSS Selector:", anchor='w').pack(side=tk.LEFT, padx=5)

        self.selector_display = tk.Label(self.selector_frame, width=50, anchor='w', bg='Black', fg='white')
        self.selector_display.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.copy_status_label = tk.Label(self.master, text="", fg="green")
        self.copy_status_label.pack(pady=5)

    def create_action_frame(self):
        self.action_frame = tk.Frame(self.master)
        self.action_frame.pack(pady=10, padx=10, fill=tk.X)

        self.action_types = ["Click", "Dropdown", "Input", "URL", "Sleep", "Keypress"]
        self.special_keys = ["", "Enter", "Tab", "Shift", "Ctrl", "Alt", "Esc", "Backspace", "Delete", "PageUp", "PageDown", "Home", "End", "Insert", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]

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

        self.add_button = ttk.Button(self.action_frame, text="Add", command=self.add_action)
        self.add_button.grid(row=4, column=0, columnspan=2, pady=10)

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
                service = Service(SELENIUM_DRIVER_PATH)
                self.driver = webdriver.Edge(service=service, options=options)
                self.driver.get(url)
                self.is_detecting = True
                self.start_button.config(text="Stop Detection")
                self.inject_mouse_move_script()
                self.detect_css()
                keyboard.on_press_key('shift', self.copy_to_clipboard)
            else:
                tk.messagebox.showerror("Invalid URL", "Please enter a valid URL.")
        except WebDriverException as e:
            tk.messagebox.showerror("WebDriver Error", f"Error initializing WebDriver: {str(e)}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")

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
        self.selector_display.config(text="")
        self.copy_status_label.config(text="")

    def inject_mouse_move_script(self):
        script = """
        window.lastElement = null;
        window.addEventListener('mousemove', function(e) {
            var element = document.elementFromPoint(e.clientX, e.clientY);
            if (element !== window.lastElement) {
                window.lastElement = element;
                var selector = '';
                if (element.id) {
                    selector = element.tagName.toLowerCase() + '#' + element.id;
                } else if (element.className) {
                    selector = element.tagName.toLowerCase() + '.' + element.className.replace(/\s+/g, '.');
                } else {
                    selector = element.tagName.toLowerCase();
                }
                window.currentSelector = selector;
            }
        });
        """
        self.driver.execute_script(script)

    def detect_css(self):
        if not self.is_detecting:
            return

        try:
            current_url = self.driver.current_url
            if current_url != self.url_entry.get():
                self.url_entry.delete(0, tk.END)
                self.url_entry.insert(0, current_url)
                self.inject_mouse_move_script()

            selector = self.driver.execute_script("return window.currentSelector || '';")
            if selector:
                self.selector_display.config(text=selector)
                self.current_selector = selector

        except NoSuchWindowException:
            self.stop_detection()
        except Exception as e:
            print(f"Error in detect_css: {e}")

        if self.is_detecting:
            self.master.after(100, self.detect_css)

    def copy_to_clipboard(self, e):
        if self.current_selector:
            pyperclip.copy(self.current_selector)
            self.selector.delete(0, tk.END)
            self.selector.insert(0, self.current_selector)
            self.copy_status_label.config(text="CSS selector copied to clipboard!")
            self.master.after(2000, lambda: self.copy_status_label.config(text=""))

    def on_action_type_change(self, event):
        action_type = self.action_type.get()
        self.selector_label.grid(row=1, column=0, padx=5, pady=5)
        self.selector.grid(row=1, column=1, padx=5, pady=5)
        self.text_label.grid(row=2, column=0, padx=5, pady=5)
        self.text.grid(row=2, column=1, padx=5, pady=5)
        self.special_key_label.grid_remove()
        self.special_key.grid_remove()
        self.special_key_input.grid_remove()
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
            self.text_label.config(text="Key:")
            self.special_key_label.grid(row=3, column=0, padx=5, pady=5)
            self.special_key.grid(row=3, column=1, padx=5, pady=5)
            self.special_key.set("")
        else:
            self.text_label.config(text="Text:")

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

        if action_type == "url":
            action = {"type": action_type, "url": text}
            display_text = f"URL: {text}"
        elif action_type == "click":
            action = {"type": action_type, "selector": selector}
            display_text = f"Click: {selector}"
        elif action_type == "dropdown":
            action = {"type": action_type, "selector": selector, "text": text}
            display_text = f"Dropdown: {selector} (Reference: {text})"
        elif action_type == "sleep":
            action = {"type": action_type, "duration": int(text)}
            display_text = f"Sleep: {text} ms"
        elif action_type == "keypress":
            key = special_key if special_key else text
            if special_key in ["Shift", "Ctrl"] and special_key_input:
                key = f"{special_key}+{special_key_input}"
            action = {"type": action_type, "key": key}
            display_text = f"Keypress: {key}"
        else:
            action = {"type": action_type, "selector": selector, "text": text}
            display_text = f"{action_type.capitalize()}: {selector} {text}"
        
        insert_position = self.insert_position.get()
        if insert_position:
            try:
                position = int(insert_position) - 1
                self.actions.insert(position, action)
                self.action_list.insert(position, display_text)
            except ValueError:
                messagebox.showwarning("Invalid Position", "Please enter a valid integer for the insert position.")
                return
        else:
            self.actions.append(action)
            self.action_list.insert(tk.END, display_text)
        
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
        except IndexError:
            messagebox.showwarning("Warning", "No action selected!")

    def complete(self):
        if not self.actions:
            messagebox.showwarning("Warning", "No actions added!")
            return
        
        code = self.generate_code()
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", OUTPUT_FILE_NAME)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(code)
        
        messagebox.showinfo("Success", f"Script saved to {file_path}")

    def generate_code(self):
        actions_json = json.dumps(self.actions, indent=4)
        code = """
import logging
import time
import traceback
import keyboard
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# Configuration
WEBDRIVER_PATH = '{0}'  # Update this path if needed
WAIT_TIME = 1
RETRY_ATTEMPTS = 5
DEBUG = True

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Setup webdriver
service = Service(executable_path=WEBDRIVER_PATH)
edge_options = Options()
edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
edge_options.add_experimental_option('useAutomationExtension', False)
edge_options.add_argument("inprivate")
edge_options.add_experimental_option("prefs", {{"profile.default_content_setting_values.notifications": 2}})

driver = webdriver.Edge(service=service, options=edge_options)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {{get: () => undefined}})")

# Action definitions
actions = {1}

def wait_for_element(selector, condition):
    logger.debug(f"Waiting for element: {{selector}}")
    return WebDriverWait(driver, WAIT_TIME).until(condition((By.CSS_SELECTOR, selector)))

def scroll_into_view(selector, by=By.CSS_SELECTOR):
    logger.debug(f"Scrolling element into view: {{selector}}")
    element = driver.find_element(by, selector)
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(0.1)

def retry_on_stale(func):
    def wrapper(*args, **kwargs):
        for attempt in range(RETRY_ATTEMPTS):
            try:
                return func(*args, **kwargs)
            except StaleElementReferenceException:
                logger.warning(f"Stale element reference exception. Retrying (attempt {{attempt + 1}}/{{RETRY_ATTEMPTS}})")
                if attempt == RETRY_ATTEMPTS - 1:
                    raise
    return wrapper

@retry_on_stale
def wait_and_click(action):
    selector = action["selector"]
    by = action.get("by", By.CSS_SELECTOR)
    timeout = action.get("timeout", WAIT_TIME)
    scroll = action.get("scroll", True)

    if scroll:
        scroll_into_view(selector, by)

    try:
        if by == By.XPATH:
            element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, selector)))
        else:
            element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, selector)))
        logger.info(f"Clicking element: {{selector}}")
        element.click()
    except TimeoutException:
        logger.error(f"Element not found: {{selector}}")

@retry_on_stale
def wait_and_input(selector, text):
    scroll_into_view(selector)
    element = wait_for_element(selector, EC.element_to_be_clickable)
    logger.info(f"Inputting text: {{text}}")
    element.clear()
    element.send_keys(text)

@retry_on_stale
def wait_and_select(selector, text):
    scroll_into_view(selector)
    dropdown = wait_for_element(selector, EC.element_to_be_clickable)
    select = Select(dropdown)
    logger.info(f"Selecting option: {{text}}")
    select.select_by_visible_text(text)

def send_key(key):
    special_keys = {{
        'Enter': Keys.ENTER, 'Tab': Keys.TAB, 'Shift': Keys.SHIFT,
        'Ctrl': Keys.CONTROL, 'Alt': Keys.ALT, 'Esc': Keys.ESCAPE,
        'Backspace': Keys.BACKSPACE, 'Delete': Keys.DELETE,
        'PageUp': Keys.PAGE_UP, 'PageDown': Keys.PAGE_DOWN,
        'Home': Keys.HOME, 'End': Keys.END, 'Insert': Keys.INSERT,
        **{{f'F{{i}}': getattr(Keys, f'F{{i}}') for i in range(1, 13)}}
    }}
    if '+' in key:
        modifier, letter = key.split('+')
        return getattr(Keys, modifier.upper()) + letter.lower()
    return special_keys.get(key, key)

def perform_action(action):
    action_type = action["type"]
    logger.info(f"Performing action: {{action_type}}")
    if action_type == "url":
        driver.get(action["url"])
        time.sleep(1)  # Wait for page to load
    elif action_type == "dropdown":
        wait_and_select(action["selector"], action["text"])
    elif action_type == "click":
        wait_and_click(action)
    elif action_type == "input":
        wait_and_input(action["selector"], action["text"])
    elif action_type == "sleep":
        duration = action["duration"] / 1000
        logger.info(f"Sleeping for {{duration}} seconds")
        time.sleep(duration)
    elif action_type == "keypress":
        key = send_key(action["key"])
        logger.info(f"Pressing key: {{action['key']}}")
        ActionChains(driver).send_keys(key).perform()
    time.sleep(0.2)  # Small delay between actions

def main():
    try:
        for action in actions:
            perform_action(action)
        logger.info("Script execution completed successfully.")
    except Exception as e:
        logger.error(f"An error occurred: {{e}}")
        if DEBUG:
            logger.error(traceback.format_exc())
    finally:
        print("Press 'F10' to close the browser...")
        while True:
            if keyboard.is_pressed('F10'):
                break
            time.sleep(0.2)
        driver.quit()

if __name__ == "__main__":
    main()
"""
        return code.format(SELENIUM_DRIVER_PATH, actions_json)

    def on_closing(self):
        self.stop_detection()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedAutomationApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
