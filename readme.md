# Web Automation Tool

This project is a Python-based web automation tool with a graphical user interface (GUI). It allows users to create, manage, and execute web automation scripts without writing code.

## Features

- User-friendly GUI for creating web automation scripts
- Supports various action types: Click, Input, Dropdown selection, URL navigation, Sleep, and Keypress
- Real-time CSS selector detection
- Easy-to-use action management (add, remove, reorder)
- Generates executable Python scripts
- Uses Selenium WebDriver for browser automation

## Requirements

- Python 3.x
- Tkinter
- Selenium WebDriver
- Microsoft Edge WebDriver
- pyperclip
- keyboard

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/web-automation-tool.git
   ```

2. Install the required Python packages:
   ```
   pip install selenium pyperclip keyboard
   ```

3. Download the Microsoft Edge WebDriver that matches your Edge browser version and place it in a known location on your computer.

4. Update the `SELENIUM_DRIVER_PATH` variable in the script with the path to your Edge WebDriver.

## Usage

1. Run the script:
   ```
   python web_automation_tool.py
   ```

2. Enter the URL of the website you want to automate in the URL field.

3. Click "Start Detection" to open the browser and start detecting CSS selectors.

4. Hover over elements on the webpage to see their CSS selectors.

5. Press Shift to copy the current selector to the clipboard.

6. Use the GUI to add actions (Click, Input, Dropdown, etc.) to your automation script.

7. Click "Complete" to generate and save the Python script for your automation.

8. The generated script will be saved on your Desktop as "web_automation_script.txt".

## Action Types

- **Click**: Click on an element
- **Input**: Enter text into an input field
- **Dropdown**: Select an option from a dropdown menu
- **URL**: Navigate to a specific URL
- **Sleep**: Pause the script for a specified duration
- **Keypress**: Simulate a keyboard key press

## Notes

- The tool uses Microsoft Edge in InPrivate mode for automation.
- Generated scripts include error handling and logging for better debugging.
- Press F10 to close the browser when running the generated script.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check [issues page](https://github.com/yourusername/web-automation-tool/issues) if you want to contribute.

## License

[MIT](https://choosealicense.com/licenses/mit/)
```

This README provides an overview of the project, its features, installation instructions, usage guide, and other relevant information. You may want to customize it further based on your specific project details, such as adding your GitHub username, adjusting the license if needed, or adding more specific instructions or examples.
