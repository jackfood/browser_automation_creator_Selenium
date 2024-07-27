# Web Automation Tool v1.7.86

## Description

Web Automation Tool v1.7.86 is a powerful and user-friendly application designed to automate web interactions using Selenium WebDriver. It provides a graphical interface for creating, managing, and executing web automation scripts without the need for manual coding.

## Features

- **Dynamic Action Recording**: Capture clicks, input, dropdown selections, and more as you interact with web pages.
- **Flexible Action Types**: Supports various action types including URL navigation, clicking, input, dropdown selection, keyboard input, relative clicking, window switching, and file dialog handling.
- **Visual Element Selection**: Uses CSS selectors with real-time highlighting for accurate element identification.
- **Action Visualization**: Option to visually highlight elements during script execution for better debugging.
- **Looping Functionality**: Create loops within your automation scripts for repetitive tasks.
- **Error Handling**: Robust error handling with user prompts for element not found scenarios.
- **Logging**: Comprehensive logging system for easy troubleshooting.
- **Import/Export**: Ability to import existing scripts and export created scripts for sharing or later use.
- **Cross-browser Support**: Primarily designed for Microsoft Edge, but can be adapted for other browsers.

## Requirements

- Python 3.7+
- Selenium WebDriver
- Microsoft Edge WebDriver
- Additional Python libraries: tkinter, pygetwindow, pyperclip, pywinauto, pyautogui, keyboard

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/web-automation-tool.git
   ```
2. Install required Python packages:
   ```
   pip install selenium pygetwindow pyperclip pywinauto pyautogui keyboard
   ```
3. Download and install the appropriate Microsoft Edge WebDriver for your Edge version.

## Usage

1. Run the script:
   ```
   python web_automation_tool.py
   ```
2. Enter the URL you want to automate in the URL field.
3. Click "Start Detection" to begin capturing actions.
4. Interact with the webpage. The tool will record your actions.
5. Use the action type dropdown and input fields to add or modify actions.
6. Click "Add" to include an action in your script.
7. Use "Remove" to delete unwanted actions.
8. Adjust the order of actions using the "Insert Position" field.
9. Click "Complete" to generate and save your automation script.

## Action Types

- **URL**: Navigate to a specific URL
- **Click**: Click on an element
- **Dropdown**: Select an option from a dropdown menu
- **Input**: Enter text into an input field
- **Sleep**: Pause execution for a specified duration
- **Keypress**: Simulate keyboard input
- **Relative Click**: Click on an element relative to another element
- **Windows Selector**: Switch between different browser windows
- **Ask and Input**: Prompt for user input during script execution
- **MouseOver**: Simulate mouse hover over an element
- **File Dialog**: Handle file upload dialogs

## Advanced Features

- **Looping**: Use "Start Loop" and "End Loop" buttons to create loops in your script.
- **Visualization**: Enable action visualization for better script debugging.
- **Alternate Selectors**: The tool captures both primary and alternate selectors for more robust element identification.

## Troubleshooting

- Check the generated log file on your desktop for detailed error information.
- If elements are not being found, try using alternate selectors or adjusting wait times.
- For file dialog issues, ensure you have the correct file paths and permissions.

## Contributing

Contributions, issues, and feature requests are welcome.

## Disclaimer

This tool is for educational and personal use only. Ensure you have permission to automate interactions with any website you target. The developers are not responsible for any misuse of this tool.
This tool and Readme are created by Claude.ai Sonnet 3.5
