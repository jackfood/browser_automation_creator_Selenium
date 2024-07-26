# Web Automation Tool v1.7.7

## Overview

This Web Automation Tool is a powerful Python-based application that allows users to create, manage, and execute web automation scripts with ease. It combines a user-friendly GUI for script creation with a robust backend for script execution, making it suitable for both beginners and advanced users.

## Features

- **User-friendly GUI**: Easy-to-use interface for creating and managing automation scripts in python.
- **Multiple Action Types**: Supports various action types including clicks, input, dropdown selection, URL navigation, and more.
- **Relative Element Selection**: Ability to select elements relative to others, enhancing flexibility.
- **Window Management**: Switch between different browser windows during automation.
- **Loop Support**: Create loops within your automation scripts for repetitive tasks.
- **Dynamic User Input**: Ability to prompt for user input during script execution.
- **Error Handling**: Robust error handling with user-friendly popups and detailed logging.
- **Visualization**: Option to visually highlight elements during script execution for better debugging.
- **Extensible**: Easy to add new action types and extend functionality.

## Requirements

- Python 3.6+
- Selenium WebDriver
- Microsoft Edge WebDriver
- Additional Python packages: tkinter, keyboard, pyperclip

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/web-automation-tool.git
   ```

2. Install the required Python packages:
   ```
   pip install selenium keyboard pyperclip
   ```

3. Download and install the Microsoft Edge WebDriver that matches your Edge browser version.

4. Update the `SELENIUM_DRIVER_PATH` in the script to point to your Edge WebDriver location.

## Usage

1. Run the script:
   ```
   python web_automation_tool.py
   ```

2. Use the GUI to create your automation script:
   - Enter the starting URL
   - Add actions (click, input, select, etc.)
   - Arrange and modify actions as needed

3. Click "Start Detection" to begin the element detection process.

4. Use 'Ctrl' to copy element selectors.

5. When your script is ready, click "Complete" to generate the Python script. It will save the Python script to your desktop.

6. Run the generated python script to execute your automation.

## Action Types

- **URL**: Navigate to a specific URL
- **Click**: Click on an element
- **Dropdown**: Select an option from a dropdown
- **Input**: Enter text into an input field
- **Sleep**: Wait for a specified duration
- **Keypress**: Simulate a keyboard key press
- **Relative Click**: Click on an element relative to another
- **Switch Window**: Switch to a different browser window
- **Ask and Input**: Prompt for user input during execution
- **MouseOver**: Hover over an element

## Advanced Features

- **Loop Control**: Create loops in your automation script for repetitive tasks.
- **Relative Element Selection**: Select elements based on their position relative to other elements.
- **Dynamic User Input**: Prompt for user input during script execution for more flexible automation.
- **Visual Debugging**: Option to highlight elements as they're interacted with during script execution.

## Error Handling

The tool includes robust error handling:
- Detailed logging of all actions and errors
- User-friendly popups for common issues like element not found
- Options to skip, retry, or exit when encountering errors

## Customization

You can easily extend the tool by adding new action types or modifying existing ones in the `WebAutomation` class.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## Disclaimer

This tool is for educational and personal use only. Be sure to comply with the terms of service of any websites you automate. The authors are not responsible for any misuse of this tool.
This form is created by Claude.ai Sonnet 3.5.
