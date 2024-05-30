
# Bloons6 Bot

This project is an automated bot for Bloons Tower Defense 6, designed to play the game by reading instructions from an Excel sheet and executing the actions using Python scripts and various libraries.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Script Execution](#script-execution)

## Prerequisites
Make sure you have the following installed:
- Python 3.x
- `pandas` library
- `pydirectinput` library
- `PyAutoGUI` library
- `pywin32` library

You can install the required Python packages using the `requirements.txt` file.

```bash
pip install -r requirements.txt
```

Make sure you have the following settings:
- Screen Size 1024x768 for your BloonsTD6, however your screen should be bigger to press on the control.xlsm play button
- Disable Auto Start rounds. PyAutoGui needs to recognize the beginning and end of a round and has to start the round
- During the execution, don't move your mouse; to abort the Script press Ctrl+C on the Sript window

## Installation
1. Clone the repository to your local machine.
2. Create a virtual environment and activate it.
   ```bash
   python -m venv venv
   ```
   - On Windows:
     ```bash
     venv\Scriptsctivate
     ```
   - On macOS and Linux:
     ```bash
     source venv/bin/activate
     ```
3. Install the required libraries using the `requirements.txt` file.
   ```bash
   pip install -r requirements.txt
   ```

## Usage
To run the bot, use the following command:

```bash
"<python.exe in your venv>" "<main.py on your local repo> " --mode "<mode>" --control_path "<path_to_control_file>"
C:/Users/david/Desktop/BloonsTD6_BOT/venv/Scripts/python.exe C:/Users/david/Desktop/BloonsTD6_BOT/src/main.py --mode "<mode>" --control_path "<path_to_control_file>"control.xlsm"
```

### Example
```bash
C:/Users/david/Desktop/BloonsTD6_BOT/venv/Scripts/python.exe C:/Users/david/Desktop/BloonsTD6_BOT/src/main.py --mode "chimps" --control_path "C:/Users/david/Desktop/bloonsTD6_BOT/src/control.xlsm"
```

### Modes
The bot supports different modes as follows:
- `easy|standard`
- `easy|primary_only`
- `medium|standard`
- `medium|military_only`
- `medium|apopalypse`
- `medium|reverse`
- `hard|standard`
- `hard|magic_monkeys_only`
- `hard|double_hp_moabs`
- `hard|half_cash`
- `hard|alternate_bloons_rounds`
- `impoppable`
- `chimps`
- `easy|deflation`

## Script Execution
The script can also be started via a Play Button in VBA, which calls the Python script. Make sure to update the path to your Python executable.

## Notes
- Ensure the towers are defined pixel-perfect in the Excel sheet.
- The script reads actions from the Excel sheet and simulates key presses and mouse movements/clicks to play the game.

---

This project is intended for educational and personal use. Use at your own risk.
