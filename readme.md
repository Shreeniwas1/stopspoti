# StopSpoti

## Overview

StopSpoti is a Python-based utility designed to monitor and control Spotify's audio sessions on your computer. It automatically pauses Spotify when other audio applications are active and resumes playback when no other audio is detected. This ensures an uninterrupted listening experience across different applications.

## Features

- **Automatic Audio Management**: Detects active audio sessions and manages Spotify playback accordingly.
- **Process Monitoring**: Continuously monitors Spotify and other audio processes to make real-time adjustments.
- **User-Friendly**: Simple command-line interface for easy usage.

## Usage

### Prerequisites

- Python 3.x installed on your system.
- Required Python packages:
  - `psutil`
  - `comtypes`
  - `pycaw`
  - `pyautogui`
  - `pywin32`
  - 'others try to understand using the terminal' 



### Running the Script

It's recommended to use Visual Studio Code (VSCode) for running the script. Open the project in VSCode and execute the following command in the terminal:

```bash
python stopspoti.py
```
### Stopping the Script

To stop the script while it's running, simply press `Ctrl+C` in the terminal.
## Important Warning


**DO NOT compile this script into an executable (`.exe`). Running it as a compiled program can cause recursive launches, system instability, and crashes. For safety and reliability, ONLY use this script within VSCode or a similar controlled development environment until proper safeguards are implemented.**


if you are developing exe and face recursion problem read this: 

If such a scenario occurs, the program is designed to stop the loop after closing of spotify. However, to ensure safety and stability, it's best to test and run the script directly using Python in a controlled environment like VSCode before considering compilation.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your enhancements or bug fixes.
if possible share the exe with GUI if you know how to do it. :| thanks 
i will be learning it but I dont have time right now :(

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
