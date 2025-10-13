# StopSpoti

## Overview

StopSpoti is a Python utility that intelligently manages Spotify playback based on system audio activity. It automatically pauses Spotify when other applications produce sound and resumes playback when they stop, providing a seamless audio experience.

## Features

- **Intelligent Audio Detection**: Uses advanced audio session monitoring to accurately detect active audio streams
- **Automatic Playback Control**: Pauses/resumes Spotify automatically based on other applications' audio activity
- **Modern GUI**: User-friendly interface with customizable settings and real-time logging
- **Performance Optimized**: 
  - Cached audio session monitoring (configurable refresh rate)
  - Efficient process management with minimal CPU usage
  - Balanced timing intervals for responsiveness and resource usage
- **System-Friendly**:
  - Runs at below-normal priority to minimize system impact
  - Safe COM object handling and cleanup
  - Graceful error recovery and shutdown

## Technical Requirements

### Prerequisites

- Python 3.x
- Windows OS (uses Windows-specific audio APIs)

### Required Python Packages

```bash
pip install psutil
pip install comtypes
pip install pycaw
pip install pyautogui
pip install pywin32
pip install customtkinter
pip install pystray
pip install Pillow
```

### Key Dependencies Explanation

- `psutil`: Process and system monitoring
- `comtypes` & `pycaw`: Windows Core Audio API interface
- `pyautogui`: Spotify playback control
- `pywin32`: Windows API access for window management
- `customtkinter`: Modern GUI framework
- `pystray` & `Pillow`: System tray icon support

## Usage

### Running the Application

1. Open the project in Visual Studio Code
2. Open an integrated terminal
3. Run the script:
```bash
python stopspotv1.py
```

### Running the Executable (Standalone)

For easier distribution, the application can be built as a standalone executable:

1. **Build the executable:**
```bash
pip install pyinstaller
pyinstaller --onedir --noconsole --icon=icon.ico --exclude-module multiprocessing --name="SpotifyAutoController" stopspotv1.py
```

2. **Run the executable:**
   - Navigate to the `dist/SpotifyAutoController` folder
   - Double-click `SpotifyAutoController.exe`
   - Or run from command line: `SpotifyAutoController.exe`

3. **Command line options:**
   - `--version`: Show version information
   - `--test`: Run a 10-second resource usage test

**Note:** The build creates a directory structure instead of a single file to avoid subprocess creation that can cause cursor loading issues on Windows. Using `--onefile` will create subprocesses and cause the cursor loading problem.

### GUI Interface

The application launches a modern GUI with the following features:

- **Settings Panel**: Configure audio detection parameters
  - Peak Threshold: Sensitivity for audio detection
  - Cache Timeout: How often to refresh audio sessions
  - Log Interval: Frequency of debug logging
  - Action Cooldown: Minimum time between pause/resume actions
  - Debug Mode: Enable/disable detailed logging
  - Ignored Processes: List of processes to exclude from audio monitoring

- **Control Buttons**: Start/Stop monitoring with visual feedback
- **Status Display**: Real-time status updates
- **Log Window**: Live logging of all actions and events

### Program Behavior

- Monitors audio sessions at configurable intervals
- Pauses Spotify when other audio is detected (configurable cooldown)
- Resumes Spotify when other audio stops
- Minimizes Spotify window after each interaction
- Logs all events with timestamps in the GUI

### Stopping the Program

Click "Stop Monitoring" in the GUI or close the window to safely stop the program.

## Configuration Options

### Audio Detection Settings

- **Peak Threshold** (default: 0.0005): Minimum audio level to consider as active
- **Cache Timeout** (default: 2s): How often to refresh audio session data
- **Log Interval** (default: 5s): Time between debug log entries
- **Action Cooldown** (default: 1.0s): Minimum delay between control actions

### Process Filtering

- **Ignored Processes**: System processes that should not trigger Spotify pausing
- Default ignores: system processes, audio tools, screen recording software

## Important Notes

1. **GUI Theme**: Features a black background with purple accents and green text for optimal visibility

2. **Performance Settings**:
   - Configurable audio cache refresh
   - Adjustable main loop intervals
   - Customizable action cooldowns
   - Flexible error recovery timing

3. **Debug Mode**:
   - Real-time logging in GUI
   - Timestamped event tracking
   - Configurable log verbosity

## Troubleshooting

- If Spotify control isn't working:
  - Ensure Spotify is running and visible
  - Check if Python has permission to control windows
  - Verify all required packages are installed

- If high CPU usage occurs:
  - Adjust cache timeout and loop intervals
  - Verify process priority settings
  - Check for multiple instances
  - Ensure proper COM object cleanup

- If GUI doesn't start:
  - Ensure all dependencies are installed
  - Check Python version compatibility
  - Try running with virtual environment

## Contributing

Feel free to contribute improvements or report issues. Some areas for potential enhancement:

- Additional GUI themes
- Configuration file persistence
- System tray minimization
- Advanced audio filtering options
- Multi-language support

## License

This project is licensed under the MIT License.
