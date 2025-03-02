# StopSpoti

## Overview

StopSpoti is a Python utility that intelligently manages Spotify playback based on system audio activity. It automatically pauses Spotify when other applications produce sound and resumes playback when they stop, providing a seamless audio experience.

## Features

- **Intelligent Audio Detection**: Uses advanced audio session monitoring to accurately detect active audio streams
- **Automatic Playback Control**: Pauses/resumes Spotify automatically based on other applications' audio activity
- **Performance Optimized**: 
  - Cached audio session monitoring (2-second refresh rate)
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
```

### Key Dependencies Explanation

- `psutil`: Process and system monitoring
- `comtypes` & `pycaw`: Windows Core Audio API interface
- `pyautogui`: Spotify playback control
- `pywin32`: Windows API access for window management

## Usage

### Running in Development Mode

1. Open the project in Visual Studio Code
2. Open an integrated terminal
3. Run the script:
```bash
python stopspotv1.py
```

### Program Behavior

- Monitors audio sessions every 0.5 seconds
- Pauses Spotify when other audio is detected (1-second cooldown between actions)
- Resumes Spotify when other audio stops
- Minimizes Spotify window after each interaction
- Logs important events with timestamps

### Stopping the Program

Press `Ctrl+C` in the terminal to safely stop the program.

## Important Notes

1. **Do Not Compile Warning**: 
   - Not recommended to compile as an executable
   - May cause recursive launches and system instability
   - Use directly with Python interpreter for safety

2. **Performance Settings**:
   - Audio cache refresh: 2 seconds
   - Main loop interval: 0.5 seconds
   - Action cooldown: 1.0 seconds
   - Error recovery delay: 2 seconds

3. **Debug Mode**:
   - Debug logging is enabled by default
   - Logs are displayed with timestamps
   - Log refresh interval: 5 seconds

## Troubleshooting

- If Spotify control isn't working:
  - Ensure Spotify is running and visible
  - Check if Python has permission to control windows
  - Verify all required packages are installed

- If high CPU usage occurs:
  - Verify process priority settings
  - Check for multiple instances
  - Ensure proper COM object cleanup

## Contributing

Feel free to contribute improvements or report issues. Some areas for potential enhancement:

- GUI implementation
- Configuration file support
- Additional audio source controls
- Better executable compilation support

## License

This project is licensed under the MIT License.
