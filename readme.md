# StopSpoti

Auto-pauses Spotify when other apps play audio, resumes when they stop.

## Requirements

- Windows OS
- Python 3.x

```bash
pip install psutil comtypes pycaw pywin32 customtkinter pystray Pillow
```

## Usage

```bash
python stopspotv1.py
```

### Build Executable

```bash
pyinstaller --onedir --noconsole --name="SpotifyAutoController" stopspotv1.py
```

### CLI Options

- `--version` - Version info
- `--test` - Resource usage test

## License

MIT
