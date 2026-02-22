# StopSpoti 🎧

**A lightweight Windows utility that automatically pauses Spotify whenever you watch a video, play a game, or trigger audio in another app—and resumes your music as soon as they stop.**

---

## 💡 Why use this?
If you listen to Spotify in the background while browsing the web or gaming, it's frustrating to constantly manually pause your music every time you open a YouTube video, click a Twitter clip, or enter a game cutscene. **StopSpoti** solves this by detecting when *any other application* starts making noise, instantly pausing your Spotify for you.

When the other app goes quiet, StopSpoti waits a polite 1.5 seconds and seamlessly resumes your music!

## ✨ Features
- **Zero-Configuration:** Start the script and click "Start Monitoring" in the GUI.
- **Intelligent Resumption:** It knows the difference between a pause in dialogue and a finished video, using a smart 1.5s silence threshold to prevent stuttering.
- **Resource Safe:** Optimized for efficiency. By utilizing garbage collection techniques and safely resetting Windows COM pointers every 5 minutes, it will never build up memory leaks, even if left running for months.
- **Customizable:** You can explicitly define which programs (like Discord or OBS) it should ignore.

---

## 🚀 Installation & Usage

1. **Prerequisites:** Ensure you have Windows and Python 3.x installed.
2. **Clone the repository:**
   ```bash
   git clone https://github.com/Shreeniwas1/stopspoti.git
   cd stopspoti
   ```
3. **Install dependencies:**
   ```bash
   pip install psutil comtypes pycaw pywin32 customtkinter pystray Pillow
   ```
4. **Run the application:**
   ```bash
   python stopspotiv1.py
   ```
5. Click **"Start Monitoring"** in the window that appears!

---

## ❓ Why no .exe?
I have explicitly dropped support for standalone compiled `.exe` files (such as those generated via PyInstaller).  

**Reasoning:**
Interacting with the deep Windows Audio Session APIs (pycaw/comtypes) requires careful memory and COM thread management. When compiling this program into an executable—especially using `--noconsole` flags—Windows introduces opaque threading states and hides `sys.stdout` streams, which leads to intermittent, hard-to-pinpoint crashes and rapid toggling bugs that simply don't exist when interpreting the code normally. 

To ensure the best, most stable experience for audio polling, running the raw Python script via the terminal is the only officially supported method.

---

## 📝 License
MIT License
