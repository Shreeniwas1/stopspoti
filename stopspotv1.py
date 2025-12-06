import psutil
import time
import os
import ctypes
from comtypes import *
import pycaw.pycaw as pycaw
import pythoncom
# from threading import Lock
from threading import RLock
import sys
import signal
import atexit
import threading
import customtkinter as ctk
from PIL import Image
import pystray

# Add audio session state constants
AUDCLNT_SESSIONSTATE_ACTIVE = 1
AUDCLNT_SESSIONSTATE_INACTIVE = 2
AUDCLNT_SESSIONSTATE_EXPIRED = 3

from comtypes import CLSCTX_ALL  # Ensure CLSCTX_ALL is imported

# Add audio session state constants
AUDCLNT_SESSIONSTATE_ACTIVE = 1
AUDCLNT_SESSIONSTATE_INACTIVE = 2
AUDCLNT_SESSIONSTATE_EXPIRED = 3

from comtypes import CLSCTX_ALL  # Ensure CLSCTX_ALL is imported

class AudioSessionManager:
    def __init__(self, peak_threshold=0.0005, cache_timeout=2, log_interval=5, debug=True, ignored_processes=None):
        self._lock = RLock()  # Initialize RLock
        self._devices = None
        self._interface = None
        self._session_manager = None
        self._sessions = None
        self._last_check = 0
        self._cache_timeout = cache_timeout
        self._debug = debug
        self._peak_threshold = peak_threshold
        self._initialized = False
        self._com_initialized = False
        self._last_log_time = 0
        self._log_interval = log_interval
        self._ignored_processes = set(proc.lower() for proc in (ignored_processes or [
            'system idle process', 'system', 'explorer.exe',
            'FxSound.exe', 'FxSound', 'fxsound.exe', 
            'obs64.exe', 'obs32.exe', 'obs.exe', 'obs-browser-page.exe',
            'SnippingTool.exe', 'ScreenClippingHost.exe',
            'ScreenClipping.exe',
            'audiodg.exe',  # Windows Audio Device Graph - proxy for all audio, ignore it
        ]))
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception as e:
            print(f"COM initialization failed: {e}")

    def _safe_release(self, com_object):
        if com_object:
            try:
                com_object.Release()
            except Exception:
                pass  # Silently handle release errors

    def _initialize_if_needed(self):
        with self._lock:
            try:
                current_time = time.time()
                if not self._initialized or current_time - self._last_check > self._cache_timeout:
                    if self._debug and (current_time - self._last_log_time) > self._log_interval:
                        print(f"{time.strftime('%H:%M:%S')} - Initializing audio session manager...", flush=True)
                        self._last_log_time = current_time
                    self._cleanup()
                    # Retry COM initialization up to 3 times
                    for attempt in range(3):
                        try:
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Initialization attempt {attempt + 1}", flush=True)
                                self._last_log_time = current_time
                            self._devices = pycaw.AudioUtilities.GetSpeakers()
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Retrieved speakers: {self._devices}", flush=True)
                                self._last_log_time = current_time
                            
                            self._interface = self._devices.Activate(
                                pycaw.IAudioSessionManager2._iid_, 
                                CLSCTX_ALL, 
                                None
                            )
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Activated IAudioSessionManager2 interface: {self._interface}", flush=True)
                                self._last_log_time = current_time
                            
                            self._session_manager = cast(self._interface, POINTER(pycaw.IAudioSessionManager2))
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Casted to IAudioSessionManager2: {self._session_manager}", flush=True)
                                self._last_log_time = current_time
                            
                            self._sessions = self._session_manager.GetSessionEnumerator()
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Retrieved session enumerator: {self._sessions}", flush=True)
                                self._last_log_time = current_time
                            
                            self._last_check = current_time
                            self._initialized = True
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Initialization successful", flush=True)
                                self._last_log_time = current_time
                            break
                        except Exception as e:
                            if self._debug and (current_time - self._last_log_time) > self._log_interval:
                                print(f"{time.strftime('%H:%M:%S')} - Initialization attempt {attempt + 1} failed: {e}", flush=True)
                                self._last_log_time = current_time
                            time.sleep(0.1)
                    
                    if not self._initialized:
                        raise Exception("Failed to initialize COM objects after 3 attempts")
            except Exception as e:
                if self._debug and (current_time - self._last_log_time) > self._log_interval:
                    print(f"{time.strftime('%H:%M:%S')} - Critical initialization error: {e}", flush=True)
                    self._last_log_time = current_time
                self._cleanup()
                raise

    def _cleanup(self):
        with self._lock:
            try:
                if self._sessions:
                    for i in range(self._sessions.GetCount()):
                        try:
                            session = self._sessions.GetSession(i)
                            self._safe_release(session)
                        except Exception:
                            pass

                self._safe_release(self._sessions)
                self._safe_release(self._session_manager)
                self._safe_release(self._interface)
            finally:
                self._sessions = None
                self._session_manager = None
                self._interface = None
                self._devices = None
                self._initialized = False

    def close(self):
        try:
            self._cleanup()
            if self._com_initialized:
                pythoncom.CoUninitialize()
        except Exception:
            pass  # Silently ignore exceptions during cleanup

    def check_audio_sessions(self, check_spotify=False):
        if not self._com_initialized:
            if self._debug:
                print(f"{time.strftime('%H:%M:%S')} - COM not initialized", flush=True)
            return False

        try:
            if self._debug:
                print(f"{time.strftime('%H:%M:%S')} - Initializing audio sessions...", flush=True)
            self._initialize_if_needed()
            if not self._initialized or not self._sessions:
                if self._debug:
                    print(f"{time.strftime('%H:%M:%S')} - Audio sessions not initialized properly", flush=True)
                return False

            # Only ignore system processes
            count = self._sessions.GetCount()
            if self._debug:
                print(f"{time.strftime('%H:%M:%S')} - Number of audio sessions: {count}", flush=True)
            found_active = False
            
            # Define identifiers for Spotify and Spotify Premium
            SPOTIFY_IDENTIFIERS = ['spotify', 'spotify premium']
            
            if self._debug:
                print(f"\n{'=' * 50}", flush=True)
                print(f"Checking {'Spotify' if check_spotify else 'other apps'} audio:", flush=True)
            
            for i in range(count):
                if self._debug:
                    print(f"{time.strftime('%H:%M:%S')} - Checking session {i+1}/{count}", flush=True)
                session = None
                audio_session = None
                volume = None
                meter = None
                
                try:
                    session = self._sessions.GetSession(i)
                    if self._debug:
                        print(f"{time.strftime('%H:%M:%S')} - Retrieved session object for session {i+1}", flush=True)
                    if not session:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Session {i+1} is None", flush=True)
                        continue

                    try:
                        audio_session = session.QueryInterface(pycaw.IAudioSessionControl2)
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Queried IAudioSessionControl2 for session {i+1}", flush=True)
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Failed to query IAudioSessionControl2 for session {i+1}: {e}", flush=True)
                        continue

                    try:
                        volume = session.QueryInterface(pycaw.ISimpleAudioVolume)
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Queried ISimpleAudioVolume for session {i+1}", flush=True)
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Failed to query ISimpleAudioVolume for session {i+1}: {e}", flush=True)

                    try:
                        meter = session.QueryInterface(pycaw.IAudioMeterInformation)
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Queried IAudioMeterInformation for session {i+1}", flush=True)
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Failed to query IAudioMeterInformation for session {i+1}: {e}", flush=True)
                    
                    try:
                        process_id = audio_session.GetProcessId()
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Process ID: {process_id}", flush=True)
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Failed to get Process ID for session {i+1}: {e}", flush=True)
                        continue

                    try:
                        process = psutil.Process(process_id)
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Retrieved process for PID {process_id}", flush=True)
                        process_name = process.name().lower()
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Process name: {process_name}", flush=True)
                    except psutil.NoSuchProcess:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - No such process with PID: {process_id}", flush=True)
                        continue
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Error retrieving process for PID {process_id}: {e}", flush=True)
                        continue
                    
                    if process_name in self._ignored_processes:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Ignored process: {process_name}", flush=True)
                        continue
                        
                    # Check for both Spotify and Spotify Premium
                    is_spotify = any(identifier in process_name for identifier in SPOTIFY_IDENTIFIERS)
                    
                    if check_spotify != is_spotify:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Skipping {'Spotify' if not is_spotify else 'other app'}: {process_name}", flush=True)
                        continue
                        
                    try:
                        state = audio_session.GetState()
                        peak = meter.GetPeakValue() if meter else 0
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - State: {state}, Peak: {peak:.6f}", flush=True)
                    except Exception as e:
                        if self._debug:
                            print(f"{time.strftime('%H:%M:%S')} - Failed to get state or peak for session {i+1}: {e}", flush=True)
                        continue
                    
                    if self._debug and (peak > self._peak_threshold or state == AUDCLNT_SESSIONSTATE_ACTIVE):
                        print(f"{time.strftime('%H:%M:%S')} - {process_name}:", flush=True)
                        print(f"  Peak: {peak:.6f} | State: {state}", flush=True)
                    
                    # Simplified audio detection - just check state and peak
                    if state == AUDCLNT_SESSIONSTATE_ACTIVE and peak > self._peak_threshold:
                        found_active = True
                        if self._debug:
                            print(f"  ** ACTIVE AUDIO **", flush=True)
                        break
                        
                finally:
                    for obj in (meter, volume, audio_session, session):
                        self._safe_release(obj)
                    if self._debug:
                        print(f"{time.strftime('%H:%M:%S')} - Released COM objects for session {i+1}", flush=True)
            
            if self._debug:
                print(f"\nFound active audio: {found_active}", flush=True)
                print('=' * 50, flush=True)
            return found_active
            
        except Exception as e:
            if self._debug:
                print(f"{time.strftime('%H:%M:%S')} - Error checking audio sessions: {e}", flush=True)
            self._cleanup()
            return False

# Modify to store all Spotify PIDs at the start
def get_spotify_processes():
    # Define identifiers for Spotify and Spotify Premium
    SPOTIFY_IDENTIFIERS = ['spotify', 'spotify premium', 'spotify.exe', 'Spotify.exe', 'Spotify Premium']
    
    spotify_processes = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name']:
            proc_name_lower = proc.info['name'].lower()
            if any(identifier in proc_name_lower for identifier in SPOTIFY_IDENTIFIERS):
                spotify_processes.append(proc)
    return spotify_processes

# Initialize SPOTIFY_PIDS
SPOTIFY_PIDS = [proc.pid for proc in get_spotify_processes()]

def get_spotify_process():
    # Define identifiers for Spotify and Spotify Premium
    SPOTIFY_IDENTIFIERS = ['spotify', 'spotify premium','spotify.exe','Spotify.exe','Spotify Premium']
    
    for proc in psutil.process_iter(['name']):
        if proc.info['name']:
            proc_name_lower = proc.info['name'].lower()
            if any(identifier in proc_name_lower for identifier in SPOTIFY_IDENTIFIERS):
                return proc
    return None

# Update focus_spotify to use stored PIDs
def focus_spotify(debug=False):
    try:
        import win32gui
        import win32process
        import win32con
        import win32api  # Ensure win32api is imported

        spotify_hwnd = None

        def callback(hwnd, pid):
            nonlocal spotify_hwnd
            if not win32gui.IsWindowVisible(hwnd):
                return True
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid in SPOTIFY_PIDS:
                spotify_hwnd = hwnd
                return False
            return True

        for pid in SPOTIFY_PIDS:
            try:
                win32gui.EnumWindows(callback, pid)
                if spotify_hwnd:
                    fg_window = win32gui.GetForegroundWindow()
                    current_thread_id = win32api.GetCurrentThreadId()
                    fg_thread_id, _ = win32process.GetWindowThreadProcessId(fg_window)
                    target_thread_id = win32process.GetWindowThreadProcessId(spotify_hwnd)[0]

                    win32process.AttachThreadInput(current_thread_id, fg_thread_id, True)
                    win32process.AttachThreadInput(current_thread_id, target_thread_id, True)

                    win32gui.ShowWindow(spotify_hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(spotify_hwnd)

                    win32process.AttachThreadInput(current_thread_id, fg_thread_id, False)
                    win32process.AttachThreadInput(current_thread_id, target_thread_id, False)

                    if debug:
                        print(f"{time.strftime('%H:%M:%S')} - Spotify window focused for PID {pid}.", flush=True)
                    return True
            except Exception as e:
                if debug:
                    print(f"{time.strftime('%H:%M:%S')} - EnumWindows failed for PID {pid}: {e}", flush=True)
        if debug:
            print(f"{time.strftime('%H:%M:%S')} - Spotify windows not found for any stored PIDs.", flush=True)
        return False

    except ImportError as e:
        if debug:
            print(f"Error importing win32 modules: {e}", flush=True)
            print("Please install pywin32 and ensure it is correctly configured.", flush=True)
        return False
    except Exception as e:
        if debug:
            print(f"Error focusing Spotify: {e}", flush=True)
        return False

# Windows AppCommand constants for media control
APPCOMMAND_MEDIA_PLAY_PAUSE = 14
APPCOMMAND_MEDIA_PLAY = 46
APPCOMMAND_MEDIA_PAUSE = 47
WM_APPCOMMAND = 0x0319

_debug_mode = False  # Global debug flag for standalone functions

def get_spotify_hwnd():
    """Get Spotify's main window handle without focusing it"""
    try:
        import win32gui
        import win32process
        
        spotify_hwnd = None
        
        def callback(hwnd, _):
            nonlocal spotify_hwnd
            if not win32gui.IsWindowVisible(hwnd):
                return True
            try:
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid in SPOTIFY_PIDS:
                    # Check if it's the main Spotify window (has a title)
                    title = win32gui.GetWindowText(hwnd)
                    if title and 'GDI+' not in title:  # Filter out helper windows
                        spotify_hwnd = hwnd
                        return False
            except:
                pass
            return True
        
        win32gui.EnumWindows(callback, None)
        return spotify_hwnd
    except Exception as e:
        if _debug_mode:
            print(f"Error getting Spotify hwnd: {e}", flush=True)
        return None

def send_appcommand_to_spotify(command):
    """Send media command directly to Spotify window without stealing focus"""
    try:
        import win32api
        import win32gui
        
        hwnd = get_spotify_hwnd()
        if hwnd:
            # WM_APPCOMMAND: wParam = hwnd, lParam = command << 16
            lparam = command << 16
            win32api.SendMessage(hwnd, WM_APPCOMMAND, hwnd, lparam)
            return True
        else:
            if _debug_mode:
                print(f"{time.strftime('%H:%M:%S')} - Spotify window not found", flush=True)
            return False
    except Exception as e:
        if _debug_mode:
            print(f"Error sending appcommand: {e}", flush=True)
        return False

def pause_spotify():
    """Pause Spotify by sending APPCOMMAND directly to its window"""
    try:
        if send_appcommand_to_spotify(APPCOMMAND_MEDIA_PAUSE):
            if _debug_mode:
                print(f"{time.strftime('%H:%M:%S')} - Paused Spotify via AppCommand", flush=True)
            return True
        # Fallback: try play/pause toggle
        if send_appcommand_to_spotify(APPCOMMAND_MEDIA_PLAY_PAUSE):
            if _debug_mode:
                print(f"{time.strftime('%H:%M:%S')} - Paused Spotify via Play/Pause toggle", flush=True)
            return True
        return False
    except Exception as e:
        if _debug_mode:
            print(f"Error pausing Spotify: {e}", flush=True)
        return False

def play_spotify():
    """Resume Spotify by sending APPCOMMAND directly to its window"""
    try:
        if send_appcommand_to_spotify(APPCOMMAND_MEDIA_PLAY):
            if _debug_mode:
                print(f"{time.strftime('%H:%M:%S')} - Resumed Spotify via AppCommand", flush=True)
            return True
        # Fallback: try play/pause toggle
        if send_appcommand_to_spotify(APPCOMMAND_MEDIA_PLAY_PAUSE):
            if _debug_mode:
                print(f"{time.strftime('%H:%M:%S')} - Resumed Spotify via Play/Pause toggle", flush=True)
            return True
        return False
    except Exception as e:
        if _debug_mode:
            print(f"Error resuming Spotify: {e}", flush=True)
        return False

class SpotifyControllerGUI:
    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  # We'll customize colors
        
        self.root = ctk.CTk()
        self.root.title("Spotify Auto Controller")
        self.root.geometry("600x700")
        self.root.configure(bg="#000000")  # Black background
        
        # Set window icon if available
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass  # Icon not found, continue without
        
        # Custom colors
        self.bg_color = "#000000"  # Black
        self.fg_color = "#00FF00"  # Green
        self.accent_color = "#800080"  # Purple
        
        # Settings
        self.peak_threshold = ctk.DoubleVar(value=0.0005)
        self.cache_timeout = ctk.IntVar(value=2)
        self.log_interval = ctk.IntVar(value=5)
        self.action_cooldown = ctk.DoubleVar(value=2.0)  # Increased cooldown to prevent rapid switching
        self.debug = ctk.BooleanVar(value=False)  # Debug mode off by default
        self.ignored_processes = [
            'system idle process', 'system', 'explorer.exe',
            'FxSound.exe', 'FxSound', 'fxsound.exe', 
            'obs64.exe', 'obs32.exe', 'obs.exe', 'obs-browser-page.exe',
            'SnippingTool.exe', 'ScreenClippingHost.exe',
            'ScreenClipping.exe',
            'audiodg.exe'
        ]
        
        self.monitoring = False
        self.monitor_thread = None
        self.audio_manager = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        title_label = ctk.CTkLabel(self.root, text="Spotify Auto Controller", font=ctk.CTkFont(size=20, weight="bold"), text_color=self.fg_color)
        title_label.pack(pady=10)
        
        # Settings Frame
        settings_frame = ctk.CTkFrame(self.root, fg_color=self.bg_color, border_color=self.accent_color, border_width=2)
        settings_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(settings_frame, text="Settings", font=ctk.CTkFont(size=16, weight="bold"), text_color=self.fg_color).pack(pady=5)
        
        # Peak Threshold
        threshold_frame = ctk.CTkFrame(settings_frame, fg_color=self.bg_color)
        threshold_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(threshold_frame, text="Peak Threshold:", text_color=self.fg_color).pack(side="left")
        threshold_entry = ctk.CTkEntry(threshold_frame, textvariable=self.peak_threshold, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        threshold_entry.pack(side="right", padx=(10,0))
        
        # Cache Timeout
        cache_frame = ctk.CTkFrame(settings_frame, fg_color=self.bg_color)
        cache_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(cache_frame, text="Cache Timeout (s):", text_color=self.fg_color).pack(side="left")
        cache_entry = ctk.CTkEntry(cache_frame, textvariable=self.cache_timeout, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        cache_entry.pack(side="right", padx=(10,0))
        
        # Log Interval
        log_frame = ctk.CTkFrame(settings_frame, fg_color=self.bg_color)
        log_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(log_frame, text="Log Interval (s):", text_color=self.fg_color).pack(side="left")
        log_entry = ctk.CTkEntry(log_frame, textvariable=self.log_interval, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        log_entry.pack(side="right", padx=(10,0))
        
        # Action Cooldown
        cooldown_frame = ctk.CTkFrame(settings_frame, fg_color=self.bg_color)
        cooldown_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(cooldown_frame, text="Action Cooldown (s):", text_color=self.fg_color).pack(side="left")
        cooldown_entry = ctk.CTkEntry(cooldown_frame, textvariable=self.action_cooldown, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        cooldown_entry.pack(side="right", padx=(10,0))
        
        # Debug Mode
        debug_check = ctk.CTkCheckBox(settings_frame, text="Debug Mode", variable=self.debug, fg_color=self.accent_color, text_color=self.fg_color)
        debug_check.pack(pady=5)
        
        # Ignored Processes
        ignored_label = ctk.CTkLabel(settings_frame, text="Ignored Processes:", text_color=self.fg_color)
        ignored_label.pack(pady=5)
        self.ignored_text = ctk.CTkTextbox(settings_frame, height=100, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        self.ignored_text.pack(fill="x", padx=10, pady=5)
        self.ignored_text.insert("0.0", "\n".join(self.ignored_processes))
        
        # Control Buttons
        control_frame = ctk.CTkFrame(self.root, fg_color=self.bg_color, border_color=self.accent_color, border_width=2)
        control_frame.pack(pady=10, padx=20, fill="x")
        
        self.start_button = ctk.CTkButton(control_frame, text="Start Monitoring", command=self.start_monitoring, fg_color=self.accent_color, text_color=self.fg_color)
        self.start_button.pack(side="left", padx=10, pady=10)
        
        self.stop_button = ctk.CTkButton(control_frame, text="Stop Monitoring", command=self.stop_monitoring, state="disabled", fg_color=self.accent_color, text_color=self.fg_color)
        self.stop_button.pack(side="right", padx=10, pady=10)
        
        # Status
        self.status_label = ctk.CTkLabel(self.root, text="Status: Stopped", text_color=self.fg_color)
        self.status_label.pack(pady=10)
        
        # Log
        log_label = ctk.CTkLabel(self.root, text="Log:", text_color=self.fg_color)
        log_label.pack()
        self.log_text = ctk.CTkTextbox(self.root, height=200, fg_color=self.bg_color, text_color=self.fg_color, border_color=self.accent_color)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0,20))
        
    def log(self, message):
        timestamp = time.strftime('%H:%M:%S')
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        
    def start_monitoring(self):
        if self.monitoring:
            return
        
        # Update settings
        self.ignored_processes = [line.strip() for line in self.ignored_text.get("0.0", "end").split("\n") if line.strip()]
        
        self.audio_manager = AudioSessionManager(
            peak_threshold=self.peak_threshold.get(),
            cache_timeout=self.cache_timeout.get(),
            log_interval=self.log_interval.get(),
            debug=self.debug.get(),
            ignored_processes=self.ignored_processes
        )
        
        if not self.audio_manager._com_initialized:
            self.log("Failed to initialize COM")
            return
        
        self.monitoring = True
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_label.configure(text="Status: Running")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
        
        self.log("Monitoring started")
        
    def stop_monitoring(self):
        if not self.monitoring:
            return
        
        self.monitoring = False
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)  # Wait up to 2 seconds for thread to finish
            if self.monitor_thread.is_alive() and self.debug.get():
                print("Warning: Monitoring thread did not stop gracefully", flush=True)
        
        if self.audio_manager:
            try:
                self.audio_manager.close()
            except Exception as e:
                if self.debug.get():
                    print(f"Error closing audio manager: {e}", flush=True)
            self.audio_manager = None
        
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.status_label.configure(text="Status: Stopped")
        
        self.log("Monitoring stopped")
        
    def monitor_loop(self):
        global _debug_mode
        _debug_mode = self.debug.get()  # Set global debug flag
        
        spotify_paused_by_us = False  # Tracks if WE paused Spotify
        last_action_time = 0
        action_cooldown = self.action_cooldown.get()
        silence_start_time = None  # Track when other audio stopped
        silence_threshold = 1.5  # Wait 1.5 seconds of silence before resuming
        
        while self.monitoring:
            try:
                current_time = time.time()
                spotify_process = get_spotify_process()
                
                if spotify_process:
                    # Only check other apps audio
                    other_apps_playing = get_audio_session_result(
                        check_spotify=False,
                        peak_threshold=self.peak_threshold.get(),
                        cache_timeout=self.cache_timeout.get(),
                        log_interval=self.log_interval.get(),
                        debug=self.debug.get(),
                        ignored_processes=self.ignored_processes
                    )
                    
                    # Respect cooldown to prevent rapid switching
                    if current_time - last_action_time >= action_cooldown:
                        if other_apps_playing and not spotify_paused_by_us:
                            # Other app started playing - pause Spotify
                            # Reset silence timer since other audio is playing
                            silence_start_time = None
                            
                            time.sleep(0.1)
                            spotify_playing = get_audio_session_result(
                                check_spotify=True,
                                peak_threshold=self.peak_threshold.get(),
                                cache_timeout=self.cache_timeout.get(),
                                log_interval=self.log_interval.get(),
                                debug=self.debug.get(),
                                ignored_processes=self.ignored_processes
                            )
                            if spotify_playing:
                                if pause_spotify():
                                    spotify_paused_by_us = True
                                    last_action_time = current_time
                                    self.log("Paused Spotify (other audio detected)")
                        
                        elif spotify_paused_by_us and not other_apps_playing:
                            # Other app stopped - track silence duration
                            if silence_start_time is None:
                                silence_start_time = current_time
                                if self.debug.get():
                                    print(f"{time.strftime('%H:%M:%S')} - Other audio stopped, waiting {silence_threshold}s before resuming...", flush=True)
                            
                            # Wait for sustained silence before resuming
                            elif current_time - silence_start_time >= silence_threshold:
                                if self.debug.get():
                                    print(f"{time.strftime('%H:%M:%S')} - Silence confirmed, resuming Spotify...", flush=True)
                                if play_spotify():
                                    spotify_paused_by_us = False
                                    last_action_time = current_time
                                    silence_start_time = None
                                    self.log("Resumed Spotify (other audio stopped)")
                                else:
                                    # Retry on next loop
                                    if self.debug.get():
                                        print(f"{time.strftime('%H:%M:%S')} - Resume failed, will retry...", flush=True)
                        
                        elif other_apps_playing and spotify_paused_by_us:
                            # Other audio still playing, reset silence timer
                            silence_start_time = None
                
                time.sleep(0.5)
                
            except Exception as e:
                error_msg = f"Error in monitoring loop: {e}"
                if self.debug.get():
                    print(error_msg, flush=True)  # Also print to console for debugging
                self.log(error_msg)
                time.sleep(2)  # Wait longer on error to prevent rapid restarts
        
    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            if self.debug.get():
                print(f"Critical GUI error: {e}", flush=True)
            # Don't restart, just exit gracefully
            sys.exit(1)

def get_audio_session_result(check_spotify, peak_threshold=0.0005, cache_timeout=2, log_interval=5, debug=False, ignored_processes=None):
    """Get audio session result using direct call instead of subprocess to avoid cursor loading"""
    try:
        audio_manager = AudioSessionManager(
            peak_threshold=peak_threshold,
            cache_timeout=cache_timeout,
            log_interval=log_interval,
            debug=debug,
            ignored_processes=ignored_processes
        )
        
        if not audio_manager._com_initialized:
            if debug:
                print(f"{time.strftime('%H:%M:%S')} - COM not initialized in direct call", flush=True)
            return False
            
        result = audio_manager.check_audio_sessions(check_spotify)
        audio_manager.close()
        return result
        
    except Exception as e:
        if debug:
            print(f"{time.strftime('%H:%M:%S')} - Direct audio check error: {e}", flush=True)
        return False

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def main():
    # Set process priority to below normal to reduce resource usage
    if hasattr(sys, 'frozen'):
        import win32process
        import win32api
        win32api.SetConsoleCtrlHandler(None, True)
        win32process.SetPriorityClass(win32api.GetCurrentProcess(), win32process.BELOW_NORMAL_PRIORITY_CLASS)

    # Check if running in test mode
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        print("Running in test mode for 10 seconds...")
        test_resource_usage()
    elif len(sys.argv) > 1 and sys.argv[1] == "--gui-only":
        print("Running GUI only (no monitoring)...")
        app = SpotifyControllerGUI()
        app.run()
    elif len(sys.argv) > 1 and sys.argv[1] == "--version":
        print("Spotify Auto Controller v1.0")
        print("Built with PyInstaller")
        sys.exit(0)
    else:
        app = SpotifyControllerGUI()
        app.run()

def test_resource_usage():
    """Test function to run monitoring for 30 seconds and measure resource usage"""
    import psutil as psutil_monitor
    import os
    
    print(f"Process ID: {os.getpid()}")
    print("Starting resource monitoring test...")
    
    audio_manager = AudioSessionManager()
    if not audio_manager._com_initialized:
        print("Failed to initialize COM")
        return
    
    spotify_was_playing = False
    last_action_time = 0
    action_cooldown = 1.0
    
    start_time = time.time()
    end_time = start_time + 30  # Run for 30 seconds
    
    process = psutil_monitor.Process(os.getpid())
    
    while time.time() < end_time:
        try:
            current_time = time.time()
            
            # Get current resource usage
            cpu_percent = process.cpu_percent(interval=0.1)
            memory_mb = process.memory_info().rss / 1024 / 1024
            
            if int(current_time - start_time) % 5 == 0:  # Log every 5 seconds
                print(f"Time: {current_time - start_time:.1f}s | CPU: {cpu_percent:.1f}% | Memory: {memory_mb:.1f} MB")
            
            spotify_process = get_spotify_process()
            
            if spotify_process:
                other_apps_playing = get_audio_session_result(check_spotify=False)
                time.sleep(0.25)
                spotify_playing = get_audio_session_result(check_spotify=True)
                
                if current_time - last_action_time >= action_cooldown:
                    if other_apps_playing and spotify_playing:
                        if not spotify_was_playing:
                            # Don't actually pause, just simulate
                            spotify_was_playing = True
                            last_action_time = current_time
                            print("Would pause Spotify")
                    elif spotify_was_playing and not other_apps_playing:
                        if not other_apps_playing:  # Only resume if no other audio
                            spotify_was_playing = False
                            last_action_time = current_time
                            print("Would resume Spotify")
            
            time.sleep(0.5)
            
        except Exception as e:
            print(f"Error: {e}")
            time.sleep(2)
    
    # Final resource usage
    final_cpu = process.cpu_percent(interval=0.1)
    final_memory = process.memory_info().rss / 1024 / 1024
    print(f"\nFinal resource usage after 30 seconds:")
    print(f"CPU: {final_cpu:.1f}%")
    print(f"Memory: {final_memory:.1f} MB")
    
    audio_manager.close()
    print("Test completed.")

if __name__ == "__main__":
    main()
