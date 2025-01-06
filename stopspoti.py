import psutil
import time
from comtypes import *
import pycaw.pycaw as pycaw
import pyautogui
import pythoncom
# from threading import Lock
from threading import RLock
import sys
import signal
import atexit
import multiprocessing
from multiprocessing import Process, Queue

# Add audio session state constants
AUDCLNT_SESSIONSTATE_ACTIVE = 1
AUDCLNT_SESSIONSTATE_INACTIVE = 2
AUDCLNT_SESSIONSTATE_EXPIRED = 3

from comtypes import CLSCTX_ALL  # Ensure CLSCTX_ALL is imported

class AudioSessionManager:
    def __init__(self):
        self._lock = RLock()  # Initialize RLock
        self._devices = None
        self._interface = None
        self._session_manager = None
        self._sessions = None
        self._last_check = 0
        self._cache_timeout = 1  # Refresh cache every 1 second
        self._debug = True  # Add debug flag
        self._peak_threshold = 0.0005  # Increased threshold for better detection
        self._initialized = False
        self._com_initialized = False
        self._last_log_time = 0  # Add last log time
        self._log_interval = 1  # Set log interval to 1 second
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
            print(f"{time.strftime('%H:%M:%S')} - COM not initialized", flush=True)
            return False

        try:
            print(f"{time.strftime('%H:%M:%S')} - Initializing audio sessions...", flush=True)
            self._initialize_if_needed()
            if not self._initialized or not self._sessions:
                print(f"{time.strftime('%H:%M:%S')} - Audio sessions not initialized properly", flush=True)
                return False

            count = self._sessions.GetCount()
            print(f"{time.strftime('%H:%M:%S')} - Number of audio sessions: {count}", flush=True)
            found_active = False
            
            # Only ignore system processes
            IGNORED_PROCESSES = {'system idle process', 'system', 'explorer.exe'}  # Using set for faster lookup
            
            # Define identifiers for Spotify and Spotify Premium
            SPOTIFY_IDENTIFIERS = ['spotify', 'spotify premium']
            
            if self._debug:
                print(f"\n{'=' * 50}", flush=True)
                print(f"Checking {'Spotify' if check_spotify else 'other apps'} audio:", flush=True)
            
            for i in range(count):
                print(f"{time.strftime('%H:%M:%S')} - Checking session {i+1}/{count}", flush=True)
                session = None
                audio_session = None
                volume = None
                meter = None
                
                try:
                    session = self._sessions.GetSession(i)
                    print(f"{time.strftime('%H:%M:%S')} - Retrieved session object for session {i+1}", flush=True)
                    if not session:
                        print(f"{time.strftime('%H:%M:%S')} - Session {i+1} is None", flush=True)
                        continue

                    try:
                        audio_session = session.QueryInterface(pycaw.IAudioSessionControl2)
                        print(f"{time.strftime('%H:%M:%S')} - Queried IAudioSessionControl2 for session {i+1}", flush=True)
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Failed to query IAudioSessionControl2 for session {i+1}: {e}", flush=True)
                        continue

                    try:
                        volume = session.QueryInterface(pycaw.ISimpleAudioVolume)
                        print(f"{time.strftime('%H:%M:%S')} - Queried ISimpleAudioVolume for session {i+1}", flush=True)
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Failed to query ISimpleAudioVolume for session {i+1}: {e}", flush=True)

                    try:
                        meter = session.QueryInterface(pycaw.IAudioMeterInformation)
                        print(f"{time.strftime('%H:%M:%S')} - Queried IAudioMeterInformation for session {i+1}", flush=True)
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Failed to query IAudioMeterInformation for session {i+1}: {e}", flush=True)
                    
                    try:
                        process_id = audio_session.GetProcessId()
                        print(f"{time.strftime('%H:%M:%S')} - Process ID: {process_id}", flush=True)
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Failed to get Process ID for session {i+1}: {e}", flush=True)
                        continue

                    try:
                        process = psutil.Process(process_id)
                        print(f"{time.strftime('%H:%M:%S')} - Retrieved process for PID {process_id}", flush=True)
                        process_name = process.name().lower()
                        print(f"{time.strftime('%H:%M:%S')} - Process name: {process_name}", flush=True)
                    except psutil.NoSuchProcess:
                        print(f"{time.strftime('%H:%M:%S')} - No such process with PID: {process_id}", flush=True)
                        continue
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Error retrieving process for PID {process_id}: {e}", flush=True)
                        continue
                    
                    if process_name in IGNORED_PROCESSES:
                        print(f"{time.strftime('%H:%M:%S')} - Ignored process: {process_name}", flush=True)
                        continue
                        
                    # Check for both Spotify and Spotify Premium
                    is_spotify = any(identifier in process_name for identifier in SPOTIFY_IDENTIFIERS)
                    
                    if check_spotify != is_spotify:
                        print(f"{time.strftime('%H:%M:%S')} - Skipping {'Spotify' if not is_spotify else 'other app'}: {process_name}", flush=True)
                        continue
                        
                    try:
                        state = audio_session.GetState()
                        peak = meter.GetPeakValue() if meter else 0
                        print(f"{time.strftime('%H:%M:%S')} - State: {state}, Peak: {peak:.6f}", flush=True)
                    except Exception as e:
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
                        print(f"{time.strftime('%H:%M:%S')} - Released COM objects for session {i+1}", flush=True)
            
            if self._debug:
                print(f"\nFound active audio: {found_active}", flush=True)
                print('=' * 50, flush=True)
            return found_active
            
        except Exception as e:
            print(f"{time.strftime('%H:%M:%S')} - Error checking audio sessions: {e}", flush=True)
            self._cleanup()
            return False

def get_spotify_process():
    # Define identifiers for Spotify and Spotify Premium
    SPOTIFY_IDENTIFIERS = ['spotify', 'spotify premium']
    
    for proc in psutil.process_iter(['name']):
        if proc.info['name']:
            proc_name_lower = proc.info['name'].lower()
            if any(identifier in proc_name_lower for identifier in SPOTIFY_IDENTIFIERS):
                return proc
    return None

def focus_spotify():
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
            window_title = win32gui.GetWindowText(hwnd).lower()
            if window_pid == pid and 'spotify' in window_title:
                spotify_hwnd = hwnd
                return False
            return True

        spotify_process = get_spotify_process()
        if spotify_process:
            try:
                win32gui.EnumWindows(callback, spotify_process.pid)
            except Exception as e:
                if isinstance(e, OSError):
                    if hasattr(e, 'winerror'):
                        if e.winerror == 126:
                            print(f"{time.strftime('%H:%M:%S')} - EnumWindows failed: {e}", flush=True)
                            print("Please ensure that pywin32 is correctly installed and all dependencies are present.", flush=True)
                        elif e.winerror == 2:
                            print(f"{time.strftime('%H:%M:%S')} - EnumWindows failed: {e}", flush=True)
                            print("The system cannot find the file specified. Please verify your pywin32 installation.", flush=True)
                        else:
                            print(f"{time.strftime('%H:%M:%S')} - EnumWindows exception: {e}", flush=True)
                    else:
                        print(f"{time.strftime('%H:%M:%S')} - EnumWindows exception: {e}", flush=True)
                else:
                    print(f"{time.strftime('%H:%M:%S')} - EnumWindows exception: {e}", flush=True)
                return False

            if spotify_hwnd:
                fg_window = win32gui.GetForegroundWindow()
                current_thread_id = win32api.GetCurrentThreadId()  # Updated to use win32api
                fg_thread_id, _ = win32process.GetWindowThreadProcessId(fg_window)
                target_thread_id = win32process.GetWindowThreadProcessId(spotify_hwnd)[0]

                # Attach threads to allow SetForegroundWindow to succeed
                win32process.AttachThreadInput(current_thread_id, fg_thread_id, True)
                win32process.AttachThreadInput(current_thread_id, target_thread_id, True)

                win32gui.ShowWindow(spotify_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(spotify_hwnd)

                # Detach threads after setting foreground
                win32process.AttachThreadInput(current_thread_id, fg_thread_id, False)
                win32process.AttachThreadInput(current_thread_id, target_thread_id, False)

                print(f"{time.strftime('%H:%M:%S')} - Spotify window focused.", flush=True)
                return True
            else:
                print(f"{time.strftime('%H:%M:%S')} - Spotify window not found.", flush=True)
                # Additional debugging: List all visible windows
                def list_all_windows():
                    def enum_callback(hwnd, results):
                        if win32gui.IsWindowVisible(hwnd):
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            title = win32gui.GetWindowText(hwnd)
                            results.append((hwnd, pid, title))
                        return True
                    windows = []
                    win32gui.EnumWindows(enum_callback, windows)
                    for hwnd, pid, title in windows:
                        print(f"Window Handle: {hwnd}, PID: {pid}, Title: {title}", flush=True)

                print("Listing all windows for debugging:", flush=True)
                list_all_windows()
                return False
        else:
            print(f"{time.strftime('%H:%M:%S')} - Spotify process not found.", flush=True)
            return False
    except ImportError as e:
        print(f"Error importing win32 modules: {e}", flush=True)
        print("Please install pywin32 and ensure it is correctly configured.", flush=True)
        return False
    except Exception as e:
        print(f"Error focusing Spotify: {e}", flush=True)
        return False

def pause_spotify():
    try:
        pyautogui.PAUSE = 0.1
        if focus_spotify():
            pyautogui.press('space')
            print("Paused Spotify")
            return True
        else:
            print("Failed to focus Spotify before pausing.")
            return False
    except Exception as e:
        print(f"Error pausing Spotify: {e}")
        return False

def play_spotify():
    try:
        pyautogui.PAUSE = 0.1
        if focus_spotify():
            pyautogui.press('space')
            print("Resumed Spotify")
            return True
        else:
            print("Failed to focus Spotify before resuming.")
            return False
    except Exception as e:
        print(f"Error resuming Spotify: {e}")
        return False

def check_audio_sessions_helper(check_spotify, queue):
    try:
        audio_manager = AudioSessionManager()  # Instantiate within the child process
        result = audio_manager.check_audio_sessions(check_spotify)
        queue.put(result)
        audio_manager.close()  # Explicitly call close to cleanup
    except Exception as e:
        queue.put(e)

def get_audio_session_result(check_spotify):
    queue = Queue()
    p = Process(target=check_audio_sessions_helper, args=(check_spotify, queue))
    p.start()
    p.join(timeout=5)  # Set a timeout (e.g., 5 seconds)
    if p.is_alive():
        p.terminate()
        print(f"{time.strftime('%H:%M:%S')} - check_audio_sessions timed out", flush=True)
        return False  # Assume no active audio if timeout
    if not queue.empty():
        result = queue.get()
        if isinstance(result, Exception):
            print(f"{time.strftime('%H:%M:%S')} - Error in audio session check: {result}", flush=True)
            return False
        return result
    return False

def main():
    audio_manager = None
    
    def cleanup():
        nonlocal audio_manager
        if audio_manager:
            print("\nCleaning up resources...", flush=True)
            del audio_manager
            audio_manager = None

    def signal_handler(signum, frame):
        print("\nSignal received, shutting down...", flush=True)
        cleanup()
        sys.exit(0)

    # Register cleanup handlers
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    atexit.register(cleanup)

    print(f"\n{time.strftime('%H:%M:%S')} - Starting Spotify Audio Monitor...", flush=True)
    print("Press Ctrl+C to exit", flush=True)
    
    spotify_was_playing = False
    last_action_time = 0
    action_cooldown = 0.5  # Reduced cooldown
    last_status_time = 0
    status_interval = 5  # Show status every 5 seconds

    while True:
        try:
            if audio_manager is None:
                audio_manager = AudioSessionManager()
                if not audio_manager._com_initialized:
                    raise Exception("Failed to initialize COM")
                print(f"{time.strftime('%H:%M:%S')} - Audio manager initialized", flush=True)
                
                # Force initial state check
                spotify_process = get_spotify_process()
                if spotify_process:
                    print(f"{time.strftime('%H:%M:%S')} - Found Spotify process", flush=True)
                    try:
                        print(f"{time.strftime('%H:%M:%S')} - Checking other apps playing status", flush=True)
                        other_apps_playing = get_audio_session_result(check_spotify=False)  # Updated call
                        time.sleep(0.1)
                        print(f"{time.strftime('%H:%M:%S')} - Checking Spotify playing status", flush=True)
                        spotify_playing = get_audio_session_result(check_spotify=True)  # Updated call
                        print(f"{time.strftime('%H:%M:%S')} - Initial state:", flush=True)
                        print(f"  Spotify playing: {spotify_playing}", flush=True)
                        print(f"  Other audio playing: {other_apps_playing}\n", flush=True)
                    except Exception as e:
                        print(f"{time.strftime('%H:%M:%S')} - Error during initial audio check: {e}", flush=True)
                        raise
                else:
                    print(f"{time.strftime('%H:%M:%S')} - Spotify is not running\n", flush=True)
                
                time.sleep(0.5)
            
            current_time = time.time()
            spotify_process = get_spotify_process()
            
            if spotify_process:
                other_apps_playing = get_audio_session_result(check_spotify=False)  # Updated call
                time.sleep(0.1)
                spotify_playing = get_audio_session_result(check_spotify=True)  # Updated call
                
                # Add clear debug start tag
                if audio_manager._debug:
                    print("=== DEBUG START ===", flush=True)
                    print(f"{time.strftime('%H:%M:%S')} - other_apps_playing: {other_apps_playing}", flush=True)
                    print(f"{time.strftime('%H:%M:%S')} - spotify_playing: {spotify_playing}", flush=True)
                
                # Show periodic status
                if current_time - last_status_time >= status_interval:
                    print(f"\n{time.strftime('%H:%M:%S')} - Status:", flush=True)
                    print(f"  Spotify playing: {spotify_playing}", flush=True)
                    print(f"  Other audio: {other_apps_playing}", flush=True)
                    print(f"  Was playing: {spotify_was_playing}\n", flush=True)
                    last_status_time = current_time
                
                if current_time - last_action_time >= action_cooldown:
                    if other_apps_playing and spotify_playing:
                        if not spotify_was_playing:
                            print(f"\n{time.strftime('%H:%M:%S')} - Other audio detected, pausing Spotify...", flush=True)
                            if pause_spotify():
                                spotify_was_playing = True
                                print(f"{time.strftime('%H:%M:%S')} - Spotify paused successfully", flush=True)
                                last_action_time = current_time
                    elif spotify_was_playing and not other_apps_playing:
                        print(f"\n{time.strftime('%H:%M:%S')} - No other audio, resuming Spotify...", flush=True)
                        if play_spotify():
                            spotify_was_playing = False
                            print(f"{time.strftime('%H:%M:%S')} - Spotify resumed successfully", flush=True)
                            last_action_time = current_time
                
                # Add clear debug end tag
                if audio_manager._debug:
                    print("=== DEBUG END ===", flush=True)
            
            time.sleep(0.2)
            sys.stdout.flush()  # Force flush output
            
        except Exception as e:
            print(f"\n{time.strftime('%H:%M:%S')} - Error: {e}", flush=True)
            if audio_manager:
                del audio_manager
                audio_manager = None
            time.sleep(1)

if __name__ == "__main__":
    main()
