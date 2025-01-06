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
                    print(f"{time.strftime('%H:%M:%S')} - Initializing audio session manager...", flush=True)
                    self._cleanup()
                    # Retry COM initialization up to 3 times
                    for attempt in range(3):
                        try:
                            print(f"{time.strftime('%H:%M:%S')} - Initialization attempt {attempt + 1}", flush=True)
                            self._devices = pycaw.AudioUtilities.GetSpeakers()
                            print(f"{time.strftime('%H:%M:%S')} - Retrieved speakers: {self._devices}", flush=True)
                            
                            self._interface = self._devices.Activate(
                                pycaw.IAudioSessionManager2._iid_, 
                                CLSCTX_ALL, 
                                None
                            )
                            print(f"{time.strftime('%H:%M:%S')} - Activated IAudioSessionManager2 interface: {self._interface}", flush=True)
                            
                            self._session_manager = cast(self._interface, POINTER(pycaw.IAudioSessionManager2))
                            print(f"{time.strftime('%H:%M:%S')} - Casted to IAudioSessionManager2: {self._session_manager}", flush=True)
                            
                            self._sessions = self._session_manager.GetSessionEnumerator()
                            print(f"{time.strftime('%H:%M:%S')} - Retrieved session enumerator: {self._sessions}", flush=True)
                            
                            self._last_check = current_time
                            self._initialized = True
                            print(f"{time.strftime('%H:%M:%S')} - Initialization successful", flush=True)
                            break
                        except Exception as e:
                            print(f"{time.strftime('%H:%M:%S')} - Initialization attempt {attempt + 1} failed: {e}", flush=True)
                            time.sleep(0.1)
                    
                    if not self._initialized:
                        raise Exception("Failed to initialize COM objects after 3 attempts")
            except Exception as e:
                print(f"{time.strftime('%H:%M:%S')} - Critical initialization error: {e}", flush=True)
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

    def __del__(self):
        self._cleanup()
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

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
                        
                    is_spotify = 'spotify' in process_name
                    
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
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'spotify' in proc.info['name'].lower():
            return proc
    return None

def focus_spotify():
    try:
        import win32gui
        import win32process
        import win32con
        
        def callback(hwnd, data):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            
            tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if 'spotify' in win32gui.GetWindowText(hwnd).lower():
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return False
            return True
            
        win32gui.EnumWindows(callback, None)
        time.sleep(0.1)  # Give Windows time to focus
        return True
    except Exception as e:
        print(f"Error focusing Spotify: {e}")
        return False

def pause_spotify():
    try:
        pyautogui.PAUSE = 0.1
        focus_spotify()
        pyautogui.press('space')
        print("Paused Spotify")
        return True
    except Exception as e:
        print(f"Error pausing Spotify: {e}")
        return False

def play_spotify():
    try:
        pyautogui.PAUSE = 0.1
        focus_spotify()
        pyautogui.press('space')
        print("Resumed Spotify")
        return True
    except Exception as e:
        print(f"Error resuming Spotify: {e}")
        return False

def check_audio_sessions_helper(audio_manager, check_spotify, queue):
    try:
        result = audio_manager.check_audio_sessions(check_spotify)
        queue.put(result)
    except Exception as e:
        queue.put(e)

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

    def get_audio_session_result(audio_manager, check_spotify):
        queue = Queue()
        p = Process(target=check_audio_sessions_helper, args=(audio_manager, check_spotify, queue))
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
                        other_apps_playing = get_audio_session_result(audio_manager, check_spotify=False)
                        time.sleep(0.1)
                        print(f"{time.strftime('%H:%M:%S')} - Checking Spotify playing status", flush=True)
                        spotify_playing = get_audio_session_result(audio_manager, check_spotify=True)
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
                other_apps_playing = get_audio_session_result(audio_manager, check_spotify=False)
                time.sleep(0.1)
                spotify_playing = get_audio_session_result(audio_manager, check_spotify=True)
                
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
                                last_action_time = current_time
                    elif spotify_was_playing and not other_apps_playing:
                        print(f"\n{time.strftime('%H:%M:%S')} - No other audio, resuming Spotify...", flush=True)
                        if play_spotify():
                            spotify_was_playing = False
                            last_action_time = current_time
            
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
