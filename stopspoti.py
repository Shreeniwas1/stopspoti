import psutil
import time
from comtypes import *
import pycaw.pycaw as pycaw
import pyautogui
import pythoncom
from threading import Lock
import sys

# Add audio session state constants
AUDCLNT_SESSIONSTATE_ACTIVE = 1
AUDCLNT_SESSIONSTATE_INACTIVE = 2
AUDCLNT_SESSIONSTATE_EXPIRED = 3

class AudioSessionManager:
    def __init__(self):
        self._lock = Lock()
        self._devices = None
        self._interface = None
        self._session_manager = None
        self._sessions = None
        self._last_check = 0
        self._cache_timeout = 1  # Refresh cache every 1 second
        self._debug = True  # Add debug flag
        self._peak_threshold = 0.0005  # Increased threshold for better detection
        self._initialized = False
        # Initialize COM for this thread
        pythoncom.CoInitialize()

    def _safe_release(self, com_object):
        try:
            if com_object:
                com_object.Release()
        except:
            pass

    def _initialize_if_needed(self):
        with self._lock:
            try:
                current_time = time.time()
                if not self._initialized or current_time - self._last_check > self._cache_timeout:
                    self._cleanup()
                    # Retry COM initialization up to 3 times
                    for _ in range(3):
                        try:
                            self._devices = pycaw.AudioUtilities.GetSpeakers()
                            self._interface = self._devices.Activate(
                                pycaw.IAudioSessionManager2._iid_, 
                                CLSCTX_ALL, 
                                None
                            )
                            self._session_manager = cast(self._interface, POINTER(pycaw.IAudioSessionManager2))
                            self._sessions = self._session_manager.GetSessionEnumerator()
                            self._last_check = current_time
                            self._initialized = True
                            break
                        except Exception as e:
                            print(f"Retry COM initialization: {e}")
                            time.sleep(0.1)
                    
                    if not self._initialized:
                        raise Exception("Failed to initialize COM objects")
            except Exception as e:
                print(f"Error in initialization: {e}")
                self._cleanup()
                raise

    def _cleanup(self):
        with self._lock:
            self._safe_release(self._sessions)
            self._safe_release(self._session_manager)
            self._safe_release(self._interface)
            self._sessions = None
            self._session_manager = None
            self._interface = None
            self._initialized = False

    def __del__(self):
        self._cleanup()
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def check_audio_sessions(self, check_spotify=False):
        try:
            self._initialize_if_needed()
            count = self._sessions.GetCount()
            found_active = False
            
            # Only ignore system processes
            IGNORED_PROCESSES = ['system idle process', 'system', 'explorer.exe']
            
            if self._debug:
                print(f"\n{'=' * 50}", flush=True)
                print(f"Checking {'Spotify' if check_spotify else 'other apps'} audio:", flush=True)
            
            for i in range(count):
                session = self._sessions.GetSession(i)
                try:
                    audio_session = session.QueryInterface(pycaw.IAudioSessionControl2)
                    volume = session.QueryInterface(pycaw.ISimpleAudioVolume)
                    meter = session.QueryInterface(pycaw.IAudioMeterInformation)
                    
                    try:
                        process_id = audio_session.GetProcessId()
                        try:
                            process = psutil.Process(process_id)
                            process_name = process.name().lower()
                            
                            if process_name in IGNORED_PROCESSES:
                                continue
                                
                            is_spotify = 'spotify' in process_name
                            
                            if check_spotify != is_spotify:
                                continue
                                
                            state = audio_session.GetState()
                            peak = meter.GetPeakValue()
                            
                            if self._debug and (peak > self._peak_threshold or state == AUDCLNT_SESSIONSTATE_ACTIVE):
                                print(f"{time.strftime('%H:%M:%S')} - {process_name}:", flush=True)
                                print(f"  Peak: {peak:.6f} | State: {state}", flush=True)
                            
                            # Simplified audio detection - just check state and peak
                            if state == AUDCLNT_SESSIONSTATE_ACTIVE and peak > self._peak_threshold:
                                found_active = True
                                if self._debug:
                                    print(f"  ** ACTIVE AUDIO **", flush=True)
                                break
                                
                        except psutil.NoSuchProcess:
                            continue
                    finally:
                        meter.Release()
                        volume.Release()
                        audio_session.Release()
                finally:
                    session.Release()
            
            if self._debug:
                print(f"\nFound active audio: {found_active}", flush=True)
                print('=' * 50, flush=True)
            return found_active
            
        except Exception as e:
            print(f"Error checking audio sessions: {e}")
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

def main():
    print(f"\n{time.strftime('%H:%M:%S')} - Starting Spotify Audio Monitor...", flush=True)
    print("Press Ctrl+C to exit", flush=True)
    spotify_was_playing = False
    audio_manager = None
    last_action_time = 0
    action_cooldown = 0.5  # Reduced cooldown
    last_status_time = 0
    status_interval = 5  # Show status every 5 seconds
    
    while True:
        try:
            if audio_manager is None:
                audio_manager = AudioSessionManager()
                print(f"{time.strftime('%H:%M:%S')} - Audio manager initialized", flush=True)
                time.sleep(0.5)
            
            current_time = time.time()
            spotify_process = get_spotify_process()
            
            if spotify_process:
                other_apps_playing = audio_manager.check_audio_sessions(check_spotify=False)
                time.sleep(0.1)
                spotify_playing = audio_manager.check_audio_sessions(check_spotify=True)
                
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
