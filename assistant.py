import speech_recognition as sr
import os
import subprocess
import time
import threading
import difflib
from send2trash import send2trash
import ctypes
import sys
import winreg
import win32com.client
import pythoncom
#import google.generativeai as genai
#genai.configure(api_key="")
from groq import Groq
import json
import re
import urllib.parse
import urllib.request
import webbrowser
from dotenv import load_dotenv
load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))
# ==========================================
# 1. AI BRAIN SETUP (FUNCTION CALLING)
# ==========================================


# --- Define the Tools for the AI ---
# The AI reads these descriptions to instantly know how to route your sentence.
def tool_launch_software(app_name: str):
    """Opens, launches, or starts a software application or program."""
    pass

def tool_close_software(app_name: str):
    """Closes, shuts down, or exits a software application (like Brave or Chrome). DO NOT use this for turning off the PC."""
    pass

def tool_shutdown_or_restart_pc(action: str, delay_in_minutes: int = 0):
    """Use this to shut down, turn off, restart, or cancel a shutdown of the physical computer. Action MUST be 'shutdown', 'restart', or 'cancel_shutdown'. Can handle requests with a time delay."""
    pass

def tool_setup_environment_variable():
    """Sets a Windows environment path variable."""
    pass

def tool_create_file(file_name: str = ""):
    """Use this to create a new empty file. Extracts the file name and extension if the user provides one (e.g., 'index.html' or 'notes.txt')."""
    pass

def tool_delete_file(file_name: str = ""):
    """Use this to delete or remove a specific file. Extracts the target file name."""
    pass

def tool_rename_file(old_name: str = "", new_name: str = ""):
    """Use this to rename a file. Extracts both the current file name and the new desired name."""
    pass

def tool_change_system_directory(target_path: str):
    """Changes the current working directory or goes to a different folder."""
    pass

def tool_get_current_directory():
    """Tells the user what the current working directory is."""
    pass



# ==========================================
# 2. SYSTEM UTILITIES & CONSTANTS
# ==========================================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

number_words = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
}

SEARCH_ROOT = os.path.expanduser("~")
context = {"current_directory": os.getcwd()}
doc_cache = {}
env_path_cache = []
app_cache = {}
folder_cache = {}
is_indexing = True

speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 1

def speak(text):
    print("Assistant:", text)
    speaker.Speak(text)
    time.sleep(0.2)

recognizer = sr.Recognizer()
recognizer.pause_threshold = 0.8

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        # We REMOVED the 0.5s ambient noise delay here so it catches fast replies!
        try:
            # Added a short timeout so it doesn't hang forever if you say nothing
            audio = recognizer.listen(source, timeout=4, phrase_time_limit=6)
            command = recognizer.recognize_google(audio)
            print("You said:", command)
            return command.lower()
        except sr.WaitTimeoutError:
            # Handles if you just stay completely silent
            return None
        except Exception:
            # Handles if it couldn't understand the audio
            return None

def get_confirmation():
    text = listen() 
    if not text:
        return False
    text = text.lower()
    print(f"[Confirmation Heard]: '{text}'") 
    yes_words = [
        "yes", "yeah", "yep", "ya", "yas", "yup", 
        "sure", "ok", "okay", "do it", "go ahead", "proceed"
    ]
    
    if any(word in text for word in yes_words):
        return True
    return False

def normalize(text):
    return text.replace("-", "").replace("_", "").replace(" ", "").lower()

# ==========================================
# 3. CORE FEATURES (YOUR HARD WORK)
# ==========================================
def change_directory(command):
    global folder_cache
    home = os.path.expanduser("~")

    if "recycle bin" in command:
        os.system("start shell:RecycleBinFolder")
        speak("Opening Recycle Bin")
        return
    if "control panel" in command:
        os.system("start control")
        speak("Opening Control Panel")
        return
    if "task manager" in command:
        os.system("start taskmgr")
        speak("Opening Task Manager")
        return

    if "desktop" in command:
        context["current_directory"] = os.path.join(home, "Desktop")
        speak("Changed directory to Desktop")
        return
    if "downloads" in command:
        context["current_directory"] = os.path.join(home, "Downloads")
        speak("Changed directory to Downloads")
        return
    if "documents" in command:
        context["current_directory"] = os.path.join(home, "Documents")
        speak("Changed directory to Documents")
        return
    if "go back" in command:
        parent = os.path.dirname(context["current_directory"])
        context["current_directory"] = parent
        speak("Moved back")
        return

    target_folder = command.replace("change directory to", "").replace("change directory", "").replace("go to", "").strip().lower()

    try:
        local_folders = [f for f in os.listdir(context["current_directory"]) if os.path.isdir(os.path.join(context["current_directory"], f))]
        lower_local = [f.lower() for f in local_folders]
        
        if target_folder in lower_local:
            idx = lower_local.index(target_folder)
            context["current_directory"] = os.path.join(context["current_directory"], local_folders[idx])
            speak(f"Entered folder {local_folders[idx]}")
            return
            
        local_match = difflib.get_close_matches(target_folder, lower_local, n=1, cutoff=0.6)
        if local_match:
            idx = lower_local.index(local_match[0])
            context["current_directory"] = os.path.join(context["current_directory"], local_folders[idx])
            speak(f"Entered folder {local_folders[idx]}")
            return
    except PermissionError:
        pass 

    if target_folder in folder_cache:
        context["current_directory"] = folder_cache[target_folder]
        real_folder_name = os.path.basename(folder_cache[target_folder])
        speak(f"Teleported to {real_folder_name}")
        return
        
    global_names = list(folder_cache.keys())
    global_match = difflib.get_close_matches(target_folder, global_names, n=1, cutoff=0.7)
    if global_match:
        matched_name = global_match[0]
        context["current_directory"] = folder_cache[matched_name]
        real_folder_name = os.path.basename(folder_cache[matched_name])
        speak(f"Teleported to {real_folder_name}")
        return

    speak(f"Could not find a folder sounding like {target_folder}")

def current_directory():
    readable_path = context['current_directory'].replace("\\", " ")
    speak(f"You are in {readable_path}")

def search_files(filename, max_depth=3):
    matches = []
    base_depth = SEARCH_ROOT.count(os.sep)
    for root, dirs, files in os.walk(SEARCH_ROOT):
        current_depth = root.count(os.sep) - base_depth
        if current_depth >= max_depth:
            del dirs[:] 
        if filename in files:
            matches.append(os.path.join(root, filename))
    return matches

def create_file(file_name=""):
    # --- FAST PASS: Did the AI already extract the name? ---
    if file_name != "":
        # Split "notes.txt" into name="notes", ext="txt"
        if "." in file_name:
            name, ext = file_name.rsplit(".", 1)
        else:
            name = file_name
            ext = "txt" # Default fallback
    else:
        # --- OLD RELIABLE: Ask out loud ---
        speak("Tell file name")
        name = listen()
        speak("Tell extension")
        ext = listen()
        if not name or not ext:
            speak("Invalid input, cancelling.")
            return
            
    # --- YOUR ORIGINAL LOGIC REMAINS UNTOUCHED ---
    name = name.replace(" ", "")
    ext = ext.replace(" ", "")
    path = os.path.join(context["current_directory"], f"{name}.{ext}")

    speak(f"Proceed to create {name} dot {ext}?")
    if not get_confirmation():
        speak("Cancelled.")
        return
    try:
        with open(path, "w") as f:
            f.write("Created by Voice Assistant")
        speak("File created")
    except Exception as e:
        speak("Failed to create file.")
        print(e)

def delete_file(file_name=""):
    if file_name == "":
        speak("Tell the name of the file you want to delete.")
        file_name = listen()
        if not file_name: return
        
    # THE FIX: Point exactly to Jarvis's current virtual folder
    full_path = os.path.join(context["current_directory"], file_name.replace(" ", ""))
    
    if os.path.exists(full_path):
        speak(f"Are you sure you want to delete {file_name}? This cannot be undone.")
        if not get_confirmation():
            speak("Cancelled.")
            return
        try:
            os.remove(full_path)
            speak(f"Deleted file {file_name}")
            print(f"[Success] Deleted: {full_path}")
        except Exception as e:
            speak("Could not delete file. It might be open.")
            print(f"[Error] {e}")
    else:
        speak("Could not find that file here.")
        print(f"[Error] File not found: {full_path}")

def rename_file(old_name="", new_name=""):
    if old_name == "":
        speak("Tell the current name of the file.")
        old_name = listen()
        if not old_name: return
        
    if new_name == "":
        speak("Tell the new name for the file.")
        new_name = listen()
        if not new_name: return
        
    # THE FIX: Apply the virtual folder path to both names
    old_full_path = os.path.join(context["current_directory"], old_name.replace(" ", ""))
    new_full_path = os.path.join(context["current_directory"], new_name.replace(" ", ""))
    
    if os.path.exists(old_full_path):
        try:
            os.rename(old_full_path, new_full_path)
            speak("File renamed successfully.")
            print(f"[Success] Renamed to: {new_full_path}")
        except Exception as e:
            speak("Could not rename the file.")
            print(f"[Error] {e}")
    else:
        speak("Could not find the original file.")
        print(f"[Error] File not found: {old_full_path}")

def shutdown_or_restart_pc(intent, command):
    if intent == "cancel_shutdown":
        os.system("shutdown /a")
        speak("Shutdown cancelled")
        return

    minutes = None
    for w in command.split():
        if w.isdigit():
            minutes = int(w)
        elif w in number_words:
            minutes = number_words[w]

    if intent == "shutdown":
        if minutes:
            os.system(f"shutdown /s /t {minutes*60}")
            speak(f"Shutting down in {minutes} minutes.")
        else:
            os.system("shutdown /s /t 0")

    if intent == "restart":
        if minutes:
            os.system(f"shutdown /r /t {minutes*60}")
            speak(f"Restarting in {minutes} minutes.")
        else:
            os.system("shutdown /r /t 0")

def build_system_caches():
    global app_cache, folder_cache, doc_cache, env_path_cache, is_indexing
    print("\n[System] Background: Indexing Apps, Folders, Docs, and Dev Paths...")
    
    user_start_menu = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
    system_start_menu = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    temp_app_cache = {}
    for base in [user_start_menu, system_start_menu]:
        if os.path.exists(base):
            for root, dirs, files in os.walk(base):
                for f in files:
                    if f.endswith(".lnk"):
                        temp_app_cache[f.replace(".lnk", "").lower()] = os.path.join(root, f)
    app_cache = temp_app_cache
    
    home = os.path.expanduser("~")
    base_depth = home.count(os.sep)
    temp_folder_cache = {}
    ignore = ['appdata', 'application data', 'local settings', 'temp', 'node_modules', '.git']
    for root, dirs, files in os.walk(home):
        if root.count(os.sep) - base_depth >= 4:
            del dirs[:]
        dirs[:] = [d for d in dirs if d.lower() not in ignore and not d.startswith('.')]
        for d in dirs:
            temp_folder_cache[d.lower()] = os.path.join(root, d)
    folder_cache = temp_folder_cache

    temp_doc_cache = {}
    doc_exts = ('.pdf', '.pptx', '.ppt', '.docx', '.doc', '.xlsx', '.txt')
    search_dirs = [os.path.join(home, "Desktop"), os.path.join(home, "Documents"), os.path.join(home, "Downloads")]
    for base in search_dirs:
        if os.path.exists(base):
            for root, dirs, files in os.walk(base):
                for f in files:
                    if f.lower().endswith(doc_exts):
                        temp_doc_cache[os.path.splitext(f)[0].lower()] = os.path.join(root, f)
    doc_cache = temp_doc_cache

    temp_env_paths = []
    search_locations = ["C:\\Program Files", "C:\\Program Files (x86)"]
    local_app_data = os.environ.get('LOCALAPPDATA')
    if local_app_data:
        search_locations.append(os.path.join(local_app_data, 'Programs'))
        
    for base in search_locations:
        if os.path.exists(base):
            for root, dirs, files in os.walk(base):
                for d in dirs:
                    temp_env_paths.append((d.lower(), os.path.join(root, d)))
                for f in files:
                    temp_env_paths.append((f.lower(), root))
                    
    env_path_cache = temp_env_paths
    is_indexing = False
    print(f"\n[System] Background: Complete! Found {len(app_cache)} apps, {len(folder_cache)} folders, {len(doc_cache)} docs, and {len(env_path_cache)} dev paths.")

def open_application(command):
    global app_cache, doc_cache, is_indexing
    app_name = command.replace("open", "").strip().lower()
    
    if "task manager" in app_name:
        speak("Opening Task Manager")
        os.system("start taskmgr")
        return
    if "control panel" in app_name:
        speak("Opening Control Panel")
        os.system("start control")
        return
        
    if is_indexing:
        speak("I'm still organizing your files in the background. Please try again in a few seconds.")
        return
    if not app_name:
        speak("Please tell me what to open.")
        return

    app_names = list(app_cache.keys())
    doc_names = list(doc_cache.keys())
    
    if app_name in app_names:
        speak(f"Opening {app_name}")
        os.startfile(app_cache[app_name])
        return

    if app_name in doc_names:
        speak(f"Opening document {app_name}")
        os.startfile(doc_cache[app_name])
        return

    app_close = difflib.get_close_matches(app_name, app_names, n=3, cutoff=0.6)
    if len(app_close) == 1:
        matched_name = app_close[0]
        speak(f"Opening {matched_name}")
        os.startfile(app_cache[matched_name])
        return
    elif len(app_close) > 1:
        speak("Multiple application matches found.")
        for i, n in enumerate(app_close):
            speak(f"Option {i+1}: {n}")
        speak("Say the option number.")
        choice = listen() 
        if not choice:
            return
        try:
            if choice in number_words:
                idx = number_words[choice] - 1
            else:
                idx = int(choice) - 1
        except ValueError:
            speak("Invalid choice.")
            return

        if 0 <= idx < len(app_close):
            selected = app_close[idx]
            speak(f"Opening {selected}")
            os.startfile(app_cache[selected])
            return
        else:
            speak("Choice out of range.")
            return

    doc_close = difflib.get_close_matches(app_name, doc_names, n=1, cutoff=0.6)
    if doc_close:
        matched_doc = doc_close[0]
        speak(f"Opening document {matched_doc}")
        os.startfile(doc_cache[matched_doc])
        return

    speak(f"Could not find an app or document sounding like {app_name}.")

def set_environment_variable_auto():
    global env_path_cache, is_indexing
    if is_indexing or not env_path_cache:
        speak("I am still mapping your system files in the background. Please wait a few seconds and try again.")
        return

    speak("Tell the new variable name")
    var_name = listen()
    if not var_name:
        return
        
    speak("Tell the folder keyword to search for")
    keyword = listen()
    if not var_name or not keyword:
        speak("Invalid input")
        return
        
    var_name = var_name.upper().replace(" ", "_")
    keyword = keyword.replace(" ", "").lower()

    speak(f"Searching for {keyword}...")
    matches = [path for name, path in env_path_cache if keyword in name]
    matches = list(set(matches)) 

    if not matches:
        speak(f"Could not find any developer paths containing {keyword}.")
        return

    speak("I found these paths:")
    for i, p in enumerate(matches[:5]):
        print(f"[{i+1}] {p}")
        speak(f"Option {i+1}")

    speak("Please say Option 1, Option 2, or your preferred number.")
    choice = listen()
    if not choice:
        speak("I didn't hear a choice. Canceling.")
        return

    choice = choice.lower().replace(".", "").replace("option", "").replace("number", "").strip()
    robust_number_map = {
        "one": 1, "won": 1, "1": 1, "two": 2, "to": 2, "too": 2, "2": 2,
        "three": 3, "tree": 3, "3": 3, "four": 4, "for": 4, "4": 4, "five": 5, "5": 5
    }

    try:
        if choice in robust_number_map:
            idx = robust_number_map[choice] - 1
        else:
            idx = int(choice) - 1
    except ValueError:
        speak("I couldn't understand the option number. Canceling.")
        return

    if idx < 0 or idx >= len(matches[:5]):
        speak("Choice is out of range.")
        return

    selected_path = matches[idx]
    os.system(f'setx {var_name} "{selected_path}"')
    speak(f"Variable {var_name} set.")

    speak("Should I add this to your system path as well?")
    if get_confirmation():
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_ALL_ACCESS)
            try:
                current_path, _ = winreg.QueryValueEx(key, "PATH")
            except FileNotFoundError:
                current_path = ""
            
            if selected_path not in current_path:
                updated_path = current_path + ";" + selected_path if current_path else selected_path
                winreg.SetValueEx(key, "PATH", 0, winreg.REG_EXPAND_SZ, updated_path)
            
            winreg.CloseKey(key)
            speak("Path updated safely.")
        except Exception as e:
            speak("Failed to update path registry.")
            print(e)

def close_application(command):
    global app_cache
    app_name = command.replace("close", "").replace("exit", "").replace("shut", "").strip().lower()

    if not app_name:
        speak("Please tell me what to close.")
        return

    if "task manager" in app_name:
        os.system("taskkill /F /IM Taskmgr.exe /T >nul 2>&1")
        speak("Closed Task Manager")
        return

    app_names = list(app_cache.keys())
    target_lnk = None
    matched_app_name = app_name

    if app_name in app_names:
        target_lnk = app_cache[app_name]
    else:
        close_matches = difflib.get_close_matches(app_name, app_names, n=1, cutoff=0.6)
        if close_matches:
            matched_app_name = close_matches[0]
            target_lnk = app_cache[matched_app_name]

    if target_lnk:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(target_lnk)
            target_path = shortcut.TargetPath
            exe_name = os.path.basename(target_path)
            
            if exe_name.endswith('.exe'):
                os.system(f"taskkill /F /IM {exe_name} /T >nul 2>&1")
                speak(f"Closed {matched_app_name}")
                return
        except Exception:
            pass 

    guess_exe = f"{app_name.replace(' ', '')}.exe"
    os.system(f"taskkill /F /IM {guess_exe} /T >nul 2>&1")
    speak(f"Attempted to close {app_name}")

# ==========================================
# 3b. BROWSER, YOUTUBE & MEDIA CONTROLS
# ==========================================
BRAVE_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"BraveSoftware\Brave-Browser\Application\brave.exe"),
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
]

_last_youtube_results = []
_last_search_results = []

def _find_brave():
    """Locate the Brave browser executable."""
    for path in BRAVE_PATHS:
        if os.path.exists(path):
            return path
    return None

def _open_in_brave(url):
    """Open a URL in Brave browser, falling back to default browser."""
    brave = _find_brave()
    if brave:
        subprocess.Popen([brave, url])
    else:
        webbrowser.open(url)

def _scrape_youtube_results(query):
    """Fetch YouTube search page and extract video IDs."""
    global _last_youtube_results
    try:
        url = f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
        })
        response = urllib.request.urlopen(req, timeout=8)
        html = response.read().decode("utf-8", errors="ignore")
        video_ids = re.findall(r'"videoId":"([a-zA-Z0-9_-]{11})"', html)
        seen = set()
        unique_ids = []
        for vid in video_ids:
            if vid not in seen:
                seen.add(vid)
                unique_ids.append(vid)
        _last_youtube_results = unique_ids[:10]
        print(f"[YouTube] Cached {len(_last_youtube_results)} video results")
    except Exception as e:
        print(f"[YouTube Scrape Error] {e}")
        _last_youtube_results = []

def _scrape_google_results(query):
    """Fetch Google search page and extract result URLs."""
    global _last_search_results
    try:
        url = f"https://www.google.com/search?q={urllib.parse.quote_plus(query)}"
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
        })
        response = urllib.request.urlopen(req, timeout=8)
        html = response.read().decode("utf-8", errors="ignore")
        urls = re.findall(r'<a href="/url\?q=(https?://[^&"]+)', html)
        filtered = [u for u in urls
                    if not any(d in u for d in ["accounts.google", "support.google", "policies.google", "maps.google"])]
        _last_search_results = filtered[:10]
        print(f"[Google] Cached {len(_last_search_results)} search results")
    except Exception as e:
        print(f"[Google Scrape Error] {e}")
        _last_search_results = []

def search_web(query):
    """Search Google in Brave browser and cache results."""
    if not query:
        speak("Please tell me what to search for.")
        return
    url = f"https://www.google.com/search?q={urllib.parse.quote_plus(query)}"
    _open_in_brave(url)
    speak(f"Searching for {query}")
    threading.Thread(target=_scrape_google_results, args=(query,), daemon=True).start()

def open_website(url):
    """Open a website URL in Brave browser."""
    if not url:
        speak("Please provide a website URL.")
        return
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    _open_in_brave(url)
    speak(f"Opening {url}")

def play_youtube(query):
    """Search YouTube in Brave browser and cache results."""
    if not query:
        speak("Please tell me what to play on YouTube.")
        return
    url = f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"
    _open_in_brave(url)
    speak(f"Searching YouTube for {query}")
    threading.Thread(target=_scrape_youtube_results, args=(query,), daemon=True).start()

def play_youtube_result(position):
    """Play the Nth video from the last YouTube search results."""
    global _last_youtube_results
    if not _last_youtube_results:
        speak("No YouTube search results available. Please search YouTube first.")
        return
    try:
        idx = int(position) - 1
    except (ValueError, TypeError):
        speak("Please provide a valid number.")
        return
    if idx < 0 or idx >= len(_last_youtube_results):
        speak(f"Please choose a number between 1 and {len(_last_youtube_results)}.")
        return
    video_id = _last_youtube_results[idx]
    url = f"https://www.youtube.com/watch?v={video_id}"
    _open_in_brave(url)
    speak(f"Playing video number {position}")

def open_search_result(position):
    """Open the Nth result from the last Google search."""
    global _last_search_results
    if not _last_search_results:
        speak("No search results available. Please search the web first.")
        return
    try:
        idx = int(position) - 1
    except (ValueError, TypeError):
        speak("Please provide a valid number.")
        return
    if idx < 0 or idx >= len(_last_search_results):
        speak(f"Please choose a number between 1 and {len(_last_search_results)}.")
        return
    result_url = _last_search_results[idx]
    _open_in_brave(result_url)
    speak(f"Opening result number {position}")

def media_control(action, amount=0):
    """Control media playback using Windows media keys."""
    if not action:
        speak("Please specify a media action like play, pause, or next.")
        return
    KEYEVENTF_EXTENDEDKEY = 0x0001
    KEYEVENTF_KEYUP = 0x0002
    media_keys = {
        "play": 0xB3, "pause": 0xB3, "play_pause": 0xB3,
        "next": 0xB0, "previous": 0xB1,
        "volume_up": 0xAF, "volume_down": 0xAE, "mute": 0xAD,
    }
    action = action.lower().strip().replace(" ", "_")
    vk = media_keys.get(action)
    if vk is None:
        speak(f"Unknown media action: {action}")
        return
    try:
        amount = int(amount) if amount else 0
    except (ValueError, TypeError):
        amount = 0
    if action in ("volume_up", "volume_down") and amount > 0:
        presses = max(1, amount // 2)
    else:
        presses = 1
    for _ in range(presses):
        ctypes.windll.user32.keybd_event(vk, 0, KEYEVENTF_EXTENDEDKEY, 0)
        ctypes.windll.user32.keybd_event(vk, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
    action_names = {
        "play": "Playing", "pause": "Paused", "play_pause": "Toggled play pause",
        "next": "Skipped to next track", "previous": "Went to previous track",
        "volume_up": f"Volume increased by {amount if amount > 0 else 2}",
        "volume_down": f"Volume decreased by {amount if amount > 0 else 2}",
        "mute": "Toggled mute",
    }
    speak(action_names.get(action, f"Performed {action}"))

# ==========================================
# 4. THE MAIN LOOP (AI ROUTING)
# ==========================================
def run_jarvis():
    # Start building caches in the background
    pythoncom.CoInitialize()
    threading.Thread(target=build_system_caches, daemon=True).start()
    speak("Assistant started and ready.")

    while True:
        command = listen()
        if not command:
            continue

        if any(w in command for w in ["exit", "quit", "stop listening", "goodbye", "bye"]):
            speak("Goodbye!")
            break

        print(f"[System] Sending to AI: '{command}'")

        try:
            # ---------- GROQ AI CALL ----------
            response = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "system",
                        "content": """
You are a command router AI.

Convert user command into JSON.

Available functions:
- create_file(file_name)
- delete_file(file_name)
- rename_file(old_name, new_name)
- change_system_directory(folder_name)
- get_current_directory()
- launch_software(app_name)
- close_software(app_name)
- shutdown_or_restart_pc(action, delay_in_minutes)
- setup_environment_variable()
- search_web(query): Search Google for something in the browser
- open_website(url): Open a specific website or URL
- play_youtube(query): Search and play a video or song on YouTube
- play_youtube_result(position): After a YouTube search, play a specific result. position is a number: 1 for first, 2 for second, etc. Use when user says 'play first video', 'play second song', etc.
- open_search_result(position): After a Google search, open a specific result. position is a number: 1 for first, 2 for second, etc. Use when user says 'open first result', 'open first site', 'click second link', etc.
- media_control(action, amount): Control media playback. action: play, pause, next, previous, volume_up, volume_down, mute. amount is optional integer for volume (e.g. 10 means change by 10 percent). If user says 'decrease volume by 20' then action=volume_down, amount=20.
- casual_response(message)

Respond ONLY in JSON like:
{
  "function": "function_name",
  "args": { "argument_name": "value" }
}
"""
                    },
                    {
                        "role": "user",
                        "content": command
                    }
                ]
            )

            ai_output = response.choices[0].message.content
            print(f"[AI RAW]: {ai_output}")

            data = json.loads(ai_output)

            func_name = data.get("function")
            args = data.get("args", {})

            print(f"[AI Decided] Target Function: {func_name} | Variables: {args}")

            # ---------- FUNCTION ROUTING (UNCHANGED) ----------
            if func_name == "create_file":
                target_file = args.get("file_name", "")
                create_file(target_file)

            elif func_name == "delete_file":
                target_file = args.get("file_name", "")
                delete_file(target_file)

            elif func_name == "rename_file":
                old_target = args.get("old_name", "")
                new_target = args.get("new_name", "")
                rename_file(old_target, new_target)

            elif func_name == "change_system_directory":
                change_directory(args.get("target_path", command))

            elif func_name == "get_current_directory":
                current_directory()

            elif func_name == "launch_software":
                open_application(args.get("app_name", command))

            elif func_name == "close_software":
                close_application(args.get("app_name", command))

            elif func_name == "shutdown_or_restart_pc":
                import os
                action = args.get("action", "").lower() 
                raw_delay = args.get("delay_in_minutes", 0)
                
                # THE FIX: Safely handle empty strings or text so Python doesn't crash
                try:
                    delay_in_minutes = int(raw_delay) if raw_delay != "" else 0
                except ValueError:
                    delay_in_minutes = 0
                    
                delay_seconds = delay_in_minutes * 60

                # Catch ANY variation of cancel
                if "cancel" in action:
                    print("[Execution] Canceling pending power actions...")
                    os.system("shutdown /a")
                    speak("Power action cancelled.")

                elif "restart" in action:
                    speak(f"Restart your PC in {delay_in_minutes} minutes?")
                    if not get_confirmation():
                        speak("Restart cancelled.")
                    else:
                        print(f"[Execution] Restarting PC in {delay_in_minutes} minutes...")
                        os.system(f"shutdown /r /t {delay_seconds}")
                        speak(f"Restart scheduled in {delay_in_minutes} minutes.")

                elif "shutdown" in action:
                    speak(f"Shut down your PC in {delay_in_minutes} minutes?")
                    if not get_confirmation():
                        speak("Shutdown cancelled.")
                    else:
                        print(f"[Execution] Shutting down PC in {delay_in_minutes} minutes...")
                        os.system(f"shutdown /s /t {delay_seconds}")
                        speak(f"Shutdown scheduled in {delay_in_minutes} minutes.")
                
                else:
                    speak("I wasn't sure what power action you wanted. Cancelling for safety.")

            elif func_name == "setup_environment_variable":
                set_environment_variable_auto()
            
            elif func_name == "casual_response":
                reply = args.get("message", "Alright.")
                speak(reply)

            elif func_name == "search_web":
                search_web(args.get("query", ""))

            elif func_name == "open_website":
                open_website(args.get("url", ""))

            elif func_name == "play_youtube":
                play_youtube(args.get("query", ""))

            elif func_name == "play_youtube_result":
                play_youtube_result(args.get("position", 1))

            elif func_name == "open_search_result":
                open_search_result(args.get("position", 1))

            elif func_name == "media_control":
                media_control(args.get("action", ""), args.get("amount", 0))

            else:
                speak("I understood, but no matching function found.")

        except Exception as e:
            print(f"[AI Error] {e}")
            speak("AI response error. Trying again shortly.")
            time.sleep(5)

# ==========================================
# THREAD LAUNCHER (This replaces your old main loop)
# ==========================================
if __name__ == "__main__":
    # 1. Create the thread and point it to the new run_jarvis function
    jarvis_thread = threading.Thread(target=run_jarvis, daemon=True)
    
    # 2. Start the background thread
    jarvis_thread.start()
    
    # 3. Keep the main script alive so the thread can run
    # (We will replace this input with our UI loop next!)
    input("Press Enter to kill the system...\n")

