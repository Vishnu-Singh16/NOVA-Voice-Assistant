"""
NOVA Voice Assistant — Core Logic
Extracted from assistant.py for use by the GUI frontend.
The original assistant.py remains untouched for terminal usage.
"""

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
from groq import Groq
import json
import re
import urllib.parse
import urllib.request
import webbrowser

# ==========================================
# 1. AI CLIENT SETUP
# ==========================================
from dotenv import load_dotenv

load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))

# ==========================================
# 2. SYSTEM CONSTANTS & STATE
# ==========================================
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

# ==========================================
# 3. CALLBACK SYSTEM FOR GUI INTEGRATION
# ==========================================
# These callbacks are set by the GUI to intercept output and confirmation
_on_speak = None          # Called when assistant wants to say something
_on_status_update = None  # Called when status changes (indexing, etc.)
_confirm_callback = None  # Called to get user confirmation (returns True/False)
_input_callback = None    # Called to get text input from user
_choice_callback = None   # Called to present choices to user


def set_speak_callback(callback):
    """GUI sets this to receive spoken text. Signature: callback(text: str)"""
    global _on_speak
    _on_speak = callback


def set_status_callback(callback):
    """GUI sets this to receive status updates. Signature: callback(key: str, value: str)"""
    global _on_status_update
    _on_status_update = callback


def set_confirm_callback(callback):
    """GUI sets this to handle confirmations. Signature: callback(prompt: str) -> bool"""
    global _confirm_callback
    _confirm_callback = callback


def set_input_callback(callback):
    """GUI sets this for text input dialogs. Signature: callback(prompt: str) -> str or None"""
    global _input_callback
    _input_callback = callback


def set_choice_callback(callback):
    """GUI sets this for choice selection dialogs. Signature: callback(prompt: str, options: list) -> int or None"""
    global _choice_callback
    _choice_callback = callback


def speak(text):
    """Sends text to GUI callback AND speaks it aloud via Windows SAPI."""
    print("Assistant:", text)
    if _on_speak:
        _on_speak(text)
    # Windows SAPI TTS on a separate thread so it doesn't block the GUI
    def _say():
        try:
            pythoncom.CoInitialize()
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Rate = 1
            speaker.Speak(text)
        except Exception as e:
            print(f"[TTS Error] {e}")
        finally:
            pythoncom.CoUninitialize()
    threading.Thread(target=_say, daemon=True).start()


def get_confirmation(prompt="Proceed?"):
    """
    Dual confirmation: tries GUI callback first, falls back to voice.
    The GUI can provide a popup dialog; voice uses speech_recognition.
    """
    if _confirm_callback:
        return _confirm_callback(prompt)
    # Fallback: voice confirmation (original behavior)
    import speech_recognition as sr
    recognizer = sr.Recognizer()
    recognizer.pause_threshold = 0.8
    try:
        with sr.Microphone() as source:
            print("Listening for confirmation...")
            audio = recognizer.listen(source, timeout=4, phrase_time_limit=4)
            text = recognizer.recognize_google(audio).lower()
            print(f"[Confirmation Heard]: '{text}'")
            yes_words = ["yes", "yeah", "yep", "ya", "yas", "yup",
                         "sure", "ok", "okay", "do it", "go ahead", "proceed"]
            return any(word in text for word in yes_words)
    except Exception:
        return False


def normalize(text):
    return text.replace("-", "").replace("_", "").replace(" ", "").lower()


def _update_status(key, value):
    """Internal helper to push status to GUI."""
    if _on_status_update:
        _on_status_update(key, value)


# ==========================================
# 4. SYSTEM CACHES (BACKGROUND INDEXING)
# ==========================================
def build_system_caches():
    global app_cache, folder_cache, doc_cache, env_path_cache, is_indexing
    _update_status("indexing", "Scanning apps...")
    print("\n[System] Background: Indexing Apps, Folders, Docs, and Dev Paths...")

    # --- Apps ---
    user_start_menu = os.path.join(os.environ.get("APPDATA", ""),
                                   r"Microsoft\Windows\Start Menu\Programs")
    system_start_menu = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    temp_app_cache = {}
    for base in [user_start_menu, system_start_menu]:
        if os.path.exists(base):
            for root, dirs, files in os.walk(base):
                for f in files:
                    if f.endswith(".lnk"):
                        temp_app_cache[f.replace(".lnk", "").lower()] = os.path.join(root, f)
    app_cache = temp_app_cache
    _update_status("indexing", "Scanning folders...")

    # --- Folders ---
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
    _update_status("indexing", "Scanning documents...")

    # --- Documents ---
    temp_doc_cache = {}
    doc_exts = ('.pdf', '.pptx', '.ppt', '.docx', '.doc', '.xlsx', '.txt')
    search_dirs = [os.path.join(home, "Desktop"), os.path.join(home, "Documents"),
                   os.path.join(home, "Downloads")]
    for base in search_dirs:
        if os.path.exists(base):
            for root, dirs, files in os.walk(base):
                for f in files:
                    if f.lower().endswith(doc_exts):
                        temp_doc_cache[os.path.splitext(f)[0].lower()] = os.path.join(root, f)
    doc_cache = temp_doc_cache
    _update_status("indexing", "Scanning dev paths...")

    # --- Dev Paths ---
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
    _update_status("indexing", "Done")
    _update_status("stats", f"{len(app_cache)} apps, {len(folder_cache)} folders, {len(doc_cache)} docs")
    print(f"\n[System] Background: Complete! Found {len(app_cache)} apps, "
          f"{len(folder_cache)} folders, {len(doc_cache)} docs, "
          f"and {len(env_path_cache)} dev paths.")


# ==========================================
# 5. CORE FEATURES
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
        os.chdir(context["current_directory"])
        speak("Changed directory to Desktop")
        _update_status("current_dir", context["current_directory"])
        return
    if "downloads" in command:
        context["current_directory"] = os.path.join(home, "Downloads")
        os.chdir(context["current_directory"])
        speak("Changed directory to Downloads")
        _update_status("current_dir", context["current_directory"])
        return
    if "documents" in command:
        context["current_directory"] = os.path.join(home, "Documents")
        os.chdir(context["current_directory"])
        speak("Changed directory to Documents")
        _update_status("current_dir", context["current_directory"])
        return
    if "go back" in command:
        parent = os.path.dirname(context["current_directory"])
        context["current_directory"] = parent
        os.chdir(context["current_directory"])
        speak("Moved back")
        _update_status("current_dir", context["current_directory"])
        return

    target_folder = (command.replace("change directory to", "")
                     .replace("change directory", "")
                     .replace("go to", "").strip().lower())

    try:
        local_folders = [f for f in os.listdir(context["current_directory"])
                         if os.path.isdir(os.path.join(context["current_directory"], f))]
        lower_local = [f.lower() for f in local_folders]

        if target_folder in lower_local:
            idx = lower_local.index(target_folder)
            context["current_directory"] = os.path.join(context["current_directory"],
                                                        local_folders[idx])
            os.chdir(context["current_directory"])
            speak(f"Entered folder {local_folders[idx]}")
            _update_status("current_dir", context["current_directory"])
            return

        local_match = difflib.get_close_matches(target_folder, lower_local, n=1, cutoff=0.6)
        if local_match:
            idx = lower_local.index(local_match[0])
            context["current_directory"] = os.path.join(context["current_directory"],
                                                        local_folders[idx])
            os.chdir(context["current_directory"])
            speak(f"Entered folder {local_folders[idx]}")
            _update_status("current_dir", context["current_directory"])
            return
    except PermissionError:
        pass

    if target_folder in folder_cache:
        context["current_directory"] = folder_cache[target_folder]
        os.chdir(context["current_directory"])
        real_folder_name = os.path.basename(folder_cache[target_folder])
        speak(f"Teleported to {real_folder_name}")
        _update_status("current_dir", context["current_directory"])
        return

    global_names = list(folder_cache.keys())
    global_match = difflib.get_close_matches(target_folder, global_names, n=1, cutoff=0.7)
    if global_match:
        matched_name = global_match[0]
        context["current_directory"] = folder_cache[matched_name]
        os.chdir(context["current_directory"])
        real_folder_name = os.path.basename(folder_cache[matched_name])
        speak(f"Teleported to {real_folder_name}")
        _update_status("current_dir", context["current_directory"])
        return

    speak(f"Could not find a folder sounding like {target_folder}")


def current_directory():
    readable_path = context['current_directory'].replace("\\", " ")
    speak(f"You are in {readable_path}")


def open_terminal():
    """Opens a command prompt at the current working directory."""
    try:
        os.system("start cmd")
        speak("Opened a terminal here.")
    except Exception as e:
        speak("Could not open terminal.")
        print(f"Terminal error: {e}")


def create_file(file_name=""):
    if file_name != "":
        if "." in file_name:
            name, ext = file_name.rsplit(".", 1)
        else:
            name = file_name
            ext = "txt"
    else:
        speak("Please provide a file name.")
        return

    name = name.replace(" ", "")
    ext = ext.replace(" ", "")
    path = os.path.join(context["current_directory"], f"{name}.{ext}")

    if not get_confirmation(f"Create {name}.{ext}?"):
        speak("Cancelled.")
        return
    try:
        with open(path, "w") as f:
            f.write("Created by NOVA Assistant")
        speak("File created")
    except Exception as e:
        speak("Failed to create file.")
        print(e)


def delete_file(file_name=""):
    if file_name == "":
        speak("Please provide the file name to delete.")
        return

    full_path = os.path.join(context["current_directory"], file_name.replace(" ", ""))

    if os.path.exists(full_path):
        if not get_confirmation(f"Delete {file_name}? This cannot be undone."):
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
        speak("Please provide the current file name.")
        return
    if new_name == "":
        speak("Please provide the new file name.")
        return

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


def shutdown_or_restart_pc(action, delay_in_minutes=0):
    if "cancel" in action:
        print("[Execution] Canceling pending power actions...")
        os.system("shutdown /a")
        speak("Power action cancelled.")
        return

    try:
        delay_in_minutes = int(delay_in_minutes) if delay_in_minutes != "" else 0
    except ValueError:
        delay_in_minutes = 0

    delay_seconds = delay_in_minutes * 60

    if "restart" in action:
        if not get_confirmation(f"Restart your PC in {delay_in_minutes} minutes?"):
            speak("Restart cancelled.")
            return
        print(f"[Execution] Restarting PC in {delay_in_minutes} minutes...")
        os.system(f"shutdown /r /t {delay_seconds}")
        speak(f"Restart scheduled in {delay_in_minutes} minutes.")
    elif "shutdown" in action:
        if not get_confirmation(f"Shut down your PC in {delay_in_minutes} minutes?"):
            speak("Shutdown cancelled.")
            return
        print(f"[Execution] Shutting down PC in {delay_in_minutes} minutes...")
        os.system(f"shutdown /s /t {delay_seconds}")
        speak(f"Shutdown scheduled in {delay_in_minutes} minutes.")
    else:
        speak("I wasn't sure what power action you wanted. Cancelling for safety.")


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
        speak("I'm still organizing your files. Please try again in a few seconds.")
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
        options_text = ", ".join([f"Option {i+1}: {n}" for i, n in enumerate(app_close)])
        speak(f"Multiple matches found. {options_text}. Say the option number.")
        # For now, open first match; multi-choice in GUI can be enhanced later
        speak(f"Opening {app_close[0]}")
        os.startfile(app_cache[app_close[0]])
        return

    doc_close = difflib.get_close_matches(app_name, doc_names, n=1, cutoff=0.6)
    if doc_close:
        matched_doc = doc_close[0]
        speak(f"Opening document {matched_doc}")
        os.startfile(doc_cache[matched_doc])
        return

    speak(f"Could not find an app or document sounding like {app_name}.")


def close_application(command):
    global app_cache
    app_name = (command.replace("close", "").replace("exit", "")
                .replace("shut", "").strip().lower())

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
            pythoncom.CoInitialize()
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


def open_document(doc_name):
    """Open a document (PDF, PPT, Word, Excel, text file) by name."""
    global doc_cache, is_indexing

    if is_indexing:
        speak("I'm still organizing your files. Please try again in a few seconds.")
        return
    if not doc_name:
        speak("Please tell me which document to open.")
        return

    doc_name = doc_name.strip().lower()
    # Remove common filler words and extensions from the query
    for word in ["my ", "the ", "a ", "file", "document", "doc ",
                 ".pdf", ".pptx", ".ppt", ".docx", ".doc", ".xlsx", ".txt",
                 "pdf", "pptx", "ppt", "docx", "xlsx"]:
        doc_name = doc_name.replace(word, "")
    doc_name = doc_name.strip()

    if not doc_name:
        speak("Please provide a document name.")
        return

    doc_names = list(doc_cache.keys())

    # Exact match
    if doc_name in doc_names:
        speak(f"Opening document {doc_name}")
        os.startfile(doc_cache[doc_name])
        return

    # Fuzzy match
    close_matches = difflib.get_close_matches(doc_name, doc_names, n=3, cutoff=0.5)
    if len(close_matches) == 1:
        matched = close_matches[0]
        speak(f"Opening document {matched}")
        os.startfile(doc_cache[matched])
        return
    elif len(close_matches) > 1:
        speak("Multiple documents found:")
        for i, m in enumerate(close_matches):
            speak(f"  [{i+1}] {m}")
        if _choice_callback:
            idx = _choice_callback("Which document would you like to open?", close_matches)
            if idx is not None and 0 <= idx < len(close_matches):
                selected = close_matches[idx]
                speak(f"Opening document {selected}")
                os.startfile(doc_cache[selected])
                return
        else:
            speak(f"Opening {close_matches[0]}")
            os.startfile(doc_cache[close_matches[0]])
            return

    # Substring search
    for name in doc_names:
        if doc_name in name or name in doc_name:
            speak(f"Opening document {name}")
            os.startfile(doc_cache[name])
            return

    speak(f"Could not find a document matching '{doc_name}'.")


def set_environment_variable_auto():
    global env_path_cache, is_indexing
    if is_indexing or not env_path_cache:
        speak("I am still mapping your system files. Please wait and try again.")
        return

    # Get variable name via GUI dialog
    if _input_callback:
        var_name = _input_callback("Enter the new environment variable name:")
    else:
        speak("Environment variable setup requires the GUI.")
        return

    if not var_name:
        speak("Cancelled.")
        return

    # Get keyword to search for
    keyword = _input_callback("Enter a keyword to search for the path\n(e.g., 'python', 'java', 'node'):")
    if not keyword:
        speak("Cancelled.")
        return

    var_name = var_name.upper().replace(" ", "_")
    keyword = keyword.replace(" ", "").lower()

    speak(f"Searching for {keyword}...")
    matches = [path for name, path in env_path_cache if keyword in name]
    matches = list(set(matches))

    if not matches:
        speak(f"Could not find any paths containing {keyword}.")
        return

    # Show matches in chat
    display_matches = matches[:8]
    speak(f"Found {len(display_matches)} matching paths:")
    for i, p in enumerate(display_matches):
        speak(f"  [{i+1}] {p}")

    # Get user choice via GUI dialog
    if _choice_callback:
        idx = _choice_callback("Select the correct path:", display_matches)
    else:
        speak("Cannot select path without GUI.")
        return

    if idx is None or idx < 0:
        speak("Cancelled.")
        return

    selected_path = display_matches[idx]
    os.system(f'setx {var_name} "{selected_path}"')
    speak(f"Variable {var_name} set to:\n{selected_path}")

    # Ask about adding to PATH
    if get_confirmation(f"Add this path to your system PATH as well?"):
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
            speak("Path updated successfully. Restart your terminal to apply.")
        except Exception as e:
            speak("Failed to update path registry.")
            print(e)
    else:
        speak("Skipped adding to PATH.")


# ==========================================
# 5b. BROWSER, YOUTUBE & MEDIA CONTROLS
# ==========================================
BRAVE_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"BraveSoftware\Brave-Browser\Application\brave.exe"),
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
]

_last_youtube_results = []   # Cached video IDs from last YouTube search
_last_search_results = []    # Cached URLs from last Google search


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
    """Fetch YouTube search page and extract video IDs for follow-up commands."""
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
    """Fetch Google search page and extract result URLs for follow-up commands."""
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
    """Search Google in Brave browser and cache results for follow-up."""
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
    """Search YouTube in Brave browser and cache results for follow-up."""
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
    """Control media playback using Windows media keys. For volume, amount sets the percentage."""
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

    # For volume, each key press = 2%. Press multiple times for larger amounts.
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
# 6. CENTRAL COMMAND ROUTER (AI-POWERED)
# ==========================================
def process_command(command):
    """
    Central routing function.
    Takes a text command, sends to Groq AI, routes to the correct function.
    Returns a dict with the result for the GUI.
    """
    print(f"[System] Processing: '{command}'")

    try:
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
- open_document(doc_name): Opens documents like PDFs, PowerPoints, Word docs, text files, spreadsheets by name
- open_terminal(): Opens a new command prompt / terminal at the current directory
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

        # ---------- FUNCTION ROUTING ----------
        if func_name == "create_file":
            create_file(args.get("file_name", ""))

        elif func_name == "delete_file":
            delete_file(args.get("file_name", ""))

        elif func_name == "rename_file":
            rename_file(args.get("old_name", ""), args.get("new_name", ""))

        elif func_name == "change_system_directory":
            change_directory(args.get("target_path", command))

        elif func_name == "get_current_directory":
            current_directory()

        elif func_name == "launch_software":
            open_application(args.get("app_name", command))

        elif func_name == "close_software":
            close_application(args.get("app_name", command))

        elif func_name == "shutdown_or_restart_pc":
            action = args.get("action", "").lower()
            raw_delay = args.get("delay_in_minutes", 0)
            shutdown_or_restart_pc(action, raw_delay)

        elif func_name == "open_document":
            open_document(args.get("doc_name", ""))

        elif func_name == "open_terminal":
            open_terminal()

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

        return {
            "success": True,
            "function": func_name,
            "args": args,
            "current_dir": context["current_directory"]
        }

    except Exception as e:
        print(f"[AI Error] {e}")
        speak("AI response error. Please try again.")
        return {
            "success": False,
            "error": str(e),
            "current_dir": context["current_directory"]
        }


def initialize():
    """
    Call this once at startup. Starts background indexing.
    """
    pythoncom.CoInitialize()
    os.chdir(context["current_directory"])
    threading.Thread(target=build_system_caches, daemon=True).start()
    _update_status("current_dir", context["current_directory"])
