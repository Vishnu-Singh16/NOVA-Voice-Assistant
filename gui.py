"""
NOVA Voice Assistant — Desktop GUI
A premium dark-themed desktop application built with CustomTkinter.
"""

import customtkinter as ctk
import threading
import time
import os
import sys
import ctypes
import winsound
import difflib
from datetime import datetime

# --- Admin elevation (same as original assistant.py) ---
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

import assistant_core

# ==========================================
# THEME & COLOR CONSTANTS
# ==========================================
COLORS = {
    "bg_darkest":     "#0d0d1a",
    "bg_dark":        "#12122a",
    "bg_card":        "#1a1a3e",
    "bg_card_hover":  "#222255",
    "bg_input":       "#15153a",
    "accent_cyan":    "#00d4ff",
    "accent_purple":  "#7c3aed",
    "accent_blue":    "#3b82f6",
    "accent_gradient_start": "#00d4ff",
    "accent_gradient_end":   "#7c3aed",
    "text_primary":   "#e8e8f0",
    "text_secondary": "#8888aa",
    "text_muted":     "#555577",
    "user_bubble":    "#2d2d6b",
    "assistant_bubble": "#1e1e44",
    "success":        "#10b981",
    "warning":        "#f59e0b",
    "danger":         "#ef4444",
    "border":         "#2a2a55",
    "mic_active":     "#ef4444",
    "mic_idle":       "#00d4ff",
    "wake_active":    "#10b981",
    "wake_inactive":  "#555577",
}

FONT_FAMILY = "Segoe UI"


# ==========================================
# CHAT BUBBLE WIDGET
# ==========================================
class ChatBubble(ctk.CTkFrame):
    """A single chat message bubble."""

    def __init__(self, parent, message, is_user=True, timestamp=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(fg_color="transparent")

        if timestamp is None:
            timestamp = datetime.now().strftime("%I:%M %p")

        bubble_color = COLORS["user_bubble"] if is_user else COLORS["assistant_bubble"]
        anchor_side = "e" if is_user else "w"
        text_color = COLORS["text_primary"]
        label_prefix = ""
        if not is_user:
            label_prefix = "🤖  "

        bubble_frame = ctk.CTkFrame(
            self, fg_color=bubble_color, corner_radius=16,
            border_width=1, border_color=COLORS["border"]
        )
        bubble_frame.pack(
            anchor=anchor_side,
            padx=(60 if not is_user else 12, 12 if not is_user else 60),
            pady=4
        )

        msg_label = ctk.CTkLabel(
            bubble_frame, text=f"{label_prefix}{message}",
            text_color=text_color, font=(FONT_FAMILY, 13),
            wraplength=420, justify="left" if not is_user else "right", anchor="w"
        )
        msg_label.pack(padx=16, pady=(10, 4))

        time_label = ctk.CTkLabel(
            bubble_frame, text=timestamp,
            text_color=COLORS["text_muted"], font=(FONT_FAMILY, 9), anchor=anchor_side
        )
        time_label.pack(padx=16, pady=(0, 8), anchor=anchor_side)


# ==========================================
# CONFIRMATION DIALOG
# ==========================================
class ConfirmDialog(ctk.CTkToplevel):
    """A modal confirmation dialog with Yes/No buttons and voice option."""

    def __init__(self, parent, prompt="Proceed?"):
        super().__init__(parent)
        self.result = False
        self.title("NOVA — Confirm")
        self.geometry("420x220")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg_dark"])
        self.attributes("-topmost", True)
        self.grab_set()

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 210
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 110
        self.geometry(f"+{x}+{y}")

        icon_label = ctk.CTkLabel(self, text="⚡", font=(FONT_FAMILY, 36), text_color=COLORS["accent_cyan"])
        icon_label.pack(pady=(20, 5))

        prompt_label = ctk.CTkLabel(
            self, text=prompt, font=(FONT_FAMILY, 14),
            text_color=COLORS["text_primary"], wraplength=360
        )
        prompt_label.pack(pady=(0, 20))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))

        yes_btn = ctk.CTkButton(
            btn_frame, text="✓  Yes", width=120, height=40,
            fg_color=COLORS["success"], hover_color="#0d9e6f",
            font=(FONT_FAMILY, 13, "bold"), corner_radius=12, command=self._yes
        )
        yes_btn.pack(side="left", padx=8)

        no_btn = ctk.CTkButton(
            btn_frame, text="✕  No", width=120, height=40,
            fg_color=COLORS["danger"], hover_color="#d13636",
            font=(FONT_FAMILY, 13, "bold"), corner_radius=12, command=self._no
        )
        no_btn.pack(side="left", padx=8)

        mic_btn = ctk.CTkButton(
            btn_frame, text="🎤", width=40, height=40,
            fg_color=COLORS["bg_card"], hover_color=COLORS["bg_card_hover"],
            font=(FONT_FAMILY, 16), corner_radius=12, command=self._voice_confirm
        )
        mic_btn.pack(side="left", padx=8)

        self.protocol("WM_DELETE_WINDOW", self._no)
        self.wait_window()

    def _yes(self):
        self.result = True
        self.destroy()

    def _no(self):
        self.result = False
        self.destroy()

    def _voice_confirm(self):
        import speech_recognition as sr
        recognizer = sr.Recognizer()
        recognizer.pause_threshold = 0.8
        try:
            with sr.Microphone() as source:
                audio = recognizer.listen(source, timeout=4, phrase_time_limit=4)
                text = recognizer.recognize_google(audio).lower()
                yes_words = ["yes", "yeah", "yep", "ya", "yas", "yup",
                             "sure", "ok", "okay", "do it", "go ahead", "proceed"]
                self.result = any(word in text for word in yes_words)
        except Exception:
            self.result = False
        self.destroy()


# ==========================================
# INPUT DIALOG
# ==========================================
class InputDialog(ctk.CTkToplevel):
    """A modal dialog with a text input field."""

    def __init__(self, parent, prompt="Enter value:", default=""):
        super().__init__(parent)
        self.result = None
        self.title("NOVA — Input")
        self.geometry("460x220")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg_dark"])
        self.attributes("-topmost", True)
        self.grab_set()

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 230
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 110
        self.geometry(f"+{x}+{y}")

        icon_label = ctk.CTkLabel(self, text="✏️", font=(FONT_FAMILY, 32), text_color=COLORS["accent_cyan"])
        icon_label.pack(pady=(20, 5))

        prompt_label = ctk.CTkLabel(
            self, text=prompt, font=(FONT_FAMILY, 13),
            text_color=COLORS["text_primary"], wraplength=400
        )
        prompt_label.pack(pady=(0, 12))

        self._entry = ctk.CTkEntry(
            self, width=360, height=40, font=(FONT_FAMILY, 13),
            fg_color=COLORS["bg_input"], border_color=COLORS["border"],
            border_width=1, text_color=COLORS["text_primary"], corner_radius=10
        )
        self._entry.pack(pady=(0, 16))
        if default:
            self._entry.insert(0, default)
        self._entry.focus_set()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=(0, 20))

        ok_btn = ctk.CTkButton(
            btn_frame, text="✓  OK", width=120, height=38,
            fg_color=COLORS["success"], hover_color="#0d9e6f",
            font=(FONT_FAMILY, 13, "bold"), corner_radius=12, command=self._ok
        )
        ok_btn.pack(side="left", padx=8)

        cancel_btn = ctk.CTkButton(
            btn_frame, text="✕  Cancel", width=120, height=38,
            fg_color=COLORS["danger"], hover_color="#d13636",
            font=(FONT_FAMILY, 13, "bold"), corner_radius=12, command=self._cancel
        )
        cancel_btn.pack(side="left", padx=8)

        self.bind("<Return>", lambda e: self._ok())
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _ok(self):
        text = self._entry.get().strip()
        if text:
            self.result = text
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ==========================================
# CHOICE DIALOG
# ==========================================
class ChoiceDialog(ctk.CTkToplevel):
    """A modal dialog showing a list of options to choose from."""

    def __init__(self, parent, prompt="Select an option:", options=None):
        super().__init__(parent)
        self.result = None
        self.title("NOVA — Select")

        num_options = len(options) if options else 0
        height = min(160 + num_options * 44, 520)
        self.geometry(f"520x{height}")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg_dark"])
        self.attributes("-topmost", True)
        self.grab_set()

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 260
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - height // 2
        self.geometry(f"+{x}+{y}")

        prompt_label = ctk.CTkLabel(
            self, text=prompt, font=(FONT_FAMILY, 14, "bold"),
            text_color=COLORS["accent_cyan"], wraplength=460
        )
        prompt_label.pack(pady=(20, 12))

        options_frame = ctk.CTkScrollableFrame(
            self, fg_color="transparent", height=min(num_options * 44, 320)
        )
        options_frame.pack(fill="x", padx=20, pady=(0, 12))

        if options:
            for i, opt in enumerate(options):
                btn = ctk.CTkButton(
                    options_frame, text=f"  [{i+1}]  {opt}",
                    font=(FONT_FAMILY, 11), fg_color=COLORS["bg_card"],
                    hover_color=COLORS["bg_card_hover"], border_width=1,
                    border_color=COLORS["border"], corner_radius=10,
                    height=38, anchor="w", command=lambda idx=i: self._select(idx)
                )
                btn.pack(fill="x", pady=2)

        cancel_btn = ctk.CTkButton(
            self, text="✕  Cancel", width=140, height=38,
            fg_color=COLORS["danger"], hover_color="#d13636",
            font=(FONT_FAMILY, 13, "bold"), corner_radius=12, command=self._cancel
        )
        cancel_btn.pack(pady=(0, 16))

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _select(self, idx):
        self.result = idx
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ==========================================
# MAIN APPLICATION WINDOW
# ==========================================
class NovaApp(ctk.CTk):
    """The main NOVA desktop application."""

    def __init__(self):
        super().__init__()

        # --- Window Setup ---
        self.title("NOVA")
        self.geometry("960x680")
        self.minsize(800, 560)
        self.configure(fg_color=COLORS["bg_darkest"])
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        try:
            self.iconbitmap(default="")
        except:
            pass

        self._is_listening = False
        self._processing = False
        self._wake_word_active = False
        self._wake_word_enabled = False

        # --- Register callbacks with the core ---
        assistant_core.set_speak_callback(self._on_assistant_speak)
        assistant_core.set_status_callback(self._on_status_update)
        assistant_core.set_confirm_callback(self._on_confirm_request)
        assistant_core.set_input_callback(self._on_input_request)
        assistant_core.set_choice_callback(self._on_choice_request)

        # --- Build UI ---
        self._build_header()
        self._build_body()
        self._build_input_bar()

        # --- Initialize assistant core ---
        assistant_core.initialize()
        self._add_assistant_message(
            'Hello! I\'m NOVA, your voice assistant. Type a command or say "OK NOVA" to speak.'
        )

        # Bind Enter key
        self.bind("<Return>", lambda e: self._send_command())

        # --- Auto-start wake word listener after a delay ---
        self.after(2500, self._start_wake_word_listener)

    # ------------------------------------------
    # HEADER
    # ------------------------------------------
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=COLORS["bg_dark"], height=60, corner_radius=0)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # Left side: logo + name
        left_frame = ctk.CTkFrame(header, fg_color="transparent")
        left_frame.pack(side="left", padx=20, fill="y")

        logo_label = ctk.CTkLabel(
            left_frame, text="⬡", font=(FONT_FAMILY, 28), text_color=COLORS["accent_cyan"]
        )
        logo_label.pack(side="left", padx=(0, 8), pady=12)

        name_label = ctk.CTkLabel(
            left_frame, text="N O V A", font=(FONT_FAMILY, 20, "bold"),
            text_color=COLORS["text_primary"]
        )
        name_label.pack(side="left", pady=12)

        subtitle = ctk.CTkLabel(
            left_frame, text="  Voice Assistant", font=(FONT_FAMILY, 11),
            text_color=COLORS["text_secondary"]
        )
        subtitle.pack(side="left", pady=12)

        # Right side: wake word toggle + status
        right_frame = ctk.CTkFrame(header, fg_color="transparent")
        right_frame.pack(side="right", padx=20, fill="y")

        self._status_dot = ctk.CTkLabel(
            right_frame, text="●", font=(FONT_FAMILY, 14), text_color=COLORS["warning"]
        )
        self._status_dot.pack(side="right", padx=(6, 0), pady=12)

        self._status_label = ctk.CTkLabel(
            right_frame, text="Indexing...", font=(FONT_FAMILY, 11),
            text_color=COLORS["text_secondary"]
        )
        self._status_label.pack(side="right", pady=12)

        # Separator
        sep = ctk.CTkLabel(
            right_frame, text="│", font=(FONT_FAMILY, 14), text_color=COLORS["text_muted"]
        )
        sep.pack(side="right", padx=12, pady=12)

        # Wake word indicator dot
        self._wake_dot = ctk.CTkLabel(
            right_frame, text="●", font=(FONT_FAMILY, 12),
            text_color=COLORS["wake_inactive"]
        )
        self._wake_dot.pack(side="right", padx=(0, 4), pady=12)

        # Wake word toggle button
        self._wake_btn = ctk.CTkButton(
            right_frame, text="🔇 OK NOVA", width=110, height=30,
            font=(FONT_FAMILY, 10, "bold"), fg_color=COLORS["bg_card"],
            hover_color=COLORS["bg_card_hover"], border_width=1,
            border_color=COLORS["wake_inactive"], corner_radius=8,
            command=self._toggle_wake_word
        )
        self._wake_btn.pack(side="right", pady=12)

    # ------------------------------------------
    # BODY (Chat + Sidebar)
    # ------------------------------------------
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=0, pady=0)

        # --- Chat area (left) ---
        chat_container = ctk.CTkFrame(body, fg_color=COLORS["bg_darkest"], corner_radius=0)
        chat_container.pack(side="left", fill="both", expand=True)

        self._chat_scroll = ctk.CTkScrollableFrame(
            chat_container, fg_color=COLORS["bg_darkest"],
            scrollbar_button_color=COLORS["bg_card"],
            scrollbar_button_hover_color=COLORS["accent_purple"],
        )
        self._chat_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # --- Sidebar (right) ---
        sidebar = ctk.CTkFrame(body, fg_color=COLORS["bg_dark"], width=260, corner_radius=0,
                               border_width=1, border_color=COLORS["border"])
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)

        # -- Current Directory section --
        dir_header = ctk.CTkLabel(
            sidebar, text="📍  CURRENT DIRECTORY", font=(FONT_FAMILY, 10, "bold"),
            text_color=COLORS["text_secondary"], anchor="w"
        )
        dir_header.pack(fill="x", padx=16, pady=(20, 4))

        self._dir_label = ctk.CTkLabel(
            sidebar, text=os.getcwd(), font=(FONT_FAMILY, 10),
            text_color=COLORS["accent_cyan"], wraplength=220, anchor="w", justify="left"
        )
        self._dir_label.pack(fill="x", padx=16, pady=(0, 16))

        sep1 = ctk.CTkFrame(sidebar, fg_color=COLORS["border"], height=1)
        sep1.pack(fill="x", padx=16, pady=4)

        # -- System Status section --
        status_header = ctk.CTkLabel(
            sidebar, text="⏳  SYSTEM STATUS", font=(FONT_FAMILY, 10, "bold"),
            text_color=COLORS["text_secondary"], anchor="w"
        )
        status_header.pack(fill="x", padx=16, pady=(16, 4))

        self._indexing_label = ctk.CTkLabel(
            sidebar, text="Indexing in progress...", font=(FONT_FAMILY, 10),
            text_color=COLORS["warning"], anchor="w"
        )
        self._indexing_label.pack(fill="x", padx=16, pady=(0, 4))

        self._stats_label = ctk.CTkLabel(
            sidebar, text="", font=(FONT_FAMILY, 10),
            text_color=COLORS["text_muted"], anchor="w", wraplength=220, justify="left"
        )
        self._stats_label.pack(fill="x", padx=16, pady=(0, 16))

        sep2 = ctk.CTkFrame(sidebar, fg_color=COLORS["border"], height=1)
        sep2.pack(fill="x", padx=16, pady=4)

        # -- Quick Actions --
        qa_header = ctk.CTkLabel(
            sidebar, text="⚡  QUICK ACTIONS", font=(FONT_FAMILY, 10, "bold"),
            text_color=COLORS["text_secondary"], anchor="w"
        )
        qa_header.pack(fill="x", padx=16, pady=(16, 10))

        quick_actions = [
            ("📂  Open App", "open_app_dialog"),
            ("📄  Open Document", "open_doc_dialog"),
            ("🏠  Go to Desktop", "go to desktop"),
            ("📥  Go to Downloads", "go to downloads"),
            ("📍  Where Am I?", "where am i"),
            ("🔙  Go Back", "go back"),
            ("🔧  Set Env Variable", "set environment variable"),
            ("⚡  Cancel Shutdown", "cancel shutdown"),
        ]

        for label, cmd in quick_actions:
            btn = ctk.CTkButton(
                sidebar, text=label, font=(FONT_FAMILY, 11),
                fg_color=COLORS["bg_card"], hover_color=COLORS["bg_card_hover"],
                border_width=1, border_color=COLORS["border"],
                corner_radius=10, height=36, anchor="w",
                command=lambda c=cmd: self._quick_action(c)
            )
            btn.pack(fill="x", padx=16, pady=3)

        # -- Bottom branding --
        brand_label = ctk.CTkLabel(
            sidebar, text="NOVA v1.0  ·  Powered by Groq AI",
            font=(FONT_FAMILY, 9), text_color=COLORS["text_muted"]
        )
        brand_label.pack(side="bottom", pady=12)

    # ------------------------------------------
    # INPUT BAR
    # ------------------------------------------
    def _build_input_bar(self):
        input_bar = ctk.CTkFrame(self, fg_color=COLORS["bg_dark"], height=70, corner_radius=0,
                                 border_width=1, border_color=COLORS["border"])
        input_bar.pack(fill="x", side="bottom")
        input_bar.pack_propagate(False)

        inner = ctk.CTkFrame(input_bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=16, pady=12)

        self._input_entry = ctk.CTkEntry(
            inner, placeholder_text="Type a command or say 'OK NOVA'...",
            font=(FONT_FAMILY, 13), fg_color=COLORS["bg_input"],
            border_color=COLORS["border"], border_width=1,
            text_color=COLORS["text_primary"],
            placeholder_text_color=COLORS["text_muted"], corner_radius=14, height=44
        )
        self._input_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self._mic_btn = ctk.CTkButton(
            inner, text="🎤", width=44, height=44, font=(FONT_FAMILY, 18),
            fg_color=COLORS["bg_card"], hover_color=COLORS["accent_cyan"],
            border_width=2, border_color=COLORS["mic_idle"], corner_radius=22,
            command=self._toggle_mic
        )
        self._mic_btn.pack(side="left", padx=(0, 8))

        send_btn = ctk.CTkButton(
            inner, text="➤", width=44, height=44, font=(FONT_FAMILY, 18),
            fg_color=COLORS["accent_purple"], hover_color=COLORS["accent_blue"],
            corner_radius=22, command=self._send_command
        )
        send_btn.pack(side="left")

    # ------------------------------------------
    # CHAT MANAGEMENT
    # ------------------------------------------
    def _add_user_message(self, text):
        bubble = ChatBubble(self._chat_scroll, text, is_user=True)
        bubble.pack(fill="x", pady=2)
        self._scroll_to_bottom()

    def _add_assistant_message(self, text):
        bubble = ChatBubble(self._chat_scroll, text, is_user=False)
        bubble.pack(fill="x", pady=2)
        self._scroll_to_bottom()

    def _scroll_to_bottom(self):
        self.after(100, lambda: self._chat_scroll._parent_canvas.yview_moveto(1.0))

    # ------------------------------------------
    # COMMAND HANDLING
    # ------------------------------------------
    def _send_command(self):
        text = self._input_entry.get().strip()
        if not text or self._processing:
            return

        self._input_entry.delete(0, "end")
        self._add_user_message(text)

        if any(w in text.lower() for w in ["exit", "quit", "goodbye", "bye"]):
            self._add_assistant_message("Goodbye! 👋")
            self.after(1500, self.destroy)
            return

        self._processing = True
        self._status_label.configure(text="Processing...", text_color=COLORS["accent_cyan"])
        self._status_dot.configure(text_color=COLORS["accent_cyan"])

        threading.Thread(target=self._process_in_background, args=(text,), daemon=True).start()

    def _process_in_background(self, text):
        try:
            result = assistant_core.process_command(text)
        except Exception as e:
            self.after(0, lambda: self._add_assistant_message(f"Error: {e}"))
        finally:
            self.after(0, self._reset_status)

    def _reset_status(self):
        self._processing = False
        self._last_command_time = time.time()  # Cooldown timer for wake word
        if assistant_core.is_indexing:
            self._status_label.configure(text="Indexing...", text_color=COLORS["warning"])
            self._status_dot.configure(text_color=COLORS["warning"])
        else:
            self._status_label.configure(text="Ready", text_color=COLORS["success"])
            self._status_dot.configure(text_color=COLORS["success"])

    def _quick_action(self, cmd):
        """Handle quick action buttons."""
        if cmd == "open_app_dialog":
            dialog = InputDialog(self, prompt="Which application would you like to open?")
            if dialog.result:
                self._input_entry.delete(0, "end")
                self._input_entry.insert(0, f"open {dialog.result}")
                self._send_command()
            return

        if cmd == "open_doc_dialog":
            dialog = InputDialog(
                self, prompt="Which document would you like to open?\n(e.g., 'resume', 'project report', 'notes')"
            )
            if dialog.result:
                self._input_entry.delete(0, "end")
                self._input_entry.insert(0, f"open document {dialog.result}")
                self._send_command()
            return

        # Otherwise send directly
        self._input_entry.delete(0, "end")
        self._input_entry.insert(0, cmd)
        self._send_command()

    # ------------------------------------------
    # MICROPHONE
    # ------------------------------------------
    @staticmethod
    def _play_listen_chime():
        """Play a short two-tone chime like Google's mic activation sound."""
        try:
            winsound.Beep(600, 120)
            winsound.Beep(900, 180)
        except Exception:
            pass

    def _toggle_mic(self):
        if self._is_listening or self._processing:
            return
        self._is_listening = True
        self._mic_btn.configure(fg_color=COLORS["mic_active"], border_color=COLORS["mic_active"])
        self._status_label.configure(text="Listening...", text_color=COLORS["mic_active"])
        self._status_dot.configure(text_color=COLORS["mic_active"])
        # Play chime on a separate thread so the UI doesn't freeze
        threading.Thread(target=self._play_listen_chime, daemon=True).start()
        threading.Thread(target=self._listen_mic, daemon=True).start()

    def _listen_mic(self):
        import speech_recognition as sr
        recognizer = sr.Recognizer()
        recognizer.pause_threshold = 0.8
        text = None
        try:
            with sr.Microphone() as source:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
                text = recognizer.recognize_google(audio)
        except Exception:
            text = None
        self.after(0, self._reset_mic, text)

    def _reset_mic(self, text=None):
        self._is_listening = False
        self._mic_btn.configure(fg_color=COLORS["bg_card"], border_color=COLORS["mic_idle"])
        self._reset_status()

        if text:
            self._input_entry.delete(0, "end")
            self._input_entry.insert(0, text)
            self._send_command()
        else:
            self._add_assistant_message("I didn't catch that. Please try again.")

    # ------------------------------------------
    # WAKE WORD LISTENER ("OK NOVA")
    # ------------------------------------------
    def _start_wake_word_listener(self):
        """Start the background wake word detection thread."""
        if self._wake_word_active:
            return
        self._wake_word_active = True
        self._wake_word_enabled = True
        self._update_wake_word_ui()
        threading.Thread(target=self._wake_word_loop, daemon=True).start()

    def _stop_wake_word_listener(self):
        """Stop the wake word detection."""
        self._wake_word_active = False
        self._wake_word_enabled = False
        self._update_wake_word_ui()

    def _toggle_wake_word(self):
        """Toggle wake word listener on/off."""
        if self._wake_word_enabled:
            self._stop_wake_word_listener()
        else:
            self._start_wake_word_listener()

    def _update_wake_word_ui(self):
        """Update the wake word button/indicator appearance."""
        if self._wake_word_enabled:
            self._wake_dot.configure(text_color=COLORS["wake_active"])
            self._wake_btn.configure(border_color=COLORS["wake_active"], text="👂 OK NOVA")
        else:
            self._wake_dot.configure(text_color=COLORS["wake_inactive"])
            self._wake_btn.configure(border_color=COLORS["wake_inactive"], text="🔇 OK NOVA")

    @staticmethod
    def _is_wake_word(text):
        """
        Check if recognized text contains the wake word 'nova'.
        Uses both direct substring matching and fuzzy matching
        on individual words to catch mis-transcriptions.
        """
        text = text.lower()

        # Direct check — covers "ok nova", "hey nova", etc.
        if "nova" in text:
            return True

        # Fuzzy check each word — catches "no va", "noba", "novaah", etc.
        target_words = ["nova", "novo"]
        words = text.split()
        for word in words:
            matches = difflib.get_close_matches(word, target_words, n=1, cutoff=0.6)
            if matches:
                return True

        return False

    def _wake_word_loop(self):
        """
        Always-on wake word detection using listen_in_background().
        The microphone stays permanently open — zero gaps, zero missed audio.
        A callback fires instantly whenever speech is detected.
        """
        import speech_recognition as sr

        recognizer = sr.Recognizer()
        recognizer.pause_threshold = 0.5
        recognizer.dynamic_energy_threshold = True

        # Calibrate ambient noise
        try:
            with sr.Microphone() as cal_source:
                recognizer.adjust_for_ambient_noise(cal_source, duration=1)
            # Lower threshold for better sensitivity
            recognizer.energy_threshold = max(recognizer.energy_threshold * 0.6, 100)
            print(f"[Wake Word] Calibrated (threshold={int(recognizer.energy_threshold)})")
        except Exception as e:
            print(f"[Wake Word] Mic calibration failed: {e}")
            return

        # Open a dedicated mic source for the background listener
        wake_mic = sr.Microphone()

        def on_audio(rec, audio):
            """Callback — fires instantly whenever the background listener captures speech."""
            # Skip if busy, disabled, or in cooldown after a command
            if not self._wake_word_active:
                return
            if self._is_listening or self._processing:
                return
            if time.time() - getattr(self, '_last_command_time', 0) < 3:
                return  # 3-second cooldown prevents TTS self-triggering

            try:
                text = rec.recognize_google(audio).lower()
                if self._is_wake_word(text):
                    print(f"[Wake Word] DETECTED: '{text}'")
                    self.after(0, self._on_wake_word_detected)
            except sr.UnknownValueError:
                pass
            except sr.RequestError:
                pass
            except Exception:
                pass

        # Start the persistent background listener (mic stays open, zero gaps)
        stop_fn = recognizer.listen_in_background(wake_mic, on_audio, phrase_time_limit=3)
        self._bg_stop_fn = stop_fn
        print("[Wake Word] ✅ Always-on listener active — say 'OK NOVA' anytime.")

        # Keep this thread alive while feature is enabled
        while self._wake_word_active:
            time.sleep(0.5)

        # Cleanup when disabled
        stop_fn(wait_for_stop=False)
        print("[Wake Word] Stopped.")

    def _on_wake_word_detected(self):
        """Called on main thread when wake word is detected."""
        if self._is_listening or self._processing:
            return
        self._add_assistant_message("👂 I heard you! Listening for your command...")
        self._toggle_mic()

    # ------------------------------------------
    # CALLBACKS FROM ASSISTANT CORE
    # ------------------------------------------
    def _on_assistant_speak(self, text):
        """Called by assistant_core.speak() — adds message to chat."""
        self.after(0, lambda: self._add_assistant_message(text))

    def _on_status_update(self, key, value):
        """Called by assistant_core when status changes."""
        if key == "current_dir":
            self.after(0, lambda: self._dir_label.configure(text=value))
        elif key == "indexing":
            if value == "Done":
                self.after(0, lambda: self._indexing_label.configure(
                    text="✅ Indexing complete", text_color=COLORS["success"]))
                self.after(0, lambda: self._reset_status())
            else:
                self.after(0, lambda: self._indexing_label.configure(
                    text=f"⏳ {value}", text_color=COLORS["warning"]))
        elif key == "stats":
            self.after(0, lambda: self._stats_label.configure(text=value))

    def _on_confirm_request(self, prompt):
        """Shows a modal confirmation dialog. Called from background thread."""
        result_holder = {"value": False, "done": False}

        def show_dialog():
            dialog = ConfirmDialog(self, prompt)
            result_holder["value"] = dialog.result
            result_holder["done"] = True

        self.after(0, show_dialog)

        while not result_holder["done"]:
            time.sleep(0.05)

        return result_holder["value"]

    def _on_input_request(self, prompt):
        """Shows a modal input dialog. Called from background thread."""
        result_holder = {"value": None, "done": False}

        def show_dialog():
            dialog = InputDialog(self, prompt)
            result_holder["value"] = dialog.result
            result_holder["done"] = True

        self.after(0, show_dialog)

        while not result_holder["done"]:
            time.sleep(0.05)

        return result_holder["value"]

    def _on_choice_request(self, prompt, options):
        """Shows a modal choice dialog. Called from background thread."""
        result_holder = {"value": None, "done": False}

        def show_dialog():
            dialog = ChoiceDialog(self, prompt, options)
            result_holder["value"] = dialog.result
            result_holder["done"] = True

        self.after(0, show_dialog)

        while not result_holder["done"]:
            time.sleep(0.05)

        return result_holder["value"]


# ==========================================
# ENTRY POINT
# ==========================================
if __name__ == "__main__":
    app = NovaApp()
    app.mainloop()
