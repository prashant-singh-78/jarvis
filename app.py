import asyncio
import ctypes
import difflib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path
from queue import Empty, Queue

import pyttsx3
import speech_recognition as sr
import tkinter as tk
import pythoncom

try:
    import pyaudio
except ImportError:
    pyaudio = None

try:
    import win32com.client
except ImportError:
    win32com = None

try:
    import win32api
    import win32event
    import winerror
except ImportError:
    win32api = None
    win32event = None
    winerror = None

try:
    import cv2
except ImportError:
    cv2 = None

try:
    import mediapipe as mp
except ImportError:
    mp = None

try:
    from ctypes import POINTER, cast

    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except ImportError:
    POINTER = None
    cast = None
    CLSCTX_ALL = None
    AudioUtilities = None
    IAudioEndpointVolume = None

APP_NAME = "Mikasa"
SINGLE_INSTANCE_MUTEX_NAME = "Local\\MikasaJarviseSingleInstance"
API_KEY_ENV = "GEMINI_API_KEY"
MODEL_ENV = "GEMINI_MODEL"
DEFAULT_MODEL = "gemini-1.5-flash"
LIVE_MODEL_ENV = "GEMINI_LIVE_MODEL"
DEFAULT_LIVE_MODEL = "gemini-3.1-flash-live-preview"
LIVE_VOICE_ENV = "GEMINI_LIVE_VOICE"
DEFAULT_LIVE_VOICE = "Kore"
LIVE_ENABLED_ENV = "GEMINI_ENABLE_LIVE"
LIVE_AUDIO_INPUT_RATE = 16000
LIVE_AUDIO_OUTPUT_RATE = 24000
WINDOW_SIZE = 280
BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
APP_INDEX_CACHE_FILE = BASE_DIR / "app_index_cache.json"

assistant_running = threading.Event()
assistant_running.set()
speaking_event = threading.Event()
speech_queue = Queue()
chat_lock = threading.Lock()
chat_session = None
chat_status = ""
app_index_lock = threading.Lock()
app_index_refresh_lock = threading.Lock()
app_index = None
live_controller = None
single_instance_mutex = None
voice_engine = None
gesture_control_enabled = threading.Event()
gesture_lock = threading.Lock()
OPEN_ACTIONS = (
    "open karo",
    "open kar",
    "open",
    "launch",
    "start",
    "run",
    "khol do",
    "kholo",
    "khol",
    "chalu karo",
    "chalao",
)
CLOSE_ACTIONS = (
    "close",
    "stop",
    "band karo",
    "band kar",
    "band",
    "exit",
)
TARGET_NOISE_PHRASES = (
    "application",
    "software",
    "app",
    "please",
    "plz",
    "jara",
    "zara",
    "mere liye",
    "mere pc me",
    "mere pc mein",
    "mere system me",
    "mere system mein",
    "mere computer me",
    "mere computer mein",
    "pc ke andar",
    "system ke andar",
    "computer ke andar",
)
APP_NAME_ALIASES = {
    "code": "visual studio code",
    "vs code": "visual studio code",
    "v s code": "visual studio code",
    "vscode": "visual studio code",
    "visual code": "visual studio code",
    "visual studio code": "visual studio code",
    "anti gravity": "antigravity",
    "ant gravity": "antigravity",
    "antgigravity": "antigravity",
    "ant gigravity": "antigravity",
    "anti gigravity": "antigravity",
    "antigravity": "antigravity",
}
robot_anim_state = {
    "blink_until": 0.0,
    "next_idle_blink": 0.0,
    "mouth_phase": 0,
    "talk_eyes_open": True,
}


def load_local_env():
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def env_flag(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def hide_console_window():
    if os.name != "nt":
        return

    try:
        console_window = ctypes.windll.kernel32.GetConsoleWindow()
        if console_window:
            ctypes.windll.user32.ShowWindow(console_window, 0)
    except Exception:
        pass


def acquire_single_instance():
    if os.name != "nt":
        return True

    global single_instance_mutex
    if win32event and win32api and winerror:
        try:
            mutex = win32event.CreateMutex(None, False, SINGLE_INSTANCE_MUTEX_NAME)
            already_exists = win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
            if already_exists:
                win32api.CloseHandle(mutex)
                return False
            single_instance_mutex = mutex
            return True
        except Exception:
            pass

    try:
        kernel32 = ctypes.windll.kernel32
        mutex = kernel32.CreateMutexW(None, False, SINGLE_INSTANCE_MUTEX_NAME)
        already_exists = kernel32.GetLastError() == 183
        if already_exists:
            if mutex:
                kernel32.CloseHandle(mutex)
            return False
        single_instance_mutex = mutex
        return True
    except Exception:
        return True


def release_single_instance():
    if os.name != "nt":
        return

    global single_instance_mutex
    if not single_instance_mutex:
        return

    try:
        if win32api:
            win32api.CloseHandle(single_instance_mutex)
        else:
            ctypes.windll.kernel32.CloseHandle(single_instance_mutex)
    except Exception:
        pass
    finally:
        single_instance_mutex = None


def focus_existing_window():
    if os.name != "nt":
        return False

    try:
        user32 = ctypes.windll.user32
        window_handle = user32.FindWindowW(None, f"{APP_NAME} Assistant")
        if not window_handle:
            return False
        user32.ShowWindow(window_handle, 9)
        user32.SetForegroundWindow(window_handle)
        return True
    except Exception:
        return False


def init_voice_engine():
    if pyttsx3 is None:
        return None

    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        engine = pyttsx3.init()
        voices = engine.getProperty("voices")
    except Exception as exc:
        print(f"[voice-init] {exc}")
        return None

    preferred_voice = None
    for voice in voices:
        voice_name = f"{getattr(voice, 'name', '')} {getattr(voice, 'id', '')}".lower()
        if "female" in voice_name or "zira" in voice_name or "hazel" in voice_name:
            preferred_voice = voice.id
            break

    try:
        if preferred_voice:
            engine.setProperty("voice", preferred_voice)
        elif voices:
            engine.setProperty("voice", voices[0].id)

        engine.setProperty("rate", 165)
    except Exception as exc:
        print(f"[voice-config] {exc}")

    return engine


def speaker_loop():
    while True:
        text = speech_queue.get()
        if text is None:
            speech_queue.task_done()
            break

        if voice_engine is None:
            speech_queue.task_done()
            continue

        speaking_event.set()
        try:
            voice_engine.say(text)
            voice_engine.runAndWait()
        except Exception as exc:
            print(f"[voice] {exc}")
        finally:
            speaking_event.clear()
            speech_queue.task_done()


def speak(text):
    cleaned = " ".join(str(text).split())
    if cleaned and voice_engine is not None:
        speech_queue.put(cleaned)


def create_chat_session():
    api_key = os.getenv(API_KEY_ENV)
    if not api_key:
        return None, "Offline friend mode is on. Add a Gemini key in .env for full AI chat."

    try:
        import google.generativeai as genai
    except ImportError:
        return None, "Offline friend mode is on because google-generativeai is missing."

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(
            model_name=os.getenv(MODEL_ENV, DEFAULT_MODEL),
            system_instruction=(
                "Your name is Mikasa. You are Seenu's close AI friend. "
                "Talk warmly, naturally, supportively, and like a caring best friend. "
                "Keep replies short, smooth, and voice-friendly."
            ),
        )
        return model.start_chat(history=[]), "AI chat is ready."
    except Exception as exc:
        return None, f"Offline friend mode is on because Gemini could not start: {exc}"


class GeminiLiveController:
    def __init__(self):
        self.loop = None
        self.thread = None
        self.playback_thread = None
        self.session = None
        self.types_module = None
        self.ready_event = threading.Event()
        self.failed_event = threading.Event()
        self.stop_event = threading.Event()
        self.output_audio_queue = Queue()
        self.pending_input_transcript = ""
        self.pending_output_transcript = ""
        self.last_input_message = ""
        self.last_output_message = ""
        self.failure_reason = ""

    def start(self):
        api_key = os.getenv(API_KEY_ENV)
        if not api_key:
            return False, "Gemini Live needs a Gemini API key in .env."
        if pyaudio is None:
            return False, "Gemini Live needs PyAudio for native audio playback."

        try:
            from google import genai  # noqa: F401
            from google.genai import types  # noqa: F401
        except ImportError:
            return False, "Gemini Live needs the google-genai package."

        self.playback_thread = threading.Thread(target=self._playback_loop, daemon=True)
        self.playback_thread.start()

        self.thread = threading.Thread(target=self._run_event_loop, daemon=True)
        self.thread.start()

        if self.ready_event.wait(timeout=10):
            return True, "Gemini Live voice is ready."

        if self.failed_event.is_set():
            return False, self.failure_reason or "Gemini Live could not start."

        return False, "Gemini Live is taking too long to start."

    def is_active(self):
        return self.ready_event.is_set() and not self.stop_event.is_set()

    def send_audio_turn(self, audio_bytes):
        if not audio_bytes or not self.is_active() or not self.loop:
            return False

        future = asyncio.run_coroutine_threadsafe(self._send_audio_turn(audio_bytes), self.loop)
        try:
            future.result(timeout=0.5)
        except Exception as exc:
            print(f"[live] send failed: {exc}")
            return False
        return True

    def stop(self):
        if self.stop_event.is_set():
            return

        self.stop_event.set()
        self.ready_event.clear()
        self._clear_output_audio_queue()
        self.output_audio_queue.put(None)

        if self.loop and self.session:
            try:
                future = asyncio.run_coroutine_threadsafe(self.session.close(), self.loop)
                future.result(timeout=2)
            except Exception:
                pass

    def _run_event_loop(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_until_complete(self._live_session_loop())
        except Exception as exc:
            self.failure_reason = f"Gemini Live error: {exc}"
            self.failed_event.set()
            self._post_status(self.failure_reason)
        finally:
            self.ready_event.clear()
            try:
                pending = asyncio.all_tasks(self.loop)
                for task in pending:
                    task.cancel()
            except Exception:
                pass
            self.loop.close()

    async def _live_session_loop(self):
        from google import genai
        from google.genai import types

        self.types_module = types
        voice_name = os.getenv(LIVE_VOICE_ENV, DEFAULT_LIVE_VOICE)
        model_name = os.getenv(LIVE_MODEL_ENV, DEFAULT_LIVE_MODEL)
        config = {
            "response_modalities": ["AUDIO"],
            "input_audio_transcription": {},
            "output_audio_transcription": {},
            "speech_config": {
                "voice_config": {"prebuilt_voice_config": {"voice_name": voice_name}}
            },
        }

        client = genai.Client(api_key=os.getenv(API_KEY_ENV))
        async with client.aio.live.connect(model=model_name, config=config) as session:
            self.session = session
            self.ready_event.set()
            self._post_status("Gemini Live voice ready. Talk to me.")

            async for response in session.receive():
                if self.stop_event.is_set():
                    break
                await self._handle_live_response(response)

    async def _handle_live_response(self, response):
        content = getattr(response, "server_content", None)
        if not content:
            return

        if getattr(content, "interrupted", False):
            self._clear_output_audio_queue()
            self.pending_output_transcript = ""
            self._post_status("Interrupted. Listening again.")

        input_transcription = getattr(content, "input_transcription", None)
        if input_transcription and getattr(input_transcription, "text", None):
            transcript = input_transcription.text.strip()
            if transcript:
                self.pending_input_transcript = transcript
                self._post_status(f"Heard: {transcript}")

        output_transcription = getattr(content, "output_transcription", None)
        if output_transcription and getattr(output_transcription, "text", None):
            transcript = output_transcription.text.strip()
            if transcript:
                self.pending_output_transcript = transcript
                self._post_status(transcript)

        model_turn = getattr(content, "model_turn", None)
        if model_turn and getattr(model_turn, "parts", None):
            for part in model_turn.parts:
                inline_data = getattr(part, "inline_data", None)
                chunk = getattr(inline_data, "data", None) if inline_data else None
                if chunk:
                    self.output_audio_queue.put(chunk)

        if getattr(content, "turn_complete", False):
            self._flush_turn_messages()

    async def _send_audio_turn(self, audio_bytes):
        if not self.session or not self.types_module:
            return

        await self.session.send_realtime_input(
            audio=self.types_module.Blob(
                data=audio_bytes,
                mime_type=f"audio/pcm;rate={LIVE_AUDIO_INPUT_RATE}",
            )
        )
        await self.session.send_realtime_input(audio_stream_end=True)

    def _flush_turn_messages(self):
        if self.pending_input_transcript and self.pending_input_transcript != self.last_input_message:
            add_message("You", self.pending_input_transcript)
            self.last_input_message = self.pending_input_transcript

        if self.pending_output_transcript and self.pending_output_transcript != self.last_output_message:
            add_message(APP_NAME, self.pending_output_transcript)
            self.last_output_message = self.pending_output_transcript

        self.pending_input_transcript = ""
        self.pending_output_transcript = ""
        self._post_status("Listening...")

    def _post_status(self, message):
        if "root" in globals():
            root.after(0, lambda text=message: update_status(text))

    def _clear_output_audio_queue(self):
        while True:
            try:
                queued = self.output_audio_queue.get_nowait()
            except Empty:
                break
            else:
                if queued is not None:
                    self.output_audio_queue.task_done()

    def _playback_loop(self):
        audio_interface = None
        audio_stream = None
        try:
            audio_interface = pyaudio.PyAudio()
            audio_stream = audio_interface.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=LIVE_AUDIO_OUTPUT_RATE,
                output=True,
            )

            while not self.stop_event.is_set():
                try:
                    chunk = self.output_audio_queue.get(timeout=0.2)
                except Empty:
                    speaking_event.clear()
                    continue

                if chunk is None:
                    self.output_audio_queue.task_done()
                    break

                try:
                    speaking_event.set()
                    audio_stream.write(chunk)
                except Exception as exc:
                    print(f"[live-audio] {exc}")
                finally:
                    self.output_audio_queue.task_done()

        except Exception as exc:
            self.failure_reason = f"Gemini Live audio output failed: {exc}"
            self.failed_event.set()
            self._post_status(self.failure_reason)
        finally:
            speaking_event.clear()
            if audio_stream is not None:
                try:
                    audio_stream.stop_stream()
                    audio_stream.close()
                except Exception:
                    pass
            if audio_interface is not None:
                try:
                    audio_interface.terminate()
                except Exception:
                    pass


def create_live_controller():
    if not env_flag(LIVE_ENABLED_ENV, default=True):
        return None, "Gemini Live is turned off in .env."

    controller = GeminiLiveController()
    ready, message = controller.start()
    if ready:
        return controller, message
    return None, message


def add_message(sender, message):
    timestamp = time.strftime("%H:%M:%S")
    print(f"[{timestamp}] {sender}: {message}")


def update_status(message):
    status_var.set(message)


def normalize_lookup_name(name):
    cleaned = re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()
    return " ".join(cleaned.split())


def compact_lookup_name(name):
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def register_app(app_map, display_name, launch_target):
    normalized = normalize_lookup_name(display_name)
    if normalized and normalized not in app_map:
        app_map[normalized] = {"name": display_name, "target": launch_target}

    compact = compact_lookup_name(display_name)
    if compact and compact not in app_map:
        app_map[compact] = {"name": display_name, "target": launch_target}


def load_cached_app_index():
    if not APP_INDEX_CACHE_FILE.exists():
        return None

    try:
        payload = json.loads(APP_INDEX_CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None

    app_map = payload.get("apps")
    if not isinstance(app_map, dict):
        return None

    cleaned_map = {}
    for key, value in app_map.items():
        if not isinstance(key, str) or not isinstance(value, dict):
            continue

        name = value.get("name")
        target = value.get("target")
        if isinstance(name, str) and isinstance(target, str):
            cleaned_map[key] = {"name": name, "target": target}
            compact = compact_lookup_name(name)
            if compact and compact not in cleaned_map:
                cleaned_map[compact] = {"name": name, "target": target}

    return cleaned_map or None


def save_app_index_cache(app_map):
    payload = {"generated_at": time.time(), "apps": app_map}
    temp_path = APP_INDEX_CACHE_FILE.with_suffix(".tmp")
    try:
        temp_path.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        temp_path.replace(APP_INDEX_CACHE_FILE)
    except OSError:
        pass


def build_app_index():
    app_map = {}

    builtin_apps = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "calc": "calc.exe",
        "paint": "mspaint.exe",
        "command prompt": "cmd.exe",
        "cmd": "cmd.exe",
        "powershell": "powershell.exe",
        "task manager": "taskmgr.exe",
        "file explorer": "explorer.exe",
        "explorer": "explorer.exe",
        "settings": "ms-settings:",
        "camera": "microsoft.windows.camera:",
        "photos": "ms-photos:",
    }
    for name, target in builtin_apps.items():
        register_app(app_map, name, target)

    search_roots = [
        Path(os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs")),
        Path(os.path.expandvars(r"%ProgramData%\Microsoft\Windows\Start Menu\Programs")),
        Path(os.path.expandvars(r"%USERPROFILE%\Desktop")),
        Path(os.path.expandvars(r"%PUBLIC%\Desktop")),
    ]

    allowed_suffixes = {".lnk", ".url", ".appref-ms", ".exe"}
    ignored_dirs = {"startup", "accessories", "administrative tools", "windows powershell"}

    for root in search_roots:
        if not root.exists():
            continue

        for path in root.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in allowed_suffixes:
                continue

            parent_name = path.parent.name.lower()
            if parent_name in ignored_dirs and path.suffix.lower() != ".exe":
                continue

            display_name = path.stem.replace("_", " ").replace("-", " ").strip()
            register_app(app_map, display_name, str(path))

    return app_map


def get_app_index():
    global app_index

    if app_index is not None:
        return app_index

    with app_index_lock:
        if app_index is None:
            app_index = load_cached_app_index()
            if app_index is None:
                app_index = build_app_index()
                save_app_index_cache(app_index)
    return app_index


def refresh_app_index():
    global app_index

    if not app_index_refresh_lock.acquire(blocking=False):
        return

    try:
        fresh_index = build_app_index()
        with app_index_lock:
            app_index = fresh_index
        save_app_index_cache(fresh_index)
    finally:
        app_index_refresh_lock.release()


def warm_app_index():
    threading.Thread(target=refresh_app_index, daemon=True).start()


def resolve_app_alias(target_name):
    normalized = normalize_lookup_name(target_name)
    compact = compact_lookup_name(target_name)
    alias = APP_NAME_ALIASES.get(normalized) or APP_NAME_ALIASES.get(compact)
    return alias or target_name


def find_app_match(target_name):
    lookup = normalize_lookup_name(resolve_app_alias(target_name))
    if not lookup:
        return None

    compact_lookup = compact_lookup_name(lookup)
    app_map = get_app_index()
    if lookup in app_map:
        return app_map[lookup]
    if compact_lookup in app_map:
        return app_map[compact_lookup]

    partial_matches = []
    seen_targets = set()
    lookup_tokens = set(lookup.split())
    for key, app in app_map.items():
        signature = app["target"]
        if signature in seen_targets:
            continue

        key_normalized = normalize_lookup_name(key)
        key_compact = compact_lookup_name(key)
        score = 0

        if lookup in key_normalized or key_normalized in lookup:
            score = max(score, 3)
        if compact_lookup and (compact_lookup in key_compact or key_compact in compact_lookup):
            score = max(score, 4)

        shared_tokens = len(lookup_tokens & set(key_normalized.split()))
        if shared_tokens:
            score = max(score, shared_tokens)

        if score:
            partial_matches.append((score, abs(len(key_compact) - len(compact_lookup)), app))
            seen_targets.add(signature)

    if partial_matches:
        partial_matches.sort(key=lambda item: (-item[0], item[1], len(item[2]["name"])))
        return partial_matches[0][2]

    close = difflib.get_close_matches(compact_lookup or lookup, list(app_map.keys()), n=1, cutoff=0.58)
    if close:
        return app_map[close[0]]

    return None


def contains_any_phrase(command, phrases):
    return any(phrase in command for phrase in phrases)


def strip_target_noise(target):
    cleaned = normalize_command(target)
    for phrase in sorted(TARGET_NOISE_PHRASES, key=len, reverse=True):
        cleaned = re.sub(rf"\b{re.escape(phrase)}\b", " ", cleaned)

    cleaned = re.sub(r"\b(?:ko|bhi)\b", " ", cleaned)
    return " ".join(cleaned.split())


def split_requested_targets(target_text):
    cleaned = strip_target_noise(target_text)
    if not cleaned:
        return []

    targets = []
    for part in re.split(r"\s*(?:,| and | aur | & )\s*", cleaned):
        current = strip_target_noise(part)
        if current and current not in targets:
            targets.append(current)
    return targets


def extract_action_target(command, action_phrases):
    normalized_command = normalize_command(command)
    for phrase in sorted(action_phrases, key=len, reverse=True):
        prefix = f"{phrase} "
        suffix = f" {phrase}"
        if normalized_command.startswith(prefix):
            return strip_target_noise(normalized_command[len(prefix):])
        if normalized_command.endswith(suffix):
            return strip_target_noise(normalized_command[: -len(suffix)])
    return ""


def looks_like_shell_target(target):
    lowered = str(target).strip().lower()
    if not lowered:
        return False

    if lowered.startswith(("http://", "https://")):
        return True
    if re.match(r"^[a-z][a-z0-9+.-]*:", lowered):
        return True
    return lowered.endswith((".lnk", ".url", ".appref-ms"))


def launch_windows_target(target, args=None):
    if not target:
        return False

    args = [str(arg) for arg in (args or [])]
    expanded = os.path.expandvars(str(target).strip())
    if not expanded:
        return False

    target_exists = os.path.exists(expanded)
    target_is_dir = os.path.isdir(expanded)
    shell_target = target_is_dir or looks_like_shell_target(expanded)
    executable = expanded if target_exists else shutil.which(expanded)

    if shell_target or (target_exists and not args):
        try:
            os.startfile(expanded)
            return True
        except OSError:
            pass

    launch_target = executable or expanded
    try:
        subprocess.Popen(
            [launch_target, *args],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except OSError:
        pass

    powershell_command = [
        "powershell",
        "-NoProfile",
        "-Command",
        "Start-Process",
        "-FilePath",
        launch_target,
    ]
    if args:
        powershell_command.extend(["-ArgumentList", subprocess.list2cmdline(args)])

    try:
        subprocess.Popen(
            powershell_command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except OSError:
        return False


def open_installed_app(target_name):
    match = find_app_match(target_name)
    if not match:
        return None

    launch_target = match["target"]
    if launch_windows_target(launch_target):
        return f"Opening {match['name']}."
    return f"I found {match['name']}, but I could not open it."


def open_requested_apps(target_names):
    messages = []
    for target_name in target_names:
        message = open_installed_app(target_name)
        messages.append(message or f"I could not find {target_name} on this PC.")
    return " ".join(messages)


def resolve_shortcut_target(path_text):
    path = Path(path_text)
    if path.suffix.lower() != ".lnk" or win32com is None:
        return path_text

    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(path))
        return shortcut.TargetPath or path_text
    except Exception as exc:
        print(f"[shortcut] {exc}")
        return path_text
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def get_process_names_for_target(launch_target):
    resolved_target = resolve_shortcut_target(launch_target)
    lower_target = str(resolved_target).lower()

    if lower_target.endswith(".exe"):
        return [Path(resolved_target).name]

    if lower_target.startswith(("http://", "https://")) or lower_target.endswith(".url"):
        return ["chrome.exe", "msedge.exe", "firefox.exe"]

    return []


def close_installed_app(target_name):
    match = find_app_match(target_name)
    if not match:
        return None

    process_names = get_process_names_for_target(match["target"])
    if not process_names:
        return f"I found {match['name']}, but I could not determine how to close it."

    if close_processes(process_names):
        return f"Closing {match['name']}."

    return f"{match['name']} abhi open nahin hai."


def close_requested_apps(target_names):
    messages = []
    for target_name in target_names:
        message = close_installed_app(target_name)
        messages.append(message or f"I could not find {target_name} on this PC.")
    return " ".join(messages)


def deliver_system_reply(message):
    add_message(APP_NAME, message)
    update_status(message)
    speak(message)


def open_installed_app_async(target_name):
    targets = target_name if isinstance(target_name, list) else [target_name]
    message = open_requested_apps(targets)

    root.after(0, lambda text=message: deliver_system_reply(text))


def close_installed_app_async(target_name):
    targets = target_name if isinstance(target_name, list) else [target_name]
    message = close_requested_apps(targets)

    root.after(0, lambda text=message: deliver_system_reply(text))


def local_friend_reply(command):
    text = command.lower()

    if any(word in text for word in ("hello", "hi", "hey", "hii")):
        return "Hi Seenu, main yahin hoon. Batao kya karna hai?"
    if "how are you" in text or "kaise ho" in text:
        return "Main theek hoon, aur tumhare saath hoon. Tum kaisa feel kar rahe ho?"
    if "your name" in text or "who are you" in text or "tumhara naam" in text:
        return "Main Mikasa hoon, tumhari AI friend."
    if "thank" in text or "shukriya" in text:
        return "Hamesha, Seenu. Main help ke liye ready hoon."
    if "love" in text or "miss you" in text:
        return "Aww, main bhi tumhare saath hoon. Ab bolo kya help chahiye?"
    if "time" in text:
        return f"Abhi {time.strftime('%I:%M %p')} ho raha hai."
    if "date" in text or "day" in text:
        return f"Aaj {time.strftime('%d %B %Y')} hai."
    if "sad" in text or "upset" in text:
        return "Main yahin hoon. Aaram se bolo, jo bhi dil mein hai."
    if "open" in text:
        return "Main open commands samajh rahi hoon. Bas app ya website ka naam bolo."
    if "close" in text or "band" in text:
        return "Main close commands bhi samajh rahi hoon. Bas kis cheez ko band karna hai woh bolo."

    return "Main sun rahi hoon, Seenu. Thoda simple bolo, main help karti hoon."


def launch_executable(candidates, args=None):
    args = args or []
    for candidate in candidates:
        if not candidate:
            continue

        expanded = os.path.expandvars(candidate)
        executable = expanded if os.path.exists(expanded) else shutil.which(expanded)
        if not executable:
            continue

        if launch_windows_target(executable, args):
            return True

    return False


def close_processes(process_names):
    closed_any = False
    for process_name in process_names:
        try:
            result = subprocess.run(
                ["taskkill", "/IM", process_name, "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
            if result.returncode == 0:
                closed_any = True
        except OSError:
            continue
    return closed_any


def open_url(url):
    webbrowser.open(url)
    return True


def open_site(name, url):
    open_url(url)
    return f"Opening {name}."


def open_chrome():
    chrome_candidates = [
        shutil.which("chrome"),
        r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
        r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
        r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
    ]
    if launch_executable(chrome_candidates):
        return "Opening Chrome."

    installed_message = open_installed_app("google chrome")
    if installed_message:
        return installed_message

    open_url("https://www.google.com")
    return "Chrome nahin mila, isliye maine default browser khol diya."


def close_chrome():
    if close_processes(["chrome.exe"]):
        return "Closing Chrome."
    return "Chrome abhi open nahin hai."


def open_vscode():
    vscode_candidates = [
        shutil.which("code"),
        r"%LocalAppData%\Programs\Microsoft VS Code\Code.exe",
        r"%ProgramFiles%\Microsoft VS Code\Code.exe",
        r"%ProgramFiles(x86)%\Microsoft VS Code\Code.exe",
    ]
    if launch_executable(vscode_candidates):
        return "Opening VS Code."

    installed_message = open_installed_app("visual studio code")
    if installed_message:
        return installed_message

    return "VS Code is not installed on this PC."


def close_vscode():
    if close_processes(["Code.exe"]):
        return "Closing VS Code."
    return "VS Code abhi open nahin hai."


def open_antigravity():
    message = open_installed_app("antigravity")
    if message:
        return message
    return "Antigravity is not installed on this PC."


def close_antigravity():
    message = close_installed_app("antigravity")
    if message:
        return message
    return "Antigravity abhi open nahin hai."


def open_spotify():
    try:
        os.startfile("spotify:")
        return "Opening Spotify."
    except OSError:
        installed_message = open_installed_app("spotify")
        if installed_message:
            return installed_message
        open_url("https://open.spotify.com")
        return "Spotify app nahin mili, maine web version khol diya."


def close_spotify():
    if close_processes(["Spotify.exe"]):
        return "Closing Spotify."
    return "Spotify abhi open nahin hai."


def close_browser():
    if close_processes(["chrome.exe", "msedge.exe", "firefox.exe"]):
        return "Closing the browser."
    return "Mujhe koi open browser nahin mila."


def normalize_command(command):
    return " ".join(command.lower().strip().split())


def matches_command(command, triggers, action_words):
    return any(trigger in command for trigger in triggers) and (
        any(word in command for word in action_words) or command in triggers
    )


def gesture_support_ready():
    return all(
        dependency is not None
        for dependency in (cv2, mp, AudioUtilities, IAudioEndpointVolume, POINTER, cast, CLSCTX_ALL)
    )


def clamp(value, minimum, maximum):
    return max(minimum, min(value, maximum))


def get_volume_controller():
    if not gesture_support_ready():
        return None

    try:
        speakers = AudioUtilities.GetSpeakers()
        endpoint_volume = getattr(speakers, "EndpointVolume", None)
        if endpoint_volume is not None:
            return endpoint_volume

        interface = speakers.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception as exc:
        print(f"[gesture] Audio init failed: {exc}")
        return None


def start_gesture_volume_control():
    if gesture_control_enabled.is_set():
        return "Gesture volume control is already running."

    if not gesture_support_ready():
        return "Gesture volume control ke liye opencv-python, mediapipe, aur pycaw install hone chahiye."

    gesture_control_enabled.set()
    threading.Thread(target=gesture_volume_loop, daemon=True).start()
    return "Gesture volume control on hai. Thumb aur index finger ki distance se volume control hoga."


def stop_gesture_volume_control():
    if not gesture_control_enabled.is_set():
        return "Gesture volume control pehle se band hai."

    gesture_control_enabled.clear()
    return "Gesture volume control band kar diya."


def open_camera_with_fallbacks():
    backend_candidates = []
    if hasattr(cv2, "CAP_DSHOW"):
        backend_candidates.append(cv2.CAP_DSHOW)
    if hasattr(cv2, "CAP_MSMF"):
        backend_candidates.append(cv2.CAP_MSMF)
    backend_candidates.append(None)

    for _ in range(3):
        for backend in backend_candidates:
            camera = cv2.VideoCapture(0, backend) if backend is not None else cv2.VideoCapture(0)
            if camera.isOpened():
                return camera
            camera.release()
        time.sleep(0.5)

    return None


def gesture_volume_loop():
    with gesture_lock:
        camera = None
        try:
            pythoncom.CoInitialize()
            volume_controller = get_volume_controller()
            if volume_controller is None:
                gesture_control_enabled.clear()
                root.after(0, lambda: update_status("Gesture volume audio controller nahin mila."))
                return

            camera = open_camera_with_fallbacks()
            if camera is None:
                gesture_control_enabled.clear()
                root.after(0, lambda: update_status("Gesture volume ke liye webcam access nahin mila."))
                return

            smoothed_level = None
            last_applied_level = None
            last_feedback_time = 0.0

            with mp.solutions.hands.Hands(
                static_image_mode=False,
                max_num_hands=1,
                min_detection_confidence=0.65,
                min_tracking_confidence=0.6,
            ) as hands:
                while assistant_running.is_set() and gesture_control_enabled.is_set():
                    ok, frame = camera.read()
                    if not ok:
                        time.sleep(0.08)
                        continue

                    frame = cv2.flip(frame, 1)
                    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    result = hands.process(rgb_frame)

                    if not result.multi_hand_landmarks:
                        continue

                    landmarks = result.multi_hand_landmarks[0].landmark
                    thumb_tip = landmarks[4]
                    index_tip = landmarks[8]
                    wrist = landmarks[0]
                    middle_base = landmarks[9]

                    pinch_distance = math.dist((thumb_tip.x, thumb_tip.y), (index_tip.x, index_tip.y))
                    palm_size = max(math.dist((wrist.x, wrist.y), (middle_base.x, middle_base.y)), 0.05)
                    normalized_distance = clamp((pinch_distance / palm_size - 0.35) / 0.95, 0.0, 1.0)

                    if smoothed_level is None:
                        smoothed_level = normalized_distance
                    else:
                        smoothed_level = (smoothed_level * 0.72) + (normalized_distance * 0.28)

                    if last_applied_level is None or abs(smoothed_level - last_applied_level) >= 0.03:
                        volume_controller.SetMasterVolumeLevelScalar(float(smoothed_level), None)
                        last_applied_level = smoothed_level

                        now = time.time()
                        if now - last_feedback_time >= 1.0:
                            volume_percent = int(smoothed_level * 100)
                            root.after(
                                0,
                                lambda value=volume_percent: update_status(
                                    f"Gesture volume active: {value}%"
                                ),
                            )
                            last_feedback_time = now
        except Exception as exc:
            print(f"[gesture] {exc}")
            root.after(0, lambda text=str(exc): update_status(f"Gesture error: {text}"))
        finally:
            if camera is not None:
                camera.release()
            gesture_control_enabled.clear()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            if assistant_running.is_set():
                root.after(0, lambda: update_status("Listening..."))


def shutdown():
    if not assistant_running.is_set():
        return

    assistant_running.clear()
    release_single_instance()
    if live_controller:
        live_controller.stop()
    update_status("Shutting down...")
    speak("Goodbye Seenu. Main phir milungi.")
    root.after(1800, root.destroy)


def handle_system_command(command):
    open_command_map = [
        (("chrome", "google chrome"), open_chrome),
        (("vs code", "vscode", "visual studio code"), open_vscode),
        (("antigravity", "anti gravity", "antgigravity", "ant gigravity"), open_antigravity),
        (("whatsapp",), lambda: open_site("WhatsApp", "https://web.whatsapp.com")),
        (("instagram",), lambda: open_site("Instagram", "https://www.instagram.com")),
        (("facebook",), lambda: open_site("Facebook", "https://www.facebook.com")),
        (("github",), lambda: open_site("GitHub", "https://github.com")),
        (("linkedin",), lambda: open_site("LinkedIn", "https://www.linkedin.com")),
        (("youtube",), lambda: open_site("YouTube", "https://www.youtube.com")),
        (("spotify",), open_spotify),
    ]
    close_command_map = [
        (("chrome", "google chrome", "youtube"), close_chrome),
        (("vs code", "vscode", "visual studio code"), close_vscode),
        (("antigravity", "anti gravity", "antgigravity", "ant gigravity"), close_antigravity),
        (("spotify",), close_spotify),
        (("browser", "tab"), close_browser),
    ]

    exit_triggers = ("exit", "quit", "close assistant", "stop assistant", "goodbye mikasa")
    if any(trigger in command for trigger in exit_triggers):
        shutdown()
        return True

    if "gesture" in command and contains_any_phrase(command, ("enable", "start", "open", "on", "chalu")):
        message = start_gesture_volume_control()
        add_message(APP_NAME, message)
        update_status(message)
        speak(message)
        return True

    if "gesture" in command and contains_any_phrase(command, ("disable", "close", "stop", "off", "band")):
        message = stop_gesture_volume_control()
        add_message(APP_NAME, message)
        update_status(message)
        speak(message)
        return True

    for triggers, handler in close_command_map:
        if matches_command(command, triggers, CLOSE_ACTIONS):
            message = handler()
            add_message(APP_NAME, message)
            update_status(message)
            speak(message)
            return True

    close_target = extract_action_target(command, CLOSE_ACTIONS)
    requested_close_targets = split_requested_targets(close_target)
    if requested_close_targets:
        update_status("Closing requested apps...")
        threading.Thread(target=close_installed_app_async, args=(requested_close_targets,), daemon=True).start()
        return True

    generic_target = extract_action_target(command, OPEN_ACTIONS)
    requested_targets = split_requested_targets(generic_target)
    if len(requested_targets) > 1:
        update_status("Opening requested apps...")
        threading.Thread(target=open_installed_app_async, args=(requested_targets,), daemon=True).start()
        return True

    for triggers, handler in open_command_map:
        if matches_command(command, triggers, OPEN_ACTIONS):
            message = handler()
            add_message(APP_NAME, message)
            update_status(message)
            speak(message)
            return True

    if generic_target:
        update_status(f"Opening {generic_target}...")
        threading.Thread(target=open_installed_app_async, args=(generic_target,), daemon=True).start()
        return True

    return False


def deliver_chat_reply(reply):
    add_message(APP_NAME, reply)
    update_status(reply)
    speak(reply)


def get_chat_reply(command):
    if not chat_session:
        root.after(0, lambda text=local_friend_reply(command): deliver_chat_reply(text))
        return

    try:
        with chat_lock:
            response = chat_session.send_message(command)
        reply = (response.text or "").strip() or local_friend_reply(command)
    except Exception as exc:
        print(f"[chat] {exc}")
        reply = local_friend_reply(command)

    root.after(0, lambda text=reply: deliver_chat_reply(text))


def process_command(command):
    normalized = normalize_command(command)
    if not normalized:
        return

    add_message("You", normalized)

    if handle_system_command(normalized):
        return

    update_status("Thinking...")
    threading.Thread(target=get_chat_reply, args=(normalized,), daemon=True).start()


def listen_loop():
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.55

    try:
        microphone = sr.Microphone()
    except OSError:
        root.after(0, lambda: update_status("No microphone detected."))
        speak("Mujhe microphone nahin mil raha.")
        return

    try:
        with microphone as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.4)
    except Exception:
        pass

    while assistant_running.is_set():
        if speaking_event.is_set():
            time.sleep(0.2)
            continue

        try:
            with microphone as source:
                audio = recognizer.listen(source, timeout=1, phrase_time_limit=6)
        except sr.WaitTimeoutError:
            continue
        except Exception as exc:
            root.after(0, lambda text=f"Microphone issue: {exc}": update_status(text))
            time.sleep(1)
            continue

        live_audio = None
        if live_controller and live_controller.is_active():
            try:
                live_audio = audio.get_raw_data(convert_rate=LIVE_AUDIO_INPUT_RATE, convert_width=2)
            except Exception as exc:
                print(f"[live-input] {exc}")

        recognized_text = None
        try:
            recognized_text = recognizer.recognize_google(audio)
        except sr.UnknownValueError:
            recognized_text = None
        except sr.RequestError:
            root.after(0, lambda: update_status("Speech recognition is unavailable right now."))
            if not (live_controller and live_controller.is_active() and live_audio):
                time.sleep(2)
                continue

        if recognized_text:
            normalized = normalize_command(recognized_text)
            if normalized:
                if handle_system_command(normalized):
                    add_message("You", normalized)
                    continue

        if live_controller and live_controller.is_active() and live_audio:
            root.after(0, lambda: update_status("Sending your voice to Gemini Live..."))
            if live_controller.send_audio_turn(live_audio):
                continue

        if recognized_text:
            root.after(0, lambda text=recognized_text: process_command(text))


def start_move(event):
    root.drag_x = event.x_root - root.winfo_x()
    root.drag_y = event.y_root - root.winfo_y()


def do_move(event):
    x_pos = event.x_root - root.drag_x
    y_pos = event.y_root - root.drag_y
    root.geometry(f"+{x_pos}+{y_pos}")


def set_robot_eyes(opened):
    robot_eyes_var.set("◉   ◉" if opened else "─   ─")


def set_robot_mouth(talking):
    if talking:
        robot_anim_state["mouth_phase"] = 1 - robot_anim_state["mouth_phase"]
        robot_mouth_var.set("▅" if robot_anim_state["mouth_phase"] else "▂")
    else:
        robot_anim_state["mouth_phase"] = 0
        robot_mouth_var.set("▁")


def animate_robot_face():
    if not assistant_running.is_set():
        return

    now = time.time()
    robot_emoji_var.set("🤖")

    if speaking_event.is_set():
        robot_anim_state["talk_eyes_open"] = not robot_anim_state["talk_eyes_open"]
        set_robot_eyes(robot_anim_state["talk_eyes_open"])
        set_robot_mouth(True)
        delay_ms = 180
    else:
        robot_anim_state["talk_eyes_open"] = True
        if robot_anim_state["next_idle_blink"] == 0.0:
            robot_anim_state["next_idle_blink"] = now + 3.0

        if now >= robot_anim_state["next_idle_blink"]:
            set_robot_eyes(False)
            robot_anim_state["blink_until"] = now + 0.15
            robot_anim_state["next_idle_blink"] = now + 3.0
        elif robot_anim_state["blink_until"] and now >= robot_anim_state["blink_until"]:
            set_robot_eyes(True)
            robot_anim_state["blink_until"] = 0.0
        else:
            set_robot_eyes(True)

        set_robot_mouth(False)
        delay_ms = 220

    root.after(delay_ms, animate_robot_face)


load_local_env()
hide_console_window()
if not acquire_single_instance():
    focus_existing_window()
    sys.exit(0)
voice_engine = init_voice_engine()
threading.Thread(target=speaker_loop, daemon=True).start()
warm_app_index()

root = tk.Tk()
root.title(f"{APP_NAME} Assistant")
root.overrideredirect(True)
root.attributes("-topmost", True)
root.configure(bg="#08131c")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_pos = screen_width - WINDOW_SIZE - 40
y_pos = screen_height - WINDOW_SIZE - 80
root.geometry(f"{WINDOW_SIZE}x{WINDOW_SIZE}+{x_pos}+{y_pos}")

card = tk.Frame(root, bg="#08131c", highlightbackground="#28d7ff", highlightthickness=2)
card.pack(fill="both", expand=True)

robot_emoji_var = tk.StringVar(value="🤖")
robot_eyes_var = tk.StringVar(value="◉   ◉")
robot_mouth_var = tk.StringVar(value="▁")

robot_emoji_label = tk.Label(
    card,
    textvariable=robot_emoji_var,
    bg="#08131c",
    fg="#f2fdff",
    font=("Segoe UI Emoji", 48),
)
robot_emoji_label.pack(pady=(14, 0))

robot_eyes_label = tk.Label(
    card,
    textvariable=robot_eyes_var,
    bg="#08131c",
    fg="#8ae8ff",
    font=("Consolas", 20, "bold"),
)
robot_eyes_label.pack(pady=(0, 0))

robot_mouth_label = tk.Label(
    card,
    textvariable=robot_mouth_var,
    bg="#08131c",
    fg="#8ae8ff",
    font=("Consolas", 22, "bold"),
)
robot_mouth_label.pack(pady=(0, 6))

name_label = tk.Label(card, text=APP_NAME, bg="#08131c", fg="#eaffff", font=("Segoe UI", 18, "bold"))
name_label.pack()

status_var = tk.StringVar(value="Listening...")
status_label = tk.Label(
    card,
    textvariable=status_var,
    bg="#08131c",
    fg="#9bd9e5",
    wraplength=220,
    justify="center",
    font=("Segoe UI", 9),
)
status_label.pack(padx=16, pady=(8, 6))

hint_label = tk.Label(
    card,
    text="Say 'open VS Code', 'open antigravity', or use thumb-index gesture for volume. Press Esc to close.",
    bg="#08131c",
    fg="#6ba8b5",
    wraplength=220,
    justify="center",
    font=("Segoe UI", 8),
)
hint_label.pack(padx=16, pady=(0, 12))

for widget in (
    card,
    robot_emoji_label,
    robot_eyes_label,
    robot_mouth_label,
    name_label,
    status_label,
    hint_label,
):
    widget.bind("<ButtonPress-1>", start_move)
    widget.bind("<B1-Motion>", do_move)

root.bind("<Escape>", lambda event: shutdown())

chat_session, chat_status = create_chat_session()
live_status = ""
if os.getenv(API_KEY_ENV):
    live_controller, live_status = create_live_controller()
    if live_controller:
        chat_status = live_status
    elif chat_session:
        chat_status = f"{chat_status} {live_status}"
    else:
        chat_status = live_status

threading.Thread(target=listen_loop, daemon=True).start()
root.after(250, animate_robot_face)

initial_message = "Hi Seenu, I am online and ready. Talk to me like a friend."
if live_controller:
    initial_message = "Hi Seenu, Gemini Live voice mode is ready. Talk to me like a friend."
if live_controller:
    update_status(chat_status or "Listening...")
else:
    update_status("Listening..." if chat_session else chat_status)
add_message(APP_NAME, initial_message)
if not live_controller:
    speak(initial_message)

if gesture_support_ready():
    update_status(start_gesture_volume_control())

root.mainloop()
