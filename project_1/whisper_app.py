#!/usr/bin/env python3
"""
WinWhisper - Voice to text using OpenAI Whisper API
Press Ctrl+Space to start recording, press again to stop and type the text.
Uses waveIn* API directly — no external DLLs, no codec loading.
"""

import sys
import os
import json
import time
import wave
import tempfile
import threading
import ctypes
import ctypes.wintypes as wt
import tkinter as tk
from tkinter import simpledialog
from pathlib import Path

import keyboard
import pyperclip
from openai import OpenAI


CONFIG_FILE = Path.home() / ".winwhisper_config.json"

# ── Windows API constants & helpers ──────────────────────────────────────────

user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32


# ── waveIn recorder (pure winmm.dll, no codecs) ───────────────────────────────

SAMPLE_RATE     = 16000
CHANNELS        = 1
BITS            = 16
WAVE_FORMAT_PCM = 1
WAVE_MAPPER     = ctypes.c_uint(-1).value   # 0xFFFFFFFF
CALLBACK_NULL   = 0
WHDR_DONE       = 0x00000001
WHDR_PREPARED   = 0x00000002
WHDR_INQUEUE    = 0x00000010

winmm = ctypes.windll.winmm


class WAVEFORMATEX(ctypes.Structure):
    _fields_ = [
        ("wFormatTag",      wt.WORD),
        ("nChannels",       wt.WORD),
        ("nSamplesPerSec",  wt.DWORD),
        ("nAvgBytesPerSec", wt.DWORD),
        ("nBlockAlign",     wt.WORD),
        ("wBitsPerSample",  wt.WORD),
        ("cbSize",          wt.WORD),
    ]


class WAVEHDR(ctypes.Structure):
    _fields_ = [
        ("lpData",          ctypes.c_void_p),
        ("dwBufferLength",  wt.DWORD),
        ("dwBytesRecorded", wt.DWORD),
        ("dwUser",          ctypes.c_size_t),   # DWORD_PTR — pointer-sized
        ("dwFlags",         wt.DWORD),
        ("dwLoops",         wt.DWORD),
        ("lpNext",          ctypes.c_void_p),
        ("reserved",        ctypes.c_size_t),   # DWORD_PTR — pointer-sized
    ]


# Small buffers for low-latency drain loop; device stays open permanently
_BUF_BYTES = SAMPLE_RATE * CHANNELS * (BITS // 8) // 20  # 50 ms chunks = 1600 bytes
_N_BUFS    = 8                                             # 8 rotating buffers


class WaveInRecorder:
    """
    Device is opened ONCE at startup and runs continuously.
    This avoids the ~500 ms hardware warmup penalty on every recording.
    begin_capture() / end_capture() mark which audio to save.
    """

    def __init__(self):
        self._hwi        = None
        self._bufs: list = []
        self._hdrs: list = []
        self._active     = False   # device is running
        self._capturing  = False   # user is holding the hotkey
        self._capture_data: bytearray = bytearray()
        self._lock       = threading.Lock()
        self._drain_thr: threading.Thread | None = None

    # ── device lifetime ───────────────────────────────────────────────────

    def open_device(self) -> bool:
        """Open the mic once at startup and keep it warm."""
        fmt = WAVEFORMATEX()
        fmt.wFormatTag      = WAVE_FORMAT_PCM
        fmt.nChannels       = CHANNELS
        fmt.nSamplesPerSec  = SAMPLE_RATE
        fmt.wBitsPerSample  = BITS
        fmt.nBlockAlign     = CHANNELS * BITS // 8
        fmt.nAvgBytesPerSec = SAMPLE_RATE * fmt.nBlockAlign
        fmt.cbSize          = 0

        hwi = wt.HANDLE()
        r = winmm.waveInOpen(ctypes.byref(hwi), WAVE_MAPPER, ctypes.byref(fmt),
                              0, 0, CALLBACK_NULL)
        if r != 0:
            print(f"waveInOpen failed: {r}")
            return False

        self._hwi  = hwi
        self._bufs = []
        self._hdrs = []

        for _ in range(_N_BUFS):
            buf = ctypes.create_string_buffer(_BUF_BYTES)
            hdr = WAVEHDR()
            hdr.lpData         = ctypes.addressof(buf)
            hdr.dwBufferLength = _BUF_BYTES
            hdr.dwFlags        = 0
            self._bufs.append(buf)
            self._hdrs.append(hdr)
            winmm.waveInPrepareHeader(hwi, ctypes.byref(hdr), ctypes.sizeof(WAVEHDR))
            winmm.waveInAddBuffer    (hwi, ctypes.byref(hdr), ctypes.sizeof(WAVEHDR))

        self._active = True
        winmm.waveInStart(hwi)

        self._drain_thr = threading.Thread(target=self._drain_loop, daemon=True)
        self._drain_thr.start()
        return True

    # ── capture control ───────────────────────────────────────────────────

    @property
    def is_recording(self) -> bool:
        return self._capturing

    def begin_capture(self):
        with self._lock:
            self._capture_data = bytearray()
            self._capturing    = True

    def end_capture(self, wav_path: str) -> bool:
        """Stop capturing, write WAV. Returns True if audio was recorded."""
        with self._lock:
            self._capturing = False
            data = bytes(self._capture_data)

        if not data:
            return False

        with wave.open(wav_path, "wb") as wf:
            wf.setnchannels(CHANNELS)
            wf.setsampwidth(BITS // 8)
            wf.setframerate(SAMPLE_RATE)
            wf.writeframes(data)
        return True

    # ── internal drain loop ───────────────────────────────────────────────

    def _drain_loop(self):
        """Continuously drain filled buffers and re-queue them."""
        while self._active:
            for i, hdr in enumerate(self._hdrs):
                if hdr.dwFlags & WHDR_DONE:
                    if self._capturing and hdr.dwBytesRecorded > 0:
                        with self._lock:
                            self._capture_data.extend(
                                ctypes.string_at(ctypes.addressof(self._bufs[i]),
                                                 hdr.dwBytesRecorded)
                            )
                    # Re-queue: keep PREPARED, clear DONE/INQUEUE
                    hdr.dwFlags         = WHDR_PREPARED
                    hdr.dwBytesRecorded = 0
                    winmm.waveInAddBuffer(self._hwi, ctypes.byref(hdr),
                                          ctypes.sizeof(WAVEHDR))
            time.sleep(0.01)  # 10 ms poll


# ── Config ────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(config: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)


def prompt_api_key() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    key = simpledialog.askstring(
        "OpenAI API Key",
        "Enter your OpenAI API key:\n(saved to ~/.winwhisper_config.json)",
        parent=root,
    )
    root.destroy()
    return key or ""


def get_api_key() -> str:
    config = load_config()
    if config.get("api_key"):
        return config["api_key"]
    env_key = os.environ.get("OPENAI_API_KEY", "")
    if env_key:
        return env_key
    key = prompt_api_key()
    if not key:
        print("No API key provided. Exiting.")
        sys.exit(1)
    config["api_key"] = key
    save_config(config)
    return key


# ── Status overlay ────────────────────────────────────────────────────────────

class StatusOverlay:
    def __init__(self):
        self._root = None

    def _run(self, text: str, color: str):
        try:
            root = tk.Tk()
            self._root = root
            root.overrideredirect(True)
            root.attributes("-topmost", True)
            root.attributes("-alpha", 0.90)
            root.configure(bg="#1e1e1e")
            tk.Label(
                root, text=text, font=("Segoe UI", 13, "bold"),
                fg=color, bg="#1e1e1e", padx=18, pady=10,
            ).pack()
            root.update_idletasks()
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            w  = root.winfo_width()
            h  = root.winfo_height()
            root.geometry(f"+{sw - w - 30}+{sh - h - 60}")
            root.mainloop()
        except Exception:
            pass

    def show(self, text: str, color: str = "#ff4444"):
        self.hide()
        threading.Thread(target=self._run, args=(text, color), daemon=True).start()

    def hide(self):
        r = self._root
        if r:
            self._root = None
            try:
                r.quit()
                r.destroy()
            except Exception:
                pass


# ── Main app ──────────────────────────────────────────────────────────────────

class WinWhisper:
    def __init__(self):
        self._lock    = threading.Lock()
        self.recorder = WaveInRecorder()
        self.overlay  = StatusOverlay()
        api_key       = get_api_key()
        self.client   = OpenAI(api_key=api_key)

        print("WinWhisper: opening mic…", end=" ", flush=True)
        if not self.recorder.open_device():
            print("FAILED — no audio device found.")
            sys.exit(1)
        print("ready.")
        print( "  Hotkey : Hold CTRL+SPACE to record, release to transcribe")
        print( "  Quit   : Ctrl+Shift+Q\n")

    # ── Recording ─────────────────────────────────────────────────────────

    def _start(self):
        with self._lock:
            if self.recorder.is_recording:
                return
            # Remember which window had focus so we can paste there
            self._target_hwnd = user32.GetForegroundWindow()
            self.recorder.begin_capture()
        self.overlay.show("  Recording...", "#ff4444")
        print("[REC] Recording...")

    def _stop(self):
        with self._lock:
            if not self.recorder.is_recording:
                return
            wav_path = tempfile.mktemp(suffix=".wav")
            ok = self.recorder.end_capture(wav_path)

        if not ok:
            print("Nothing recorded.")
            self.overlay.hide()
            return

        self.overlay.show("  Transcribing...", "#f0a500")
        print("[...] Transcribing...")
        threading.Thread(target=self._transcribe, args=(wav_path,), daemon=True).start()

    # ── Transcription ─────────────────────────────────────────────────────

    def _transcribe(self, wav_path: str):
        try:
            with open(wav_path, "rb") as f:
                result = self.client.audio.transcriptions.create(
                    model="whisper-1",
                    file=f,
                    response_format="text",
                    language="en",
                )
            text = (result if isinstance(result, str) else result.text).strip()
            if text:
                print(f"[OK] {text}\n")
                self.overlay.hide()
                self._paste_text(text)
            else:
                print("(no speech detected)\n")
                self.overlay.hide()
        except Exception as exc:
            print(f"[ERR] {exc}\n")
            self.overlay.hide()
        finally:
            try:
                os.unlink(wav_path)
            except Exception:
                pass

    # ── Paste ─────────────────────────────────────────────────────────────

    def _paste_text(self, text: str):
        hwnd = getattr(self, "_target_hwnd", None)
        if hwnd:
            # AttachThreadInput trick — the only reliable way to force
            # SetForegroundWindow when our process isn't the foreground process
            cur_tid = kernel32.GetCurrentThreadId()
            fg_hwnd = user32.GetForegroundWindow()
            fg_tid  = user32.GetWindowThreadProcessId(fg_hwnd, None)
            if fg_tid and fg_tid != cur_tid:
                user32.AttachThreadInput(cur_tid, fg_tid, True)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
            if fg_tid and fg_tid != cur_tid:
                user32.AttachThreadInput(cur_tid, fg_tid, False)
            time.sleep(0.15)

        try:
            old = pyperclip.paste()
        except Exception:
            old = ""
        pyperclip.copy(text)
        keyboard.send("ctrl+v")
        time.sleep(0.2)
        try:
            pyperclip.copy(old)
        except Exception:
            pass

    # ── Hotkey ────────────────────────────────────────────────────────────

    def run(self):
        def on_space_down(e):
            if keyboard.is_pressed("ctrl"):
                threading.Thread(target=self._start, daemon=True).start()

        def on_space_up(e):
            if self.recorder.is_recording:
                threading.Thread(target=self._stop, daemon=True).start()

        # suppress=False so plain Space is NEVER blocked
        keyboard.on_press_key("space",   on_space_down, suppress=False)
        keyboard.on_release_key("space", on_space_up)
        keyboard.add_hotkey("ctrl+shift+q", lambda: os.kill(os.getpid(), 9))
        keyboard.wait()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = WinWhisper()
    app.run()
