# WinWhisper

A lightweight Windows voice-to-text tool that listens while you hold **Ctrl+Space**, transcribes your speech via the OpenAI Whisper API, and types the result wherever your cursor is — in any app.

---

## How It Works

```
Hold Ctrl+Space  →  mic records  →  release  →  Whisper API transcribes  →  text pasted at cursor
```

1. The microphone opens **once at startup** and stays warm (avoids the ~500 ms hardware warmup penalty that causes silent recordings on some USB/webcam mics).
2. While you hold **Ctrl+Space**, raw PCM audio is captured via the Windows `waveIn` API directly (no numpy, no sounddevice — bypasses Application Control policies on managed machines).
3. On release, the audio is saved as a temporary WAV file and sent to OpenAI's `whisper-1` model.
4. The transcribed text is copied to the clipboard and pasted into whatever window had focus when you started speaking, using `SetForegroundWindow` + `AttachThreadInput` to reliably restore focus.

---

## Requirements

- **Windows 10 / 11** (uses Win32 `waveIn` API and `user32.dll`)
- **Python 3.10+** (tested on 3.14)
- **OpenAI account** with Whisper API access and billing credits
  → [platform.openai.com/billing](https://platform.openai.com/billing)

---

## Installation

```bash
# 1. Clone the repo
git clone https://github.com/spatil1985/EAGV3.git
cd EAGV3/project_1

# 2. Install dependencies
pip install openai keyboard pyperclip

# 3. Run
python whisper_app.py
```

On first launch a dialog will prompt for your OpenAI API key.

---

## Configuration

### Option 1 — First-run dialog (recommended)
Run the app and enter your key when prompted. It is saved to:
```
%USERPROFILE%\.winwhisper_config.json
```

### Option 2 — Environment variable
```powershell
$env:OPENAI_API_KEY = "sk-..."
python whisper_app.py
```

### Option 3 — Config file
Create `%USERPROFILE%\.winwhisper_config.json`:
```json
{
  "api_key": "sk-..."
}
```

> **Never commit your API key.** The config file lives outside the project directory by design. `config.json` and `.env` files are blocked by `.gitignore`.

### Changing the hotkey or language

Open `whisper_app.py` and edit the top of the file:

| Setting | Where | Default |
|---|---|---|
| Transcription language | `language=` in `_transcribe()` | `"en"` |
| Audio sample rate | `SAMPLE_RATE` constant | `16000` |
| Buffer size | `_BUF_BYTES` / `_N_BUFS` | 50 ms × 8 |

---

## Usage

| Action | Result |
|---|---|
| Hold **Ctrl+Space** | Start recording (red overlay appears bottom-right) |
| Release **Ctrl+Space** | Stop → transcribe → paste into active window |
| **Ctrl+Shift+Q** | Quit |

---

## Architecture

```
whisper_app.py
│
├── WaveInRecorder          Windows waveIn API via ctypes (winmm.dll)
│   ├── open_device()       Opens mic once at startup, starts drain loop
│   ├── begin_capture()     Starts saving audio to buffer
│   └── end_capture()       Stops saving, writes WAV file
│
├── WinWhisper              Main application
│   ├── _start()            Saves target window HWND, begins capture
│   ├── _stop()             Ends capture, triggers transcription
│   ├── _transcribe()       Sends WAV to OpenAI Whisper API
│   └── _paste_text()       Restores focus via AttachThreadInput, pastes
│
└── StatusOverlay           Always-on-top Tkinter status widget
```

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `waveInOpen failed` | No microphone found. Check Windows sound settings. |
| `Nothing recorded` | Hold the key for at least 0.5 s. Some USB mics need warmup. |
| `429 insufficient_quota` | Add billing credits at platform.openai.com/billing |
| Text pastes into wrong window | Click into your target window **before** pressing Ctrl+Space |
| Keyboard stops working | Kill all `python.exe` processes in Task Manager |

---

## Dependencies

| Package | Purpose |
|---|---|
| `openai` | Whisper API client |
| `keyboard` | Global hotkey detection (Ctrl+Space) |
| `pyperclip` | Clipboard read/write for paste |

Audio recording uses only Windows built-in `winmm.dll` via `ctypes` — no additional audio packages required.

---

## License

MIT
