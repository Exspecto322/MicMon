# MicMon (Windows “Listen to this device” toggler)

A small Windows-only utility that enables/disables/toggles the **“Listen to this device”**
setting for a recording device (microphone) using **pycaw** + Core Audio’s property store.

Includes helper batch wrappers for Hotkey usage and an optional alternative using NirSoft SoundVolumeView.

## Requirements

- Windows 10/11
- Python 3.9+ (or similar)
- Install dependencies:

```bash
pip install -r requirements.txt
```
## Usage

### 1) List devices (copy/paste the exact names)

```bash
python micmon.py --list-devices
```

### 2) One-time setup (optional): save your device names (no manual file editing required)

This lets you run `--on / --off / --toggle` without typing device names every time.

Create the config file:

```bash
python micmon.py --init-config
```

Set your microphone (required):

```bash
python micmon.py --set-microphone "YOUR MICROPHONE NAME"
```

Set your playback device (optional):

```bash
python micmon.py --set-playback "YOUR SPEAKERS NAME"
```

Or set playback to default:

```bash
python micmon.py --set-default-playback
```

(Optional) Show current config:

```bash
python micmon.py --show-config
```

### 3) Turn “Listen to this device” ON/OFF/Toggle

#### Option A: Using the config (recommended)

```bash
python micmon.py --on
python micmon.py --off
python micmon.py --toggle
```

#### Option B: Passing device names via CLI (no config required)

```bash
python micmon.py --microphone "YOUR MICROPHONE NAME" --on
python micmon.py --microphone "YOUR MICROPHONE NAME" --off
python micmon.py --microphone "YOUR MICROPHONE NAME" --toggle
```

### 4) Optionally set “Playback through this device”

#### With config (use your saved `playback_device`)

```bash
python micmon.py --on
python micmon.py --toggle
```

#### Or via CLI (override for this run)

```bash
python micmon.py --microphone "YOUR MICROPHONE NAME" --on --playback-device "YOUR SPEAKERS NAME"
```

#### Force default playback for this run (ignore config playback)

```bash
python micmon.py --on --default-playback
```

> Tip: CLI arguments override config values.

## Permissions / Admin note (important)

On many systems, writing these Core Audio properties requires administrator privileges.

If you see errors like `Access is denied` or `E_ACCESSDENIED`, run the command from an elevated terminal:

* Start → search **PowerShell**
* Right click → **Run as administrator**

Some IDE terminals can be sandboxed and may not have sufficient privileges even if your user is admin.

---

## Hotkey usage (recommended)

If your launcher (e.g., Elgato Stream Deck) runs elevated, it can execute the Python script with the correct privileges.

Use the batch wrappers in `batch/`:

* `batch/MicMonOnPython.bat`
* `batch/MicMonOffPython.bat`

Edit them once and replace:

* `YOUR_MIC_NAME_HERE`
* optional: `YOUR_SPEAKERS_NAME_HERE`

Then bind those `.bat` files to Stream Deck buttons.

---

## Alternative: SoundVolumeView (batch)

Also included:

* `batch/MicMonOn_SoundVolumeView.bat`
* `batch/MicMonOff_SoundVolumeView.bat`

These require `SoundVolumeView.exe` to be installed by the user.

### Note on device name format (Python vs SoundVolumeView)

Device names may differ between the Python method (pycaw/Core Audio) and NirSoft SoundVolumeView.

- **Python (`micmon.py`)** expects the exact *Windows friendly name* shown in the script output from:
  ```bash
  python micmon.py --list-devices
  ```


Example (common on many systems):

* `Microphone (HyperX QuadCast S)`

* **SoundVolumeView (`.bat`)** uses the device name format shown inside SoundVolumeView itself (or its CLI listing),
  which may omit the `Microphone (...)` prefix.
  Example:

  * `HyperX QuadCast S`

If the SoundVolumeView batch files don’t work with the same string you used for Python, open SoundVolumeView and copy the device name exactly as it appears there, then paste it into the `.bat`.
