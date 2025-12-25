#!/usr/bin/env python
"""
micmon.py - Toggle Windows "Listen to this device" for a given microphone.

Uses pycaw + Windows Core Audio property store to programmatically control
the "Listen to this device" checkbox and the "Playback through this device"
dropdown for a recording device.

Keeps the original CLI behavior:
- --microphone / --playback-device / --on|--off|--toggle
- --list-devices

Adds optional config support (set once, then just run --on/--off/--toggle):
- --init-config
- --show-config
- --set-microphone
- --set-playback
- --set-default-playback
- --config (path)

Windows 10/11 only.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional

import comtypes
from comtypes import GUID
from comtypes.automation import VT_BOOL, VT_LPWSTR, VT_EMPTY
from comtypes.persist import STGM_READWRITE

from pycaw.api.mmdeviceapi import PROPERTYKEY
from pycaw.api.mmdeviceapi.depend import PROPVARIANT
from pycaw.constants import CLSID_MMDeviceEnumerator
from pycaw.pycaw import (
    AudioUtilities as CoreAudioUtilities,
    IMMDeviceEnumerator,
    EDataFlow,
    DEVICE_STATE,
)
from pycaw.utils import AudioUtilities as UtilsAudioUtilities

# Hard-coded Core Audio property IDs used by the "Listen" tab
LISTEN_SETTING_GUID = "{24DBB0FC-9311-4B3D-9CF0-18FF155639D4}"
CHECKBOX_PID = 1           # "Listen to this device" checkbox
LISTENING_DEVICE_PID = 0   # "Playback through this device" dropdown

DEFAULT_CONFIG_PATH = Path(__file__).with_name("micmon.config.json")


@dataclass
class DeviceInfo:
    id: str
    friendly_name: str
    direction: str  # "input" or "output"


# -----------------------------
# Config helpers
# -----------------------------

def load_config(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(path: Path, cfg: dict) -> None:
    path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def write_config_template(path: Path) -> None:
    template = {
        "microphone": "YOUR MICROPHONE NAME",
        "playback_device": None
    }
    save_config(path, template)


# -----------------------------
# Device enumeration
# -----------------------------

def get_active_devices(direction: str) -> List[DeviceInfo]:
    if direction not in {"input", "output"}:
        raise ValueError("direction must be 'input' or 'output'")

    e_dataflow = EDataFlow.eRender.value if direction == "output" else EDataFlow.eCapture.value

    device_enumerator = comtypes.CoCreateInstance(
        CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER,
    )
    if device_enumerator is None:
        raise RuntimeError("Couldn't create IMMDeviceEnumerator")

    collection = device_enumerator.EnumAudioEndpoints(e_dataflow, DEVICE_STATE.ACTIVE.value)
    if collection is None:
        raise RuntimeError("Couldn't enumerate audio endpoints")

    devices: List[DeviceInfo] = []
    count = collection.GetCount()
    for i in range(count):
        dev = collection.Item(i)
        if dev is None:
            continue

        device = CoreAudioUtilities.CreateDevice(dev)
        if ": None" in str(device):  # Filter out ghost "None" devices
            continue

        devices.append(DeviceInfo(id=device.id, friendly_name=device.FriendlyName, direction=direction))

    return devices


def validate_device_name(name: str, direction: str) -> None:
    """
    Validate that a name exists among active devices of the given direction.
    direction: "input" or "output"
    """
    names = {d.friendly_name for d in get_active_devices(direction)}
    if name not in names:
        raise ValueError(
            f"{direction.capitalize()} device not found: {name!r}. "
            f"Run: python micmon.py --list-devices"
        )


def find_device_guid_by_name(name: str) -> str:
    for dev in get_active_devices("input") + get_active_devices("output"):
        if dev.friendly_name == name:
            return dev.id
    raise ValueError(f"Audio device not found: {name!r}")


# -----------------------------
# Core property store logic
# -----------------------------

def open_property_store_for_device(device_name: str):
    device_guid = find_device_guid_by_name(device_name)
    enumerator = UtilsAudioUtilities.GetDeviceEnumerator()
    device = enumerator.GetDevice(device_guid)
    if device is None:
        raise RuntimeError(f"Failed to open device for GUID {device_guid}")
    return device.OpenPropertyStore(STGM_READWRITE)


def _listen_checkbox_property_key() -> PROPERTYKEY:
    pk = PROPERTYKEY()
    pk.fmtid = GUID(LISTEN_SETTING_GUID)
    pk.pid = CHECKBOX_PID
    return pk


def _listen_playback_property_key() -> PROPERTYKEY:
    pk = PROPERTYKEY()
    pk.fmtid = GUID(LISTEN_SETTING_GUID)
    pk.pid = LISTENING_DEVICE_PID
    return pk


def get_listen_enabled(property_store) -> Optional[bool]:
    try:
        current = property_store.GetValue(_listen_checkbox_property_key())
        return bool(current.union.boolVal)
    except Exception:
        return None


def set_listen_enabled(property_store, enabled: bool) -> None:
    new_value = PROPVARIANT(VT_BOOL)
    new_value.union.boolVal = bool(enabled)
    property_store.SetValue(_listen_checkbox_property_key(), new_value)


def set_listen_playback_device(property_store, playback_device_name: Optional[str]) -> None:
    if playback_device_name is not None:
        listening_device_guid = find_device_guid_by_name(playback_device_name)
        new_value = PROPVARIANT(VT_LPWSTR)
        new_value.union.pwszVal = listening_device_guid
    else:
        new_value = PROPVARIANT(VT_EMPTY)

    property_store.SetValue(_listen_playback_property_key(), new_value)


def apply_listen_settings(
    microphone_name: str,
    enabled: Optional[bool],
    playback_device_name: Optional[str],
    verbose: bool = True,
) -> None:
    store = open_property_store_for_device(microphone_name)
    current = get_listen_enabled(store)

    if enabled is None:
        enabled_to_set = (not bool(current)) if current is not None else True
    else:
        enabled_to_set = enabled

    set_listen_enabled(store, enabled_to_set)
    set_listen_playback_device(store, playback_device_name)

    try:
        store.Commit()
    except Exception:
        pass

    if verbose:
        state = "ON" if enabled_to_set else "OFF"
        target = playback_device_name or "Default playback device"
        print(f"[micmon] Listen to this device: {state}")
        print(f"[micmon] Playback through this device: {target}")


def print_devices() -> None:
    inputs = get_active_devices("input")
    outputs = get_active_devices("output")

    print("Input devices (recording):")
    if not inputs:
        print("  (none)")
    for dev in inputs:
        print(f"  - {dev.friendly_name}")

    print("\nOutput devices (playback):")
    if not outputs:
        print("  (none)")
    for dev in outputs:
        print(f"  - {dev.friendly_name}")


# -----------------------------
# CLI
# -----------------------------

def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Toggle Windows 'Listen to this device' for a microphone using pycaw.\n\n"
            "Examples:\n"
            "  micmon.py --list-devices\n"
            "  micmon.py --init-config\n"
            "  micmon.py --set-microphone \"...\"\n"
            "  micmon.py --set-playback \"...\"\n"
            "  micmon.py --on   (uses config if present)\n"
            "  micmon.py --microphone \"...\" --on   (CLI overrides config)\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # Runtime device args (original behavior)
    parser.add_argument("--microphone", "-m", required=False,
                        help="Recording device friendly name. If omitted, uses config (if present).")
    parser.add_argument("--playback-device", "-p", default=None,
                        help="Playback device friendly name for this run. If omitted, uses config (if present).")

    # Runtime override to force default playback (ignore config playback)
    parser.add_argument("--default-playback", action="store_true",
                        help="For this run, force 'Default playback device' (ignore config playback_device).")

    # Main action flags
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--on", action="store_true", help="Enable 'Listen to this device'.")
    group.add_argument("--off", action="store_true", help="Disable 'Listen to this device'.")
    group.add_argument("--toggle", action="store_true", help="Toggle the current state.")

    # Utility flags
    parser.add_argument("--list-devices", action="store_true",
                        help="List active input/output devices and exit.")

    # Config management
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH),
                        help="Path to config JSON file (default: micmon.config.json next to micmon.py).")
    parser.add_argument("--init-config", action="store_true",
                        help="Create a starter config file if missing, then exit.")
    parser.add_argument("--show-config", action="store_true",
                        help="Print the current config (if present), then exit.")
    parser.add_argument("--set-microphone", metavar="NAME", default=None,
                        help="Write 'microphone' to config (input device name), then exit.")
    parser.add_argument("--set-playback", metavar="NAME", default=None,
                        help="Write 'playback_device' to config (output device name), then exit.")
    parser.add_argument("--set-default-playback", action="store_true",
                        help="Set config playback_device to null (default playback), then exit.")

    parser.add_argument("--quiet", "-q", action="store_true",
                        help="Suppress informational output (errors are still shown).")

    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)
    verbose = not args.quiet

    config_path = Path(args.config)
    cfg = load_config(config_path)

    # --- Config management mode (CLI-only setup, no file editing required) ---
    if args.init_config or args.show_config or args.set_microphone or args.set_playback or args.set_default_playback:
        try:
            if args.init_config:
                if config_path.exists():
                    if verbose:
                        print(f"[micmon] config already exists: {config_path}")
                else:
                    write_config_template(config_path)
                    if verbose:
                        print(f"[micmon] wrote config template: {config_path}")
                return 0

            if args.show_config:
                if not config_path.exists():
                    if verbose:
                        print(f"[micmon] config not found: {config_path}")
                    return 1
                if verbose:
                    print(f"[micmon] config path: {config_path}")
                print(json.dumps(cfg, indent=2))
                return 0

            # Ensure we have a dict and create file if missing when setting values
            if not isinstance(cfg, dict):
                cfg = {}
            if not config_path.exists():
                # create empty base so we can write updates
                save_config(config_path, cfg)

            changed = False

            if args.set_microphone:
                validate_device_name(args.set_microphone, "input")
                cfg["microphone"] = args.set_microphone
                changed = True

            if args.set_playback:
                validate_device_name(args.set_playback, "output")
                cfg["playback_device"] = args.set_playback
                changed = True

            if args.set_default_playback:
                cfg["playback_device"] = None
                changed = True

            if changed:
                save_config(config_path, cfg)
                if verbose:
                    print(f"[micmon] updated config: {config_path}")
                    print(json.dumps(cfg, indent=2))
                return 0

            return 0

        except ValueError as exc:
            if verbose:
                print(f"[micmon] error: {exc}")
            return 1
        except Exception as exc:
            if verbose:
                print(f"[micmon] unexpected error: {exc}")
            return 2

    # --- Device listing ---
    if args.list_devices:
        try:
            print_devices()
            return 0
        except Exception as exc:
            if verbose:
                print(f"[micmon] error listing devices: {exc}")
            return 1

    # --- Toggle mode (original behavior preserved) ---
    if not (args.on or args.off or args.toggle):
        if verbose:
            print("[micmon] error: one of --on/--off/--toggle is required (or use a config command like --set-microphone).")
        return 1

    microphone = args.microphone or cfg.get("microphone")
    if not microphone:
        if verbose:
            print("[micmon] error: microphone not provided. Use --microphone or set it via --set-microphone / config.")
        return 1

    # playback resolution order:
    # 1) --default-playback (force None)
    # 2) --playback-device "NAME" (runtime override)
    # 3) config playback_device (if any)
    # 4) None (default playback)
    if args.default_playback:
        playback_device = None
    elif args.playback_device is not None:
        playback_device = args.playback_device
    else:
        playback_device = cfg.get("playback_device")

    try:
        enabled = True if args.on else False if args.off else None
        apply_listen_settings(
            microphone_name=microphone,
            enabled=enabled,
            playback_device_name=playback_device,
            verbose=verbose,
        )
        return 0
    except ValueError as exc:
        if verbose:
            print(f"[micmon] error: {exc}")
        return 1
    except Exception as exc:
        if verbose:
            print(f"[micmon] unexpected error: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())