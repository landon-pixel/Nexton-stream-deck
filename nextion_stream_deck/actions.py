from __future__ import annotations

import ctypes
import os
from pathlib import Path
import re
import subprocess
import time
import webbrowser


user32 = ctypes.WinDLL("user32", use_last_error=True)

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

VK_CODES = {
    "alt": 0x12,
    "backspace": 0x08,
    "ctrl": 0x11,
    "delete": 0x2E,
    "down": 0x28,
    "enter": 0x0D,
    "esc": 0x1B,
    "f1": 0x70,
    "f2": 0x71,
    "f3": 0x72,
    "f4": 0x73,
    "f5": 0x74,
    "f6": 0x75,
    "f7": 0x76,
    "f8": 0x77,
    "f9": 0x78,
    "f10": 0x79,
    "f11": 0x7A,
    "f12": 0x7B,
    "left": 0x25,
    "right": 0x27,
    "shift": 0x10,
    "space": 0x20,
    "tab": 0x09,
    "up": 0x26,
    "win": 0x5B,
    "media_next": 0xB0,
    "media_prev": 0xB1,
    "media_stop": 0xB2,
    "media_play_pause": 0xB3,
    "next_track": 0xB0,
    "previous_track": 0xB1,
    "play_pause": 0xB3,
}

for code in range(ord("A"), ord("Z") + 1):
    VK_CODES[chr(code).lower()] = code

for number in range(10):
    VK_CODES[str(number)] = ord(str(number))


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class INPUT(ctypes.Structure):
    class _INPUTUNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.c_ulong), ("u", _INPUTUNION)]


def run_action(action_type: str, payload: str) -> str:
    action_type = action_type.strip().lower()
    payload = payload.strip()
    if not payload:
        raise ValueError("Payload is empty.")

    if action_type == "launch":
        target = os.path.expandvars(payload)
        if _looks_like_uri(target):
            os.startfile(target)
            return f"Launched {target}"
        if ".exe" in target.lower() and " " in target.strip():
            subprocess.Popen(target, shell=True)
            return f"Started {target}"
        path_target = Path(target).expanduser()
        os.startfile(str(path_target))
        return f"Launched {path_target}"

    if action_type == "url":
        webbrowser.open(payload)
        return f"Opened {payload}"

    if action_type == "command":
        subprocess.Popen(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", payload],
            cwd=str(Path.cwd()),
        )
        return f"Started command: {payload}"

    if action_type == "hotkey":
        send_hotkey(payload)
        return f"Sent hotkey: {payload}"

    raise ValueError(f"Unsupported action type: {action_type}")


def run_mapping(action_type: str, payload: str, shortcut_keys: str = "") -> str:
    result = run_action(action_type, payload)
    shortcut_keys = shortcut_keys.strip()
    if shortcut_keys:
        if action_type.strip().lower() == "launch":
            time.sleep(0.35)
        send_hotkey(shortcut_keys)
        return f"{result} + shortcut {shortcut_keys}"
    return result


def send_hotkey(combo: str) -> None:
    keys = [part.strip().lower() for part in combo.split("+") if part.strip()]
    if not keys:
        raise ValueError("Hotkey is empty.")
    virtual_keys = []
    for key in keys:
        if key not in VK_CODES:
            raise ValueError(f"Unknown hotkey key: {key}")
        virtual_keys.append(VK_CODES[key])
    for vk in virtual_keys:
        _key_event(vk, 0)
    for vk in reversed(virtual_keys):
        _key_event(vk, KEYEVENTF_KEYUP)


def _key_event(vk_code: int, flags: int) -> None:
    extra = ctypes.c_ulong(0)
    keyboard_input = KEYBDINPUT(
        wVk=vk_code,
        wScan=0,
        dwFlags=flags,
        time=0,
        dwExtraInfo=ctypes.pointer(extra),
    )
    input_record = INPUT(type=INPUT_KEYBOARD, ki=keyboard_input)
    sent = user32.SendInput(1, ctypes.byref(input_record), ctypes.sizeof(INPUT))
    if sent != 1:
        raise ctypes.WinError(ctypes.get_last_error())


def _looks_like_uri(value: str) -> bool:
    return bool(re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:(//)?", value))
