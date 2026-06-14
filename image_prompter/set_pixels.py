"""
set_pixels.py  —  Interactive coordinate setup for image_main.py

Run this once per machine (or whenever your screen layout changes).
It walks you through clicking three positions on screen, then saves
them to pixels.json next to this file.  image_main.py reads that
file at startup instead of using hardcoded coordinates.

Usage:
    python set_pixels.py
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import pyautogui

PIXELS_FILE = Path(__file__).resolve().parent / "pixels.json"

STEPS = [
    {
        "key": "chat_neutral_click",
        "label": "Neutral chat area (for Ctrl+A copy)",
        "instructions": (
            "Click somewhere in the MIDDLE OF THE CHAT MESSAGES area "
            "— NOT on the input box, NOT on a button.\n"
            "This is the spot the script clicks before doing Ctrl+A "
            "to make sure the input box loses focus."
        ),
    },
    {
        "key": "prompt_box_initial",
        "label": "Prompt input box — INITIAL (boot focus)",
        "instructions": (
            "Click INSIDE the ChatGPT prompt/input box.\n"
            "This is used at startup to focus the box before pasting "
            "the first prompt."
        ),
    },
    {
        "key": "prompt_box_post_injection",
        "label": "Prompt input box — POST INJECTION (stuck recovery)",
        "instructions": (
            "Click INSIDE the ChatGPT prompt/input box again.\n"
            "This can be the same spot or a slightly different position "
            "used when the script detects the prompt is stuck and needs "
            "to re-press Enter."
        ),
    },
]

COUNTDOWN_SECONDS = 3


def countdown(label: str) -> None:
    print(f"\n  Move your mouse to: {label}")
    for i in range(COUNTDOWN_SECONDS, 0, -1):
        print(f"  Capturing in {i}...", end="\r", flush=True)
        time.sleep(1)
    print()


def capture_position(step: dict[str, str]) -> tuple[int, int]:
    print()
    print(f"{'─' * 60}")
    print(f"  STEP: {step['label']}")
    print(f"{'─' * 60}")
    print(f"  {step['instructions']}")
    print()
    input("  Press ENTER when you are ready to position your mouse, then DON'T MOVE IT.")
    countdown(step["label"])
    x, y = pyautogui.position()
    print(f"  ✓  Captured: ({x}, {y})")
    return (x, y)


def verify_and_confirm(results: dict[str, tuple[int, int]]) -> bool:
    print()
    print(f"{'═' * 60}")
    print("  CAPTURED COORDINATES SUMMARY")
    print(f"{'═' * 60}")
    for step in STEPS:
        key = step["key"]
        x, y = results[key]
        print(f"  {step['label']}")
        print(f"      → ({x}, {y})")
        print()
    answer = input("  Save these coordinates? [Y/n]: ").strip().lower()
    return answer in ("", "y", "yes")


def save_pixels(results: dict[str, tuple[int, int]]) -> None:
    payload = {key: {"x": x, "y": y} for key, (x, y) in results.items()}
    PIXELS_FILE.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"\n  ✓  Saved to: {PIXELS_FILE}")


def main() -> None:
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║          image_main.py  —  Coordinate Setup             ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()
    print("  This tool captures the screen positions that image_main.py")
    print("  needs.  Open Firefox with ChatGPT BEFORE continuing so you")
    print("  can click the correct spots.")
    print()
    print("  You will be asked to position your mouse over 3 targets.")
    print(f"  Each capture has a {COUNTDOWN_SECONDS}-second countdown after you press ENTER.")
    print()
    input("  Press ENTER to begin  ▶")

    results: dict[str, tuple[int, int]] = {}
    for step in STEPS:
        results[step["key"]] = capture_position(step)

    if verify_and_confirm(results):
        save_pixels(results)
        print()
        print("  All done! Run image_main.py normally now.")
    else:
        print()
        print("  Nothing saved. Run set_pixels.py again to redo.")


if __name__ == "__main__":
    main()
