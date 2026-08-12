from __future__ import annotations

import ctypes

from PIL import ImageGrab
from pynput import keyboard, mouse


def enable_windows_dpi_awareness() -> None:
    """Keep listener coordinates aligned with physical screenshot pixels."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def get_pixel_rgb(x: int, y: int) -> tuple[int, int, int]:
    """Return the RGB color of the physical screen pixel at (x, y)."""
    pixel = ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).convert("RGB")
    red, green, blue = pixel.getpixel((0, 0))
    return int(red), int(green), int(blue)


def on_click(x: int, y: int, button: mouse.Button, pressed: bool) -> None:
    if not pressed:
        return

    pixel_x, pixel_y = int(x), int(y)
    try:
        rgb = get_pixel_rgb(pixel_x, pixel_y)
    except Exception as exc:
        print(f"Click: ({pixel_x}, {pixel_y}) | RGB detection failed: {exc}")
        return

    print(
        f"Position: ({pixel_x}, {pixel_y}) | RGB: {rgb} | "
        f'JSON: {{"position": [{pixel_x}, {pixel_y}], "rgb": {list(rgb)}}}'
    )


def on_key_press(key: keyboard.Key | keyboard.KeyCode) -> bool | None:
    if key == keyboard.Key.esc:
        print("\nPixel detector stopped.")
        return False
    return None


def main() -> None:
    enable_windows_dpi_awareness()
    print("Pixel Color Detector")
    print("Click anywhere to print that pixel's position and RGB color.")
    print("Press Esc to stop.\n")

    with mouse.Listener(on_click=on_click) as mouse_listener:
        with keyboard.Listener(on_press=on_key_press) as keyboard_listener:
            keyboard_listener.join()
        mouse_listener.stop()
        mouse_listener.join()


if __name__ == "__main__":
    main()