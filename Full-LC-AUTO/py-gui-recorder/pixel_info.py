from pynput import mouse, keyboard
from PIL import ImageGrab

def get_pixel_color(x, y):
    try:
        # Grabs the RGB color at the coordinate
        return ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).getpixel((0, 0))
    except Exception:
        return "N/A"

def on_click(x, y, button, pressed):
    if pressed:
        color = get_pixel_color(x, y)
        print(f"Clicked at: ({x}, {y}) | RGB: {color}")

def on_press(key):
    if key == keyboard.Key.esc:
        print("\nStopping Logger...")
        return False

print("--- CLICK LOGGER ---")
print("1. Click anywhere to print coordinates and color.")
print("2. Press 'ESC' to exit.")
print("-" * 40)

# Listen for clicks and keyboard (to stop)
with mouse.Listener(on_click=on_click) as m_listener, \
     keyboard.Listener(on_press=on_press) as k_listener:
    k_listener.join()