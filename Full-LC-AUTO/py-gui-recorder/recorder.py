import time
import json
import os
from pynput import mouse, keyboard
from PIL import ImageGrab

recording = []
last_event_time = None
shift_pressed = False 

def get_delay():
    global last_event_time
    current_time = time.time()
    delay = current_time - last_event_time
    last_event_time = current_time
    return round(delay, 3)

def get_pixel_color(x, y):  
    try:
        return ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).getpixel((0, 0))
    except:
        return (0, 0, 0)

def on_click(x, y, button, pressed):
    if pressed and last_event_time:
        delay = get_delay()
        
        # Shift + Click = pixel_watcher
        if shift_pressed:
            color = get_pixel_color(x, y)
            event_type = "pixel_watcher"
            print(f"[PIXEL WATCHER] Target: ({x}, {y}) | Color: {color} | Recorded Delay: {delay}s")
        else:
            color = None
            event_type = "click"
            print(f"[CLICK] ({x}, {y}) | Delay: {delay}s")

        recording.append({
            "type": event_type, 
            "x": x, 
            "y": y, 
            "target_color": color, 
            "delay": delay
        })

def on_press(key):
    global last_event_time, shift_pressed
    if key == keyboard.Key.esc: return False
    
    if key == keyboard.Key.shift or key == keyboard.Key.shift_r:
        shift_pressed = True
    
    if last_event_time:
        try: k = key.char
        except AttributeError: k = str(key)
        recording.append({"type": "key", "value": k, "delay": get_delay()})

def on_release(key):
    global shift_pressed
    if key == keyboard.Key.shift or key == keyboard.Key.shift_r:
        shift_pressed = False

print("--- SELECTIVE DELAY RECORDER ---")
print("S: Start | ESC: Save | Shift + Click: Pixel Watcher (Instant)")

with keyboard.Listener(on_press=lambda k: False if hasattr(k, 'char') and k.char.lower() == 's' else True) as s: s.join()
last_event_time = time.time()

with mouse.Listener(on_click=on_click) as m_listener, \
     keyboard.Listener(on_press=on_press, on_release=on_release) as k_listener:
    k_listener.join()

with open("recording.json", "w") as f:
    json.dump(recording, f, indent=4)