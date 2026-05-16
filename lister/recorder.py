import time
import json
import os
from pynput import mouse, keyboard
from PIL import ImageGrab

recording = []
last_event_time = None
shift_pressed = False 
ctrl_pressed = False

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
    if last_event_time:
        delay = get_delay()
        
        if pressed:
            # Shift + Click = pixel_watcher
            if shift_pressed:
                color = get_pixel_color(x, y)
                print(f"[PIXEL WATCHER] Target: ({x}, {y}) | Color: {color} | Recorded Delay: {delay}s")
                recording.append({
                    "type": "pixel_watcher", 
                    "x": x, 
                    "y": y, 
                    "target_color": color, 
                    "delay": delay
                })
            else:
                print(f"[MOUSE DOWN] ({x}, {y}) | Delay: {delay}s")
                recording.append({
                    "type": "mouse_down", 
                    "x": x, 
                    "y": y, 
                    "delay": delay
                })
        else:
            # Record mouse up to complete clicks or drag selections
            if not shift_pressed:
                print(f"[MOUSE UP] ({x}, {y}) | Delay: {delay}s")
                recording.append({
                    "type": "mouse_up", 
                    "x": x, 
                    "y": y, 
                    "delay": delay
                })

def on_press(key):
    global last_event_time, shift_pressed, ctrl_pressed
    if key == keyboard.Key.esc: return False
    
    if key in [keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
        ctrl_pressed = True
        return
    
    if key in [keyboard.Key.shift, keyboard.Key.shift_r]:
        shift_pressed = True
        return
    
    if last_event_time:
        delay = get_delay()
        try: 
            k = key.char
        except AttributeError: 
            k = str(key)
        
        # Intercept and log hotkey combinations
        if ctrl_pressed and k:
            if k in ['c', 'C', '\x03']: # '\x03' is the ASCII control code for Ctrl+C
                recording.append({"type": "hotkey", "modifier": "ctrl", "key": "c", "delay": delay})
                print(f"[HOTKEY] ctrl + c | Delay: {delay}s")
                return
            elif k in ['v', 'V', '\x16']: # '\x16' is Ctrl+V
                recording.append({"type": "hotkey", "modifier": "ctrl", "key": "v", "delay": delay})
                print(f"[HOTKEY] ctrl + v | Delay: {delay}s")
                return
            elif k in ['x', 'X', '\x18']: # '\x18' is Ctrl+X
                recording.append({"type": "hotkey", "modifier": "ctrl", "key": "x", "delay": delay})
                print(f"[HOTKEY] ctrl + x | Delay: {delay}s")
                return
            elif k in ['a', 'A', '\x01']: # '\x01' is Ctrl+A
                recording.append({"type": "hotkey", "modifier": "ctrl", "key": "a", "delay": delay})
                print(f"[HOTKEY] ctrl + a | Delay: {delay}s")
                return

        # Prevent standalone modifier key strings from cluttering output
        if "ctrl" in str(key).lower() or "shift" in str(key).lower():
            return

        recording.append({"type": "key", "value": k, "delay": delay})
        print(f"[KEY] {k} | Delay: {delay}s")

def on_release(key):
    global shift_pressed, ctrl_pressed
    if key in [keyboard.Key.shift, keyboard.Key.shift_r]:
        shift_pressed = False
    if key in [keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
        ctrl_pressed = False

print("--- SELECTIVE DELAY RECORDER WITH DRAG & COPY/PASTE ---")
print("S: Start | ESC: Save | Shift + Click: Pixel Watcher (Instant)")

with keyboard.Listener(on_press=lambda k: False if hasattr(k, 'char') and k.char.lower() == 's' else True) as s: s.join()
last_event_time = time.time()

with mouse.Listener(on_click=on_click) as m_listener, \
     keyboard.Listener(on_press=on_press, on_release=on_release) as k_listener:
    k_listener.join()

with open("recording.json", "w") as f:
    json.dump(recording, f, indent=4)