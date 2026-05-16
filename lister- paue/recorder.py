import time
import json
import os
from pynput import mouse, keyboard
from PIL import ImageGrab

recording = []
last_event_time = None
shift_pressed = False 
ctrl_pressed = False
is_paused = False
pause_start_time = None

ocr_clicks = []
selecting_ocr = False

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

def handle_ocr_selection(x, y):
    global selecting_ocr, last_event_time
    ocr_clicks.append((x, y))
    print(f"[OCR SELECTION] Corner {len(ocr_clicks)} recorded at: ({x}, {y})")
    
    if len(ocr_clicks) == 2:
        x1, y1 = ocr_clicks[0]
        x2, y2 = ocr_clicks[1]
        left, top, right, bottom = min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)
        if left == right: right += 100
        if top == bottom: bottom += 40

        target_text = input("\nEnter target text to wait for during playback: ").strip()
        recording.append({
            "type": "ocr_watcher",
            "bbox": [int(left), int(top), int(right), int(bottom)],
            "target_text": target_text,
            "delay": 0.0
        })
        print(f"[OCR WATCHER ADDED] Region Saved: [{int(left)}, {int(top)}, {int(right)}, {int(bottom)}] | Target: '{target_text}'")
        ocr_clicks.clear()
        selecting_ocr = False
        resume_recording()

def resume_recording():
    global is_paused, last_event_time
    is_paused = False
    last_event_time = time.time()
    print("=== RECORDING RESUMED ===\n")

def on_click(x, y, button, pressed):
    global selecting_ocr
    if pressed and selecting_ocr:
        handle_ocr_selection(x, y)
        return
    if is_paused: return

    if last_event_time:
        delay = get_delay()
        if pressed:
            if shift_pressed:
                color = get_pixel_color(x, y)
                print(f"[PIXEL WATCHER] Target: ({x}, {y}) | Color: {color} | Delay: {delay}s")
                recording.append({"type": "pixel_watcher", "x": x, "y": y, "target_color": color, "delay": delay})
            else:
                print(f"[MOUSE DOWN] ({x}, {y}) | Delay: {delay}s")
                recording.append({"type": "mouse_down", "x": x, "y": y, "delay": delay})
        else:
            if not shift_pressed:
                print(f"[MOUSE UP] ({x}, {y}) | Delay: {delay}s")
                recording.append({"type": "mouse_up", "x": x, "y": y, "delay": delay})

def on_press(key):
    global last_event_time, shift_pressed, ctrl_pressed, is_paused, pause_start_time, selecting_ocr
    
    if key == keyboard.Key.esc: return False
    
    # --- PAUSE / RESUME TOGGLE ---
    if key == keyboard.Key.caps_lock and not selecting_ocr:
        is_paused = not is_paused
        current_time = time.time()
        if is_paused:
            pause_start_time = current_time
            print("\n=== RECORDING PAUSED (Press '\\' for OCR) ===")
        else:
            paused_duration = current_time - pause_start_time
            last_event_time += paused_duration
            print("=== RECORDING RESUMED ===\n")
        return

    # --- OCR TRIGGER IF PAUSED ---
    if is_paused and not selecting_ocr:
        try:
            if key.char == '\\':
                selecting_ocr = True
                print("\n-> OCR Selection Mode. Click Corner 1...")
                return
        except AttributeError: pass

    if is_paused or selecting_ocr: return
    
    # --- CONDITIONAL SELECTOR TRIGGER ( ' Key ) ---
    try:
        if key.char == "'":
            delay = get_delay()
            # Use pynput's controller to dynamically fetch current active cursor coordinates
            from pynput.mouse import Controller
            m_ctrl = Controller()
            mx, my = int(m_ctrl.position[0]), int(m_ctrl.position[1])
            color = get_pixel_color(mx, my)
            
            recording.append({
                "type": "conditional_selector",
                "x": mx,
                "y": my,
                "expected_color": color,
                "accepted" : "accepted.json",
                "rejected" : "rejected.json",
                "delay": delay
            })
            print(f"\n[CONDITIONAL SELECTOR ADDED] Pos: ({mx}, {my}) | Color: {color} | Delay: {delay}s")
            return
    except AttributeError: pass

    if key in [keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
        ctrl_pressed = True
        return
    if key in [keyboard.Key.shift, keyboard.Key.shift_r]:
        shift_pressed = True
        return
    
    if last_event_time:
        delay = get_delay()
        try: k = key.char
        except AttributeError: k = str(key)
        
        if ctrl_pressed and k:
            if k in ['c', 'C', '\x03']: recording.append({"type": "hotkey", "modifier": "ctrl", "key": "c", "delay": delay}); return
            elif k in ['v', 'V', '\x16']: recording.append({"type": "hotkey", "modifier": "ctrl", "key": "v", "delay": delay}); return
            elif k in ['x', 'X', '\x18']: recording.append({"type": "hotkey", "modifier": "ctrl", "key": "x", "delay": delay}); return
            elif k in ['a', 'A', '\x01']: recording.append({"type": "hotkey", "modifier": "ctrl", "key": "a", "delay": delay}); return

        if "ctrl" in str(key).lower() or "shift" in str(key).lower(): return
        recording.append({"type": "key", "value": k, "delay": delay})
        print(f"[KEY] {k} | Delay: {delay}s")

def on_release(key):
    global shift_pressed, ctrl_pressed
    if is_paused or selecting_ocr: return
    if key in [keyboard.Key.shift, keyboard.Key.shift_r]: shift_pressed = False
    if key in [keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]: ctrl_pressed = False

print("--- RECORDER WITH BRANCHING SELECTOR ---")
print("S: Start | Caps Lock: Pause | ' Key: Conditional Selector | ESC: Save")
print("-" * 65)

with keyboard.Listener(on_press=lambda k: False if hasattr(k, 'char') and k.char.lower() == 's' else True) as s: s.join()
last_event_time = time.time()

with mouse.Listener(on_click=on_click) as m_listener, \
     keyboard.Listener(on_press=on_press, on_release=on_release) as k_listener:
    k_listener.join()

with open("recording.json", "w") as f: json.dump(recording, f, indent=4)
print(f"\nSaved {len(recording)} events to recording.json")