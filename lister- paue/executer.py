import json
import time
import random
import string
import pyautogui
from PIL import ImageGrab
import pytesseract
import ctypes

# --- FORCE HIGH-DPI AWARENESS (Fixes screen scaling coordinate offsets) ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

# --- CONFIG ---
FILE_NAME = "recording.json"
CODE_LENGTH = 5
TOLERANCE = 15
TIMEOUT_LIMIT = 45
HOVER_DELAY = 0.1 # 100ms breathing room for UI hover animations to render
HOVER_Condition_DELAY = 1  # 100ms breathing room for UI hover animations to render
pyautogui.PAUSE = 0

# If Tesseract is not in your environment variables, uncomment and configure the line below:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def generate_random_code(length):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def get_current_color(x, y):
    try: return ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).getpixel((0, 0))
    except: return (0, 0, 0)

def color_matches(current, target, tolerance):
    return all(abs(current[i] - target[i]) <= tolerance for i in range(3))

def extract_text_from_bbox(bbox):
    try:
        coords = tuple(int(c) for c in bbox)
        screenshot = ImageGrab.grab(bbox=coords, all_screens=True)
        screenshot = screenshot.convert('L') 
        screenshot = screenshot.resize((screenshot.width * 2, screenshot.height * 2))
        return pytesseract.image_to_string(screenshot, config='--psm 6').strip()
    except: return ""

def execute_event_list(events):
    """Helper runner to execute an event sequence (handles standard and nested components)"""
    aborted = False
    for event in events:
        if aborted: break
        etype = event['type']
        
        # Apply delay for sequential steps
        if etype not in ["pixel_watcher", "ocr_watcher"]:
            time.sleep(event.get('delay', 0))

        # --- NESTED/INLINE CONDITIONAL SELECTOR RUNNER ---
        if etype == "conditional_selector":
            x, y = event['x'], event['y']
            
            # Move to target to trigger hover animation states before sampling color
            pyautogui.moveTo(x, y)
            time.sleep(HOVER_Condition_DELAY)
            
            expected = tuple(event['expected_color'])
            current = get_current_color(x, y)
            print('curr-col',current)
            
            file_true = event.get('file_if_true', event['accepted'])
            file_false = event.get('file_if_false', event['rejected'])
            
            target_file = file_true if color_matches(current, expected, TOLERANCE) else file_false
            print(f"  [SELECTOR] Hovered & Matched ({x}, {y}) -> Routing to: '{target_file}'")
            
            try:
                with open(target_file, "r") as f: sub_events = json.load(f)
                if not execute_event_list(sub_events): aborted = True
            except Exception as e:
                print(f"  [!] Failed to execute sub-macro '{target_file}': {e}")
                aborted = True

        elif etype == "ocr_watcher":
            bbox = event['bbox']
            target_text = event['target_text'].lower()
            start_wait_time = time.time()
            while True:
                if target_text in extract_text_from_bbox(bbox).lower(): break
                if (time.time() - start_wait_time) > TIMEOUT_LIMIT: aborted = True; break
                time.sleep(0.5)

        elif etype == "pixel_watcher":
            x, y = event['x'], event['y']
            
            # Move to coordinates first to force hover rules to catch up
            pyautogui.moveTo(x, y)
            time.sleep(HOVER_DELAY)
            
            target = tuple(event['target_color'])
            start_wait_time = time.time()
            while not color_matches(get_current_color(x, y), target, TOLERANCE):
                if (time.time() - start_wait_time) > TIMEOUT_LIMIT: aborted = True; break
                time.sleep(0.01)
            if not aborted: 
                time.sleep(0.5)
                pyautogui.click(x, y)

        elif etype == "click": 
            pyautogui.moveTo(event['x'], event['y'])
            time.sleep(HOVER_DELAY)
            pyautogui.click(event['x'], event['y'])
            
        elif etype == "mouse_down": 
            pyautogui.moveTo(event['x'], event['y'])
            time.sleep(HOVER_DELAY)
            pyautogui.mouseDown(event['x'], event['y'])
            
        elif etype == "mouse_up": 
            pyautogui.moveTo(event['x'], event['y'])
            time.sleep(HOVER_DELAY)
            pyautogui.mouseUp(event['x'], event['y'])
            
        elif etype == "hotkey": pyautogui.hotkey(event['modifier'], event['key'])
        elif etype == "key":
            val = str(event.get('value', ''))
            if val in ["None", ";"] or "Key.shift" in val or "Key.ctrl" in val: continue
            if val == "_" or "underscore" in val:
                pyautogui.write(generate_random_code(CODE_LENGTH), interval=0)
            else:
                try: pyautogui.press(val.replace("Key.", ""))
                except: pass
    return not aborted

def run_playback():
    try:
        with open(FILE_NAME, "r") as f: main_events = json.load(f)
    except Exception as e:
        print(f"Error loading main macro file {FILE_NAME}: {e}")
        return

    loops = int(input("Loops: ") or 1)
    time.sleep(3)
    
    for run_idx in range(loops):
        print(f"\n>>> RUN #{run_idx + 1} STARTING")
        success = execute_event_list(main_events)
        
        if not success:
            print(f">>> RUN #{run_idx + 1} FAILED / ABORTED")
        else:
            print(f">>> RUN #{run_idx + 1} SUCCESSFUL")
        time.sleep(2)

if __name__ == "__main__":
    run_playback()