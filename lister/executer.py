import json
import time
import random
import string
import pyautogui
from PIL import ImageGrab

# --- CONFIG ---
FILE_NAME = "recording.json"
CODE_LENGTH = 5
TOLERANCE = 15
TIMEOUT_LIMIT = 20  # Seconds to wait before aborting the loop
pyautogui.PAUSE = 0

def generate_random_code(length):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def get_current_color(x, y):
    try: return ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).getpixel((0, 0))
    except: return (0, 0, 0)

def color_matches(current, target, tolerance):
    return all(abs(current[i] - target[i]) <= tolerance for i in range(3))

def run_playback():
    try:
        with open(FILE_NAME, "r") as f:
            events = json.load(f)
    except: return

    loops = int(input("Loops: ") or 1)
    time.sleep(3)
    
    for run_idx in range(loops):
        print(f"\n>>> RUN #{run_idx + 1} STARTING")
        aborted = False

        for event in events:
            if aborted: break # Break out of the event list to start next loop

            etype = event['type']
            
            # --- DELAY LOGIC ---
            if etype != "pixel_watcher":
                time.sleep(event.get('delay', 0))

            # --- ACTION LOGIC WITH TIMEOUT ---
            if etype in ["pixel_watcher", "smart_click"]:
                x, y = event['x'], event['y']
                target = tuple(event['target_color'])
                
                start_wait_time = time.time()
                print(f"Waiting for color at ({x}, {y})...")
                
                while not color_matches(get_current_color(x, y), target, TOLERANCE):
                    # Check if we have exceeded the 20-second limit
                    if (time.time() - start_wait_time) > TIMEOUT_LIMIT:
                        print(f"[!] TIMEOUT: Color not found at ({x}, {y}) after {TIMEOUT_LIMIT}s. Aborting run.")
                        aborted = True
                        break
                    time.sleep(0.01)
                
                if not aborted:
                    pyautogui.click(x, y)

            elif etype == "click":
                pyautogui.click(event['x'], event['y'])
            
            elif etype == "key":
                val = str(event.get('value', ''))
                if val in ["None", ";"] or "Key.shift" in val: continue
                
                if val == "_" or "underscore" in val:
                    pyautogui.write(generate_random_code(CODE_LENGTH), interval=0)
                else:
                    try: pyautogui.press(val.replace("Key.", ""))
                    except: pass
        
        if aborted:
            print(f">>> RUN #{run_idx + 1} FAILED (Aborted)")
        else:
            print(f">>> RUN #{run_idx + 1} SUCCESSFUL")
        
        time.sleep(2) # Breath time between loops

if __name__ == "__main__":
    run_playback()