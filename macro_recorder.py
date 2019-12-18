from pynput import keyboard
import time
import json

class KeyButton(object):
    def __init__(self, keyID):
        self.keyID = keyID
        self.timeKeyPress = 0
        self.timeKeyRelease = 0

    def setTimeKeyPress(self, time):
        self.timeKeyPress = time

    def setTimeKeyRelease(self, time):
        self.timeKeyRelease = time - self.timeKeyPress

    def __repr__(self):
        return json.dumps(self.__dict__)

    def to_dict(self):
        return {"key": str(self.keyID).upper(), "action": "virtual_key", "time_press": self.timeKeyRelease, "wait": 0, "times": 1}

def current_milli_time():
    return int(round(time.time() * 1000))

keys = []
    
def on_press(key):
    if str(key) != 'Key.esc':
        if not keys:
            keys.append(KeyButton(key))
            keys[-1].setTimeKeyPress(current_milli_time())
            print('Key {} pressed.'.format(key))
        elif keys[-1].timeKeyRelease != 0:
            keys.append(KeyButton(key))
            keys[-1].setTimeKeyPress(current_milli_time())
            print('Key {} pressed.'.format(key))

def on_release(key):
    if str(key) == 'Key.esc':
        print('Exporting data...')
        data = [kb.to_dict() for kb in keys]
        json_data = json.dumps({"script": data})
        result = {
            "title": "Script Macro",
            "macros": [
                {
                    "title": "Recorded Macro",
                    "shortcut": "F8",
                    "times": 1,
                    "script": data
                }
            ]
        }
        with open('script_recorded_macro.json', 'w') as outfile:
            json.dump(result, outfile)
        # print(json.dumps(result, indent=4))
        print('Exiting...')
        return False

    if keys and str(keys[-1].keyID) == str(key):
        keys[-1].setTimeKeyRelease(current_milli_time())
        print('Key {} released in time {}.'.format(key, keys[-1].timeKeyRelease))

print('Starting listener...')
with keyboard.Listener(
    on_press = on_press,
    on_release = on_release) as listener:
    listener.join()