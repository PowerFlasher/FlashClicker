# from pynput import keyboard
import time
import json
import os
from datetime import datetime
import winsound

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
        return {"key": str(self.keyID).upper().replace('\'', ''), "action": "virtual_key", "time_press": self.timeKeyRelease, "wait": 0, "times": 1}

def current_milli_time():
    return int(round(time.time() * 1000))
    
class MacroRecorder(object):
    """docstring for MacroRecorder"""
    def __init__(self):
        self.keys = []
     
    def on_press(self, key):
        # if str(key) != 'Key.esc':
        if str(key) != 'VK_ESCAPE':
            if not self.keys:
                winsound.Beep(1000, 100)  # Beep at 1000 Hz for 100 ms
                self.keys.append(KeyButton(key))
                self.keys[-1].setTimeKeyPress(current_milli_time())
                print('Key {} pressed.'.format(key))
            elif self.keys[-1].timeKeyRelease != 0:
                winsound.Beep(1000, 100)  # Beep at 1000 Hz for 100 ms
                self.keys.append(KeyButton(key))
                self.keys[-1].setTimeKeyPress(current_milli_time())
                print('Key {} pressed.'.format(key))

    def on_release(self, key, callback):
        # if str(key) == 'Key.esc':
        if str(key) == 'VK_ESCAPE':
            print('Exporting data...')
            data = [kb.to_dict() for kb in self.keys]
            json_data = json.dumps({"script": data})
            result = {
                "title": "Script Macro",
                "macros": [
                    {
                        "title": "Recorded Macro",
                        "shortcut": "VK_F8",
                        "times": 1,
                        "script": data
                    }
                ]
            }
            file_path = os.path.join('scripts', 'Record_' + datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + '.json')
            with open(file_path, 'w') as outfile:
                json.dump(result, outfile, indent=4)
                outfile.close()
                callback(file_path)
            # print(json.dumps(result, indent=4))
            print('File exported...')
            self.keys = []
            winsound.Beep(4000, 1000)  # Beep at 1000 Hz for 100 ms
            return False

        if self.keys and str(self.keys[-1].keyID) == str(key):
            winsound.Beep(3000, 100)  # Beep at 1000 Hz for 100 ms
            self.keys[-1].setTimeKeyRelease(current_milli_time())
            print('Key {} released in time {}.'.format(key, self.keys[-1].timeKeyRelease))