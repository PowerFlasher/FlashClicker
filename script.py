import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading

class Script:
    def __init__ (self):
        self.data = None
        self.title = None
        self.shortcuts = []
        self.thread = None
        self.stop_event = None

    def load_json(self, file):
        if file != None and file != '':
            with open(file) as data_file:
                self.data = json.load(data_file)
                self.title = self.data['title']
                for idx, macro in enumerate(self.data['macros']):
                    self.shortcuts.append((macro['shortcut'], idx))

    def key_press(self, hwnd, key, time_press, wait):
        if key == "F5":
            key = win32con.VK_F5
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
        time.sleep(time_press / 1000.0)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, key, 0)
        time.sleep(wait / 1000.0)

    def press_shortcut(self, hwnd, idx):
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self.run_macro, args=(hwnd, idx, self.stop_event), daemon=True)
        self.thread.start()

    def stop_macro(self):
        self.stop_event.set()

    def run_macro(self, hwnd, idx, stop_event):
        for action in self.data['macros'][idx]['script']:
            if not stop_event.is_set():
                for i in range(action['times']):
                    if not stop_event.is_set():
                        self.key_press(hwnd, action['key'], action['time_press'], action['wait'])
