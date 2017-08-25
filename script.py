import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading
import ctypes
from ctypes import wintypes
from hotkey import Hotkey
from hotkey import keys
# from pynput import keyboard

class Script:
    def __init__ (self):
        self.__hwnd = None
        self.data = None
        self.title = None
        self.shortcuts = []
        self.thread = None
        self.stop_event = None
        self.hks = []
        #
        self.HOTKEYS = {
          1 : (win32con.VK_F1, None)
        }
        # 
        # self.hookerThread = threading.Thread(target=self.hooker, daemon=True)
        # self.hookerThread.start()


    def set_handle(self, hwnd):
        self._handle = hwnd


    def get_handle(self):
        return self._handle

    def register_hotkey(self):
        for idx, shortcut in enumerate(self.shortcuts):
            print(idx, shortcut, keys.get(shortcut[0]))
            self.hks.append(Hotkey(idx + 1, None, keys.get(shortcut[0])))
            print(self.hks[-1].register(None))
        self.hooker()
        # self.hookerThread = threading.Thread(target=self.hooker, daemon=True)
        # self.hookerThread.start()

    def unregister_hotkey(self):
        for hk in self.hks:
            hk.unregister(None)


    def load_json(self, file):
        if file != None and file != '':
            with open(file) as data_file:
                self.data = json.load(data_file)
                self.title = self.data['title']
                for idx, macro in enumerate(self.data['macros']):
                    self.shortcuts.append((macro['shortcut'], idx))
                self.unregister_hotkey()
                self.register_hotkey()


    def hooker(self):
        while True:
            try:
                msg = wintypes.MSG()
                if ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                    if msg.message == win32con.WM_HOTKEY:
                        print('test', msg.message)
                        print(msg.wParam)
                        ctypes.windll.user32.PostQuitMessage(0)

                    ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
                    ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))

            finally:
                pass
                # self.unregister_hotkey()

        # with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as listener:
        #     listener.join()


    def on_press(self, key):
        pass
        # print(key.char)
        # for shortcut in self.shortcuts:
        #     print(key.char)
        #     if key.char == 'Key.f5':
        #         print('success')
        #     # if key.char == shortcut[0]:
        #         self.press_shortcut(self.__hwnd, shortcut[1])
        # try:
        #     print('alphanumeric key {0} pressed'.format(key.char))
        #     print(self.shortcuts)
        #     for shortcut in self.shortcuts:
        #         print(key.char)
        #         if key.char == 'Key.f5':
        #             print('success')
        #         # if key.char == shortcut[0]:
        #             self.press_shortcut(self.__hwnd, shortcut[1])

        # except AttributeError:
        #     print('special key {0} pressed'.format(key))


    def on_release(self, key):
        print('{0} released'.format(key))
        for shortcut in self.shortcuts:
            print(key)
            if str(key) == 'Key.f5':
                print('success')
            # if key.char == shortcut[0]:
                self.press_shortcut(self.__hwnd, shortcut[1])
        if key == keyboard.Key.esc:
            # Stop listener
            return False


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
        print('test')
        for action in self.data['macros'][idx]['script']:
            if not stop_event.is_set():
                for i in range(action['times']):
                    if not stop_event.is_set():
                        self.key_press(hwnd, action['key'], action['time_press'], action['wait'])
