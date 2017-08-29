import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading
import ctypes
from ctypes import wintypes, c_void_p, c_int, windll, CFUNCTYPE, POINTER
from hotkey import Hotkey, keys, values
# from pynput import keyboard

WH_KEYBOARD_LL = 13                                                                 
WM_KEYDOWN = 0x0100
CTRL_CODE = 162

class Script(threading.Thread):
    def __init__ (self):
        threading.Thread.__init__(self)
        self.__hwnd = None
        self.data = None
        self.title = None
        self.shortcuts = []
        self.thread = None
        self.stop_event = None
        # self.hks = []
        #
        self.progress_bar = None
        # 
        self.daemon = True
        self.hooked  = None
        self.start()
        # print('Keys', keys)
        # print('Values', values)
        # print(self.get_key_text(65))
        # print(self.get_text_key('VK_F1'))
        # print(self.get_text_key('A'))


    def get_key_text(self, key):
        return keys.get(key, chr(key))


    def get_text_key(self, text):
        key = values.get(text)
        if key == None and len(text) == 1:
            key = ord(text)
        return key


    def set_handle(self, hwnd):
        self.__hwnd = hwnd


    def get_handle(self):
        return self.__hwnd


    # def register_hotkey(self):
    #     for idx, shortcut in enumerate(self.shortcuts):
    #         self.hks.append(Hotkey(idx + 1, None, keys.get(shortcut[0])))
    #         print(shortcut, keys.get(shortcut[0]), self.hks[-1].register(None))


    # def unregister_hotkey(self):
    #     for hk in self.hks:
    #         hk.unregister(None)


    def load_json(self, file):
        if file:
            with open(file, 'r') as data_file:
                self.data = json.load(data_file)
                self.title = self.data['title']
                for idx, macro in enumerate(self.data['macros']):
                    self.shortcuts.append((macro['shortcut'], idx))
                data_file.close()
                # self.unregister_hotkey()
                # self.register_hotkey()
            # self.hooker()


    def installHookProc(self, pointer):                                           
        self.hooked = ctypes.windll.user32.SetWindowsHookExA( 
                        WH_KEYBOARD_LL, 
                        pointer, 
                        windll.kernel32.GetModuleHandleW(None), 
                        0
        )

        if not self.hooked:
            return False
        return True


    def uninstallHookProc(self):                                                  
        if self.hooked is None:
            return
        ctypes.windll.user32.UnhookWindowsHookEx(self.hooked)
        self.hooked = None


    def getFPTR(self, fn):                                                                  
        CMPFUNC = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))
        return CMPFUNC(fn)


    def hookProc(self, nCode, wParam, lParam):                                              
        if wParam is not WM_KEYDOWN:
            return ctypes.windll.user32.CallNextHookEx(self.hooked, nCode, wParam, lParam)

        for shortcut in self.shortcuts:
            if keys.get(lParam[0]) and shortcut[0] in keys.get(lParam[0]):
                self.press_shortcut(shortcut[1])

        return ctypes.windll.user32.CallNextHookEx(self.hooked, nCode, wParam, lParam)


    def run(self):                                 
        pointer = self.getFPTR(self.hookProc)

        if self.installHookProc(pointer):
            print('Start listening')
            self.startKeyLog()


    def startKeyLog(self):                                                                
        msg = wintypes.MSG()
        ctypes.windll.user32.GetMessageA(ctypes.byref(msg),0,0,0)


    def key_press(self, key, time_press, wait, action):
        if action == 'keyboard_write':
            for ch in key:
                k = self.get_text_key(ch)
                win32api.PostMessage(self.__hwnd, win32con.WM_CHAR, k, 0)
        elif action == 'keyboard_control':
            key = self.get_text_key(key)
            win32api.PostMessage(self.__hwnd, win32con.WM_KEYDOWN, key, 0)
            time.sleep(time_press / 1000.0)
            win32api.PostMessage(self.__hwnd, win32con.WM_KEYUP, key, 0)
        time.sleep(wait / 1000.0)


    def press_shortcut(self, idx):
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self.run_macro, args=(idx, self.stop_event), daemon=True)
        self.thread.start()


    def stop_macro(self):
        self.stop_event.set()


    def get_max_progress(self, idx):
        count = 0
        for action in self.data['macros'][idx]['script']:
            for i in range(action['times']):
                count += 1
        return count


    def run_macro(self, idx, stop_event):
        self.progress_bar["maximum"] = self.get_max_progress(idx)
        self.progress_bar["value"] = 0
        for action in self.data['macros'][idx]['script']:
            if not stop_event.is_set():
                for i in range(action['times']):
                    if not stop_event.is_set():
                        self.key_press(action['key'], action['time_press'], action['wait'], action['action'])
                        self.progress_bar.update()
                        self.progress_bar["value"] += 1