import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading
import ctypes
from ctypes import wintypes, c_void_p, c_int, windll, CFUNCTYPE, POINTER
from hotkey import Dirkey, keys, values, dik
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
        self.thread = threading.Thread()
        self.dirkey = Dirkey()
        self.stop_event = None
        self.current_shortkey = None
        #
        self.progress_bar = None
        self.infinity = False
        self.macro_title = None
        self.macro_key = None
        self.status_text = None
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
        if 'DIK' in text:
            key = dik.get(text)
        else:
            key = values.get(text)
            if key == None and len(text) == 1:
                key = ord(text)
        return key


    def set_handle(self, hwnd):
        self.__hwnd = hwnd


    def get_handle(self):
        return self.__hwnd


    def set_status_info(self, title, key, status):
        self.macro_title = title
        self.macro_key = key
        self.status_text = status


    def load_json(self, file):
        try:
            if file:
                with open(file, 'r') as data_file:
                    self.shortcuts = []
                    self.data = json.load(data_file)
                    self.title = self.data['title']
                    for idx, macro in enumerate(self.data['macros']):
                        self.shortcuts.append((macro['shortcut'], idx))
                    data_file.close()
        except Exception: 
            pass


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
                if not self.thread.is_alive():
                    self.current_shortkey = shortcut[0]
                    self.press_shortcut(shortcut[1])
                elif shortcut[0] == self.current_shortkey:
                    self.current_shortkey = None
                    self.stop_macro()

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
        elif action == 'direct_input':
            key = self.get_text_key(key)
            self.dirkey.press_key(key)
            time.sleep(time_press / 1000.0)
            self.dirkey.release_key(key)
        time.sleep(wait / 1000.0)


    def press_shortcut(self, idx):
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self.run_macro, args=(idx, self.stop_event), daemon=True)
        self.thread.start()


    def stop_macro(self):
        self.stop_event.set()


    def get_max_progress(self, idx):
        count = 0
        for c in range(self.data['macros'][idx]['times']):
            for action in self.data['macros'][idx]['script']:
                for i in range(action['times']):
                    count += 1
        return count


    def run_macro(self, idx, stop_event):
        while True and not stop_event.is_set():
            self.progress_bar["maximum"] = self.get_max_progress(idx)
            self.progress_bar["value"] = 0
            self.macro_title.set((self.data['macros'][idx]['title'][:23] + '..') if len(self.data['macros'][idx]['title']) > 23 else self.data['macros'][idx]['title'])
            self.macro_key.set(self.current_shortkey.replace("VK_", ""))
            self.status_text.set('RUN')
            for count in range(self.data['macros'][idx]['times']):
                if not stop_event.is_set():
                    for action in self.data['macros'][idx]['script']:
                        if not stop_event.is_set():
                            for i in range(action['times']):
                                if not stop_event.is_set():
                                    self.key_press(action['key'], action['time_press'], action['wait'], action['action'])
                                    self.progress_bar.update()
                                    self.progress_bar["value"] += 1
                                else:
                                    self.status_text.set('STOPPED')
                                    break
                        else:
                            self.status_text.set('STOPPED')
                            break
                else:
                    self.status_text.set('STOPPED')
                    break
            if not stop_event.is_set() and not self.infinity:
                self.status_text.set('FINISH')
                break
        else:
            self.status_text.set('STOPPED')