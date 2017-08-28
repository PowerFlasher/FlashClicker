import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading
import ctypes
from ctypes import wintypes, c_void_p, c_int, windll, CFUNCTYPE, POINTER
from hotkey import Hotkey
from hotkey import keys
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
        self.hks = []
        #
        self.progress_bar = None
        # 
        self.daemon = True
        self.hooked  = None
        self.start()
        # 
        # self.hookerThread = threading.Thread(target=self.hooker, daemon=True)
        # self.hookerThread.start()

    def set_handle(self, hwnd):
        self._handle = hwnd


    def get_handle(self):
        return self._handle


    def register_hotkey(self):
        for idx, shortcut in enumerate(self.shortcuts):
            self.hks.append(Hotkey(idx + 1, None, keys.get(shortcut[0])))
            print(shortcut, keys.get(shortcut[0]), self.hks[-1].register(None))
        # self.hookerThread = threading.Thread(target=self.hooker, daemon=True)
        # self.hookerThread.start()


    def unregister_hotkey(self):
        for hk in self.hks:
            hk.unregister(None)


    def load_json(self, file):
        if file:
            with open(file) as data_file:
                self.data = json.load(data_file)
                self.title = self.data['title']
                for idx, macro in enumerate(self.data['macros']):
                    self.shortcuts.append((macro['shortcut'], idx))
                self.unregister_hotkey()
                self.register_hotkey()
            # self.hooker()


    def hooker(self):
        while True:
            try:
                msg = wintypes.MSG()
                if ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                    if msg.message == win32con.WM_HOTKEY:
                        print('test', msg.message)
                        print(msg.wParam)
                        self.press_shortcut(msg.wParam - 1)
                        ctypes.windll.user32.PostQuitMessage(0)

                    ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
                    ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))

            finally:
                pass
                # self.unregister_hotkey()

        # with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as listener:
        #     listener.join()


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

        key = chr(lParam[0])
        print(keys)
        print(key, nCode, lParam[0])

        if len(key) > 100:
            print(key)

        # if (CTRL_CODE == int(lParam[0])):
        #     self.uninstallHookProc()

        return ctypes.windll.user32.CallNextHookEx(self.hooked, nCode, wParam, lParam)


    def run(self):                                 
        pointer = self.getFPTR(self.hookProc)

        if self.installHookProc(pointer):
            print('Start listening')
            self.startKeyLog()


    def startKeyLog(self):                                                                
        msg = wintypes.MSG()
        ctypes.windll.user32.GetMessageA(ctypes.byref(msg),0,0,0)


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
        #             self.press_shortcut(shortcut[1])

        # except AttributeError:
        #     print('special key {0} pressed'.format(key))


    def on_release(self, key):
        print('{0} released'.format(key))
        for shortcut in self.shortcuts:
            print(key)
            if str(key) == 'Key.f5':
                print('success')
            # if key.char == shortcut[0]:
                self.press_shortcut(shortcut[1])
        if key == keyboard.Key.esc:
            # Stop listener
            return False


    def key_press(self, key, time_press, wait):
        if key == "F5":
            key = win32con.VK_F5
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


    def run_macro(self, idx, stop_event):
        print('test')
        for count, action in enumerate(self.data['macros'][idx]['script']):
            self.progress_bar.maximum = len(self.data['macros'][idx]['script'])
            if not stop_event.is_set():
                for i in range(action['times']):
                    if not stop_event.is_set():
                        self.key_press( action['key'], action['time_press'], action['wait'])
                        self.progress_bar.value = count + 1
                        print(count + 1)
