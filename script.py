import json
import win32con
import win32api
import win32gui
import win32process
import time
import threading
import ctypes
import os
from ctypes import wintypes, c_void_p, c_int, windll, CFUNCTYPE, POINTER
from hotkey import Dirkey, keys, values, dik, MouseInput, mouse_buttons_keys, mouse_buttons_values
from macro_recorder import MacroRecorder
# from pynput import keyboard

WH_KEYBOARD_LL = win32con.WH_KEYBOARD_LL     
WH_MOUSE_LL = win32con.WH_MOUSE_LL                                                            
WM_KEYUP = win32con.WM_KEYUP
WM_KEYDOWN = win32con.WM_KEYDOWN
HC_ACTION = win32con.HC_ACTION
WM_SYSKEYDOWN = win32con.WM_SYSKEYDOWN
WM_SYSKEYUP = win32con.WM_SYSKEYUP

#: The messages that correspond to a key press
PRESS_MESSAGES = (WM_KEYDOWN, WM_SYSKEYDOWN)
#: The messages that correspond to a key release
RELEASE_MESSAGES = (WM_KEYUP, WM_SYSKEYUP)

WM_QUIT        = 0x0012
WM_MOUSEMOVE   = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP   = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP   = 0x0205
WM_MBUTTONDOWN = 0x0207
WM_MBUTTONUP   = 0x0208
WM_MOUSEWHEEL  = 0x020A
WM_MOUSEHWHEEL = 0x020E
WM_XBUTTONDOWN = 0x020B
WM_XBUTTONUP = 0x020C

MSG_TEXT = {WM_MOUSEMOVE:   'WM_MOUSEMOVE',
            WM_LBUTTONDOWN: 'WM_LBUTTONDOWN',
            WM_LBUTTONUP:   'WM_LBUTTONUP',
            WM_RBUTTONDOWN: 'WM_RBUTTONDOWN',
            WM_RBUTTONUP:   'WM_RBUTTONUP',
            WM_MBUTTONDOWN: 'WM_MBUTTONDOWN',
            WM_MBUTTONUP:   'WM_MBUTTONUP',
            WM_MOUSEWHEEL:  'WM_MOUSEWHEEL',
            WM_MOUSEHWHEEL: 'WM_MOUSEHWHEEL',
            WM_XBUTTONDOWN: 'WM_XBUTTONDOWN',
            WM_XBUTTONUP: 'WM_XBUTTONUP'}


class Script(threading.Thread):
    def __init__ (self):
        threading.Thread.__init__(self)
        self.__hwnd = None
        self.load_recorded_macro = None
        self.set_action_text = None
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
        self.record_macro = False
        self.macro_recorder = MacroRecorder()
        # 
        self.daemon = True
        self.hooked = [None, None]
        self.start()
        # print('Keys', keys)
        # print('Values', values)
        # print(self.get_key_text(65))
        # print(self.get_text_key('VK_F1'))
        # print(self.get_text_key('A'))


    def get_key_text(self, key):
        return keys.get(hex(key), chr(key))


    def get_text_key(self, text):
        if 'DIK' in str(text):
            key = dik.get(str(text))
        elif 'MK' in str(text):
            key = mouse_buttons_values.get(str(text))
        else:
            key = values.get(str(text))
            if key == None and len(str(text)) == 1:
                key = ord(str(text))
        return key


    def set_handle(self, hwnd):
        self.__hwnd = hwnd


    def get_handle(self):
        return self.__hwnd


    def set_status_info(self, title, key, status):
        self.macro_title = title
        self.macro_key = key
        self.status_text = status

    def get_statistics(self):
        statistics = []
        time_keys = 0
        if self.data:
            for idx, macro in enumerate(self.data['macros']):
                for key in macro['script']:
                    time_keys += (key['time_press'] + key['wait']) * key['times']
                time_keys *= macro['times']
                millis = int(time_keys)
                seconds=(millis/1000)%60
                seconds = round(seconds,2)
                minutes=(millis/(1000*60))%60
                minutes = round(minutes,2)
                hours=(millis/(1000*60*60))%24
                hours=round(hours,2)
                days=(millis/(1000*60*60*24))
                days=round(days,2)
                statistics.append((macro['title'], millis, seconds, minutes, hours, days))
        return statistics

    def load_json(self, file):
        if file:
            try:
                self.shortcuts = []
                while not os.path.exists(os.path.join(os.getcwd(),file)):
                    time.sleep(1)
                with open(file, 'r') as data_file:
                    self.data = json.load(data_file)
                    # data_file.close()
                    self.title = self.data['title']
                    self.infinity = self.data['infinity']
                    for idx, macro in enumerate(self.data['macros']):
                        self.shortcuts.append((str(macro['shortcut']), idx))
            except FileNotFoundError():
                print(e)


    def installHookProc(self, key_pointer, mouse_pointer):                                           
        self.hooked[0] = ctypes.windll.user32.SetWindowsHookExA( 
                        WH_KEYBOARD_LL, 
                        key_pointer, 
                        windll.kernel32.GetModuleHandleW(None), 
                        0
        )
        self.hooked[1] = ctypes.windll.user32.SetWindowsHookExA( 
                        WH_MOUSE_LL, 
                        mouse_pointer, 
                        windll.kernel32.GetModuleHandleW(None), 
                        0
        )
        if not self.hooked[0] and not self.hooked[1]:
            return False
        return True


    def uninstallHookProc(self):                                                  
        if self.hooked[0] is None:
            return
        ctypes.windll.user32.UnhookWindowsHookEx(self.hooked[0])
        self.hooked[0] = None
        if self.hooked[1] is None:
            return
        ctypes.windll.user32.UnhookWindowsHookEx(self.hooked[1])
        self.hooked[1] = None


    def getFPTR(self, fn):                                                                  
        CMPFUNC = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))
        return CMPFUNC(fn)

    def mouse_hook_proc(self, nCode, wParam, lParam):
        if nCode == HC_ACTION:
            msg = ctypes.cast(lParam, POINTER(MouseInput))[0]
            msgid = MSG_TEXT.get(wParam, str(wParam))
            # msg = ((msg.dx, msg.dy),
            #     msg.mouseData, msg.dwFlags,
            #     msg.time, msg.dwExtraInfo)
            # print('{:15s}: {}'.format(msgid, msg))
            for shortcut in self.shortcuts:
                if mouse_buttons_keys.get(msg.mouseData) and shortcut[0] == mouse_buttons_keys.get(msg.mouseData) and 'DOWN' in msgid:
                    if not self.thread.is_alive():
                        self.current_shortkey = shortcut[0]
                        self.press_shortcut(shortcut[1])
                    elif shortcut[0] == self.current_shortkey:
                        self.current_shortkey = None
                        self.stop_macro()

        return ctypes.windll.user32.CallNextHookEx(self.hooked[1], nCode, wParam, lParam)

    def key_hook_proc(self, nCode, wParam, lParam):
        if self.record_macro:
            if wParam in PRESS_MESSAGES:
                self.macro_recorder.on_press(chr(lParam[0]) if keys.get(hex(lParam[0])) is None else keys.get(hex(lParam[0])))
            if wParam in RELEASE_MESSAGES:
                self.macro_recorder.on_release(chr(lParam[0]) if keys.get(hex(lParam[0])) is None else keys.get(hex(lParam[0])), self.load_recorded_macro)
        if wParam is not WM_KEYDOWN:
            return ctypes.windll.user32.CallNextHookEx(self.hooked[0], nCode, wParam, lParam)
        if self.get_key_text(lParam[0]) and "VK_F10" == self.get_key_text(lParam[0]):
            self.get_cursor_pos()
        for shortcut in self.shortcuts:
            if self.get_key_text(lParam[0]) and shortcut[0] == self.get_key_text(lParam[0]):
                if not self.thread.is_alive():
                    self.current_shortkey = shortcut[0]
                    self.press_shortcut(shortcut[1])
                elif shortcut[0] == self.current_shortkey:
                    self.current_shortkey = None
                    self.stop_macro()

        return ctypes.windll.user32.CallNextHookEx(self.hooked[0], nCode, wParam, lParam)


    def run(self):                                 
        key_pointer = self.getFPTR(self.key_hook_proc)
        mouse_pointer = self.getFPTR(self.mouse_hook_proc)

        if self.installHookProc(key_pointer, mouse_pointer):
            print('Start listening')
            self.startKeyLog()


    def startKeyLog(self):                                                                
        msg = wintypes.MSG()
        ctypes.windll.user32.GetMessageA(ctypes.byref(msg),0,0,0)


    def key_press(self, key, time_press, wait, action):
        self.set_action_text('Press key: ' + str(key))
        if action == 'chat':
            for ch in key:
                k = self.get_text_key(ch)
                print('Key press: {}'.format(k))
                win32api.PostMessage(self.__hwnd, win32con.WM_CHAR, k, 0)
        elif action == 'virtual_key':
            print('Key press: {}'.format(key))
            key = self.get_text_key(key)
            win32api.PostMessage(self.__hwnd, win32con.WM_KEYDOWN, key, 0)
            time.sleep(time_press / 1000.0)
            win32api.PostMessage(self.__hwnd, win32con.WM_KEYUP, key, 0)
        elif action == 'direct_key':
            print('Key press: {}'.format(key))
            key = self.get_text_key(key)
            self.dirkey.press_key(key)
            time.sleep(time_press / 1000.0)
            self.dirkey.release_key(key)
        self.set_action_text('Waiting')
        time.sleep(wait / 1000.0)

    def mouse_click(self, button, time_press, wait, point):
        # client_pos = (point['x'], point['y'])
        client_pos = win32gui.ScreenToClient(self.__hwnd, (point['x'], point['y']))
        lParam = win32api.MAKELONG(client_pos[0], client_pos[1])
        cur_pos = win32api.GetCursorPos()
        win32api.SetCursorPos(client_pos)
        print('Mouse click: {} at position x={}, y={}, with client position by window x={}, y={}'.format(button, point['x'], point['y'], client_pos[0], client_pos[1]))
        self.set_action_text('Press ' + str(button) + ' mouse key at position: x=' + str(point['x']) + ', y=' + str(point['y']))
        if button == "left":
            win32api.PostMessage(self.__hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
            win32api.PostMessage(self.__hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(time_press / 1000.0)
            win32api.PostMessage(self.__hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)
        elif button == "right":
            win32api.PostMessage(self.__hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
            win32api.PostMessage(self.__hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lParam)
            time.sleep(time_press / 1000.0)
            win32api.PostMessage(self.__hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, lParam)
        win32api.SetCursorPos(cur_pos)
        self.set_action_text('Waiting')
        time.sleep(wait / 1000.0)

    def get_cursor_pos(self):
        pos = win32gui.GetCursorPos()
        print('Cursor position at x={}, y={}'.format(pos[0], pos[1]))


    def press_shortcut(self, idx):
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self.run_macro, args=(idx, self.stop_event), daemon=True)
        self.thread.start()


    def stop_macro(self):
        if self.stop_event:
            self.stop_event.set()
            self.progress_bar["value"] = 0


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
            self.macro_title.set((self.data['macros'][idx]['title'][:35] + '..') if len(self.data['macros'][idx]['title']) > 35 else self.data['macros'][idx]['title'])
            self.macro_key.set(self.current_shortkey.replace("VK_", ""))
            self.status_text.set('RUN')
            for count in range(self.data['macros'][idx]['times']):
                if not stop_event.is_set():
                    for action in self.data['macros'][idx]['script']:
                        if not stop_event.is_set():
                            for i in range(action['times']):
                                if not stop_event.is_set():
                                    if action['action'] == "mouse":
                                        self.mouse_click(action['key'], action['time_press'], action['wait'], action['point'])
                                    else:
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
        self.set_action_text('')