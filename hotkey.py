import ctypes
import win32con
  
keys = {
    "shift": win32con.MOD_SHIFT
    , "control": win32con.MOD_CONTROL
    , "ctrl": win32con.MOD_CONTROL
    , "alt": win32con.MOD_ALT
    , "win": win32con.MOD_WIN
    , "up": win32con.VK_UP
    , "down": win32con.VK_DOWN
    , "left": win32con.VK_LEFT
    , "right": win32con.VK_RIGHT
    , "pgup": win32con.VK_PRIOR
    , "pgdown": win32con.VK_NEXT
    , "home": win32con.VK_HOME
    , "end": win32con.VK_END
    , "insert": win32con.VK_INSERT
    , "enter": win32con.VK_RETURN
    , "return": win32con.VK_RETURN
    , "tab": win32con.VK_TAB
    , "space": win32con.VK_SPACE
    , "backspace": win32con.VK_BACK
    , "delete": win32con.VK_DELETE
    , "del": win32con.VK_DELETE
    , "apps": win32con.VK_APPS
    , "popup": win32con.VK_APPS
    , "escape": win32con.VK_ESCAPE
    , "npmul": win32con.VK_MULTIPLY
    , "npadd": win32con.VK_ADD
    , "npsep": win32con.VK_SEPARATOR
    , "npsub": win32con.VK_SUBTRACT
    , "npdec": win32con.VK_DECIMAL
    , "npdiv": win32con.VK_DIVIDE
    , "np0": win32con.VK_NUMPAD0
    , "numpad0": win32con.VK_NUMPAD0
    , "np1": win32con.VK_NUMPAD1
    , "numpad1": win32con.VK_NUMPAD1
    , "np2": win32con.VK_NUMPAD2
    , "numpad2": win32con.VK_NUMPAD2
    , "np3": win32con.VK_NUMPAD3
    , "numpad3": win32con.VK_NUMPAD3
    , "np4": win32con.VK_NUMPAD4
    , "numpad4": win32con.VK_NUMPAD4
    , "np5": win32con.VK_NUMPAD5
    , "numpad5": win32con.VK_NUMPAD5
    , "np6": win32con.VK_NUMPAD6
    , "numpad6": win32con.VK_NUMPAD6
    , "np7": win32con.VK_NUMPAD7
    , "numpad7": win32con.VK_NUMPAD7
    , "np8": win32con.VK_NUMPAD8
    , "numpad8": win32con.VK_NUMPAD8
    , "np9": win32con.VK_NUMPAD9
    , "numpad9": win32con.VK_NUMPAD9
    , "F1": win32con.VK_F1
    , "F2": win32con.VK_F2
    , "F3": win32con.VK_F3
    , "F4": win32con.VK_F4
    , "F5": win32con.VK_F5
    , "F6": win32con.VK_F6
    , "F7": win32con.VK_F7
    , "F8": win32con.VK_F8
    , "F9": win32con.VK_F9
    , "F10": win32con.VK_F10
    , "F11": win32con.VK_F11
    , "F12": win32con.VK_F12
    , "F13": win32con.VK_F13
    , "F14": win32con.VK_F14
    , "F15": win32con.VK_F15
    , "F16": win32con.VK_F16
    , "F17": win32con.VK_F17
    , "F18": win32con.VK_F18
    , "F19": win32con.VK_F19
    , "F20": win32con.VK_F20
    , "F21": win32con.VK_F21
    , "F22": win32con.VK_F22
    , "F23": win32con.VK_F23
    , "F24": win32con.VK_F24
}
  
class Hotkey(object):
  
    def __init__(self, keyId, modifiers, virtualkeys):
  
        self.keyId = keyId 
        self.modifiers = modifiers
        self.virtualkeys = virtualkeys
  
    def register(self, hwnd):
        """
        Registers the hotkeys into windows
        Returns true on success
        Returns false on error
        """
  
        if ctypes.windll.user32.RegisterHotKey(hwnd, self.keyId, self.modifiers, self.virtualkeys):
            return True
        else:
            return False
  
    def unregister(self, hwnd):
        """
        Unregisters the hotkeys that are created on initialization
        """
  
        ctypes.windll.user32.UnregisterHotKey(hwnd, self.keyId)