import win32gui
import win32process
import win32con

class WindowManager:
    """Encapsulates some calls to the winapi for window management"""
    def __init__ (self, hwnd = None):
        self._handle = hwnd

    def set_handle(self, hwnd):
        self._handle = hwnd

    def get_handle(self):
        return self._handle

    def set_foreground(self):
        """put the window in the foreground"""
        if self._handle is not None:
            win32gui.ShowWindow(self._handle, True)
            win32gui.SetForegroundWindow(self._handle)

    def isRealWindow(self, hWnd):
        '''Return True iff given handler corespond to a real visible window on the desktop.'''
        if not win32gui.IsWindowVisible(hWnd):
            return False
        if win32gui.GetParent(hWnd) != 0:
            return False
        hasNoOwner = win32gui.GetWindow(hWnd, win32con.GW_OWNER) == 0
        lExStyle = win32gui.GetWindowLong(hWnd, win32con.GWL_EXSTYLE)
        if (((lExStyle & win32con.WS_EX_TOOLWINDOW) == 0 and hasNoOwner)
          or ((lExStyle & win32con.WS_EX_APPWINDOW != 0) and not hasNoOwner)):
            if win32gui.GetWindowText(hWnd):
                return True
        return False
  
    def getWindows(self):
        '''
        Return a list of tuples (handler, (width, height)) for each real window.
        '''
        def callback(hWnd, windows):
            if not self.isRealWindow(hWnd):
                return
            text = win32gui.GetWindowText(hWnd)
            windows.append((hWnd, text))
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows