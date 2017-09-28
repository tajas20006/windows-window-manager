# based on https://gist.github.com/ethanhs/80f0a7cc5c7881f5921f

import constants as C
from ctypes import windll, CFUNCTYPE, POINTER, c_int, c_void_p, byref
import win32con
import win32api
import win32gui
import atexit
from collections import namedtuple

KeyboardEvent = namedtuple('KeyboardEvent', ['event_type', 'key_code'])

class KeyHandler:
    def __init__(self, mod_key=C.NOCONVERT):
        self.handlers = []
        self.mod_flag = 0
        self.mod_key = mod_key

    def listen(self):
        """
        Calls `handlers` for each keyboard event received.
        This is a blocking call.
        """
        # Adapted from http://www.hackerthreads.org/Topic-42395

        event_types = {
                win32con.WM_KEYDOWN: 'key down',
                win32con.WM_KEYUP: 'key up',
                win32con.WM_SYSKEYDOWN: 'key down',
                win32con.WM_SYSKEYUP: 'key up'
                }

        def low_level_handler(nCode, wParam, lParam):
            """
            Processes a low level Windows keyboard event.
            """
            key_code = lParam[0] & 0xFFFFFFFF
            event = KeyboardEvent(event_types[wParam], key_code)

            if key_code == 29:     # noconvert down
                if event_types[wParam] == "key down":
                    self.mod_flag |= C.NOCONVERT
                elif event_types[wParam] == "key up":
                    self.mod_flag -= C.NOCONVERT
            elif key_code == 164 or key_code == 165:    # alt
                if event_types[wParam] == "key down":
                    self.mod_flag |= C.ALT
                elif event_types[wParam] == "key up":
                    self.mod_flag -= C.ALT
            elif key_code == 91:    # win
                if event_types[wParam] == "key down":
                    self.mod_flag |= C.WIN
                elif event_types[wParam] == "key up":
                    self.mod_flag -= C.WIN
            elif key_code == 162 or key_code == 163:    # ctrl
                if event_types[wParam] == "key down":
                    self.mod_flag |= C.CTRL
                elif event_types[wParam] == "key up":
                    self.mod_flag -= C.CTRL
            elif key_code == 160 or key_code == 161:    # shift
                if event_types[wParam] == "key down":
                    self.mod_flag |= C.SHIFT
                elif event_types[wParam] == "key up":
                    self.mod_flag -= C.SHIFT

            if self.mod_flag & self.mod_key:
                for handler in self.handlers:
                    handler(event, self.mod_flag)
                # capture the key press
                return 1
            else:
                # let the key press go
                return windll.user32.CallNextHookEx(hook_id, nCode,
                                                    wParam, lParam)

        # Our low level handler signature.
        CMPFUNC = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))
        # Convert the Python handler into C pointer.
        pointer = CMPFUNC(low_level_handler)

        # Hook both key up and key down events for common keys (non-system).
        hook_id = windll.user32.SetWindowsHookExA(
                win32con.WH_KEYBOARD_LL,
                pointer,
                win32api.GetModuleHandle(None),
                0
                )

        # Register to remove the hook when the interpreter exits.
        # Unfortunately a try/finally block doesn't seem to work here.
        atexit.register(windll.user32.UnhookWindowsHookEx, hook_id)

        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            try:
                win32gui.TranslateMessage(byref(msg))
                win32gui.DispatchMessage(byref(msg))
            except Exception:
                print("exit keyhandle")
                break
