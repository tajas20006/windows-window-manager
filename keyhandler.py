# based on https://gist.github.com/ethanhs/80f0a7cc5c7881f5921f
#

from collections import namedtuple

KeyboardEvent = namedtuple('KeyboardEvent', ['event_type', 'key_code'])

class KeyHandler:
    def __init__(self):
        self.handlers = []
        self.mod_flag = 0
        self.noconvert = 1
        self.win = 2
        self.ctrl = 4
        self.shift = 8
        self.alt = 16

    def listen(self):
        """
        Calls `handlers` for each keyboard event received.
        This is a blocking call.
        """
        # Adapted from http://www.hackerthreads.org/Topic-42395
        from ctypes import windll, CFUNCTYPE, POINTER, c_int, c_void_p, byref
        import win32con, win32api, win32gui, atexit

        event_types = {win32con.WM_KEYDOWN: 'key down',
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
                    self.mod_flag |= self.noconvert
                elif event_types[wParam] == "key up":
                    self.mod_flag -= self.noconvert
            elif key_code == 164 or key_code == 165:    # alt
                if event_types[wParam] == "key down":
                    self.mod_flag |= self.alt
                elif event_types[wParam] == "key up":
                    self.mod_flag -= self.alt
            elif key_code == 91:    # win
                if event_types[wParam] == "key down":
                    self.mod_flag |= self.win
                elif event_types[wParam] == "key up":
                    self.mod_flag -= self.win
            elif key_code == 162 or key_code == 163:    # ctrl
                if event_types[wParam] == "key down":
                    self.mod_flag |= self.ctrl
                elif event_types[wParam] == "key up":
                    self.mod_flag -= self.ctrl
            elif key_code == 160 or key_code == 161:    # shift
                if event_types[wParam] == "key down":
                    self.mod_flag |= self.shift
                elif event_types[wParam] == "key up":
                    self.mod_flag -= self.shift

            if self.mod_flag & self.noconvert:
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
        hook_id = windll.user32.SetWindowsHookExA(win32con.WH_KEYBOARD_LL,
                                                pointer,
                                                win32api.GetModuleHandle(None),
                                                0
                                                )

        # Register to remove the hook when the interpreter exits.
        # Unfortunately a try/finally block doesn't seem to work here.
        atexit.register(windll.user32.UnhookWindowsHookEx, hook_id)

        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            print(msg)
            try:
                win32gui.TranslateMessage(byref(msg))
                win32gui.DispatchMessage(byref(msg))
            except Exception:
                print("exit keyhandle")
                break

if __name__ == '__main__':
    def print_event(e, mod_flag):
        print(e)

    handle = KeyHandler()
    handle.handlers.append(print_event)
    handle.listen()
