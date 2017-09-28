import keyhandler
import windowhandler
import windowmanager
import constants as C
try:
    import config
    exists_config = True
except ModuleNotFoundError:
    exists_config = False


import sys
import threading
import time

import ctypes
import win32con
import win32gui


class EntryPoint():
    def __init__(self, mod_key=C.NOCONVERT):
        self.mod_key = mod_key
        self.flag = True
        self.config_list = []
        self.set_default_config()
        self.modify_config_list()

    def set_default_config(self):
        self.config_list = [
                {"mod": self.mod_key | C.SHIFT, "key": "Q",
                    "function": self.kill_program},
                {"mod": self.mod_key | C.SHIFT, "key": "C",
                    "function": windowmanager.close_window},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "J",
                    "function": windowmanager.swap_windows, "param": [+1]},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "K",
                    "function": windowmanager.swap_windows, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.ENTER,
                    "function": windowmanager.swap_master},
                {"mod": C.NOCONVERT, "key": C.COMMA,
                    "function": windowmanager.inc_master_n, "param": [+1]},
                {"mod": C.NOCONVERT, "key": C.PERIOD,
                    "function": windowmanager.inc_master_n, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.TAB,
                    "function": windowmanager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "J",
                    "function": windowmanager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "K",
                    "function": windowmanager.focus_up, "param": [-1]},
                {"mod": C.NOCONVERT, "key": "I",
                    "function": windowmanager.show_window_info},
                {"mod": C.NOCONVERT, "key": "H",
                    "function": windowmanager.expand_master, "param": [-20]},
                {"mod": C.NOCONVERT, "key": "L",
                    "function": windowmanager.expand_master, "param": [+20]},
                {"mod": C.NOCONVERT, "key": C.SPACE,
                    "function": windowmanager.next_layout},
                {"mod": C.NOCONVERT, "key": "D",
                    "function": windowmanager.toggle_caption},
                {"mod": C.NOCONVERT, "key": "A",
                    "function": windowmanager.arrange_windows},
                {"mod": C.NOCONVERT, "key": "Q",
                    "function": windowmanager.reset_windows}
                ]
        for i in range(windowmanager.workspace_n):
            self.config_list.append({
                "mod": C.NOCONVERT | C.SHIFT,
                "key": str(i+1),
                "function": windowmanager.send_to_nth_ws, "param": [i]
                })
            self.config_list.append({"mod": C.NOCONVERT, "key": str(i+1),
                "function": windowmanager.switch_to_nth_ws, "param": [i]})

    def remove_config(self, remove_list):
        r_list = []
        for remove in remove_list:
            for conf in self.config_list:
                if remove["key"] == conf["key"] and\
                        remove["mod"] == conf["mod"]:
                    r_list.append(conf)
        for conf in r_list:
            self.config_list.remove(conf)

    def append_config(self, append_list):
        self.config_list.extend(append_list)

    def modify_config_list(self):
        if exists_config:
            try:
                self.config_list = config.config_list
            except AttributeError:
                try:
                    self.remove_config(config.remove_list)
                    self.append_config(config.append_list)
                except AttributeError:
                    pass
        else:
            print("config file not found. using default")

    def print_event(self, e, mod_flag):
        if e.event_type == "key down":
            for conf in self.config_list:
                try:
                    if e.key_code == ord(conf["key"].upper()) and\
                            mod_flag == conf['mod']:
                        print(e.key_code)
                        if "param" in conf:
                            conf["function"](*conf["param"])
                        else:
                            conf["function"]()
                        return
                except AttributeError:
                    if e.key_code == conf["key"] and\
                            mod_flag == conf['mod']:
                        print(e.key_code)
                        if "param" in conf:
                            conf["function"](*conf["param"])
                        else:
                            conf["function"]()
                        return

    def kill_program(self):
        windowmanager.recover_windows()
        ctypes.windll.user32.PostThreadMessageW(
                hook_thread.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                windowmanager.thr.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                window_thread.ident,
                win32con.WM_QUIT,
                0, 0,
                )
        self.flag = False

    def window_changed(self):
        print("hello")
        if windowmanager.watch_dog():
            windowmanager.arrange_windows()

if __name__ == '__main__':
    mod_key = C.NOCONVERT
    entry = EntryPoint(mod_key)
    key_handler = keyhandler.KeyHandler(mod_key)
    key_handler.handlers.append(entry.print_event)
    window_handler = windowhandler.WindowHandler(entry.window_changed)
    window_thread = threading.Thread(target=window_handler.main)
    window_thread.start()

    hook_thread = threading.Thread(target=key_handler.listen)
    hook_thread.start()

    try:
        windomanager.first_to_do()
        while True:
            windowmanager.show_system_info()
            windowmanager._update_taskbar_text()
            time.sleep(0.5)
    except:
        hook_thread.join()
        window_thread.join()
        windowmanager.thr.join()
        win32gui.MessageBox(
                None,
                "window manager exit",
                "window information",
                win32con.MB_ICONEXCLAMATION | win32con.MB_OK
                )

