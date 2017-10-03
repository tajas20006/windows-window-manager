import keyhandler
import windowhandler
import windowmanager
import constants as C
try:
    import config
except ModuleNotFoundError:
    print("no config file found. using default")

import sys
import threading
import time

import ctypes
import win32con
import win32gui


class EntryPoint():
    def __init__(self):
        try:
            self.mod_key = config.mod_key
        except (AttributeError, NameError):
            self.mod_key = C.NOCONVERT
        try:
            ignore_list = config.ignore_list
        except (AttributeError, NameError):
            ignore_list = []
        try:
            workspace_n = config.workspace_n
        except (AttributeError, NameError):
            workspace_n = 2
        try:
            network_interface = config.network_interface
        except (AttributeError, NameError):
            network_interface = ""

        self.manager = windowmanager.WindowManager(
                ignore_list=ignore_list,
                workspace_n=workspace_n,
                network_interface=network_interface
                )

        self.config_list = [
                {"mod": self.mod_key | C.SHIFT, "key": "Q",
                    "function": self.kill_program}
                ]
        try:
            for a in config.config_list:
                a["function"] = getattr(self.manager, a["function"], None)
            self.config_list.extend(config.config_list)
        except (AttributeError, NameError):
            self.set_default_config()
            try:
                self.remove_config(config.remove_list)
            except (AttributeError, NameError):
                pass
            try:
                self.append_config(config.append_list)
            except (AttributeError, NameError):
                pass

    def set_default_config(self):
        self.config_list.extend([
                {"mod": self.mod_key | C.SHIFT, "key": "Q",
                    "function": self.kill_program},
                {"mod": self.mod_key | C.SHIFT, "key": "C",
                    "function": self.manager.close_window},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "J",
                    "function": self.manager.swap_windows, "param": [+1]},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "K",
                    "function": self.manager.swap_windows, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.ENTER,
                    "function": self.manager.swap_master},
                {"mod": C.NOCONVERT, "key": C.COMMA,
                    "function": self.manager.inc_master_n, "param": [+1]},
                {"mod": C.NOCONVERT, "key": C.PERIOD,
                    "function": self.manager.inc_master_n, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.TAB,
                    "function": self.manager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "J",
                    "function": self.manager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "K",
                    "function": self.manager.focus_up, "param": [-1]},
                {"mod": C.NOCONVERT, "key": "I",
                    "function": self.manager.show_window_info},
                {"mod": C.NOCONVERT, "key": "H",
                    "function": self.manager.expand_master, "param": [-20]},
                {"mod": C.NOCONVERT, "key": "L",
                    "function": self.manager.expand_master, "param": [+20]},
                {"mod": C.NOCONVERT, "key": C.SPACE,
                    "function": self.manager.next_layout},
                {"mod": C.NOCONVERT, "key": "D",
                    "function": self.manager.toggle_caption},
                {"mod": C.NOCONVERT, "key": "A",
                    "function": self.manager.arrange_windows},
                {"mod": C.NOCONVERT, "key": "Q",
                    "function": self.manager.reset_windows}
                ])
        for i in range(self.manager.workspace_n):
            self.config_list.append({
                "mod": C.NOCONVERT | C.SHIFT,
                "key": str(i+1),
                "function": self.manager.send_to_nth_ws, "param": [i]
                })
            self.config_list.append({"mod": C.NOCONVERT, "key": str(i+1),
                "function": self.manager.switch_to_nth_ws, "param": [i]})

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
        for a in append_list:
            a["function"] = getattr(self.manager, a["function"], None)
        self.config_list.extend(append_list)

    def print_event(self, e, mod_flag):
        if e.event_type == "key down":
            for conf in self.config_list:
                try:
                    if e.key_code == ord(conf["key"].upper()) and\
                            mod_flag == conf['mod']:
                        if "param" in conf:
                            conf["function"](*conf["param"])
                        else:
                            conf["function"]()
                        return
                except AttributeError:
                    if e.key_code == conf["key"] and\
                            mod_flag == conf['mod']:
                        if "param" in conf:
                            conf["function"](*conf["param"])
                        else:
                            conf["function"]()
                        return

    def kill_program(self):
        self.manager.recover_windows()
        ctypes.windll.user32.PostThreadMessageW(
                hook_thread.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                self.manager.thr.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                window_thread.ident,
                win32con.WM_QUIT,
                0, 0,
                )

    def print_event_orig(self, e, mod_flag):
        if e.event_type == 'key down':
            if mod_flag & self.mod_key:
                if mod_flag & C.SHIFT:
                    for i in range(self.manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            self.manager.send_to_nth_ws(i)
                    if e.key_code == ord('C'):
                        self.manager.close_window()
                    elif e.key_code == ord('Q'):
                        self.kill_program()
                    elif e.key_code == ord('J'):
                        self.manager.swap_windows(+1)
                    elif e.key_code == ord('K'):
                        self.manager.swap_windows(-1)
                else:
                    for i in range(self.manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            self.manager.switch_to_nth_ws(i)
                    if e.key_code == 0x0D:    # enter
                        self.manager.swap_master()
                    elif e.key_code == 0xBC:    # comma
                        self.manager.inc_master_n(+1)
                    elif e.key_code == 0xBE:    # period
                        self.manager.inc_master_n(-1)
                    elif e.key_code == 0x09:    # tab
                        self.manager.focus_up(+1)
                    elif e.key_code == ord('J'):
                        self.manager.focus_up(+1)
                    elif e.key_code == ord('K'):
                        self.manager.focus_up(-1)
                    elif e.key_code == ord('I'):
                        self.manager.show_window_info()
                    elif e.key_code == ord('H'):
                        self.manager.expand_master(-20)
                    elif e.key_code == ord('L'):
                        self.manager.expand_master(+20)
                    elif e.key_code == 0x20:    # space
                        self.manager.next_layout()
                    elif e.key_code == ord('D'):
                        self.manager.toggle_caption()
                    elif e.key_code == ord('A'):
                        self.manager.arrange_windows()
                    elif e.key_code == ord('Q'):
                        self.manager.reset_windows()

    def window_changed(self):
        if self.manager.watch_dog():
            self.manager.arrange_windows()

if __name__ == '__main__':

    entry = EntryPoint()
    key_handler = keyhandler.KeyHandler(entry.mod_key)
    key_handler.handlers.append(entry.print_event)
    window_handler = windowhandler.WindowHandler(entry.window_changed)
    window_thread = threading.Thread(target=window_handler.main)
    window_thread.start()

    hook_thread = threading.Thread(target=key_handler.listen)
    hook_thread.start()
    try:

        while True:
            entry.manager.show_system_info()
            entry.manager._update_taskbar_text()
            time.sleep(0.5)
    except:
        hook_thread.join()
        window_thread.join()
        entry.manager.thr.join()
        win32gui.MessageBox(
                None,
                "window manager exit",
                "window information",
                win32con.MB_ICONEXCLAMATION | win32con.MB_OK
                )

