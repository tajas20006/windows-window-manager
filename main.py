import keyhandler
import windowhandler
import windowmanager
import constants as C


import sys
import threading
import time

import ctypes
import win32con
import win32gui


class EntryPoint():
    def __init__(self, manager, mod_key=C.NOCONVERT):
        self.mod_key = mod_key
        self.manager = manager
        self.flag = True
        self.config_list = []
        self.set_default_config()

        try:
            import config
            try:
                self.config_list = config.config_list
            except NameError:
                try:
                    self.remove_config(config.remove_list)
                    self.append_config(config.append_list)
                except NameError:
                    pass
        except ModuleNotFoundError:
            print("config file not found. using default")

    def set_default_config(self):
        self.config_list = [
                {"mod": self.mod_key | C.SHIFT, "key": "Q",
                    "function": self.kill_program},
                {"mod": self.mod_key | C.SHIFT, "key": "C",
                    "function": manager.close_window},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "J",
                    "function": manager.swap_windows, "param": [+1]},
                {"mod": C.NOCONVERT | C.SHIFT, "key": "K",
                    "function": manager.swap_windows, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.ENTER,
                    "function": manager.swap_master},
                {"mod": C.NOCONVERT, "key": C.COMMA,
                    "function": manager.inc_master_n, "param": [+1]},
                {"mod": C.NOCONVERT, "key": C.PERIOD,
                    "function": manager.inc_master_n, "param": [-1]},
                {"mod": C.NOCONVERT, "key": C.TAB,
                    "function": manager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "J",
                    "function": manager.focus_up, "param": [+1]},
                {"mod": C.NOCONVERT, "key": "K",
                    "function": manager.focus_up, "param": [-1]},
                {"mod": C.NOCONVERT, "key": "I",
                    "function": manager.show_window_info},
                {"mod": C.NOCONVERT, "key": "H",
                    "function": manager.expand_master, "param": [-20]},
                {"mod": C.NOCONVERT, "key": "L",
                    "function": manager.expand_master, "param": [+20]},
                {"mod": C.NOCONVERT, "key": C.SPACE,
                    "function": manager.next_layout},
                {"mod": C.NOCONVERT, "key": "D",
                    "function": manager.toggle_caption},
                {"mod": C.NOCONVERT, "key": "A",
                    "function": manager.arrange_windows},
                {"mod": C.NOCONVERT, "key": "Q",
                    "function": manager.reset_windows}
                ]
        for i in range(manager.workspace_n):
            self.config_list.append({
                "mod": C.NOCONVERT | C.SHIFT,
                "key": str(i+1),
                "function": manager.send_to_nth_ws, "param": [i]
                })
            self.config_list.append({"mod": C.NOCONVERT, "key": str(i+1),
                "function": manager.switch_to_nth_ws, "param": [i]})

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
                except AttributeError:
                    if e.key_code == conf["key"] and\
                            mod_flag == conf['mod']:
                        if "param" in conf:
                            conf["function"](*conf["param"])
                        else:
                            conf["function"]()

    def kill_program(self):
        manager.recover_windows()
        ctypes.windll.user32.PostThreadMessageW(
                hook_thread.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                manager.thr.ident,
                win32con.WM_QUIT,
                0, 0
                )
        ctypes.windll.user32.PostThreadMessageW(
                window_thread.ident,
                win32con.WM_QUIT,
                0, 0,
                )
        self.flag = False

    def print_event_orig(self, e, mod_flag):
        if e.event_type == 'key down':
            if mod_flag & self.mod_key:
                if mod_flag & C.SHIFT:
                    for i in range(manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            manager.send_to_nth_ws(i)
                    if e.key_code == ord('C'):
                        manager.close_window()
                    elif e.key_code == ord('Q'):
                        self.kill_program()
                    elif e.key_code == ord('J'):
                        manager.swap_windows(+1)
                    elif e.key_code == ord('K'):
                        manager.swap_windows(-1)
                else:
                    for i in range(manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            manager.switch_to_nth_ws(i)
                    if e.key_code == 0x0D:    # enter
                        manager.swap_master()
                    elif e.key_code == 0xBC:    # comma
                        manager.inc_master_n(+1)
                    elif e.key_code == 0xBE:    # period
                        manager.inc_master_n(-1)
                    elif e.key_code == 0x09:    # tab
                        manager.focus_up(+1)
                    elif e.key_code == ord('J'):
                        manager.focus_up(+1)
                    elif e.key_code == ord('K'):
                        manager.focus_up(-1)
                    elif e.key_code == ord('I'):
                        manager.show_window_info()
                    elif e.key_code == ord('H'):
                        manager.expand_master(-20)
                    elif e.key_code == ord('L'):
                        manager.expand_master(+20)
                    elif e.key_code == 0x20:    # space
                        manager.next_layout()
                    elif e.key_code == ord('D'):
                        manager.toggle_caption()
                    elif e.key_code == ord('A'):
                        manager.arrange_windows()
                    elif e.key_code == ord('Q'):
                        manager.reset_windows()

    def window_changed(self):
        print("hello")
        if manager.watch_dog():
            manager.arrange_windows()

if __name__ == '__main__':
    manager = windowmanager.WindowManager(
            ignore_list=[
                {"class_name": "Windows.UI.Core.CoreWindow"},
                {"class_name": "TaskManagerWindow"},
                {"class_name": "Microsoft-Windows-SnipperToolbar"},
                {"class_name": "Qt5QWindowIcon", "title": "GtransWeb"},
                {"class_name": "screenClass"}
                ],
            workspace_n=9,
            network_interface="Dell Wireless 1820A 802.11ac"
            )

    mod_key = C.NOCONVERT
    entry = EntryPoint(manager, mod_key)
    key_handler = keyhandler.KeyHandler(mod_key)
    key_handler.handlers.append(entry.print_event)
    window_handler = windowhandler.WindowHandler(entry.window_changed)
    window_thread = threading.Thread(target=window_handler.main)
    window_thread.start()

    hook_thread = threading.Thread(target=key_handler.listen)
    hook_thread.start()
    counter = 0
        # if manager.watch_dog():
        #     manager.arrange_windows()
    try:

        while True:
            manager.show_system_info()
            manager._update_taskbar_text()
            time.sleep(0.5)
    except:
        hook_thread.join()
        window_thread.join()
        manager.thr.join()
        win32gui.MessageBox(
                None,
                "window manager exit",
                "window information",
                win32con.MB_ICONEXCLAMATION | win32con.MB_OK
                )

