import keyhandler
import windowmanager

import sys
import threading
import time

import ctypes
import win32con
import win32gui


class EntryPoint():
    def __init__(self, manager, mod_key=keyhandler.NOCONVERT):
        self.mod_key = mod_key
        self.manager = manager
        self.flag = True

    def print_event(self, e, mod_flag):
        if e.event_type == 'key down':
            if mod_flag & self.mod_key:
                if mod_flag & keyhandler.SHIFT:
                    for i in range(manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            manager.send_to_nth_ws(i)
                    if e.key_code == ord('C'):
                        manager.close_window()
                    elif e.key_code == ord('Q'):
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
                        self.flag = False
                    elif e.key_code == ord('J'):
                        manager.swap_windows(+1)
                    elif e.key_code == ord('K'):
                        manager.swap_windows(-1)
                else:
                    for i in range(manager.workspace_n):
                        if e.key_code == ord(str(i+1)):
                            manager.switch_to_nth_ws(i)
                    if e.key_code == 0x0D:    #enter
                        manager.swap_master()
                    elif e.key_code == 0xBC:    #comma
                        manager.inc_master_n(+1)
                    elif e.key_code == 0xBE:    #period
                        manager.inc_master_n(-1)
                    elif e.key_code == 0x09:    #tab
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
                    elif e.key_code == 0x20:    #space
                        manager.next_layout()
                    elif e.key_code == ord('D'):
                        manager.toggle_caption()
                    elif e.key_code == ord('A'):
                        manager.arrange_windows()
                    elif e.key_code == ord('Q'):
                        manager.reset_windows()
                    elif e.key_code == ord('N'):
                        manager.show_network_info()
        # print(e)

if __name__ == '__main__':
    manager = windowmanager.WindowManager(
            ignore_list=[{"class_name": "Windows.UI.Core.CoreWindow"}, {"class_name": "TaskManagerWindow"}, {"class_name": "Microsoft-Windows-SnipperToolbar"}, {"class_name": "Qt5QWindowIcon", "title": "GtransWeb"}],
            workspace_n=9
            )

    mod_key = keyhandler.NOCONVERT
    entry = EntryPoint(manager, mod_key)
    handle = keyhandler.KeyHandler(mod_key)
    handle.handlers.append(entry.print_event)

    hook_thread = threading.Thread(target=handle.listen)
    hook_thread.start()
    counter = 0
    while entry.flag:
        if manager.watch_dog():
            manager.arrange_windows()
        if not manager.is_window_info_on:
            manager.show_system_info()
        time.sleep(0.5)

    hook_thread.join()
    manager.thr.join()
    win32gui.MessageBox(
            None,
            "window manager exit",
            "window information",
            win32con.MB_ICONEXCLAMATION | win32con.MB_OK
            )
