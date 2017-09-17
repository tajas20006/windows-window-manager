import keyhandler
import window_manager
import sys
import threading
import time

import ctypes
import win32con
import win32gui


class EntryPoint():
    def __init__(self, manager):
        self.noconvert = 1
        self.win = 2
        self.ctrl = 4
        self.shift = 8
        self.alt = 16
        self.manager = manager
        self.flag = True

    def print_event(self, e, mod_flag):
        if e.event_type == 'key down':
            if mod_flag & self.noconvert:
                if mod_flag & self.shift:
                    if e.key_code == ord('A'):
                        manager.send_active_window_to_nth_stack(0)
                    elif e.key_code == ord('S'):
                        manager.send_active_window_to_nth_stack(1)
                    elif e.key_code == ord('C'):
                        manager.close_active_window()
                    elif e.key_code == ord('Q'):
                        manager.recover_windows()
                        ctypes.windll.user32.PostThreadMessageW(hookThread.ident, win32con.WM_QUIT, 0, 0)
                        self.flag = False
                else:
                    if e.key_code == ord('A'):
                        manager.switch_to_nth_stack(0)
                    elif e.key_code == ord('S'):
                        manager.switch_to_nth_stack(1)
                    elif e.key_code == 13:    #enter
                        manager.move_active_to_main_stack()
                    elif e.key_code == 0xBC:    #comma
                        manager.change_max_main(+1)
                    elif e.key_code == 0xBE:    #period
                        manager.change_max_main(-1)
                    elif e.key_code == 9:    #tab
                        manager.focus_next()
                    elif e.key_code == ord('I'):
                        manager.show_window_information()
        # print(e)

if __name__ == '__main__':
    manager = window_manager.WindowManager(ignore_list=['TaskManagerWindow'])

    entry = EntryPoint(manager)
    handle = keyhandler.KeyHandler()
    handle.handlers.append(entry.print_event)

    hookThread = threading.Thread(target=handle.listen)
    hookThread.start()

    while entry.flag:
        manager.move_n_resize()
        time.sleep(2)

    hookThread.join()
    win32gui.MessageBox(None, "window manager exit", "window information", win32con.MB_ICONEXCLAMATION | win32con.MB_OK)
