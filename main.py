import keyhandler
import window_manager
import sys

class EntryPoint():
    def __init__(self, manager):
        self.noconvert = 1
        self.win = 2
        self.ctrl = 4
        self.shift = 8
        self.alt = 16
        self.manager = manager

    def print_event(self, e, mod_flag):
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
                    sys.exit(0)
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
        print(e)

if __name__ == '__main__':
    manager = window_manager.WindowManager(ignore_list=['TaskManagerWindow'])
    manager.move_n_resize()

    entry = EntryPoint(manager)
    handle = keyhandler.KeyHandler()
    handle.handlers.append(entry.print_event)
    handle.listen()
