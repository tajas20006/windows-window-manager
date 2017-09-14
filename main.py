import keyhandler
# import window_manager

noconvert = 1
win = 2
ctrl = 4
shift = 8
alt = 16

if __name__ == '__main__':
    def print_event(e, mod_flag):
        if mod_flag == noconvert + shift:
            if e.key_code == ord('A') and e.event_type == "key down":
                print ("noconv shift a down")
        print(e)

    handle = keyhandler.KeyHandler()
    handle.handlers.append(print_event)
    handle.listen()