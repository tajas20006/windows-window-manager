#!/usr/bin/python
# -*- coding: utf-8 -*-
import math
import time
import json

import win32con
import win32gui
import win32process
from win32api import GetAsyncKeyState

import wmi

class WindowManager():
    def __init__(self, title="WindowManager", max_main=1, num_stacks=2,
                    y_offset=40, ignore_list=[]):
        # logger.debug('new manager is created')
        print('debug: new manager is created')

        self.c = wmi.WMI()

        dt_l,dt_t,dt_r,dt_b =\
                        win32gui.GetWindowRect(win32gui.GetDesktopWindow())
        self.screen_width = dt_r-dt_l
        self.screen_height = dt_b-dt_t

        self.windows = set()
        self.window_stack = []
        self.cur_stack_idx = 0
        self.max_main = max_main
        self.num_stacks = num_stacks
        for i in range(self.num_stacks):
            self.window_stack.append(list())

        self.cur_idx = 0
        self.cur_win = 0

        self.y_offset = y_offset
        self.ignore_list = ["Windows.UI.Core.CoreWindow",
                                "ApplicationFrameWindow"] + ignore_list

    def _isRealWindow(self, hwnd):
        '''Return True iff given window is a real Windows application window.'''
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if win32gui.GetParent(hwnd) != 0:
            return False
        hasNoOwner = win32gui.GetWindow(hwnd, win32con.GW_OWNER) == 0
        lExStyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if (((lExStyle & win32con.WS_EX_TOOLWINDOW) == 0 and hasNoOwner)
            or ((lExStyle & win32con.WS_EX_APPWINDOW != 0) and not hasNoOwner)):
            if win32gui.GetWindowText(hwnd):
                return True
        return False

    def _getWindows(self):
        def callback(hwnd, windows):
            if not self._isRealWindow(hwnd):
                return
            class_name = win32gui.GetClassName(hwnd)
            for ignore_item in self.ignore_list:
                if class_name == ignore_item:
                    return
            window = {}
            window['hwnd'] = hwnd
            window['class_name'] = class_name
            window['title'] = win32gui.GetWindowText(hwnd)

            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                for p in self.c.query('SELECT Name FROM Win32_Process WHERE\
                                        ProcessId = %s' % str(pid)):
                    window['name'] = p.Name
            except:
                window['name'] = ""
            window_string = json.dumps(window, sort_keys=True,
                                            ensure_ascii=False)
            windows.add(window_string)

        windows = set()
        win32gui.EnumWindows(callback, windows)
        return windows

    def _get_element(self, window_string, dict_key):
        window = json.loads(window_string)
        return window[dict_key]

    def manage_window_stack(self):
        old_windows = self.windows
        self.windows = self._getWindows()

        if old_windows >= self.windows:
            # some window(s) closed
            diffs = old_windows - self.windows
            for diff in diffs:
                for i in range(self.num_stacks):
                    if diff in self.window_stack[i]:
                        self.window_stack[i].remove(diff)
        else:
            # some window(s) opened
            diffs = self.windows - old_windows
            for diff in diffs:
                self.window_stack[self.cur_stack_idx].insert(0, diff)

    def move_n_resize(self):
        self.manage_window_stack()
        cur_stack = self.window_stack[self.cur_stack_idx]
        print(cur_stack)
        num_win = len(cur_stack)
        if num_win is 0:
            return
        elif num_win <= self.max_main:
            win_h = math.floor((self.screen_height-self.y_offset)/num_win)
            win_w = self.screen_width

            for i in range(num_win):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(cur_stack[i], 'hwnd'),
                            win32con.HWND_TOPMOST,
                            0, self.y_offset + win_h*i, win_w, win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + cur_stack[i])


        else:
            win_w = math.floor(self.screen_width/2)
            main_win_h = math.floor((self.screen_height-self.y_offset)/
                                                            (self.max_main))
            sub_win_h = math.floor((self.screen_height-self.y_offset)/
                                                    (num_win-self.max_main))

            for i in range(self.max_main):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(cur_stack[i], 'hwnd'),
                            win32con.HWND_TOPMOST,
                            0, self.y_offset + main_win_h*i, win_w, main_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + cur_stack[i])

            for i in range(num_win-self.max_main):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(cur_stack[i+self.max_main],
                                                'hwnd'),
                            win32con.HWND_TOPMOST,
                            win_w, self.y_offset+sub_win_h*i, win_w, sub_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + cur_stack[i])

    def close_active_window(self):
        hwnd = win32gui.GetForegroundWindow()
        self.close_window(hwnd)

    def close_window(self, hwnd):
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE,0,0)
        time.sleep(0.5)
        self.move_n_resize()

    def focus_next(self):
        cur_stack = self.window_stack[self.cur_stack_idx]
        self.cur_idx = (self.cur_idx + 1) % len(cur_stack)
        self.cur_win = cur_stack[self.cur_idx]
        try:
            win32gui.SetForegroundWindow(self._get_element(self.cur_win,
                                                            'hwnd'))
            print("debug: focus_next: " + str(self.cur_idx))
        except Exception:
            print ("error: SetForegroundWindow" + self.cur_win)

    def switch_to_nth_stack(self, stack_idx):
        self.cur_stack_idx = stack_idx
        print("debug: switch_to_nth_stack: " + str(stack_idx)
                                    + ":" + str(self.cur_stack_idx))
        self.move_n_resize()

    def send_active_window_to_nth_stack(self, stack_idx):
        cur_stack = self.window_stack[self.cur_stack_idx]
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for window in cur_stack:
            if hwnd == self._get_element(window, 'hwnd'):
                print("found hwnd")
                target_window = cur_stack.index(window)

        if target_window == -1:
            pass
        else:
            self.window_stack[stack_idx].append(cur_stack[target_window])
            self.cur_stack.pop[target_window]
            print("debug: send_to_nth_stack: " + str(stack_idx))
            self.move_n_resize()

    def change_main_stack(self, num):
        self.max_main += num
        if self.max_main <= 0:
            self.max_main = 1
        print("debug: change_main_stack: " + str(num))
        self.move_n_resize()


def isKeyDown(key):
    state = GetAsyncKeyState(key)
    if state != 0 and state != 0:
        return True
    else:
        return False

if __name__ == '__main__':
    print('INFO: Start window manager')
    manager = WindowManager(max_main=1, ignore_list=['TaskManagerWindow'])

    while True:
        manager.move_n_resize()
        ##########
        # CONFIG #
        ##########
        if isKeyDown(win32con.VK_NONCONVERT):
            if isKeyDown(ord('Q')):
                break
            elif isKeyDown(win32con.VK_TAB):
                manager.focus_next()
            elif isKeyDown(ord('C')):
                manager.close_active_window()
            elif isKeyDown(ord('A')):
                manager.send_active_window_to_nth_stack(1)
            elif isKeyDown(ord('S')):
                manager.send_active_window_to_nth_stack(0)
            elif isKeyDown(ord('Z')):
                manager.switch_to_nth_stack(1)
            elif isKeyDown(ord('X')):
                manager.switch_to_nth_stack(0)
            elif isKeyDown(0xBC):       # comma
                manager.change_main_stack(+1)
            elif isKeyDown(0xBE):       # period
                manager.change_main_stack(-1)
