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
    def __init__(self, title="WindowManager", max_main=1, y_offset=40, ignore_list=[]):
        # logger.debug('new manager is created')
        print('debug: new manager is created')

        self.c = wmi.WMI()

        dt_l,dt_t,dt_r,dt_b = win32gui.GetWindowRect(win32gui.GetDesktopWindow())
        self.screen_width = dt_r-dt_l
        self.screen_height = dt_b-dt_t

        self.windows = set()
        self.window_stack= []
        self.max_main = max_main

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
            window_string = json.dumps(window, sort_keys=True, ensure_ascii=False)
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

        if len(old_windows) > len(self.windows):
            # some window(s) closed
            diffs = old_windows - self.windows
            for diff in diffs:
                if diff in self.window_stack:
                    self.window_stack.remove(diff)
        else:
            # some window(s) opened
            diffs = self.windows - old_windows
            for diff in diffs:
                self.window_stack.insert(0, diff)

    def move_n_resize(self):
        self.manage_window_stack()
        num_win = len(self.window_stack)
        if num_win is 0:
            return
        elif num_win < self.max_main:
            win_h = math.floor((self.screen_height-self.y_offset)/num_win)
            win_w = self.screen_width

            for i in range(num_win):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(self.window_stack[i], 'hwnd'),
                            win32con.HWND_TOPMOST,
                            0, self.y_offset + win_h*i, win_w, win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + self.window_stack[i])


        else:
            win_w = math.floor(self.screen_width/2)
            main_win_h = math.floor((self.screen_height-self.y_offset)/
                                                            (self.max_main))
            sub_win_h = math.floor((self.screen_height-self.y_offset)/
                                                    (num_win-self.max_main))

            for i in range(self.max_main):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(self.window_stack[i], 'hwnd'),
                            win32con.HWND_TOPMOST,
                            0, self.y_offset + main_win_h*i, win_w, main_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + self.window_stack[i])

            for i in range(num_win-self.max_main):
                try:
                    win32gui.SetWindowPos(
                            self._get_element(self.window_stack[i+self.max_main],
                                                'hwnd'),
                            win32con.HWND_TOPMOST,
                            win_w, self.y_offset+sub_win_h*i, win_w, sub_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + self.window_stack[i])


    def close_active_window(self):
        hwnd = win32gui.GetForegroundWindow()
        self.close_window(hwnd)

    def close_window(self, hwnd):
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE,0,0)
        time.sleep(0.5)
        self.move_n_resize()

    def focus_next(self):
        self.cur_idx = (self.cur_idx + 1) % len(self.window_stack)
        self.cur_win = self.window_stack[self.cur_idx]
        try:
            win32gui.SetForegroundWindow(self._get_element(self.cur_win, 'hwnd'))
        except Exception:
            print (Exception)
            print (self.cur_win)

def isKeyDown(key):
    state = GetAsyncKeyState(key)
    if state != 0 and state != 0:
        return True
    else:
        return False

if __name__ == '__main__':
    print('INFO: Start window manager')
    manager = WindowManager(max_main=1)

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

        # time.sleep(1)
