#!/usr/bin/python
# -*- coding: utf-8 -*-
import math
import time

import win32api
import win32con
import win32gui
import win32process

import wmi

class WindowManager():
    def __init__(self, title="WindowManager", max_main=1, num_stacks=2,
                    ignore_list=[]):
        # logger.debug('new manager is created')
        print('debug: new manager is created')

        self.c = wmi.WMI()

        monitors = win32api.EnumDisplayMonitors(None, None)
        (h_first_mon, _, (_,_,_,_)) = monitors[0]
        first_mon = win32api.GetMonitorInfo(h_first_mon)
        m_left, m_top, m_right, m_bottom = first_mon['Monitor']
        w_left, w_top, w_right, w_bottom = first_mon['Work']

        self.monitor_width = m_right - m_left
        self.monitor_height = m_bottom - m_top
        self.work_width = w_right - w_left
        self.work_height = w_bottom - w_top
        self.taskbar_height = self.monitor_height - self.work_height

        self.num_stacks = num_stacks
        self.window_stack = []
        for i in range(self.num_stacks):
            self.window_stack.append(list())
        self.cur_stack_idx = 0

        self.max_main = max_main

        self.cur_idx = 0
        self.cur_win = None

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

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            window['pid'] = pid
            windows.append(window)

        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    def get_process_name(self,pid):
        try:
            for p in self.c.query('SELECT Name FROM Win32_Process WHERE\
                                    ProcessId = %s' % str(pid)):
                return p.Name
        except:
            return ""

    def manage_window_stack(self):
        old_stack = self.window_stack[self.cur_stack_idx]
        new_stack = self._getWindows()

        rm_list_old = []
        rm_list_new = []
        for old in old_stack:
            still_here = False
            for new in new_stack:
                if old['hwnd'] == new['hwnd']:
                    # if same window is still existing
                    rm_list_new.append(new)
                    still_here = True
                    continue
            if not still_here:
                # if the window is closed
                rm_list_old.append(old)

        # rm the closed windows
        for rm_item in rm_list_old:
            old_stack.remove(rm_item)
        # rm the duplicate windows
        for rm_item in rm_list_new:
            new_stack.remove(rm_item)
        print("new: " + str(new_stack))
        print("old: " + str(old_stack))
        # prepend new windows
        self.window_stack[self.cur_stack_idx] = new_stack + old_stack

    def move_n_resize(self):
        self.manage_window_stack()
        cur_stack = self.window_stack[self.cur_stack_idx]
        # print(cur_stack)
        num_win = len(cur_stack)
        if num_win is 0:
            return
        elif num_win <= self.max_main:
            win_h = math.floor((self.work_height)/num_win)
            win_w = self.work_width

            for i in range(num_win):
                try:
                    win32gui.SetWindowPos(
                            cur_stack[i]['hwnd'],
                            win32con.HWND_TOPMOST,
                            0, win_h*i + self.taskbar_height, win_w, win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + str(cur_stack[i]))
                    # # exit(0)

        else:
            win_w = math.floor(self.work_width/2)
            main_win_h = math.floor((self.work_height) / (self.max_main))
            sub_win_h = math.floor((self.work_height) / (num_win-self.max_main))

            for i in range(self.max_main):
                try:
                    win32gui.SetWindowPos(
                            cur_stack[i]['hwnd'],
                            win32con.HWND_TOPMOST,
                            0, main_win_h*i + self.taskbar_height,
                            win_w, main_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + str(cur_stack[i]))
                    # exit(0)

            for i in range(num_win-self.max_main):
                try:
                    win32gui.SetWindowPos(
                            cur_stack[i+self.max_main]['hwnd'],
                            win32con.HWND_TOPMOST,
                            win_w, sub_win_h*i + self.taskbar_height,
                            win_w, sub_win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + str(cur_stack[i]))
                    # exit(0)

    def close_active_window(self):
        hwnd = win32gui.GetForegroundWindow()
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE,0,0)
        time.sleep(0.5)
        self.move_n_resize()

    def focus_next(self, num=1):
        cur_stack = self.window_stack[self.cur_stack_idx]
        self.cur_idx = (self.cur_idx + num) % len(cur_stack)
        self.cur_win = cur_stack[self.cur_idx]
        try:
            win32gui.SetForegroundWindow(self.cur_win['hwnd'])
            print("debug: focus_next: " + str(self.cur_idx))
        except Exception:
            print ("error: SetForegroundWindow" + str(self.cur_win))
            # exit(0)

    def hide_window(self, hwnd):
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            win32gui.SetWindowPos(hwnd,
                    win32con.HWND_TOPMOST,
                    left, top, right, bottom,
                    win32con.SWP_HIDEWINDOW | win32con.SWP_NOZORDER |
                    win32con.SWP_NOMOVE
                    )
        except Exception:
            print("error: SetWindowPos" + str(hwnd))

    def show_window(self, hwnd):
        try:
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            win32gui.SetWindowPos(hwnd,
                    win32con.HWND_TOPMOST,
                    left, top, right, bottom,
                    win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER |
                    win32con.SWP_NOMOVE
                    )
        except Exception:
            print("error: SetWindowPos" + str(hwnd))

    def switch_to_nth_stack(self, stack_idx):
        if self.cur_stack_idx == stack_idx:
            return
        for window in self.window_stack[self.cur_stack_idx]:
            self.hide_window(window['hwnd'])

        self.cur_stack_idx = stack_idx
        for window in self.window_stack[self.cur_stack_idx]:
            self.show_window(window['hwnd'])

        print("debug: switch_to_nth_stack: " + str(stack_idx)
                                    + ":" + str(self.cur_stack_idx))
        self.move_n_resize()

    def send_active_window_to_nth_stack(self, stack_idx):
        if self.cur_stack_idx == stack_idx:
            return
        cur_stack = self.window_stack[self.cur_stack_idx]
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(cur_stack):
            if hwnd == window['hwnd']:
                print(i)
                target_window = i

        if target_window != -1:
            self.window_stack[stack_idx].append(cur_stack[target_window])
            self.hide_window(cur_stack[target_window]['hwnd'])
            cur_stack.pop(target_window)
            print("debug: send_to_nth_stack: " + str(stack_idx))
            self.move_n_resize()

    def move_active_to_main_stack(self):
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(self.window_stack[self.cur_stack_idx]):
            if hwnd == window['hwnd']:
                print(i)
                target_window = i
        if target_window != -1:
            window = self.window_stack[self.cur_stack_idx].pop(target_window)
            self.window_stack[self.cur_stack_idx][:0] = [window]
            self.move_n_resize()

    def change_max_main(self, num=1):
        self.max_main += num
        if self.max_main <= 0:
            self.max_main = 1
        print("debug: change_max_main: " + str(num))
        self.move_n_resize()

    def recover_windows(self):
        for i, stack in enumerate(self.window_stack):
            if i == self.cur_stack_idx:
                continue
            else:
                for window in stack:
                    self.show_window(window['hwnd'])
        self.move_n_resize()


def isKeyDown(key, isAsync=True):
    state = 0
    if isAsync:
        state = win32api.GetAsyncKeyState(key)
    else:
        state = win32api.GetKeyState(key)

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
                manager.recover_windows()
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
                manager.change_max_main(+1)
            elif isKeyDown(0xBE):       # period
                manager.change_max_main(-1)
            elif isKeyDown(win32con.VK_RETURN):
                manager.move_active_to_main_stack()
            else:
                continue
        time.sleep(0.2)
