#!/usr/bin/python
# -*- coding: utf-8 -*-
import math
import time

import win32api
import win32con
import win32gui
import win32process
import win32com.client

import ctypes
from ctypes import wintypes

import wmi

class WindowInfo():
    def __init__(self, hwnd=-1, class_name="", title="", pid=-1, proc_name=""):
        self.hwnd = hwnd
        self.class_name = class_name
        self.title = title
        self.pid = pid
        self.proc_name = proc_name

    def __eq__(self, other):
        if isinstance(other, WindowInfo):
            return self.hwnd == other.hwnd
        elif isinstance(other, int):
            return self.hwnd == other
        else:
            return False

    def __ne__(self, other):
        return not self.__eq__(other)

    def __str__(self):
        return "hwnd: {}, class: {}\ntitle: {}, pid: {}\nproc: {}".format(
                self.hwnd,
                self.class_name,
                self.title,
                self.pid,
                self.proc_name
                )


class WindowManager():
    def __init__(self, title="WindowManager", master_n=1, workspace_n=2,
                    ignore_list=[]):
        # logger.debug('new manager is created')
        print('debug: new manager is created')

        self.c = wmi.WMI()
        self.dwm = ctypes.cdll.dwmapi
        self.shell = win32com.client.Dispatch("WScript.Shell")

        monitors = win32api.EnumDisplayMonitors(None, None)
        (h_first_mon, _, (_,_,_,_)) = monitors[0]
        first_mon = win32api.GetMonitorInfo(h_first_mon)
        m_left, m_top, m_right, m_bottom = first_mon['Monitor']
        w_left, w_top, w_right, w_bottom = first_mon['Work']

        self.monitor_w = m_right - m_left
        self.monitor_h = m_bottom - m_top
        self.work_w = w_right - w_left
        self.work_h = w_bottom - w_top
        self.taskbar_h = self.monitor_h - self.work_h

        self.workspace_n = workspace_n
        self.workspaces = []
        for i in range(self.workspace_n):
            self.workspaces.append(list())
        self.workspace_idx = 0

        self.master_n = master_n
        self.layout = 0

        self.ignore_list = ignore_list

        self.offset_from_center = 0

    def _is_real_window(self, hwnd):
        '''Return True iff given window is a real Windows application window.'''
        status = ctypes.wintypes.DWORD()
        self.dwm.DwmGetWindowAttribute(
                ctypes.wintypes.HWND(hwnd),
                ctypes.wintypes.DWORD(14),
                ctypes.byref(status),
                ctypes.sizeof(status)
                )
        if status.value != 0:
            return False
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

    def _get_windows(self):
        def callback(hwnd, windows):
            if not self._is_real_window(hwnd):
                return
            class_name = win32gui.GetClassName(hwnd)
            for ignore_item in self.ignore_list:
                if class_name == ignore_item:
                    return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            window = WindowInfo(hwnd, class_name, win32gui.GetWindowText(hwnd),
                                pid, self._get_proc_name(pid))
            windows.append(window)

        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    def _get_proc_name(self,pid):
        try:
            for p in self.c.query('SELECT Name FROM Win32_Process WHERE\
                                    ProcessId = %s' % str(pid)):
                return p.Name
        except:
            return ""

    def _manage_windows(self):
        old_stack = self.workspaces[self.workspace_idx]
        new_stack = self._get_windows()

        rm_list_old = []
        rm_list_new = []
        for old in old_stack:
            still_here = False
            for new in new_stack:
                if old == new:
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
        # prepend new windows
        self.workspaces[self.workspace_idx] = new_stack + old_stack

    def arrange_windows(self):
        self._manage_windows()
        workspace = self.workspaces[self.workspace_idx]
        win_n = len(workspace)
        if win_n is 0:
            return
        if self.layout == 0:
            if win_n <= self.master_n:
                win_h = math.floor((self.work_h)/win_n)
                win_w = self.work_w

                for i in range(win_n):
                    try:
                        win32gui.SetWindowPos(
                                workspace[i].hwnd,
                                win32con.HWND_TOPMOST,
                                0, win_h*i + self.taskbar_h, win_w, win_h,
                                win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                                )
                    except Exception:
                        print("error: SetWindowPos" + str(workspace[i]))
                        # # exit(0)

            else:
                sub_w = math.floor(self.work_w/2) - self.offset_from_center
                main_w = self.work_w - sub_w
                main_h = math.floor((self.work_h) / (self.master_n))
                sub_h = math.floor((self.work_h) /\
                                                (win_n-self.master_n))

                for i in range(self.master_n):
                    try:
                        win32gui.SetWindowPos(
                                workspace[i].hwnd,
                                win32con.HWND_TOPMOST,
                                0, main_h*i + self.taskbar_h,
                                main_w, main_h,
                                win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                                )
                    except Exception:
                        print("error: SetWindowPos" + str(workspace[i]))
                        # exit(0)

                for i in range(win_n-self.master_n):
                    try:
                        win32gui.SetWindowPos(
                                workspace[i+self.master_n].hwnd,
                                win32con.HWND_TOPMOST,
                                main_w, sub_h*i + self.taskbar_h,
                                sub_w, sub_h,
                                win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                                )
                    except Exception:
                        print("error: SetWindowPos" + str(workspace[i]))
                        # exit(0)
        else:
            win_h = self.work_h
            win_w = self.work_w

            for i in range(win_n):
                try:
                    win32gui.SetWindowPos(
                            workspace[i].hwnd,
                            win32con.HWND_TOPMOST,
                            0, self.taskbar_h, win_w, win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print ("error: SetWindowpos" + str(workspace[i]))

    def next_layout(self):
        self.layout = (self.layout + 1) % 2
        self.arrange_windows()

    def close_window(self):
        hwnd = win32gui.GetForegroundWindow()
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE,0,0)
        time.sleep(0.5)
        self.arrange_windows()

    def focus_up(self, num=1):
        workspace = self.workspaces[self.workspace_idx]
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(workspace):
            if hwnd == window:
                print(i)
                target_window = i

        if target_window != -1:
            next_idx = (target_window + num) % len(workspace)
            try:
                # according to https://stackoverflow.com/questions/14295337/win32gui-setactivewindow-error-the-specified-procedure-could-not-be-found
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(workspace[next_idx].hwnd)
                print("debug: focus_up: " + str(next_idx))
            except Exception:
                print ("error: SetForegroundWindow" + str(next_idx))
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

    def switch_to_nth_vd(self, dstIdx):
        if self.workspace_idx == dstIdx:
            return
        for window in self.workspaces[self.workspace_idx]:
            self.hide_window(window.hwnd)

        self.workspace_idx = dstIdx
        for window in self.workspaces[self.workspace_idx]:
            self.show_window(window.hwnd)

        print("debug: switch_to_nth_vd: " + str(dstIdx)
                                    + ":" + str(self.workspace_idx))
        self.arrange_windows()

    def send_to_nth_vd(self, dstIdx):
        if self.workspace_idx == dstIdx:
            return
        workspace = self.workspaces[self.workspace_idx]
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(workspace):
            if hwnd == window:
                target_window = i

        if target_window != -1:
            self.workspaces[dstIdx][:0] = [workspace[target_window]]
            self.hide_window(workspace[target_window].hwnd)
            workspace.pop(target_window)
            print("debug: send_to_nth_vd: " + str(dstIdx))
            self.arrange_windows()

    def swap_master(self):
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(self.workspaces[self.workspace_idx]):
            if hwnd == window:
                target_window = i
        if target_window != -1:
            window = self.workspaces[self.workspace_idx].pop(target_window)
            self.workspaces[self.workspace_idx][:0] = [window]
            self.arrange_windows()

    def swap_windows(self, num=1):
        workspace = self.workspaces[self.workspace_idx]
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(workspace):
            if hwnd == window:
                print(i)
                target_window = i

        if target_window != -1:
            next_idx = (target_window + num) % len(workspace)
            window = workspace.pop(target_window)
            workspace[next_idx:next_idx] = [window]
            self.arrange_windows()

    def inc_master_n(self, num=1):
        self.master_n += num
        if self.master_n <= 0:
            self.master_n = 1
        print("debug: inc_master_n: " + str(num))
        self.arrange_windows()

    def expand_master(self, num=10):
        self.offset_from_center += num
        self.arrange_windows()

    def recover_windows(self):
        for i, stack in enumerate(self.workspaces):
            for window in stack:
                self.show_caption(window.hwnd)
                if i != self.workspace_idx:
                    self.show_window(window.hwnd)
        self.arrange_windows()

    def show_window_info(self):
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(self.workspaces[self.workspace_idx]):
            if hwnd == window:
                target_window = i
        if target_window != -1:
            window = self.workspaces[self.workspace_idx][target_window]
            print(window)
            # ret = win32gui.MessageBox(None, str(window), "window information", win32con.MB_YESNO)

    def show_caption(self, hwnd=-1):
        if hwnd == -1:
            hwnd = win32gui.GetForegroundWindow()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if not style & win32con.WS_CAPTION:
            style += win32con.WS_CAPTION
        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
        win32gui.SetWindowPos(
                hwnd,
                0, 0, 0, 0, 0,
                win32con.SWP_FRAMECHANGED |
                win32con.SWP_NOMOVE |
                win32con.SWP_NOSIZE |
                win32con.SWP_NOZORDER
                )

    def toggle_caption(self, hwnd=-1):
        if hwnd == -1:
            hwnd = win32gui.GetForegroundWindow()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if style & win32con.WS_CAPTION:
            style -= win32con.WS_CAPTION
        else:
            style += win32con.WS_CAPTION
        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
        win32gui.SetWindowPos(
                hwnd,
                0, 0, 0, 0, 0,
                win32con.SWP_FRAMECHANGED |
                win32con.SWP_NOMOVE |
                win32con.SWP_NOSIZE |
                win32con.SWP_NOZORDER
                )


if __name__ == '__main__':
    print('INFO: Start window manager')
    manager = WindowManager(ignore_list=['TaskManagerWindow'])

    while True:
        manager.arrange_windows()
        ##########
        # CONFIG #
        ##########
        if isKeyDown(win32con.VK_NONCONVERT):
            if isKeyDown(ord('Q')):
                manager.recover_windows()
                break
            elif isKeyDown(win32con.VK_TAB):
                manager.focus_up()
            elif isKeyDown(ord('C')):
                manager.close_window()
            elif isKeyDown(ord('A')):
                manager.send_to_nth_vd(1)
            elif isKeyDown(ord('S')):
                manager.send_to_nth_vd(0)
            elif isKeyDown(ord('Z')):
                manager.switch_to_nth_vd(1)
            elif isKeyDown(ord('X')):
                manager.switch_to_nth_vd(0)
            elif isKeyDown(0xBC):       # comma
                manager.inc_master_n(+1)
            elif isKeyDown(0xBE):       # period
                manager.inc_master_n(-1)
            elif isKeyDown(win32con.VK_RETURN):
                manager.swap_master()
            else:
                continue
        time.sleep(0.2)
