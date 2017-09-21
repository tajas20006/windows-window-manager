#!/usr/bin/python
# -*- coding: utf-8 -*-
import drawtext

import time
import threading
from os import cpu_count

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
        return "hwnd: {}  class: {}  pid: {}  \ntitle: {}  proc: {}".format(
                self.hwnd,
                self.class_name,
                self.pid,
                self.title if len(self.title)<=25 else self.title[:22] + "...",
                self.proc_name
                )


class FILETIME(ctypes.Structure):
    _fields_ = [
            ("dwLowDateTime", ctypes.wintypes.DWORD),
            ("dwHighDateTime", ctypes.wintypes.DWORD)
            ]

class WindowManager():
    def __init__(self, title="WindowManager", master_n=1, workspace_n=2,
                    ignore_list=[], layout = 0):
        # logger.debug('new manager is created')
        # print('debug: new manager is created')

        self.c = wmi.WMI()
        self.dwm = ctypes.cdll.dwmapi
        self.shell = win32com.client.Dispatch("WScript.Shell")

        self.text = drawtext.TextOnBar()
        self.thr = threading.Thread(target=self.text.create_text_box)
        self.thr.start()

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
        self.workspaces = [[]] * self.workspace_n
        self.workspace_idx = 0
        self.lock = threading.Lock()

        self.master_n = [master_n] * self.workspace_n
        self.layout = [layout] * self.workspace_n

        self.ignore_list = ignore_list

        self.offset_from_center = 0
        self.is_window_info_on = False

        self.idle_time = 0
        self.krnl_time = 0
        self.user_time = 0
        self.tick_count = 0
        self.cpu_n = cpu_count()
        self.GetSystemTimes = ctypes.windll.kernel32.GetSystemTimes

        self.first_to_do()

    def first_to_do(self):
        self.watch_dog()
        self.arrange_windows()

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

    def _in_ignore_list(self, class_name=None, title=None):
        if {"class_name": class_name, "title": title} in self.ignore_list:
            return True
        if {"class_name": class_name} in self.ignore_list:
            return True

    def _get_windows(self):
        def callback(hwnd, windows):
            if not self._is_real_window(hwnd):
                return
            class_name = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if self._in_ignore_list(class_name, title):
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            window = WindowInfo(hwnd, class_name, title,
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

    def show_system_info(self):
        up, dn = self.get_network_info()
        cpu = self.get_cpu_info()
        text = "CPU: {} %, Up: {}/s, dn: {}/s".format(cpu, up, dn)
        self.text.redraw(new_btm_text=text)

    def get_network_info(self):
        try:
            p = self.c.query('SELECT BytesReceivedPerSec, BytesSentPerSec\
                    FROM Win32_PerfFormattedData_Tcpip_NetworkInterface\
                    WHERE Name Like "Dell Wireless 1820A 802.11ac"')
            received = self._bytes_to_str(int(p[0].BytesReceivedPerSec))
            sent = self._bytes_to_str(int(p[0].BytesSentPerSec))
            return (received, sent)
        except:
            return ("0","0")
            pass

    def get_cpu_info(self):
        new_tick_count = win32api.GetTickCount()
        elapsed = new_tick_count - self.tick_count

        new_idle_time = FILETIME()
        new_krnl_time = FILETIME()
        new_user_time = FILETIME()

        success = self.GetSystemTimes(
                ctypes.byref(new_idle_time),
                ctypes.byref(new_krnl_time),
                ctypes.byref(new_user_time)
                )

        total_idle_ticks = (new_idle_time.dwLowDateTime - self.idle_time) * 0.0001/self.cpu_n
        ppu_idle = total_idle_ticks / elapsed * 100
        ppu = 100 - ppu_idle

        sysTime = int(ppu)
        # sysTime = int((1 - (new_idle_time.dwLowDateTime - self.idle_time) / (new_krnl_time.dwLowDateTime - self.krnl_time + new_user_time.dwLowDateTime - self.user_time)) * 100)

        self.tick_count = new_tick_count
        self.idle_time = new_idle_time.dwLowDateTime
        self.krnl_time = new_krnl_time.dwLowDateTime
        self.user_time = new_user_time.dwLowDateTime

        return "{:>2}".format(str(sysTime))

    def _bytes_to_str(self, b):
        unit = ""
        if b > 1047527424:
            b /= 1024*1024*1024
            unit = "GB"
        elif b > 1022976:
            b /= 1024*1024
            unit = "MB"
        elif b > 999:
            b /= 1024
            unit = "kB"
        else:
            unit = " B"

        if b > 99.9 or unit == " B":
            b = str(b)[0:3]
        else:
            b = str(b)[0:4]
        return "{:>4} {}".format(b, unit)

    def watch_dog(self):
        self.lock.acquire()
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
                    old.title = new.title
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
        self.lock.release()

        try:
            left, top, right, bottom = self.text.rect
            win32gui.SetWindowPos(
                    self.text.hwnd,
                    win32con.HWND_TOPMOST,
                    left, top, right, bottom,
                    win32con.SWP_NOMOVE
                    )
        except:
            pass

        if len(new_stack) > 0:
            print("new!true")
            return True
        if len(rm_list_old) > 0:
            print("closed!true")
            return True
        return False

    def _update_taskbar_text(self):
        layout_str = ""
        if self.layout[self.workspace_idx] == 0:
            layout_str = "tile"
        else:
            layout_str = "full"
        workspace_str = ["","",""]
        workspace_color = [(250,250,250),(250,250,0),(250,250,250)]
        i = 0
        for j, workspace in enumerate(self.workspaces):
            if j == self.workspace_idx:
                workspace_str[1] += "[{}]".format(j+1)
                i = 2
            elif len(workspace) > 0:
                workspace_str[i] += " {} ".format(j+1)
        workspace_str[2] += "{}".format(layout_str)
        self.text.redraw(
                new_top_text=workspace_str,
                new_top_color=workspace_color
                )

    def arrange_windows(self):
        workspace = self.workspaces[self.workspace_idx]
        cur_master_n = self.master_n[self.workspace_idx]
        win_n = len(workspace)
        if win_n is 0:
            return
        if self.layout[self.workspace_idx] == 0:
            if win_n <= cur_master_n:
                win_h = int(self.work_h / win_n)
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

            else:
                sub_w = int(self.work_w/2) - self.offset_from_center
                main_w = self.work_w - sub_w
                main_h = int(self.work_h / cur_master_n)
                sub_h = int(self.work_h / (win_n-cur_master_n))

                for i in range(cur_master_n):
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

                for i in range(win_n-cur_master_n):
                    try:
                        win32gui.SetWindowPos(
                                workspace[i+cur_master_n].hwnd,
                                win32con.HWND_TOPMOST,
                                main_w, sub_h*i + self.taskbar_h,
                                sub_w, sub_h,
                                win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                                )
                    except Exception:
                        print("error: SetWindowPos" + str(workspace[i]))
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
        self.layout[self.workspace_idx] =\
                (self.layout[self.workspace_idx] + 1) % 2
        self.arrange_windows()
        self._update_taskbar_text()

    def close_window(self):
        self.lock.acquire()
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(self.workspaces[self.workspace_idx]):
            if hwnd == window:
                print(i)
                target_window = i
        if target_window != -1:
            self.workspaces[self.workspace_idx].pop(target_window)
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE,0,0)
        self.lock.release()
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
                # print("debug: focus_up: " + str(next_idx))
            except Exception:
                print ("error: SetForegroundWindow" + str(next_idx))

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

    def switch_to_nth_ws(self, dstIdx):
        if self.workspace_idx == dstIdx:
            return
        self.lock.acquire()
        for window in self.workspaces[self.workspace_idx]:
            self.hide_window(window.hwnd)

        self.workspace_idx = dstIdx
        for window in self.workspaces[self.workspace_idx]:
            self.show_window(window.hwnd)

        # print("debug: switch_to_nth_ws: " + str(dstIdx)
                                    # + ":" + str(self.workspace_idx))
        self.lock.release()
        self.arrange_windows()
        self._update_taskbar_text()

    def send_to_nth_ws(self, dstIdx):
        if self.workspace_idx == dstIdx:
            return
        self.lock.acquire()
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
            self.lock.release()
            # print("debug: send_to_nth_ws: " + str(dstIdx))
            self.arrange_windows()
            self._update_taskbar_text()
        else:
            self.lock.release()

    def swap_master(self):
        self.lock.acquire()
        hwnd = win32gui.GetForegroundWindow()
        target_window = -1
        for i, window in enumerate(self.workspaces[self.workspace_idx]):
            if hwnd == window:
                target_window = i
        if target_window != -1:
            window = self.workspaces[self.workspace_idx].pop(target_window)
            self.workspaces[self.workspace_idx][:0] = [window]
            self.lock.release()
            self.arrange_windows()
        else:
            self.lock.release()

    def swap_windows(self, num=1):
        self.lock.acquire()
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
            self.lock.release()
            self.arrange_windows()
        else:
            self.lock.release()

    def inc_master_n(self, num=1):
        self.master_n[self.workspace_idx] += num
        if self.master_n[self.workspace_idx] <= 0:
            self.master_n[self.workspace_idx] = 1
        # print("debug: inc_master_n: " + str(num))
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

    def reset_windows(self):
        self.recover_windows()
        self.workspaces = [[] for _ in range(self.workspace_n)]

    def show_window_info(self):
        if self.is_window_info_on:
            self.text.redraw(new_btm_text="")
            self.is_window_info_on = False
        else:
            hwnd = win32gui.GetForegroundWindow()
            class_name = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            window_str = "hwnd: {}, class: {}, pid: {} \ntitle: {}".format(
                    hwnd,
                    class_name,
                    pid,
                    title if len(title) <= 25 else title[:22] + "..."
                    )
            print(window_str)
            self.text.redraw(new_btm_text=window_str)
            self.is_window_info_on = True

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
