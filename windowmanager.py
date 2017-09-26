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
    def init(hwnd=-1, class_name="", title="", pid=-1, proc_name=""):
        hwnd = hwnd
        class_name = class_name
        title = title
        pid = pid
        proc_name = proc_name

    def __eq__(other):
        if isinstance(other, WindowInfo):
            return hwnd == other.hwnd
        elif isinstance(other, int):
            return hwnd == other
        else:
            return False

    def __ne__(other):
        return not __eq__(other)

    def __str__():
        short_title = ""
        if len(title) <= 25:
            short_title = title
        else:
            short_title = title[:22] + "..."

        return "hwnd: {}  class: {}  pid: {}  \ntitle: {}  proc: {}".format(
                hwnd,
                class_name,
                pid,
                short_title,
                proc_name
                )


class FILETIME(ctypes.Structure):
    _fields_ = [
            ("dwLowDateTime", ctypes.wintypes.DWORD),
            ("dwHighDateTime", ctypes.wintypes.DWORD)
            ]


def init(
        title="WindowManager",
        master_n=1, workspace_n=2,
        ignore_list=[],
        layout=0,
        network_interface=""
        ):
    # logger.debug('new manager is created')
    # print('debug: new manager is created')

    c = wmi.WMI()
    dwm = ctypes.cdll.dwmapi
    shell = win32com.client.Dispatch("WScript.Shell")

    text = drawtext.TextOnBar()
    thr = threading.Thread(target=text.create_text_box)
    thr.start()

    monitors = win32api.EnumDisplayMonitors(None, None)
    (h_first_mon, _, (_, _, _, _)) = monitors[0]
    first_mon = win32api.GetMonitorInfo(h_first_mon)
    m_left, m_top, m_right, m_bottom = first_mon['Monitor']
    w_left, w_top, w_right, w_bottom = first_mon['Work']

    monitor_w = m_right - m_left
    monitor_h = m_bottom - m_top
    work_w = w_right - w_left
    work_h = w_bottom - w_top
    taskbar_h = monitor_h - work_h

    workspace_n = workspace_n
    workspaces = [[]] * workspace_n
    workspace_idx = 0
    lock = threading.Lock()

    master_n = [master_n] * workspace_n
    layout = [layout] * workspace_n

    ignore_list = ignore_list

    offset_from_center = 0
    is_window_info_on = False

    network_interface = network_interface
    idle_time = 0
    krnl_time = 0
    user_time = 0
    tick_count = 0
    cpu_n = cpu_count()
    GetSystemTimes = ctypes.windll.kernel32.GetSystemTimes

    first_to_do()

def first_to_do():
    watch_dog()
    arrange_windows()

def _is_real_window(hwnd):
    '''Return True iff given window is a real window.'''
    status = ctypes.wintypes.DWORD()
    dwm.DwmGetWindowAttribute(
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
    if (((lExStyle & win32con.WS_EX_TOOLWINDOW) == 0 and hasNoOwner) or
            ((lExStyle & win32con.WS_EX_APPWINDOW != 0) and not
                hasNoOwner)):
        if win32gui.GetWindowText(hwnd):
            return True
    return False

def _in_ignore_list(class_name=None, title=None):
    if {"class_name": class_name, "title": title} in ignore_list:
        return True
    if {"class_name": class_name} in ignore_list:
        return True

def _get_windows():
    def callback(hwnd, windows):
        if not _is_real_window(hwnd):
            return
        class_name = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if _in_ignore_list(class_name, title):
            return
        # if win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)\
        #         & win32con.WS_POPUP:
        #     return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        window = WindowInfo(hwnd, class_name, title,
                            pid, _get_proc_name(pid))
        windows.append(window)

    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows

def _get_proc_name(pid):
    try:
        for p in c.query('SELECT Name FROM Win32_Process WHERE\
                                ProcessId = %s' % str(pid)):
            return p.Name
    except:
        return ""

def show_system_info():
    up, dn = get_network_info()
    cpu = get_cpu_info()
    text = "CPU: {}%, Up: {}/s, dn: {}/s".format(cpu, up, dn)
    try:
        text.redraw(tol_text=text, tol_color=(250, 250, 250))
    except:
        pass

def get_network_info():
    try:
        p = c.query('SELECT BytesReceivedPerSec, BytesSentPerSec\
                FROM Win32_PerfFormattedData_Tcpip_NetworkInterface\
                WHERE Name Like "{}"'.format(network_interface))
        received = _bytes_to_str(int(p[0].BytesReceivedPerSec))
        sent = _bytes_to_str(int(p[0].BytesSentPerSec))
        return (received, sent)
    except:
        return (" EEE", " EEE")
        pass

def _bytes_to_str(b):
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
    return "{:>4}{}".format(b, unit)

def get_cpu_info():
    new_tick_count = win32api.GetTickCount()
    elapsed = new_tick_count - tick_count

    new_idle_time = FILETIME()
    new_krnl_time = FILETIME()
    new_user_time = FILETIME()

    success = GetSystemTimes(
            ctypes.byref(new_idle_time),
            ctypes.byref(new_krnl_time),
            ctypes.byref(new_user_time)
            )

    total_idle_ticks = (new_idle_time.dwLowDateTime - idle_time)\
        * 0.0001/cpu_n
    ppu_idle = total_idle_ticks / elapsed * 100
    ppu = 100 - ppu_idle

    sysTime = int(ppu)
    # sysTime = int((1 - (new_idle_time.dwLowDateTime - idle_time) /
    #             (new_krnl_time.dwLowDateTime - krnl_time +
    #                 new_user_time.dwLowDateTime - user_time)) * 100)

    tick_count = new_tick_count
    idle_time = new_idle_time.dwLowDateTime
    krnl_time = new_krnl_time.dwLowDateTime
    user_time = new_user_time.dwLowDateTime

    return "{:>2}".format(str(sysTime)) if sysTime < 100 else "EE"

def watch_dog():
    lock.acquire()
    old_stack = workspaces[workspace_idx]
    new_stack = _get_windows()

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
    workspaces[workspace_idx] = new_stack + old_stack
    lock.release()

    try:
        left, top, right, bottom = text.rect
        win32gui.SetWindowPos(
                text.hwnd,
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

def _update_taskbar_text():
    layout_str = ""
    if layout[workspace_idx] == 0:
        layout_str = "tile"
    else:
        layout_str = "full"
    workspace_str = ["", "", ""]
    workspace_color = [(250, 250, 250), (250, 250, 0), (250, 250, 250)]
    i = 0
    for j, workspace in enumerate(workspaces):
        if j == workspace_idx:
            workspace_str[1] += "[{}]".format(j+1)
            i = 2
        elif len(workspace) > 0:
            workspace_str[i] += " {} ".format(j+1)
    workspace_str[2] += "{}".format(layout_str)
    text.redraw(
            tor_text=workspace_str,
            tor_color=workspace_color
            )

def arrange_windows():
    workspace = workspaces[workspace_idx]
    cur_master_n = master_n[workspace_idx]
    win_n = len(workspace)
    if win_n is 0:
        return
    if layout[workspace_idx] == 0:
        if win_n <= cur_master_n:
            win_h = int(work_h / win_n)
            win_w = work_w

            for i in range(win_n):
                try:
                    win32gui.SetWindowPos(
                            workspace[i].hwnd,
                            win32con.HWND_TOPMOST,
                            0, win_h*i + taskbar_h, win_w, win_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + str(workspace[i]))

        else:
            sub_w = int(work_w/2) - offset_from_center
            main_w = work_w - sub_w
            main_h = int(work_h / cur_master_n)
            sub_h = int(work_h / (win_n-cur_master_n))

            for i in range(cur_master_n):
                try:
                    win32gui.SetWindowPos(
                            workspace[i].hwnd,
                            win32con.HWND_TOPMOST,
                            0, main_h*i + taskbar_h,
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
                            main_w, sub_h*i + taskbar_h,
                            sub_w, sub_h,
                            win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                            )
                except Exception:
                    print("error: SetWindowPos" + str(workspace[i]))
    else:
        win_h = work_h
        win_w = work_w

        for i in range(win_n):
            try:
                win32gui.SetWindowPos(
                        workspace[i].hwnd,
                        win32con.HWND_TOPMOST,
                        0, taskbar_h, win_w, win_h,
                        win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER
                        )
            except Exception:
                print("error: SetWindowpos" + str(workspace[i]))

def next_layout():
    layout[workspace_idx] =\
            (layout[workspace_idx] + 1) % 2
    arrange_windows()
    _update_taskbar_text()

def close_window():
    lock.acquire()
    hwnd = win32gui.GetForegroundWindow()
    target_window = -1
    for i, window in enumerate(workspaces[workspace_idx]):
        if hwnd == window:
            print(i)
            target_window = i
    if target_window != -1:
        workspaces[workspace_idx].pop(target_window)
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    lock.release()
    time.sleep(0.5)
    arrange_windows()

def focus_up(num=1):
    workspace = workspaces[workspace_idx]
    hwnd = win32gui.GetForegroundWindow()
    target_window = -1
    for i, window in enumerate(workspace):
        if hwnd == window:
            print(i)
            target_window = i

    if target_window != -1:
        next_idx = (target_window + num) % len(workspace)
        try:
            # according to https://stackoverflow.com/questions/14295337/\
            # win32gui-setactivewindow-error-the-specified-procedure-could\
            # -not-be-found
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(workspace[next_idx].hwnd)
            # print("debug: focus_up: " + str(next_idx))
        except Exception:
            print("error: SetForegroundWindow" + str(next_idx))

def hide_window(hwnd):
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOPMOST,
                left, top, right, bottom,
                win32con.SWP_HIDEWINDOW | win32con.SWP_NOZORDER |
                win32con.SWP_NOMOVE
                )
    except Exception:
        print("error: SetWindowPos" + str(hwnd))

def go_fullscreen(hwnd=-1):
    if hwnd == -1:
        hwnd = win32gui.GetForegroundWindow()
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    style |= win32con.WS_SYSMENU | win32con.WS_VISIBLE | win32con.WS_POPUP
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)

def ungo_fullscreen(hwnd=-1):
    if hwnd == -1:
        hwnd = win32gui.GetForegroundWindow()
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    if style & win32con.WS_SYSMENU:
        style -= win32con.WS_SYSMENU
    if style & win32con.WS_VISIBLE:
        style -= win32con.WS_VISIBLE
    if style & win32con.WS_POPUP:
        style -= win32con.WS_POPUP

def show_window(hwnd):
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        flags = win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER |\
            win32con.SWP_NOMOVE
        win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOPMOST,
                left, top, right, bottom,
                flags
                )
    except Exception:
        print("error: SetWindowPos" + str(hwnd))

def switch_to_nth_ws(dstIdx):
    if workspace_idx == dstIdx:
        return
    lock.acquire()
    for window in workspaces[workspace_idx]:
        hide_window(window.hwnd)

    workspace_idx = dstIdx
    for window in workspaces[workspace_idx]:
        show_window(window.hwnd)

    # print("debug: switch_to_nth_ws: " + str(dstIdx) +
    #       ":" + str(workspace_idx))
    lock.release()
    arrange_windows()
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(workspaces[workspace_idx][0].hwnd)
    _update_taskbar_text()

def send_to_nth_ws(dstIdx):
    if workspace_idx == dstIdx:
        return
    lock.acquire()
    workspace = workspaces[workspace_idx]
    hwnd = win32gui.GetForegroundWindow()
    target_window = -1
    for i, window in enumerate(workspace):
        if hwnd == window:
            target_window = i

    if target_window != -1:
        workspaces[dstIdx][:0] = [workspace[target_window]]
        hide_window(workspace[target_window].hwnd)
        workspace.pop(target_window)
        lock.release()
        # print("debug: send_to_nth_ws: " + str(dstIdx))
        arrange_windows()
        _update_taskbar_text()
    else:
        lock.release()

def swap_master():
    lock.acquire()
    hwnd = win32gui.GetForegroundWindow()
    target_window = -1
    for i, window in enumerate(workspaces[workspace_idx]):
        if hwnd == window:
            target_window = i
    if target_window != -1:
        window = workspaces[workspace_idx].pop(target_window)
        workspaces[workspace_idx][:0] = [window]
        lock.release()
        arrange_windows()
    else:
        lock.release()

def swap_windows(num=1):
    lock.acquire()
    workspace = workspaces[workspace_idx]
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
        lock.release()
        arrange_windows()
    else:
        lock.release()

def inc_master_n(num=1):
    master_n[workspace_idx] += num
    if master_n[workspace_idx] <= 0:
        master_n[workspace_idx] = 1
    # print("debug: inc_master_n: " + str(num))
    arrange_windows()

def expand_master(num=10):
    offset_from_center += num
    arrange_windows()

def recover_windows():
    for i, stack in enumerate(workspaces):
        for window in stack:
            show_caption(window.hwnd)
            if i != workspace_idx:
                show_window(window.hwnd)
    arrange_windows()

def reset_windows():
    recover_windows()
    workspaces = [[] for _ in range(workspace_n)]
    if watch_dog():
        arrange_windows()

def show_window_info():
    if is_window_info_on:
        text.redraw(btm_text="")
        is_window_info_on = False
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
        text.redraw(btm_text=window_str)
        is_window_info_on = True

def show_caption(hwnd=-1):
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

def toggle_caption(hwnd=-1):
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
