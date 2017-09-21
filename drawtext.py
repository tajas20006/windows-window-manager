# based on https://stackoverflow.com/questions/40614509/cant-update-text-in-\
# window-with-win32gui-drawtext

import win32api
import win32con
import win32gui
import time
import threading
import win32ui
import ctypes
import math

# Code example modified from:
# Christophe Keller
# Hello World in Python using Win32


class TextOnBar():
    def __init__(
            self,
            top_text=[""], top_color=[(250, 250, 250)],
            btm_text=[""], btm_color=[(250, 250, 250)],
            tol_text=[""], tol_color=[(250, 250, 250)],
            btl_text=[""], btl_color=[(250, 250, 250)],
            tor_text=[""], tor_color=[(250, 250, 250)],
            btr_text=[""], btr_color=[(250, 250, 250)]
            ):
        self.texts = [top_text, btm_text, tol_text, btl_text,
                      tor_text, btr_text]
        self.colors = [top_color, btm_color, tol_color, btl_color,
                       tor_color, btr_color]

        tray = win32gui.FindWindow('Shell_TrayWnd', None)
        self.tray_rect = win32gui.GetWindowRect(tray)
        self.rect = (0, 0, 1800, 900)

        lf = win32gui.LOGFONT()
        # ref. http://chokuto.ifdef.jp/urawaza/struct/LOGFONT.html
        # lf.lfFaceName = ""
        lf.lfCharSet = win32con.SHIFTJIS_CHARSET     # shift jis
        lf.lfOutPrecision = 4    # true type
        lf.lfPitchAndFamily = 1
        self.font_height = 16
        self.font_width = int(self.font_height/2) + 1
        lf.lfHeight = self.font_height
        lf.lfWeight = 900
        # Use nonantialiased to remove the white edges around the text.
        lf.lfQuality = win32con.NONANTIALIASED_QUALITY
        self.hf = win32gui.CreateFontIndirect(lf)

    def create_text_box(self):
        hInstance = win32api.GetModuleHandle()
        className = 'SimpleWin32'

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ms633576(v=vs.85).aspx
        # win32gui does not support WNDCLASSEX.
        wndClass = win32gui.WNDCLASS()
        # http://msdn.microsoft.com/en-us/library/windows/desktop/
        # ff729176(v=vs.85).aspx
        wndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wndClass.lpfnWndProc = self.wndProc
        wndClass.hInstance = hInstance
        wndClass.hCursor = win32gui.LoadCursor(None, win32con.IDC_ARROW)
        wndClass.hbrBackground = win32gui.GetStockObject(win32con.WHITE_BRUSH)
        wndClass.lpszClassName = className
        # win32gui does not support RegisterClassEx
        wndClassAtom = win32gui.RegisterClass(wndClass)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ff700543(v=vs.85).aspx
        # Consider using: WS_EX_COMPOSITED, WS_EX_LAYERED, WS_EX_NOACTIVATE,
        # WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT
        # The WS_EX_TRANSPARENT flag makes events (like mouse clicks)
        # fall through the window.
        exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED |\
            win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST |\
            win32con.WS_EX_TRANSPARENT

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ms632600(v=vs.85).aspx
        # Consider using: WS_DISABLED, WS_POPUP, WS_VISIBLE
        style = win32con.WS_DISABLED | win32con.WS_POPUP | win32con.WS_VISIBLE

        left, top, right, bottom = self.rect
        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ms632680(v=vs.85).aspx
        self.hwnd = win32gui.CreateWindowEx(
            exStyle,
            wndClassAtom,
            None,     # WindowName
            style,
            left, top, right, bottom,
            None,     # hWndParent
            None,     # hMenu
            hInstance,
            None      # lpParam
        )

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ms633540(v=vs.85).aspx
        win32gui.SetLayeredWindowAttributes(
                self.hwnd,
                0x00ffffff,
                255,
                win32con.LWA_COLORKEY | win32con.LWA_ALPHA
                )

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # dd145167(v=vs.85).aspx
        # win32gui.UpdateWindow(hwnd)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/\
        # ms633545(v=vs.85).aspx
        flags = win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE |\
            win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(
                self.hwnd,
                win32con.HWND_TOPMOST,
                0, 0, 0, 0,
                flags
                )

        # Show & update the window
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
        win32gui.UpdateWindow(self.hwnd)

        # Dispatch messages
        win32gui.PumpMessages()

    def redraw(
            self,
            top_text=None, top_color=None,
            btm_text=None, btm_color=None,
            tol_text=None, tol_color=None,
            btl_text=None, btl_color=None,
            tor_text=None, tor_color=None,
            btr_text=None, btr_color=None
            ):
        if top_text is not None:
            # if None, keep the old text
            if isinstance(top_text, str):
                # if str is given, put it in a list
                self.texts[0] = [top_text]
            else:
                self.texts[0] = top_text
        if top_color is not None:
            if isinstance(top_color, tuple):
                self.colors[0] = [top_color]
            else:
                self.colors[0] = top_color
        if btm_text is not None:
            if isinstance(btm_text, str):
                self.texts[1] = [btm_text]
            else:
                self.texts[1] = btm_text
        if btm_color is not None:
            if isinstance(btm_color, tuple):
                self.colors[1] = [btm_color]
            else:
                self.colors[1] = btm_color
        if tol_text is not None:
            if isinstance(tol_text, str):
                self.texts[2] = [tol_text]
            else:
                self.texts[2] = tol_text
        if tol_color is not None:
            if isinstance(tol_color, tuple):
                self.colors[2] = [tol_color]
            else:
                self.colors[2] = tol_color
        if btl_text is not None:
            if isinstance(btl_text, str):
                self.texts[3] = [btl_text]
            else:
                self.texts[3] = btl_text
        if btl_color is not None:
            if isinstance(btl_color, tuple):
                self.colors[3] = [btl_color]
            else:
                self.colors[3] = btl_color
        if tor_text is not None:
            if isinstance(tor_text, str):
                self.texts[4] = [tor_text]
            else:
                self.texts[4] = tor_text
        if tor_color is not None:
            if isinstance(tor_color, tuple):
                self.colors[4] = [tor_color]
            else:
                self.colors[4] = tor_color
        if btr_text is not None:
            if isinstance(btr_text, str):
                self.texts[5] = [btr_text]
            else:
                self.texts[5] = btr_text
        if btr_color is not None:
            if isinstance(btr_color, tuple):
                self.colors[5] = [btr_color]
            else:
                self.colors[5] = btr_color

        win32gui.RedrawWindow(
                self.hwnd,
                None, None,
                win32con.RDW_INVALIDATE | win32con.RDW_ERASE
                )

    def kill_proc(self):
        ctypes.windll.user32.PostThreadMessageW(
                thr.ident,
                win32con.WM_QUIT,
                0, 0
                )

    def wndProc(self, hWnd, message, wParam, lParam):

        if message == win32con.WM_PAINT:
            hDC, paintStruct = win32gui.BeginPaint(hWnd)

            win32gui.SelectObject(hDC, self.hf)
            win32gui.SetBkMode(hDC, win32con.TRANSPARENT)

            text_len = [0] * 6
            for i, text in enumerate(self.texts):
                for t in text:
                    text_len[i] += len(t) * self.font_width

            left, top, right, bottom = self.tray_rect
            vcenter = int((bottom-top) / 2)
            bgn_pos = [
                    int((right-left - text_len[0]) / 2),
                    int((right-left - text_len[1]) / 2),
                    left + 400,
                    left + 400,
                    right - 400 - text_len[4],
                    right - 400 - text_len[5]
                    ]

            print(self.texts)
            print(self.colors)

            for i, (t, c) in enumerate(zip(self.texts, self.colors)):
                for text, color in zip(t, c):
                    r, g, b = color
                    win32gui.SetTextColor(hDC, win32api.RGB(r, g, b))
                    flags = win32con.DT_SINGLELINE | win32con.DT_LEFT |\
                        win32con.DT_VCENTER
                    win32gui.DrawText(
                            hDC,
                            text,
                            -1,
                            (bgn_pos[i], top + vcenter*(i % 2),
                                right, vcenter + vcenter*(i % 2)),
                            flags
                            )
                    bgn_pos[i] += len(text) * self.font_width

            win32gui.EndPaint(hWnd, paintStruct)
            return 0

        elif message == win32con.WM_DESTROY:
            print('Being destroyed')
            win32gui.PostQuitMessage(0)
            return 0

        else:
            return win32gui.DefWindowProc(hWnd, message, wParam, lParam)


if __name__ == '__main__':
    draw = TextOnBar(
            top_text=["hello"], top_color=[(250, 250, 250)],
            btm_text=["world"], btm_color=[(250, 250, 250)],
            tol_text=["good"], tol_color=[(250, 0, 250)],
            btl_text=["mornig"], btl_color=[(0, 250, 250)],
            tor_text=["see you"], tor_color=[(0, 0, 250)],
            btr_text=["again"], btr_color=[(0, 250, 0)]
            )
    thr = threading.Thread(target=draw.create_text_box)
    thr.start()
    time.sleep(2)
    draw.redraw(
            top_text=["r", "a", "i", "n", "b", "o", "w"],
            btm_text=["white  only"],
            top_color=[(255, 0, 0), (255, 127, 0), (255, 255, 0),
                       (0, 255, 0), (0, 0, 255), (75, 0, 130), (148, 0, 211)],
            btm_color=[(250, 250, 250)]
            )
    time.sleep(2)
    draw.redraw("test", (23, 66, 200))
    time.sleep(2)
    draw.redraw([""], None, [""], None)
    time.sleep(2)
    draw.redraw(
            ["e", "n", "d"], [(255, 0, 0), (255, 127, 0), (255, 255, 0)],
            ["t", "e", "s", "t"],
            [(0, 255, 0), (0, 0, 255), (75, 0, 130), (148, 0, 211)]
            )
    time.sleep(2)
    draw.kill_proc()
