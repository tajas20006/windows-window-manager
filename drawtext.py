# based on https://stackoverflow.com/questions/40614509/cant-update-text-in-window-with-win32gui-drawtext

import win32api
import win32con
import win32gui
import time
import threading
import win32ui
import ctypes
import math

#Code example modified from:
#Christophe Keller
#Hello World in Python using Win32


class TextOnTray():
    def __init__(self, top_text=["hello"], top_color=[(250,250,250)],
                       btm_text=["world"], btm_color=[(250,250,250)]):
        self.top_text = top_text
        self.top_color = top_color
        self.btm_text = btm_text
        self.btm_color = btm_color

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
        # # # Use nonantialiased to remove the white edges around the text.
        lf.lfQuality = win32con.NONANTIALIASED_QUALITY
        # lf = win32gui.GetObject(win32gui.GetStockObject(17))
        self.hf = win32gui.CreateFontIndirect(lf)

    def create_text_box(self):
        hInstance = win32api.GetModuleHandle()
        className = 'SimpleWin32'

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633576(v=vs.85).aspx
        # win32gui does not support WNDCLASSEX.
        wndClass                = win32gui.WNDCLASS()
        # http://msdn.microsoft.com/en-us/library/windows/desktop/ff729176(v=vs.85).aspx
        wndClass.style          = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wndClass.lpfnWndProc    = self.wndProc
        wndClass.hInstance      = hInstance
        wndClass.hCursor        = win32gui.LoadCursor(None, win32con.IDC_ARROW)
        wndClass.hbrBackground  = win32gui.GetStockObject(win32con.WHITE_BRUSH)
        wndClass.lpszClassName  = className
        # win32gui does not support RegisterClassEx
        wndClassAtom = win32gui.RegisterClass(wndClass)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ff700543(v=vs.85).aspx
        # Consider using: WS_EX_COMPOSITED, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT
        # The WS_EX_TRANSPARENT flag makes events (like mouse clicks) fall through the window.
        exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632600(v=vs.85).aspx
        # Consider using: WS_DISABLED, WS_POPUP, WS_VISIBLE
        style = win32con.WS_DISABLED | win32con.WS_POPUP | win32con.WS_VISIBLE

        tray = win32gui.FindWindow('Shell_TrayWnd', None)
        self.rect  = win32gui.GetWindowRect(tray)

        left, top, right, bottom = self.rect
        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632680(v=vs.85).aspx
        self.hwnd = win32gui.CreateWindowEx(
            exStyle,
            wndClassAtom,
            None, # WindowName
            style,
            left, top, right, bottom,
            # 0, # x
            # 0, # y
            # win32api.GetSystemMetrics(win32con.SM_CXSCREEN), # width
            # win32api.GetSystemMetrics(win32con.SM_CYSCREEN), # height
            None, # hWndParent
            None, # hMenu
            hInstance,
            None # lpParam
        )

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633540(v=vs.85).aspx
        win32gui.SetLayeredWindowAttributes(self.hwnd, 0x00ffffff, 255, win32con.LWA_COLORKEY | win32con.LWA_ALPHA)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/dd145167(v=vs.85).aspx
        #win32gui.UpdateWindow(hwnd)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633545(v=vs.85).aspx
        win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
            win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)


        # Show & update the window
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
        win32gui.UpdateWindow(self.hwnd)

        # New code: Create and start the thread
        # thr = threading.Thread(target=self.customDraw)
        # thr.setDaemon(False)
        # thr.start()

        # Dispatch messages
        win32gui.PumpMessages()

# New code: Attempt to change the text 1 second later
    def customDraw(self, new_top_text=None, new_top_color=None,
                         new_btm_text=None, new_btm_color=None):
        if new_top_text is not None:
            if isinstance(new_top_text, str):
                self.top_text = [new_top_text]
            else:
                self.top_text = new_top_text
        if new_top_color is not None:
            if isinstance(new_top_color, tuple):
                self.top_color = [new_top_color]
            else:
                self.top_color = new_top_color
        if new_btm_text is not None:
            if isinstance(new_btm_text, str):
                self.btm_text = [new_btm_text]
            else:
                self.btm_text = new_btm_text
        if new_btm_color is not None:
            if isinstance(new_btm_color, tuple):
                self.btm_color = [new_btm_color]
            else:
                self.btm_color = new_btm_color
        win32gui.RedrawWindow(self.hwnd, None, None, win32con.RDW_INVALIDATE | win32con.RDW_ERASE)

    def kill_proc(self):
        ctypes.windll.user32.PostThreadMessageW(
                thr.ident,
                win32con.WM_QUIT,
                0, 0
                )

    def get_str_width(self, text_str=""):
        return len(text_str) * self.font_height/2

    def wndProc(self, hWnd, message, wParam, lParam):

        if message == win32con.WM_PAINT:
            # https://msdn.microsoft.com/en-us/library/windows/desktop/dd145096(v=vs.85).aspx
            # change color
            hDC, paintStruct = win32gui.BeginPaint(hWnd)

            win32gui.SelectObject(hDC, self.hf)
            win32gui.SetBkMode(hDC, win32con.TRANSPARENT)

            top_text_len = 0
            btm_text_len = 0
            for text in self.top_text:
                top_text_len += len(text) * self.font_width
            for text in self.btm_text:
                btm_text_len += len(text) * self.font_width

            left, top, right, bottom = self.rect
            vcenter = int((bottom-top) / 2)
            top_bgn_pos = int((right-left - top_text_len) / 2)
            btm_bgn_pos = int((right-left - btm_text_len) / 2)

            for (text, color) in zip(self.top_text, self.top_color):
                r,g,b = color
                win32gui.SetTextColor(hDC, win32api.RGB(r,g,b))
                win32gui.DrawText(
                        hDC,
                        text,
                        -1,
                        (top_bgn_pos, top, right, vcenter),
                        win32con.DT_SINGLELINE | win32con.DT_LEFT | win32con.DT_VCENTER
                        )
                top_bgn_pos += len(text) * self.font_width

            for (text, color) in zip(self.btm_text, self.btm_color):
                r,g,b = color
                win32gui.SetTextColor(hDC, win32api.RGB(r,g,b))
                win32gui.DrawText(
                        hDC,
                        text,
                        -1,
                        (btm_bgn_pos, vcenter, right, bottom),
                        win32con.DT_SINGLELINE | win32con.DT_LEFT | win32con.DT_VCENTER
                        )
                btm_bgn_pos += len(text) * self.font_width

            # win32gui.SetTextColor(hDC, win32api.RGB(255,255,0))
            #
            #
            # win32gui.DrawText(
            #     hDC,
            #     self.textLists,
            #     -1,
            #     (left, top, math.floor(right/2), bottom),
            #     win32con.DT_SINGLELINE | win32con.DT_RIGHT | win32con.DT_VCENTER)
            #
            # win32gui.SetTextColor(hDC, win32api.RGB(255,0,0))
            # win32gui.DrawText(
            #         hDC,
            #         self.textLists,
            #         -1,
            #         (math.floor(right/2), top, right, bottom),
            #         win32con.DT_SINGLELINE | win32con.DT_LEFT | win32con.DT_VCENTER)
            #
            win32gui.EndPaint(hWnd, paintStruct)
            return 0

        elif message == win32con.WM_DESTROY:
            print('Being destroyed')
            win32gui.PostQuitMessage(0)
            return 0

        else:
            return win32gui.DefWindowProc(hWnd, message, wParam, lParam)

if __name__ == '__main__':
    draw = TextOnTray()

    thr = threading.Thread(target=draw.create_text_box)
    thr.start()
    time.sleep(2)
    draw.customDraw(
            new_top_text=["r","a","i","n","b","o","w"],
            new_btm_text=["white  only"],
            new_top_color=[(255,0,0),(255,127,0),(255,255,0),
                            (0,255,0),(0,0,255),(75,0,130),(148,0,211)],
            new_btm_color=[(250,250,250)]
            )
    time.sleep(2)
    draw.customDraw("test",(23,66,200))
    time.sleep(2)
    draw.customDraw([""],None,[""],None)
    time.sleep(2)
    draw.customDraw(["e","n","d"],[(255,0,0,),(255,127,0),(255,255,0)],
                    ["t","e","s","t"],[(0,255,0),(0,0,255),(75,0,130),(148,0,211)])
    time.sleep(2)
    draw.kill_proc()
