# based on https://gist.github.com/keturn/6695625

import win32con
import win32gui

import sys
import ctypes
import ctypes.wintypes

user32 = ctypes.windll.user32
ole32 = ctypes.windll.ole32
kernel32 = ctypes.windll.kernel32

WinEventProcType = ctypes.WINFUNCTYPE(
        None,
        ctypes.wintypes.HANDLE,
        ctypes.wintypes.DWORD,
        ctypes.wintypes.HWND,
        ctypes.wintypes.LONG,
        ctypes.wintypes.LONG,
        ctypes.wintypes.DWORD,
        ctypes.wintypes.DWORD
        )

# The types of events we want to listen for, and the names we'll use for
# them in the log output. Pick from
# http://msdn.microsoft.com/en-us/library/windows/desktop/dd318066(v=vs.85).aspx
eventTypes = {
        # win32con.EVENT_SYSTEM_FOREGROUND: "Foreground",
        win32con.EVENT_OBJECT_FOCUS: "Focus",
        # win32con.EVENT_OBJECT_CREATE: "Create",
        # win32con.EVENT_OBJECT_DESTROY: "Destroy",
        win32con.EVENT_OBJECT_SHOW: "Show",
        # win32con.EVENT_OBJECT_HIDE: "Hide",
        # 0x8017: "Cloaked",
        # 0x0016: "min",
        # win32con.EVENT_OBJECT_VALUECHANGE: "VLUE",
        win32con.EVENT_OBJECT_STATECHANGE: "STATE"
        # win32con.EVENT_OBJECT_PARENTCHANGE: "parent"
        }

processFlag = getattr(win32con, 'PROCESS_QUERY_LIMITED_INFORMATION',
                        win32con.PROCESS_QUERY_INFORMATION)
threadFlag = getattr(win32con, 'THREAD_QUERY_LIMITED_INFORMATION',
                        win32con.THREAD_QUERY_INFORMATION)

class WindowHandler():
    def __init__(self, function_to_call=None):
        self.lastTime = 0
        self.function_to_call = function_to_call

    def log(self, msg):
        print(msg)

    def logError(self, msg):
        sys.stdout.write(msg + '\n')

    def getProcessID(self, dwEventThread, hwnd):
        hThread = kernel32.OpenThread(threadFlag, 0, dwEventThread)
        if hThread:
            try:
                processID = kernel32.GetProcessIdOfThread(hThread)
                if not processID:
                    self.logError("Couldn't get process for thread %s: %s" %
                            (hThread, ctypes.WinError()))
            finally:
                kernel32.CloseHandle(hThread)
        else:
            errors = ["No thread handle for %s: %s" %
                    (dwEventThread, ctypes.WinError(),)]

            if hwnd:
                processID = ctypes.wintypes.DWORD()
                threadID = user32.GetWindowThreadProcessId(
                    hwnd, ctypes.byref(processID))
                if threadID != dwEventThread:
                    self.logError("Window thread != event thread? %s != %s" %
                             (threadID, dwEventThread))
                if processID:
                    processID = processID.value
                else:
                    errors.append(
                        "GetWindowThreadProcessID(%s) didn't work either: %s" % (
                        hwnd, ctypes.WinError()))
                    processID = None
            else:
                processID = None

            if not processID:
                for err in errors:
                    self.logError(err)

        return processID


    def getProcessFilename(self, processID):
        hProcess = kernel32.OpenProcess(processFlag, 0, processID)
        if not hProcess:
            self.logError("OpenProcess(%s) failed: %s" % (processID, ctypes.WinError()))
            return None

        try:
            filenameBufferSize = ctypes.wintypes.DWORD(4096)
            filename = ctypes.create_unicode_buffer(filenameBufferSize.value)
            kernel32.QueryFullProcessImageNameW(hProcess, 0, ctypes.byref(filename),
                                                ctypes.byref(filenameBufferSize))

            return filename.value
        finally:
            kernel32.CloseHandle(hProcess)


    def isRealWindow(self, hwnd):
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

    def callback(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread,
                 dwmsEventTime):
        length = user32.GetWindowTextLengthW(hwnd)
        title = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title, length + 1)

        if title.value[-10:] == "vimrun.exe":
            return

        if self.isRealWindow(hwnd):
            processID = self.getProcessID(dwEventThread, hwnd)
        else:
            return
        shortName = '?'
        if processID:
            filename = self.getProcessFilename(processID)
            if filename:
                splitted = filename.rsplit("\\", 2)
                if splitted[-1:][0] == "ShellExperienceHost.exe":
                    return
                shortName = '\\'.join(splitted[-2:])

        if hwnd:
            hwnd = hex(hwnd)
        elif idObject == win32con.OBJID_CURSOR:
            hwnd = '<Cursor>'

        # self.log(u"%s:%04.2f\t%-10s\t"
        #     u"W:%-8s\tP:%-8d\tT:%-8d\t"
        #     u"%s\t%s" % (
        #     dwmsEventTime, float(dwmsEventTime - self.lastTime)/1000, eventTypes.get(event, hex(event)),
        #     hwnd, processID or -1, dwEventThread or -1,
        #     shortName, title.value))

        if self.function_to_call:
            self.function_to_call()
        self.lastTime = dwmsEventTime


    def setHook(self, WinEventProc, eventType):
        return user32.SetWinEventHook(
            eventType,
            eventType,
            0,
            WinEventProc,
            0,
            0,
            win32con.WINEVENT_OUTOFCONTEXT
        )


    def main(self):
        ole32.CoInitialize(0)

        WinEventProc = WinEventProcType(self.callback)
        user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE

        hookIDs = [self.setHook(WinEventProc, et) for et in eventTypes.keys()]
        if not any(hookIDs):
            print ('SetWinEventHook failed')
            sys.exit(1)

        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
            user32.TranslateMessageW(msg)
            user32.DispatchMessageW(msg)

        for hookID in hookIDs:
            user32.UnhookWinEvent(hookID)
        ole32.CoUninitialize()


if __name__ == '__main__':
    window = WindowHandler()
    window.main()
