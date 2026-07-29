import ctypes
import time
import sys

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
PostMessage = ctypes.windll.user32.PostMessageW
WM_CLOSE = 0x0010

log_file = open(r"C:\ledugma\DEMO\click_log.txt", "w", encoding="utf-8")

def click_ok():
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buff, length + 1)
                title = buff.value
                
                if "Microsoft Excel" in title or "מערכת דוחות" in title:
                    log_file.write(f"Found window: {title}\n")
                    log_file.flush()
                    # Send Enter key
                    PostMessage(hwnd, 0x0100, 0x0D, 0) # WM_KEYDOWN, VK_RETURN
                    PostMessage(hwnd, 0x0101, 0x0D, 0) # WM_KEYUP, VK_RETURN
        return True

    for _ in range(60):
        EnumWindows(EnumWindowsProc(foreach_window), 0)
        time.sleep(1)

if __name__ == "__main__":
    click_ok()
    log_file.close()
