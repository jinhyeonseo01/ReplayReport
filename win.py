import win32con
import win32api
import win32gui
import win32process

from pynput import keyboard, mouse
import pyautogui

import time

def is_main_window(hwnd):
    # 윈도우가 가시적인지 확인
    if not win32gui.IsWindowVisible(hwnd):
        return False
    # 윈도우가 최상위 레벨 윈도우인지 확인
    #if win32gui.GetWindow(hwnd, win32gui.GetWindow) != 0:
        #return False
    # 윈도우가 프로그램 창인지 확인
    if win32gui.GetWindowText(hwnd) == "":
        return False
    return True

def is_foreground_window(hwnd):
    # 윈도우의 프로세스 ID를 가져옴
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    # 윈도우가 백그라운드 프로세스인지 확인
    return pid != 0

def get_window_current():
    return win32process.GetWindowThreadProcessId(win32process.GetCurrentProcess())[1]

def get_window_rect(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    left, top, right, bottom = rect
    return (left, top, right, bottom)

def find_hwnd_by_title(window_title):
    return win32gui.FindWindow(None, window_title)

# 키 이벤트를 보내는 함수
def send_key_event(hwnd, key):
    for i in range(4):
        #print(hwnd)
        #win32api.keybd_event(key, 0, 0, 0)  # 키 누름
        #win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
        #time.sleep(0.2)  # 0.1초 대기
        #win32gui.Mess
        pyautogui.keyDown('space')
        time.sleep(0.1)
        pyautogui.keyUp('space')
        #win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)
        #win32gui.SendMessage(hwnd, win32con.WM_CHAR, key, 0)
        #time.sleep(0.1)
        #win32gui.SendMessage(hwnd, win32con.WM_KEYUP, key, 0)

def send_input(data, offset):
    if data['type'] == 'Keyboard':
        if data['pressed'] == "Down":
            pyautogui.keyDown(data['data'])
        if data['pressed'] == "Up":
            pyautogui.keyUp(data['data'])
    else:
        if data['event'] == "Click":
            if data['pressed'] == "Down":
                pyautogui.mouseDown(data['data'][0] + offset[0], data['data'][1] + offset[1], button=data['data'][2])
            else:
                pyautogui.mouseUp(data['data'][0] + offset[0], data['data'][1] + offset[1], button=data['data'][2])
        if data['event'] == "Move":
            win32api.SetCursorPos((data['data'][0] + offset[0], data['data'][1] + offset[1]))
        if data['event'] == "Scroll":
            pyautogui.scroll(3*(-1 if data['pressed'] == "Down" else 1))#, data['data'][0], data['data'][1]


def enum_windows_callback(hwnd, window_list):
    if is_foreground_window(hwnd) and is_main_window(hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        window_list.append((hwnd, window_title))
        return True
    return True

def get_top_level_windows():
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows


def OnKeyboardEvent(event):
    print('Key:', event.Key)
    return True

def getProcessInfoList():
    return [i for i in get_top_level_windows() if i[1] != '' and i != 'Default IME']



special_keys = {
    keyboard.Key.enter: win32con.VK_RETURN,
    keyboard.Key.esc: win32con.VK_ESCAPE,
    keyboard.Key.shift: win32con.VK_SHIFT,
    keyboard.Key.ctrl: win32con.VK_CONTROL,
    keyboard.Key.alt: win32con.VK_MENU,
    keyboard.Key.tab: win32con.VK_TAB,
    keyboard.Key.backspace: win32con.VK_BACK,
    keyboard.Key.delete: win32con.VK_DELETE,
    keyboard.Key.space: win32con.VK_SPACE,
    keyboard.Key.left: win32con.VK_LEFT,
    keyboard.Key.right: win32con.VK_RIGHT,
    keyboard.Key.up: win32con.VK_UP,
    keyboard.Key.down: win32con.VK_DOWN,
    keyboard.Key.home: win32con.VK_HOME,
    keyboard.Key.end: win32con.VK_END,
    keyboard.Key.page_up: win32con.VK_PRIOR,
    keyboard.Key.page_down: win32con.VK_NEXT,
    keyboard.Key.f1: win32con.VK_F1,
    keyboard.Key.f2: win32con.VK_F2,
    keyboard.Key.f3: win32con.VK_F3,
    keyboard.Key.f4: win32con.VK_F4,
    keyboard.Key.f5: win32con.VK_F5,
    keyboard.Key.f6: win32con.VK_F6,
    keyboard.Key.f7: win32con.VK_F7,
    keyboard.Key.f8: win32con.VK_F8,
    keyboard.Key.f9: win32con.VK_F9,
    keyboard.Key.f10: win32con.VK_F10,
    keyboard.Key.f11: win32con.VK_F11,
    keyboard.Key.f12: win32con.VK_F12,
    keyboard.Key.caps_lock: win32con.VK_CAPITAL,
    keyboard.Key.num_lock: win32con.VK_NUMLOCK,
    keyboard.Key.scroll_lock: win32con.VK_SCROLL,
    keyboard.Key.print_screen: win32con.VK_SNAPSHOT,
    keyboard.Key.pause: win32con.VK_PAUSE,
    keyboard.Key.insert: win32con.VK_INSERT,
    keyboard.Key.cmd: win32con.VK_LWIN,  # left Windows key
    keyboard.Key.menu: win32con.VK_APPS,
}

def ConvertKeyKeyboardToWin32(key):
    data = None
    if (key != None):
        if isinstance(key, keyboard.KeyCode):
            data = (ord(key.char), win32con.WM_CHAR)
        elif isinstance(key, keyboard.Key):
            data = (special_keys.get(key), win32con.WM_KEYDOWN)
    return data

def ConvertKeyKeyboardToAutoGUI(key):
    data = None
    if (key != None):
        if isinstance(key, keyboard.KeyCode):
            data = key.char
        elif isinstance(key, keyboard.Key):
            data = str(key).split('.')[-1]
    return data

def ConvertKeyMouseToAutoGUI(key):
    data = None
    if (key != None):
        if isinstance(key, mouse.Button):
            data = str(key).split('.')[-1]
    return data
# 후킹 설