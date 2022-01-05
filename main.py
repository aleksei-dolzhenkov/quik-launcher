import ctypes
import logging
import os
import psutil
import subprocess
import time

from win_input import send_scancode, send_scancode_down, send_scancode_up


logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] - %(levelname)s - %(message)s',
    datefmt='%d.%m.%y %H:%M:%S'
)
logger = logging.getLogger(__name__)

QUIK_EXE = 'C:\QUIK_VTB\info.exe'
QUIK_DIR = os.path.dirname(QUIK_EXE)
CTRL = 29
Q = 16

BM_CLICK = 0x00F5
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_SHOWWINDOW = 0x0040
WM_SETTEXT = 0x000C


def _str(s: str):
    return s.encode('cp1251')


def find_window_a(class_, name):
    return ctypes.windll.user32.FindWindowA(
        _str(class_) if class_ else None,
        _str(name) if name else None
    )


def check_if_quik_running() -> bool:
    result = False
    for p in psutil.process_iter():
        if p.name() == 'info.exe':
            if p.cwd().lower().startswith(QUIK_DIR.lower()):
                result = True

    logger.info('Checking QUIK running: %s', 'Yes' if result else 'No')

    return result


def run_quik():
    logger.info('Starting QUIK')
    subprocess.Popen([QUIK_EXE], cwd=QUIK_DIR)
    time.sleep(5)


def check_if_logged_in():
    result = False

    if parent_window := find_window_a('InfoClass', None):
        _size = 255
        _buffer = ctypes.create_string_buffer(255)
        ctypes.windll.user32.GetWindowTextA(parent_window, _buffer, _size)
        _title = _buffer.value.decode('cp1251')

        if 'UID:' in _title:
            result = True

    logger.info('Checking connection: %s', 'Yes' if result else 'No')

    return result


def login(username, password):
    if parent_window := find_window_a(None, 'Идентификация пользователя'):
        logger.info('Logging in')
        ctypes.windll.user32.SwitchToThisWindow(parent_window, False)

        h_login = ctypes.windll.user32.FindWindowExA(parent_window, None, _str('Edit'), None)
        h_password = ctypes.windll.user32.FindWindowExA(parent_window, h_login, _str('Edit'), None)
        h_btn_ok = ctypes.windll.user32.FindWindowExA(parent_window, h_password, _str('Button'))

        ctypes.windll.user32.SendMessageA(h_login, WM_SETTEXT, None, _str(username))
        ctypes.windll.user32.SendMessageA(h_password, WM_SETTEXT, None, _str(password))

        ctypes.windll.user32.SetFocus(h_btn_ok)
        ctypes.windll.user32.PostMessageA(h_btn_ok, BM_CLICK, None, None)


def check_if_login_dialog_opened():
    handler = find_window_a(None, 'Идентификация пользователя')
    return handler > 0


def send_ctrl_q():
    send_scancode_down(CTRL)
    send_scancode(Q)
    send_scancode_up(CTRL)


def open_login_dialog():
    logger.info('Open login dialog')

    if parent_window := find_window_a('InfoClass', None):
        ctypes.windll.user32.SwitchToThisWindow(parent_window, False)

        time.sleep(0.3)
        send_ctrl_q()
        time.sleep(0.3)


def main(username, password):
    if not check_if_quik_running():
        run_quik()

    connected = check_if_logged_in()

    if not connected:
        if not check_if_login_dialog_opened():
            open_login_dialog()

        login(username, password)


def run_forever():
    if not os.path.exists('credentials.txt'):
        print('credentials.txt')
        exit(1)

    with open('credentials.txt', 'r', encoding='utf-8') as file:
        username, password = file.read().split('\n')[:2]

    logger.info('Starting')

    try:
        while 1:
            main(username, password)
            time.sleep(10)
    except (KeyboardInterrupt, SystemExit):
        logger.info('Bye')


if __name__ == '__main__':
    run_forever()
