import sys, time, re, ctypes
import serial
from serial.tools import list_ports

PORT  = sys.argv[1] if len(sys.argv) > 1 else "COM8"
BAUD  = int(sys.argv[2]) if len(sys.argv) > 2 else 115200
MOUSE_SENS = 1.0

user32 = ctypes.windll.user32

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP  = 0x0002

MOUSEEVENTF_MOVE      = 0x0001
MOUSEEVENTF_LEFTDOWN  = 0x0002
MOUSEEVENTF_LEFTUP    = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP   = 0x0010

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_uint32),
                ("dwFlags", ctypes.c_uint32),
                ("time", ctypes.c_uint32),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_uint32),
                ("time", ctypes.c_uint32),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)))

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg", ctypes.c_uint32),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort))

class INPUT(ctypes.Structure):
    class _I(ctypes.Union):
        _fields_ = (("mi", MOUSEINPUT),
                    ("ki", KEYBDINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("i",)
    _fields_ = (("type", ctypes.c_uint32),
                ("i", _I))

def send_unicode(ch):
    arr = (INPUT * 2)()
    arr[0].type = INPUT_KEYBOARD
    arr[0].ki = KEYBDINPUT(0, ord(ch), KEYEVENTF_UNICODE, 0, None)
    arr[1].type = INPUT_KEYBOARD
    arr[1].ki = KEYBDINPUT(0, ord(ch), KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, None)
    user32.SendInput(2, ctypes.byref(arr), ctypes.sizeof(INPUT))

def send_vk(vk):
    arr = (INPUT * 2)()
    arr[0].type = INPUT_KEYBOARD
    arr[0].ki = KEYBDINPUT(vk, 0, 0, 0, None)
    arr[1].type = INPUT_KEYBOARD
    arr[1].ki = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, None)
    user32.SendInput(2, ctypes.byref(arr), ctypes.sizeof(INPUT))

def mouse_move(dx, dy):
    dx = int(dx * MOUSE_SENS)
    dy = int(dy * MOUSE_SENS)
    if dx == 0 and dy == 0:
        return
    arr = INPUT()
    arr.type = INPUT_MOUSE
    arr.mi = MOUSEINPUT(dx, dy, 0, MOUSEEVENTF_MOVE, 0, None)
    user32.SendInput(1, ctypes.byref(arr), ctypes.sizeof(INPUT))

def mouse_press(down, right=False):
    arr = INPUT()
    arr.type = INPUT_MOUSE
    flag = MOUSEEVENTF_RIGHTDOWN if right else MOUSEEVENTF_LEFTDOWN
    if not down:
        flag = MOUSEEVENTF_RIGHTUP if right else MOUSEEVENTF_LEFTUP
    arr.mi = MOUSEINPUT(0, 0, 0, flag, 0, None)
    user32.SendInput(1, ctypes.byref(arr), ctypes.sizeof(INPUT))

def open_port(port, baud):
    try:
        ser = serial.Serial(
            port, baudrate=baud,
            bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE, timeout=0.01,
            xonxoff=False, rtscts=False, dsrdtr=False
        )
        ser.dtr = False
        ser.rts = False
        time.sleep(0.15)
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        return ser
    except Exception as e:
        print(f"[ERROR] Open {port} failed: {e}")
        print("Ports:", [p.device for p in list_ports.comports()])
        sys.exit(1)

def main():
    ser = open_port(PORT, BAUD)
    print(f"[ESP32 HID Bridge Active on {PORT}]")
    print("Listening for Keyboard/Mouse... Ctrl+C to exit\n")

    mode = ""
    token = ""
    at_line_start = True
    skipping_log = False
    mouse_line = ""
    last_x = None; last_y = None
    left_down = False; right_down = False

    mouse_re = re.compile(r"X:\s*([+-]?\d+).*?Y:\s*([+-]?\d+).*?\|(.?)\|(.?)\|?")

    while True:
        try:
            n = ser.in_waiting
            if n == 0:
                time.sleep(0.002)
                continue
            chunk = ser.read(n)
            for b in chunk:
                ch = chr(b)

                if at_line_start and ch in ("I","W","E"):
                    skipping_log = True

                # 改动点：优先处理换行/回车
                if ch in ("\n", "\r"):
                    if mode == "Keyboard":
                        if ch == "\r":
                            send_vk(0x0D)  # Enter
                        at_line_start = True
                        skipping_log = False
                        continue
                    at_line_start = True
                    skipping_log = False
                    if mode == "Mouse" and mouse_line:
                        m = mouse_re.search(mouse_line)
                        mouse_line = ""
                        if m:
                            x = int(m.group(1))
                            y = int(m.group(2))
                            b1 = (m.group(3) == 'o')
                            b2 = (m.group(4) == 'o')

                            if last_x is not None and last_y is not None:
                                mouse_move(x - last_x, y - last_y)
                            last_x, last_y = x, y

                            if b1 != left_down:
                                mouse_press(b1, right=False)
                                left_down = b1
                            if b2 != right_down:
                                mouse_press(b2, right=True)
                                right_down = b2
                    continue
                else:
                    at_line_start = False

                token = (token + ch)[-9:]
                if "Keyboard" in token:
                    mode = "Keyboard"
                    last_x = last_y = None
                    continue
                if "Mouse" in token:
                    mode = "Mouse"
                    continue

                if mode == "Mouse":
                    mouse_line += ch
                    if len(mouse_line) > 200:
                        mouse_line = mouse_line[-200:]
                    continue

                if mode == "Keyboard":
                    if skipping_log:
                        continue
                    if ch == '\b':
                        send_vk(0x08)
                    elif ch == '\t':
                        send_vk(0x09)
                    elif 32 <= ord(ch) <= 126:
                        send_unicode(ch)
                    else:
                        pass
                    continue

        except KeyboardInterrupt:
            print("\n[EXIT] User cancelled.")
            break
        except Exception as e:
            print("[ERR]", e)
            time.sleep(0.05)

if __name__ == "__main__":
    main()
