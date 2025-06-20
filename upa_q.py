import tkinter as tk
import threading
import re
import serial
import serial.tools.list_ports
from pynput.keyboard import Controller, Key
import signal
import sys
import time
import os
import winreg

keyboard = Controller()
ser = None
reading = False
last_data_time = 0
retry_interval = 5

# GUI Setup
root = tk.Tk()
root.title("UPA-Q Serial Tool")

# --- GUI FIELDS ---
def refresh_ports():
    ports = serial.tools.list_ports.comports()
    port_menu['menu'].delete(0, 'end')
    for p in ports:
        port_menu['menu'].add_command(label=p.device, command=tk._setit(port_var, p.device))
    if ports:
        port_var.set(ports[0].device)
    else:
        port_var.set('')

# Port dropdown
port_var = tk.StringVar()
tk.Label(root, text="Port:").grid(row=0, column=0)
port_menu = tk.OptionMenu(root, port_var, '')
port_menu.grid(row=0, column=1, sticky='ew')
refresh_button = tk.Button(root, text="🔄", command=refresh_ports)
refresh_button.grid(row=0, column=2)

# Other serial settings
tk.Label(root, text="Baud:").grid(row=1, column=0)
baud_entry = tk.Entry(root)
baud_entry.insert(0, '9600')
baud_entry.grid(row=1, column=1)

tk.Label(root, text="Encoding:").grid(row=2, column=0)
encoding_entry = tk.Entry(root)
encoding_entry.insert(0, 'ascii')
encoding_entry.grid(row=2, column=1)

tk.Label(root, text="Regex (optional):").grid(row=3, column=0)
regex_entry = tk.Entry(root, width=40)
regex_entry.insert(0, r'^WT\.:\s+([0-9]+\.[0-9]+)')
regex_entry.grid(row=3, column=1, columnspan=2)

send_enter = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="Gửi Enter sau khi nhập", variable=send_enter).grid(row=4, column=0, columnspan=2, sticky='w')

tk.Label(root, text="Timeout (s):").grid(row=5, column=0)
timeout_entry = tk.Entry(root)
timeout_entry.insert(0, '1800')
timeout_entry.grid(row=5, column=1)

# Auto-start checkbox
autostart_var = tk.BooleanVar()
autostart_check = tk.Checkbutton(root, text="Tự động chạy khi mở máy (Windows)", variable=autostart_var, command=lambda: toggle_autostart(autostart_var.get()))
autostart_check.grid(row=6, column=0, columnspan=3, sticky='w')

def toggle_autostart(enable):
    app_name = "UPAQSerialTool"
    exe_path = os.path.realpath(sys.argv[0])
    key = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_ALL_ACCESS) as reg_key:
            if enable:
                winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
                log("✅ Đã bật tự động chạy khi khởi động Windows")
            else:
                try:
                    winreg.DeleteValue(reg_key, app_name)
                    log("❌ Đã tắt tự động chạy khi khởi động Windows")
                except FileNotFoundError:
                    log("ℹ️ Tự động chạy không được bật trước đó")
    except Exception as e:
        log(f"⚠️ Không thể thay đổi cấu hình autostart: {e}")

output = tk.Text(root, height=12, width=60)
output.grid(row=9, column=0, columnspan=3)

# --- FUNCTION ---
def log(msg):
    output.insert(tk.END, msg + "\n")
    output.see(tk.END)
    if int(output.index('end-1c').split('.')[0]) > 500:
        output.delete('1.0', '2.0')

def start_reading():
    global reading
    if reading:
        return
    reading = True
    threading.Thread(target=read_loop, daemon=True).start()

def stop_reading():
    global ser, reading
    reading = False
    if ser and ser.is_open:
        try:
            ser.close()
            log("🔌 Serial port closed")
        except Exception as e:
            log(f"⚠️ Error closing serial port: {e}")

def safe_open_serial():
    global ser
    try:
        ser = serial.Serial(
            port=port_var.get(),
            baudrate=int(baud_entry.get()),
            bytesize=8,
            parity='N',
            stopbits=1,
            timeout=1
        )
        log(f"✅ Connected to {ser.port}")
        return True
    except Exception as e:
        log(f"❌ Error opening serial port: {e}")
        return False

def read_loop():
    global ser, reading, last_data_time
    regex = regex_entry.get()
    encoding = encoding_entry.get()
    timeout_duration = int(timeout_entry.get())
    disconnected_logged = False

    while reading:
        if ser is None or not ser.is_open:
            if not safe_open_serial():
                time.sleep(retry_interval)
                continue
            last_data_time = time.time()
            disconnected_logged = False

        try:
            line = ser.readline().decode(encoding, errors='strict').strip()
            now = time.time()
            if line:
                last_data_time = now
                disconnected_logged = False
                log(f"📥 Received: {repr(line)}")
                if regex.strip():
                    match = re.search(regex, line)
                    if match:
                        try:
                            extracted = match.group(1)
                        except IndexError:
                            extracted = match.group(0)
                        log(f"🎯 Matched: {repr(extracted)}")
                    else:
                        log("⚠️ No match found.")
                        continue
                else:
                    extracted = line
                    log(f"🎯 Full line used: {repr(extracted)}")
                keyboard.type(extracted)
                if send_enter.get():
                    keyboard.press(Key.enter)
                    keyboard.release(Key.enter)
            elif now - last_data_time > timeout_duration and not disconnected_logged:
                log(f"⚠️ No data received for {timeout_duration} seconds. Closing COM.")
                try:
                    ser.close()
                except:
                    pass
                disconnected_logged = True

        except (serial.SerialException, UnicodeDecodeError) as e:
            log(f"⚠️ Serial error: {e}")
            try:
                ser.close()
            except:
                pass
            time.sleep(retry_interval)

        time.sleep(0.1)

def on_closing():
    stop_reading()
    root.destroy()

def signal_handler(sig, frame):
    stop_reading()
    sys.exit(0)

# --- BUTTONS ---
tk.Button(root, text="▶ Start", command=start_reading).grid(row=8, column=0)
tk.Button(root, text="■ Stop", command=stop_reading).grid(row=8, column=1)

# --- INIT ---
refresh_ports()
root.protocol("WM_DELETE_WINDOW", on_closing)
signal.signal(signal.SIGINT, signal_handler)

root.mainloop()
