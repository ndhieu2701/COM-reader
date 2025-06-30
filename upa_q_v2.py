import tkinter as tk
from ttkbootstrap import Style
from ttkbootstrap.constants import *
import ttkbootstrap as ttk
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
from tkinter import ttk as native_ttk

keyboard = Controller()
ser = None
reading = False
last_data_time = 0
retry_interval = 5

# Setup style
root = tk.Tk()
style = Style("flatly")
root.title("UPA-Q Serial Tool")
root.geometry("800x540")
root.minsize(800, 540)

main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill='both', expand=True)

# --- GUI FIELDS ---
def refresh_ports():
    ports = [p.device for p in serial.tools.list_ports.comports()]
    port_combo['values'] = ports
    if ports:
        port_var.set(ports[0])
    else:
        port_var.set('')

row = 0
port_var = tk.StringVar()
ttk.Label(main_frame, text="Port:").grid(row=row, column=0, sticky='e', pady=4)
port_select_frame = ttk.Frame(main_frame)
port_select_frame.grid(row=row, column=1, columnspan=2, sticky='ew', pady=4)
port_select_frame.columnconfigure(0, weight=1)
port_combo = native_ttk.Combobox(port_select_frame, textvariable=port_var, state='readonly')
port_combo.grid(row=0, column=0, sticky='ew')
ttk.Button(port_select_frame, text="🔄 Làm mới cổng COM", command=refresh_ports, bootstyle=SECONDARY).grid(row=0, column=1, padx=5)

row += 1
ttk.Label(main_frame, text="Baud:").grid(row=row, column=0, sticky='e', pady=4)
baud_entry = ttk.Entry(main_frame)
baud_entry.insert(0, '9600')
baud_entry.grid(row=row, column=1, sticky='ew', pady=4)

row += 1
ttk.Label(main_frame, text="Encoding:").grid(row=row, column=0, sticky='e', pady=4)
encoding_entry = ttk.Entry(main_frame)
encoding_entry.insert(0, 'ascii')
encoding_entry.grid(row=row, column=1, sticky='ew', pady=4)

row += 1
ttk.Label(main_frame, text="Regex (optional):").grid(row=row, column=0, sticky='e', pady=4)
regex_entry = ttk.Entry(main_frame, width=40)
regex_entry.insert(0, r'^(?:WT\.|T\.)\:\s*([0-9]+(?:\.[0-9]+)?)')
regex_entry.grid(row=row, column=1, columnspan=2, sticky='ew', pady=4)

row += 1
send_enter = tk.BooleanVar(value=False)
ttk.Checkbutton(main_frame, text="Gửi Enter sau khi nhập", variable=send_enter, bootstyle="success-round-toggle").grid(row=row, column=0, columnspan=3, sticky='w', pady=4)

row += 1
ttk.Label(main_frame, text="Timeout (s):").grid(row=row, column=0, sticky='e', pady=4)
timeout_entry = ttk.Entry(main_frame)
timeout_entry.insert(0, '1800')
timeout_entry.grid(row=row, column=1, sticky='ew', pady=4)

row += 1
autostart_var = tk.BooleanVar()
autostart_check = ttk.Checkbutton(main_frame, text="Tự động chạy khi mở máy (Windows)", variable=autostart_var, command=lambda: toggle_autostart(autostart_var.get()), bootstyle="info-round-toggle")
autostart_check.grid(row=row, column=0, columnspan=3, sticky='w', pady=4)

row += 1
btn_frame = ttk.Frame(main_frame)
btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
ttk.Button(btn_frame, text="▶ Start", command=lambda: start_reading(), bootstyle="success").pack(side='left', padx=8)
ttk.Button(btn_frame, text="■ Stop", command=lambda: stop_reading(), bootstyle="danger").pack(side='left', padx=8)
ttk.Button(btn_frame, text="🧹 Clear Log", command=lambda: clear_log(), bootstyle="secondary").pack(side='left', padx=8)

row += 1
text_frame = ttk.Frame(main_frame)
text_frame.grid(row=row, column=0, columnspan=3, sticky='nsew')
main_frame.rowconfigure(row, weight=1)
main_frame.columnconfigure(1, weight=1)

scrollbar = ttk.Scrollbar(text_frame)
scrollbar.pack(side="right", fill="y")

output = tk.Text(text_frame, height=12, wrap="word", bg="black", fg="lime", font=("Courier", 10), yscrollcommand=scrollbar.set)
output.pack(fill="both", expand=True)
output.config(state="disabled")
scrollbar.config(command=output.yview)

# --- FUNCTION ---
def log(msg):
    output.config(state="normal")
    output.insert(tk.END, msg + "\n")
    output.see(tk.END)
    if int(output.index('end-1c').split('.')[0]) > 500:
        output.delete('1.0', '2.0')
    output.config(state="disabled")

def clear_log():
    output.config(state="normal")
    output.delete("1.0", tk.END)
    output.config(state="disabled")

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
            cleaned_line = re.sub(r'\s+', ' ', line).strip()  # Chuẩn hóa khoảng trắng
            now = time.time()
            if line:
                last_data_time = now
                disconnected_logged = False
                log(f"📥 Received: {repr(cleaned_line)}")
                if regex.strip():
                    match = re.search(regex, cleaned_line)
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
                    extracted = cleaned_line
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

def check_autostart_enabled():
    app_name = "UPAQSerialTool"
    key = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ) as reg_key:
            value, _ = winreg.QueryValueEx(reg_key, app_name)
            if value:
                autostart_var.set(True)
    except FileNotFoundError:
        autostart_var.set(False)
    except Exception as e:
        log(f"⚠️ Không thể kiểm tra autostart: {e}")

def on_closing():
    stop_reading()
    root.destroy()

def signal_handler(sig, frame):
    stop_reading()
    sys.exit(0)

refresh_ports()
check_autostart_enabled()
root.protocol("WM_DELETE_WINDOW", on_closing)
signal.signal(signal.SIGINT, signal_handler)

root.mainloop()