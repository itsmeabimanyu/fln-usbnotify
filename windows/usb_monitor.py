# Built-in modules
import datetime
import getpass
import os
import platform
import socket
import subprocess
import threading
import time
import uuid
import win32serviceutil
import win32service
import win32event
import servicemanager
import wmi

# Third-party modules
import psutil
import schedule
import smtplib
from dotenv import load_dotenv

# Email modules (built-in, tapi spesifik)
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Standard library pathlib
from pathlib import Path

current_os = platform.system()

# Memuat isi dari file .env
load_dotenv()
    
def get_logged_in_user():
    try:
        if current_os == "Windows":
            try:
                for proc in psutil.process_iter(['name', 'username']):
                    if proc.info['name'].lower() == 'explorer.exe':
                        return proc.info['username'].split("\\")[-1]
                return None
            except Exception as e:
                print(f"Failed to get logged-in user: {e}")
                return None
        elif current_os in ["Linux", "Darwin"]:  # Darwin = macOS
            output = subprocess.check_output(['who']).decode().strip()
            if output:
                return output.split()[0]
            else:
                return "Unknown"
        else:
            return "Unknown OS"
        
    except Exception as e:
        print(f"Failed to get logged-in user: {e}")
        return "Unknown"

def get_info():
    hostname = socket.gethostname()
    ip = 'Not found'

    # Cari IP pertama yang bukan localhost (127.x.x.x) atau APIPA (169.x.x.x)
    for interface_name, interface_addresses in psutil.net_if_addrs().items():
        for address in interface_addresses:
            if (
                address.family == socket.AF_INET and
                not address.address.startswith('127.') and
                not address.address.startswith('169.')
            ):
                ip = address.address
                break  # Ambil IP pertama yang valid
        if ip != 'Not found':
            break

    # Ambil MAC address
    try:
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> i) & 0xff)
                        for i in range(0, 8*6, 8)][::-1])
    except:
        mac = 'Not found'

    ip_is_illegal = ip and not ip.startswith(os.getenv("VALID_IP_PREFIX"))
    return hostname, ip, mac, ip_is_illegal

# Deteksi dan Tindakan Jika IP Tidak Sah
def check_illegal_ip(prev_ip=None):
    user = get_logged_in_user()
    waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    hostname, ip, mac, ip_is_illegal = get_info()
    # Abaikan jika IP belum ditemukan
    if ip == "Not found":
        return prev_ip

    # Jika IP belum berubah, tidak perlu log ulang
    if prev_ip == ip:
        return ip

    if ip_is_illegal:
        print(f"\n[ALERT] IP Tidak Sah Terdeteksi!")
        print(f"Hostname : {hostname}")
        print(f"IP       : {ip}")
        print(f"MAC      : {mac}")
        print(f"OS       : {current_os}")
        html = f"""
        <html>
        <body>
            <p><strong>[ALERT]</strong> Unauthorized IP detected at <strong>{waktu}</strong> by user: <strong>{user}</strong></p>
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
            <tr style="background-color: #f2f2f2;">
                <th align="left">Field</th>
                <th align="left">Value</th>
            </tr>
            <tr><td>Hostname</td><td>{hostname}</td></tr>
            <tr><td>IP Address</td><td>{ip}</td></tr>
            <tr><td>MAC Address</td><td>{mac}</td></tr>
            <tr><td>OS</td><td>{current_os}</td></tr>
            </table>
        </body>
        </html>
        """
        # Kirim email alert
        send_email("🚨 Alert: Unauthorized IP Detected", html, save_if_failed=False)
    else:
        print(f"[OK] IP valid: {ip}")

    return ip

def ip_monitor_loop():
    print("[INFO] Memulai pemantauan IP (loop polling)...")
    previous_ip = None
    while True:
        previous_ip = check_illegal_ip(previous_ip)
        time.sleep(5)  # Cek setiap 5 detik

def send_email(subject, body, save_if_failed=True):
    from_email = os.getenv("EMAIL_HOST_USER")
    password = os.getenv("EMAIL_HOST_PASSWORD")
    to_email = os.getenv("TO_EMAIL")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.attach(MIMEText(body, "html"))

    try:
        server = smtplib.SMTP(os.getenv("EMAIL_HOST"), (os.getenv("EMAIL_PORT")))
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, password)
        server.sendmail(from_email, [to_email], msg.as_string())
        server.quit()
        print("Email sent successfully.")
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        if save_if_failed:
            save_failed_email(body)
        return False

def get_home_dir(username):
    if username:
        return Path(f"C:/Users/{username}")
    else:
        return Path.home()

def get_temp_dir():
    username = get_logged_in_user()
    if username:
        home_dir = get_home_dir(username)
        if home_dir:
            temp_dir = home_dir / ".cache" / "temp_dir"
            return temp_dir

    # fallback pakai Path.home() (home user proses saat ini)
    return Path.home() / ".cache" / "temp_dir"

# Usage example
TEMP_DIR = get_temp_dir()

def save_failed_email(body):
    os.makedirs(TEMP_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = TEMP_DIR / f"email_{timestamp}.txt"
    
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"Failed email content saved to: {filename}")
    except Exception as e:
        print(f"Error saving failed email: {e}")

# gabung body jadi satu email
def resend_failed_emails():
    if not os.path.exists(TEMP_DIR):
        return

    combined_body = ""
    files_to_delete = []

    for filename in os.listdir(TEMP_DIR):
        file_path = os.path.join(TEMP_DIR, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                body = f.read()
            combined_body += body + "<br><br>"
            files_to_delete.append(file_path)
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    if not combined_body:
        print("No failed emails to resend.")
        return

    # Kirim email gabungan
    success = send_email("🚨 Alert: Unauthorized IP Detected", combined_body, save_if_failed=False)
    if success:
        # Hapus file-file yang sudah berhasil dikirim
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
                print(f"Deleted {file_path}")
            except Exception as e:
                print(f"Failed to delete {file_path}: {e}")
    else:
        print("Resend failed. Files not deleted.")

def win_monitor_usb():
    c = wmi.WMI()

    watcher = c.watch_for(
        notification_type="Creation",
        wmi_class="Win32_PnPEntity"
    )

    while True:
        try:
            device = watcher()  # Tunggu perangkat baru muncul (blocking)
            device_id = device.DeviceID
            name = device.Name
            manufacturer = device.Manufacturer

            service = getattr(device, 'Service', '')
            pnp_class = getattr(device, 'PNPClass', '')

            if (
                (pnp_class == "DiskDrive" and "USBSTOR" in device_id.upper().split("\\")[0]) or
                (pnp_class in ["WPD", "Net"] and "USB" in device_id.upper().split("\\")[0])
            ):
                waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                user = get_logged_in_user()
                hostname, ip, mac, ip_is_illegal = get_info()

                print(f"USB detected at {waktu} by user: {user}")
                print("Name:", name if name else "Unknown")
                print("Manufacturer:", manufacturer if manufacturer else "Unknown")
                print("Serial:", device_id)
                print("USB Attach Event:", pnp_class)
                print(f"OS: {current_os}")
                print(f"Hostname : {hostname}")
                print(f"IP Address : {ip}")
                print(f"MAC Address: {mac}")

                html = f"""
                <html>
                <body>
                    <p>USB detected at <strong>{waktu}</strong> by user: <strong>{user}</strong></p>
                    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
                    <tr style="background-color: #f2f2f2;">
                        <th align="left">Field</th>
                        <th align="left">Value</th>
                    </tr>
                    <tr><td>Name</td><td>{name}</td></tr>
                    <tr><td>Manufacturer</td><td>{manufacturer}</td></tr>
                    <tr><td>Serial</td><td>{device_id}</td></tr>
                    <tr><td>USB Attach Event</td><td>{pnp_class.upper()}</td></tr>
                    <tr><td>Hostname</td><td>{hostname}</td></tr>
                    <tr><td>IP Address</td><td>{ip}</td></tr>
                    <tr><td>MAC Address</td><td>{mac}</td></tr>
                    <tr><td>OS</td><td>{current_os}</td></tr>
                    </table>
                </body>
                </html>
                """

                # Kirim email alert
                send_email("🔔 Alert: USB Device Connection Detected", html)

        except KeyboardInterrupt:
            print("User interrupted with Ctrl+C")
            break

def schedule_runner():
    schedule.every(1).hours.do(resend_failed_emails)

    while True:
        schedule.run_pending()
        time.sleep(1)

class USBWatcherService(win32serviceutil.ServiceFramework):
    _svc_name_ = "USBWatcher"
    _svc_display_name_ = "USB Watcher Python Service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("USB Watcher Service started.")

        threading.Thread(target=schedule_runner, daemon=True).start()
        threading.Thread(target=ip_monitor_loop, daemon=True).start()
        win_monitor_usb()

        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(USBWatcherService)
