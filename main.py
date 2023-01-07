import ctypes
import pickle
import subprocess
import time
import datetime
import wave

import pyaudio as pyaudio
import serial
import pystray
import win32con
import importlib
import threading
import win32gui

from PIL import Image

# Set up the serial connection to the COM port
while True:
    try:
        # Re-open the serial connection
        ser = serial.Serial(port='COM4', baudrate=115200, bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                            stopbits=serial.STOPBITS_ONE, timeout=1)
        break
    except serial.serialutil.SerialException:
        time.sleep(10)

# Load the icon image
icon_image = Image.open('icon.png')
open_windows_filename = 'open_windows.save'


class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", ctypes.c_long * 4),
        ("rcWork", ctypes.c_long * 4),
        ("dwFlags", ctypes.c_ulong),
    ]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def save_open_windows():
    # Get the list of open windows
    open_windows = []

    def callback(hwnd, open_windows):
        # Check if the window is visible and not minimized
        if win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd):
            # Get the window's placement
            placement = win32gui.GetWindowPlacement(hwnd)

            # Get the window's rectangle
            rect = win32gui.GetWindowRect(hwnd)

            # Create a RECT structure for the window's rectangle
            rect_struct = RECT(*rect)

            # Get the window's monitor info
            monitor_info = MONITORINFO()
            ctypes.windll.user32.GetMonitorInfoW(
                ctypes.windll.user32.MonitorFromRect(ctypes.byref(rect_struct), win32con.MONITOR_DEFAULTTONEAREST),
                ctypes.byref(monitor_info))

            # Compute the window's position and size relative to the primary monitor
            left = rect[0] - monitor_info.rcMonitor[0]
            top = rect[1] - monitor_info.rcMonitor[1]
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]

            # Save the window's handle, placement, and position and size relative to the primary monitor
            open_windows.append((hwnd, placement, (left, top, width, height)))

    win32gui.EnumWindows(callback, open_windows)

    # Save the open_windows variable to a file using pickle
    with open(open_windows_filename, 'wb') as f:
        pickle.dump(open_windows, f)


def restore_open_windows():
    # Load the saved open_windows variable from the file
    with open(open_windows_filename, 'rb') as f:
        open_windows = pickle.load(f)

    for hwnd, placement, rect in open_windows:
        # Check if the window is still open
        if win32gui.IsWindow(hwnd):

            if win32gui.IsIconic(hwnd):
                # Restore the minimized window
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            # Check if the window was maximized when it was saved
            if placement[1] == win32con.SW_MAXIMIZE:
                print("SW_MAXIMIZE")
                # Check if the window is currently maximized
                if win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_MAXIMIZE:
                    # Restore the window
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # Get the window's current size
                rect_struct = RECT()
                ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect_struct))
                width = rect_struct.right - rect_struct.left
                height = rect_struct.bottom - rect_struct.top

                # Get the window's monitor info
                monitor_info = MONITORINFO()
                ctypes.windll.user32.GetMonitorInfoW(
                    ctypes.windll.user32.MonitorFromRect(ctypes.byref(rect_struct),
                                                         win32con.MONITOR_DEFAULTTONEAREST),
                    ctypes.byref(monitor_info))

                # Set the window's position and size
                ctypes.windll.user32.MoveWindow(hwnd, rect[0] + monitor_info.rcMonitor[0],
                                                rect[1] + monitor_info.rcMonitor[1], width, height, True)
                # Maximize the window
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            else:
                # Get the window's monitor info
                rect_struct = RECT(*rect)
                monitor_info = MONITORINFO()
                ctypes.windll.user32.GetMonitorInfoW(ctypes.windll.user32.MonitorFromRect(ctypes.byref(rect_struct),
                                                                                          win32con.MONITOR_DEFAULTTONEAREST),
                                                     ctypes.byref(monitor_info))

                # Set the window's position and size
                ctypes.windll.user32.MoveWindow(hwnd, rect[0] + monitor_info.rcMonitor[0],
                                                rect[1] + monitor_info.rcMonitor[1], rect[2], rect[3], True)


def enable_ps5_mode():
    print("Enable PS5 mode")
    save_open_windows()
    subprocess.run([r"disable_primary.bat"])


def disable_ps5_mode():
    print("Disable PS5 mode")
    subprocess.run([r"enable_primary.bat"])
    restore_open_windows()


def play_sound_through_mic(sound_file, device_name="Microphone (Arctis Pro Wireless"):
    # Open the sound file
    wf = wave.open(sound_file, 'rb')

    # Open a pyaudio stream
    p = pyaudio.PyAudio()

    # Find the index of the output device with the specified name
    output_device_index = None
    for i in range(p.get_device_count()):
        device_info = p.get_device_info_by_index(i)
        if device_info['name'] == device_name:
            output_device_index = device_info['index']
            break

    if output_device_index is None:
        print("Error: Could not find an output device with the name '{}'".format(device_name))
        return

    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True,
                    output_device_index=output_device_index)

    # Read and play the sound data
    data = wf.readframes(1024)
    while len(data) > 0:
        stream.write(data)
        data = wf.readframes(1024)

    # Stop the stream and close it
    stream.stop_stream()
    stream.close()

    # Close the PyAudio object
    p.terminate()


# Define a dictionary of functions to run based on the input code
functions = {
    b'0x01': enable_ps5_mode,
    b'0x02': disable_ps5_mode
}


# Set up the pystray icon
def setup_icon(icon):
    icon.visible = True


# Listen for incoming data on the COM port and run the appropriate function
def listen_for_input(icon, log_file):
    while True:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            data = ser.read(size=4)
            if data:
                # Try to run the function associated with the input code
                try:
                    functions[data]()
                except KeyError:
                    log_file.write("{} - Invalid input code\n".format(timestamp))
        except serial.serialutil.SerialException:
            log_file.write("{} - Serial connection dropped, trying to reconnect...\n".format(timestamp))
            ser.close()
            while True:
                try:
                    # Re-open the serial connection
                    ser.open()
                    log_file.write("{} - Serial connection re-established\n".format(timestamp))
                    break
                except serial.serialutil.SerialException:
                    log_file.write(
                        "{} - Failed to reconnect serial connection, trying again in 10 seconds...\n".format(timestamp))
                    time.sleep(10)
        if not icon.visible:
            break


# Reload the script
def reload_script(icon):
    # Set the icon to not visible to stop the listen_for_input thread
    icon.visible = False
    # Close all threads
    for thread in threading.enumerate():
        if thread != threading.main_thread():
            thread.join()
    # Terminate the pystray icon
    icon.stop()
    # Import the script as a module and execute the main function
    importlib.import_module(__name__).main()


def quit_program(icon):
    # Set the icon to not visible to stop the listen_for_input thread
    icon.visible = False
    # Close all threads
    for thread in threading.enumerate():
        if thread != threading.main_thread():
            thread.join()
    # Terminate the pystray icon
    icon.stop()


def main():
    # Open the log file in append mode
    log_file = open("com_port_listener.log", "a")

    # Create the pystray menu
    menu = pystray.Menu(pystray.MenuItem("Enable PS5 Mode", enable_ps5_mode),
                        pystray.MenuItem("Disable PS5 Mode", disable_ps5_mode),
                        pystray.MenuItem("Quit", quit_program))

    # Create the pystray icon and start the event loop
    icon = pystray.Icon("MacroPad Listener", icon_image, setup_callback=setup_icon, menu=menu,
                        title="MacroPad Listener")

    # Start the listen_for_input thread
    listen_thread = threading.Thread(target=listen_for_input, args=(icon, log_file))
    listen_thread.start()

    icon.run()


if __name__ == "__main__":
    main()
