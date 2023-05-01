import ctypes
import pickle
import subprocess
import time
import wave
import numpy
import pyaudio
import serial
import win32con
import win32gui
from serial import Serial

import log_handler

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
                ctypes.windll.user32.GetMonitorInfoW(
                    ctypes.windll.user32.MonitorFromRect(ctypes.byref(rect_struct),
                                                         win32con.MONITOR_DEFAULTTONEAREST),
                    ctypes.byref(monitor_info))

                # Set the window's position and size
                ctypes.windll.user32.MoveWindow(hwnd, rect[0] + monitor_info.rcMonitor[0],
                                                rect[1] + monitor_info.rcMonitor[1], rect[2], rect[3], True)


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


class ComHandler:
    def __init__(self, ser: Serial):
        self.ser = ser

        # Define a dictionary of functions to run based on the input code
        self.functions = {
            b'0x01': self.enable_ps5_mode,
            b'0x02': self.disable_ps5_mode,
            b'0x03': self.change_audio_to_headset,
            b'0x04': self.change_audio_to_speaker
        }

    def listen_for_com_input(self, icon):

        while True:

            try:
                data = self.ser.read(size=4)
                if data:
                    log_handler.write_to_log(f"Got input from MacroPad: {data}")
                    # Try to run the function associated with the input code
                    try:
                        # hack to convert a string to function call because functions pointers are unavailable
                        self.functions[data]()
                    except KeyError:
                        log_handler.write_to_log(f"Invalid input code")

            except serial.serialutil.SerialException:
                log_handler.write_to_log(f"Serial connection dropped, trying to reconnect...")
                self.ser.close()

                while True:
                    try:
                        # Re-open the serial connection
                        self.ser.open()
                        log_handler.write_to_log(f"Serial connection re-established")
                        break

                    except serial.serialutil.SerialException:
                        log_handler.write_to_log(
                            f"Failed to reconnect serial connection, trying again in 10 seconds...")
                        time.sleep(10)

                    if not icon.visible:
                        break

            if not icon.visible:
                break

    def enable_ps5_mode(self):
        log_handler.write_to_log(f"Enable PS5 mode")
        print("Enable PS5 mode")
        save_open_windows()
        subprocess.run([r"disable_primary.bat"])

    def disable_ps5_mode(self):
        log_handler.write_to_log(f"Disable PS5 mode")
        print("Disable PS5 mode")
        subprocess.run([r"enable_primary.bat"])
        restore_open_windows()

    def change_audio_to_headset(self):
        audio_device_to_set = "Arctis Pro Wireless"
        log_handler.write_to_log(f"Changing default audio device to {audio_device_to_set}")
        print(f"Changing default audio device to {audio_device_to_set}")
        audio_devices = self.get_audio_devices()

        self.set_audio_device(audio_devices, audio_device_to_set)

    def change_audio_to_speaker(self):
        audio_device_to_set = "Realtek(R) Audio"
        log_handler.write_to_log(f"Changing default audio device to {audio_device_to_set}")
        print(f"Changing default audio device to {audio_device_to_set}")
        audio_devices = self.get_audio_devices()

        self.set_audio_device(audio_devices, audio_device_to_set)

    def get_audio_devices(self) -> dict:
        audio_devices = {}
        result = subprocess.run([r"EndPointController.exe", r"-f", "\"%d %s\""], stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE)
        outputs = [device.strip() for device in result.stdout.decode("utf-8").split("\n") if device != ""]
        print(result)
        for device in outputs:
            index = device.find("(")
            audio_devices[device[1]] = device[index:]

        return audio_devices

    def set_audio_device(self, audio_devices, audio_device_to_set):

        for key, value in audio_devices.items():
            print(f"{key}:{value}")
            if audio_device_to_set in value:
                print(value)
                result = subprocess.run([r"EndPointController.exe", key], stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE)
    def audio_datalist_set_volume(self, datalist, volume):
        """ Change value of list of audio chunks """
        sound_level = (volume / 100.)

        for i in range(len(datalist)):
            chunk = numpy.fromstring(datalist[i], numpy.int16)

            chunk = chunk * sound_level

            datalist[i] = chunk.astype(numpy.int16)

    def play_sound_through_mic(self, sound_file='test.mp3', volume=1.0, device_name="Microphone (Arctis Pro Wireless"):
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

        while data:
            stream.write(data)
            data = wf.readframes(1024)

        # Stop the stream and close it
        stream.stop_stream()
        stream.close()

        # Close the PyAudio object
        p.terminate()


