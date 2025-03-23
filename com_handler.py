import ctypes
import os
import pickle
import shutil
import subprocess
import time
import wave
import numpy
import pyaudio
import serial
import win32con
import win32gui
import pyautogui
import log_handler
import re
import pygetwindow as gw
import subprocess
from screeninfo import get_monitors

open_windows_filename = 'open_windows.save'
CREATE_NO_WINDOW = 0x08000000


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
    def __init__(self):
        self.ser = None
        self.monitor_refresh_rate_reduced = False  # maybe pickle and load on boot?
        self.ps5_mode_active = False
        self.replay_save_dir = r'Z:\Storage\Jakob\Gaming clips'


        # Define a dictionary of functions to run based on the input code
        self.functions = {
            b'0x01': self.enable_ps5_mode,
            b'0x02': self.disable_ps5_mode,
            b'0x03': self.change_audio_to_headset,
            b'0x04': self.change_audio_to_speaker,
            b'0x05': self.switch_refresh_rate,
            b'0x06': self.install_vencord,
            b'0x07': self.restart_discord_stream,
            b'0x08': self.save_replay,
            b'0x09': self.switch_monitor_mode,
            b'0x0a': self.change_audio_to_vr,
            b'0x0b': self.change_audio_to_rtx,
        }

    def setup_serial(self, icon):

        # Set up the serial connection to the COM port
        while True:
            try:
                # Re-open the serial connection
                self.ser = serial.Serial(port='COM4', baudrate=115200, bytesize=serial.EIGHTBITS,
                                         parity=serial.PARITY_NONE,
                                         stopbits=serial.STOPBITS_ONE, timeout=1)
                log_handler.write_to_log(f"Serial connection established")
                break

            except serial.serialutil.SerialException:
                log_handler.write_to_log(
                    f"Failed to reconnect serial connection, trying again in 10 seconds...")
                time.sleep(10)

            if not icon.visible:
                break

    def establish_com_connection(self, icon):
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

    def listen_for_com_input(self, icon):
        log_handler.write_to_log("hi")
        if self.ser is None:
            self.setup_serial(icon)

        while True:

            try:
                data = self.ser.read(size=4)

                if data:
                    log_handler.write_to_log(f"Got input from MacroPad: {data}")
                    # Try to run the function associated with the input code
                    try:
                        # hack to convert a string to function call because functions pointers are unavailable
                        self.functions[data]()
                        log_handler.write_to_log(f"test")
                    except KeyError:
                        log_handler.write_to_log(f"Invalid input code")

            except serial.serialutil.SerialException:
                log_handler.write_to_log(f"Serial connection dropped, trying to reconnect...")
                self.ser.close()
                self.establish_com_connection(icon)

            if not icon.visible:
                break

    def enable_ps5_mode(self):
        log_handler.write_to_log(f"Enable PS5 mode")
        print("Enable PS5 mode")
        save_open_windows()
        subprocess.run([r"disable_primary.bat"], creationflags=CREATE_NO_WINDOW)
        self.ps5_mode_active = True

    def disable_ps5_mode(self):
        log_handler.write_to_log(f"Disable PS5 mode")
        print("Disable PS5 mode")
        subprocess.run([r"enable_primary.bat"], creationflags=CREATE_NO_WINDOW)
        restore_open_windows()
        self.ps5_mode_active = False

    def tv_mode(self):
        self.switch_monitor_mode(r"\\\\.\\DISPLAY2")

    def change_audio_to_headset(self):
        microphone_name = "Headset Mic"
        audio_name = "Headset"

        self.set_default_sound_device(audio_name, microphone_name)

        log_handler.write_to_log(f"Changing default audio device to headset")
        print(f"Changing default audio device to headset")

    def change_audio_to_speaker(self):
        microphone_name = "Headset Mic"
        audio_name = "Speakers"

        self.set_default_sound_device(audio_name, microphone_name)

        log_handler.write_to_log(f"Changing default audio device to speakers")
        print(f"Changing default audio device to speakers")

    def change_audio_to_vr(self):
        microphone_name = "VR Mic"
        audio_name = "VR"

        self.set_default_sound_device(audio_name, microphone_name)

        log_handler.write_to_log(f"Changing default audio device to VR")
        print(f"Changing default audio device to VR")

    def change_audio_to_rtx(self):
        microphone_name = "RTX Mic"
        audio_name = "RTX"

        self.set_default_sound_device(audio_name, microphone_name)

        log_handler.write_to_log(f"Changing default audio device to RTX")
        print(f"Changing default audio device to RTX")

    def set_default_sound_device(self, audio_name, microphone_name):
        if audio_name is not None:
            subprocess.run(f'nircmd setdefaultsounddevice {audio_name} 1', stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
            subprocess.run(f'nircmd setdefaultsounddevice {audio_name} 2', stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)

        if microphone_name is not None:
            subprocess.run(f'nircmd setdefaultsounddevice "{microphone_name}" 1', stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
            subprocess.run(f'nircmd setdefaultsounddevice "{microphone_name}" 2', stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)


    def switch_refresh_rate(self):
        if self.monitor_refresh_rate_reduced:
            result = subprocess.run('nircmd setdisplay 2560 1440 32 240', stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
            self.monitor_refresh_rate_reduced = False
        else:
            result = subprocess.run('nircmd setdisplay 2560 1440 32 144', stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
            self.monitor_refresh_rate_reduced = True

        log_handler.write_to_log(f"Changing default audio device to speakers")
        print(f"Changing default audio device to speakers")

    def install_vencord(self):
        subprocess.run([r"install_vencord.bat"])

        log_handler.write_to_log(f"Installing Vencord")
        print(f"Installing Vencord")

    def restart_discord_stream(self):

        button1_x, button1_y = 0, 0
        button2_x, button2_y = 0, 0

        if self.ps5_mode_active is True:
            # Coordinates of the first and second button
            button1_x, button1_y = 287, 912
            button2_x, button2_y = 163, 1000

        elif self.ps5_mode_active is False:
            #Coordinates of the first and second button
            button1_x, button1_y = -1634, 274
            button2_x, button2_y = -1759, 369

        # Try to locate the Discord window
        discord_windows = [window for window in pyautogui.getWindowsWithTitle('Discord') if "- Discord" in window.title]
        # Check if Discord window found
        if discord_windows:
            discord_window = discord_windows[0]
            log_handler.write_to_log(discord_windows[0])
            #discord_window.activate()

            # Bring the Discord window to the foreground
            discord_window.maximize()  # Maximize the window (optional)
            #discord_window.activate()

            # Move the mouse and click on the first button
            pyautogui.click(button1_x, button1_y)

            # Move the mouse and click on the second button
            pyautogui.click(button2_x, button2_y)

            print(f"Restarted the Discord stream")
        else:
            log_handler.write_to_log(f"Discord window not found")
            print("Discord window not found")

    def get_audio_devices(self) -> dict:
        audio_devices = {}
        result = subprocess.run([r"EndPointController.exe", r"-f", "\"%d %s\""], stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
        outputs = [device.strip() for device in result.stdout.decode("utf-8").split("\n") if device != ""]

        for device in outputs:
            index = device.find("(")
            audio_devices[device[1]] = device[index:]

        return audio_devices

    def set_audio_device(self, audio_devices, audio_device_to_set):

        for key, value in audio_devices.items():
            if audio_device_to_set in value:
                result = subprocess.run([r"EndPointController.exe", key], stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)

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

    # Function to trigger OBS "Save Replay" shortcut
    def trigger_save_replay(self):
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('f6')
        pyautogui.keyUp('f6')
        pyautogui.keyUp('ctrl')

    # Function to get the title of the active window
    def get_active_window_title(self):
        active_window = gw.getActiveWindow()
        return active_window.title

    # Wait for OBS to save the file
    def wait_for_save_completion(self):
        max_loops = 30  # Maximum number of loops to wait for the file
        loops = 0
        while loops < max_loops:
            newest_file = max(os.listdir(self.replay_save_dir), key=lambda f: os.path.getctime(os.path.join(self.replay_save_dir, f)))
            if re.match(r'Replay \d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2}\.mkv', newest_file):
                file_path = os.path.join(self.replay_save_dir, newest_file)
                while True:
                    initial_mod_time = os.path.getmtime(file_path)
                    time.sleep(2)  # Adjust sleep duration as needed
                    current_mod_time = os.path.getmtime(file_path)
                    if current_mod_time == initial_mod_time:
                        return newest_file
            loops += 1
            time.sleep(2)  # Adjust sleep duration as needed
        return None

    # Rename the file with the active window title and move it to a folder with the same name
    def rename_and_move_file(self, file_name, new_name):
        src_path = os.path.join(self.replay_save_dir, file_name)
        dst_folder = os.path.join(self.replay_save_dir, new_name)
        dst_path = os.path.join(dst_folder, file_name)

        # Create destination folder if it doesn't exist
        if not os.path.exists(dst_folder):
            os.makedirs(dst_folder)

        # Move the file to the destination folder
        shutil.move(src_path, dst_path)

        # Extract timestamp from the filename
        timestamp_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2}', file_name)
        if timestamp_match:
            timestamp = timestamp_match.group(0).replace(' - ', ' - ').replace('-', '.').replace(' ', ' - ')
            new_file_name = f"{new_name} - {timestamp}.mkv"
            new_file_path = os.path.join(dst_folder, new_file_name)
            os.rename(dst_path, new_file_path)
        else:
            print("Timestamp not found in filename. Skipping renaming.")

    # Main function
    def save_replay(self):
        self.trigger_save_replay()
        file_name = self.wait_for_save_completion()
        if file_name:
            active_window_title = self.get_active_window_title()
            self.rename_and_move_file(file_name, active_window_title)
            log_handler.write_to_log("Saved, renamed and moved file successfully")
            print("Saved, renamed and moved file successfully")
        else:
            log_handler.write_to_log("File not found within the specified timeframe")
            print("File not found within the specified timeframe")

#--------------------------------------------------------------------------------------------------------

    def get_connected_monitors(self):
        """Returns a list of connected monitor names."""
        monitors = get_monitors()
        return [monitor.name for monitor in monitors]


    def is_monitor_enabled(self, monitor_name):
        """Check if a monitor is currently enabled via xrandr."""
        try:
            # Use xrandr to get the list of enabled monitors
            output = subprocess.check_output("xrandr --listmonitors", shell=True).decode("utf-8")
            return monitor_name in output
        except subprocess.CalledProcessError as e:
            print(f"Failed to get monitor information: {e}")
            return False


    def disable_monitor(self, monitor_name):
        """Disable the specified monitor."""
        try:
            subprocess.run(f"xrandr --output {monitor_name} --off", shell=True, check=True)
            print(f"Disabled monitor: {monitor_name}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to disable monitor {monitor_name}: {e}")


    def enable_monitor(self, monitor_name):
        """Enable the specified monitor."""
        try:
            subprocess.run(f"xrandr --output {monitor_name} --auto", shell=True, check=True)
            print(f"Enabled monitor: {monitor_name}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to enable monitor {monitor_name}: {e}")


    def switch_monitor_mode(self, monitor_name):
        monitor_to_enabled_disable = monitor_name  # Adjust based on your setup
        connected_monitors = self.get_connected_monitors()
        print(f"Connected monitors: {connected_monitors}")

        if monitor_to_enabled_disable in connected_monitors:
            # If the specific monitor is connected and enabled
            if self.is_monitor_enabled(monitor_to_enabled_disable):
                # Disable all other monitors
                for monitor in connected_monitors:
                    if monitor != monitor_to_enabled_disable:
                        self.disable_monitor(monitor)
            elif len(connected_monitors) == 1:
                # If the specific monitor is the only monitor connected, re-enable others
                for monitor in connected_monitors:
                    if monitor != monitor_to_enabled_disable:
                        self.enable_monitor(monitor)
        else:
            print(f"Monitor {monitor_to_enabled_disable} is not connected.")

