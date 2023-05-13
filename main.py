import time
import serial
import pystray
import importlib
import threading
import log_handler
from PIL import Image
from serial import Serial
from keyboard_handler import KeyboardHandler
from com_handler import ComHandler

# Load the icon image
icon_image = Image.open('icon.ico')


# Set up the pystray icon
def setup_icon(icon):
    icon.visible = True


# Listen for incoming data on the COM port and run the appropriate function
ser: Serial

keyboard_handler: KeyboardHandler


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

    keyboard_handler.listener.stop()
    ser.close()

    # Close all threads
    for thread in threading.enumerate():
        if thread != threading.main_thread():
            thread.join()

    # Close all reminding resources
    icon.stop()



def setup_serial():

    # Set up the serial connection to the COM port
    while True:
        try:
            # Re-open the serial connection
            ser = serial.Serial(port='COM4', baudrate=115200, bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                                stopbits=serial.STOPBITS_ONE, timeout=1)
            log_handler.write_to_log(f"Serial connection established")
            return ser

        except serial.serialutil.SerialException:
            log_handler.write_to_log(
                f"Failed to reconnect serial connection, trying again in 10 seconds...")
            time.sleep(10)


def main():
    global ser, keyboard_handler

    ser = setup_serial()
    com_handler = ComHandler(ser)

    keyboard_handler = KeyboardHandler(0.4)


    # Create the pystray menu
    menu = pystray.Menu(pystray.MenuItem("Enable PS5 Mode", com_handler.enable_ps5_mode),
                        pystray.MenuItem("Disable PS5 Mode", com_handler.disable_ps5_mode),
                        pystray.MenuItem("Change audio to headset", com_handler.change_audio_to_headset),
                        pystray.MenuItem("Change audio to speakers", com_handler.change_audio_to_speaker),
                        pystray.MenuItem("Quit", quit_program))

    # Create the pystray icon and start the event loop
    icon = pystray.Icon("MacroPad Listener", icon_image, setup_callback=setup_icon, menu=menu,
                        title="MacroPad Listener")



    # Start the listen_for_input thread
    listen_com_thread = threading.Thread(target=com_handler.listen_for_com_input, args=(icon,))
    listen_com_thread.start()

    print("Booted successfully")



    # Start the listen_for_input thread
    #listen_key_thread = threading.Thread(target=keyboard_handler.listen_for_keyboard_input, args=(icon,))
    #listen_key_thread.start()

    icon.run()


if __name__ == "__main__":
    main()
