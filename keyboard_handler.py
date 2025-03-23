import win32gui
from datetime import datetime
from pynput import keyboard
from pynput.keyboard import Key, Controller
from vk_codes_enum import VKCodes


class KeyboardHandler:
    def __init__(self, media_key_threshold: float, windows_key_threshold: float):
        self.media_key_threshold = media_key_threshold
        self.windows_key_threshold = windows_key_threshold
        self.last_ts = None
        self.datetime = datetime.now()
        self.timestamp = self.datetime.timestamp()
        self.listener = keyboard.Listener(
            on_press=self.on_press,
            on_release=self.on_release,
            #win32_event_filter=self.win32_event_filter,
            suppress=False
        )
        self.listener.start()
        self.keyboard = Controller()

    def handle_media_key(self):
        self.datetime = datetime.now()
        self.timestamp = self.datetime.timestamp()

        if self.last_ts is not None and self.timestamp - self.last_ts < self.media_key_threshold:
            self.keyboard.press(Key.media_next)
            self.timestamp = datetime.now().timestamp()
            self.last_ts = None

        self.last_ts = self.datetime.timestamp()

        # self.keyboard.press(Key.media_play_pause)
        return

    def handle_windows_key(self):
        self.datetime = datetime.now()
        self.timestamp = self.datetime.timestamp()

        if self.last_ts is not None and self.timestamp - self.last_ts < self.media_key_threshold:
            self.keyboard.press(Key.media_next)
            self.timestamp = datetime.now().timestamp()
            self.last_ts = None

        self.last_ts = self.datetime.timestamp()

        # self.keyboard.press(Key.media_play_pause)
        return

    def is_window_in_focus(self, window_text):
        current_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        #print(current_window)
        if window_text.lower() in current_window.lower():
            #print("true")
            return True

        # Get a list of all open windows with the text "Brave" in the title
        return False

    def win32_event_filter(self, msg, data):
        # print(data.vkCode)
        if data.vkCode == VKCodes.VK_VOLUME_DOWN or data.vkCode == VKCodes.VK_VOLUME_UP:
            if self.is_window_in_focus("brave"):
                #print("Suppressing F1 up")
                self.listener._suppress = True
                return True
            # return False # if you return False, your on_press/on_release will not be called
            else:
                #print("elif data.vkCode == 174 or data.vkCode == 175")
                self.listener._suppress = False
                return True

        self.listener._suppress = False
        return True

    def on_press(self, key):
        #print(key)
        if key is Key.media_play_pause:
            self.handle_media_key()

        # windows key
        if key is Key.cmd:
            #self.handle_windows_key()
            pass

        #if key is Key.media_volume_down:
        #    with self.keyboard.pressed(Key.ctrl):
        #        self.keyboard.press(Key.tab)
        #        self.keyboard.release(Key.tab)

        #if key is Key.media_volume_up:
        #    with self.keyboard.pressed(Key.ctrl):
        #        self.keyboard.tap(Key.tab)

    def on_release(self, key):
        pass


