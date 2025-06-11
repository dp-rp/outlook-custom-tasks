# Copyright (C) 2024  Denver Pallis
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from math import floor
from random import random
from pynput import keyboard
import threading
import ctypes
from time import sleep

# ---- CONSTANTS ----
INITIAL_SECONDS_TO_WAIT = 5.0
KEY_CANDIDATES = [
    {
        "name": "Caps Lock",
        "pnyput_key_obj": keyboard.Key.caps_lock,
        "virt_key_hex_val": 0x14
    },
    # {
    #     "name": "Scroll Lock",
    #     "pnyput_key_obj": keyboard.Key.scroll_lock,
    #     "virt_key_hex_val": 0x91
    # },
    {
        "name": "Media Volume Mute",
        "pnyput_key_obj": keyboard.Key.media_volume_mute,
        "virt_key_hex_val": 0xAD
    },
    {
        "name": "Num Lock",
        "pnyput_key_obj": keyboard.Key.num_lock,
        "virt_key_hex_val": 0x90
    }
]

# ---- GLOBALS ----
key_pressed = False

# ---- FUNCTIONS ----
def on_press(cancel_timer, target_key):
    def _on_press(key):
        global key_pressed
        try:
            # if key.char == 'n':
            if key == target_key["pnyput_key_obj"]:
                key_pressed = True
                cancel_timer()
                return False  # Stop listener
            else:
                print("wrong key!!")
                cancel_timer()
                return False
        except AttributeError:
            pass
    return _on_press

def on_too_slow():
    print("TOO SLOW!")

def choose_random_key(key_candidates):
    chosen_key = key_candidates[floor(random()*3)]
    return chosen_key

def wait_for_keypress(seconds_to_wait):
    global key_pressed
    
    key_pressed = False
    
    target_key = choose_random_key(KEY_CANDIDATES)
    
    timer = threading.Timer(seconds_to_wait, on_too_slow)

    # press the key to signify it's popped up to hit
    keyboard.Controller().press(target_key["pnyput_key_obj"])
    keyboard.Controller().release(target_key["pnyput_key_obj"])

    listener = keyboard.Listener(on_press=on_press(timer.cancel, target_key))

    listener.start()
    # print("press n!")
    print(f"press {target_key['name']}!... ", end="", flush=True)
    # Wait for 5 seconds
    timer.start()
    timer.join()
    listener.stop()

    # Check if key was pressed
    if key_pressed:
        print("nice!")
    else:
        print("nope!")

    # wait a sec before choosing a new key and toggling it
    # sleep(0.2)

def is_virt_key_on(KEY_HEX_VAL):
    hllDll = ctypes.WinDLL("User32.dll")
    return hllDll.GetKeyState(KEY_HEX_VAL) & 0x0001 != 0

def is_caps_on():
    return is_virt_key_on(KEY_CANDIDATES[0]["virt_key_hex_val"])

def is_scrlk_on():
    return is_virt_key_on(KEY_CANDIDATES[1]["virt_key_hex_val"])

def is_numlk_on():
    return is_virt_key_on(KEY_CANDIDATES[2]["virt_key_hex_val"])

def toggle_all_lock_keys_off():
    if is_caps_on():
        keyboard.Controller().press(keyboard.Key.caps_lock)
        keyboard.Controller().release(keyboard.Key.caps_lock)
    if is_scrlk_on():
        keyboard.Controller().press(keyboard.Key.scroll_lock)
        keyboard.Controller().release(keyboard.Key.scroll_lock)
    if is_numlk_on():
        keyboard.Controller().press(keyboard.Key.num_lock)
        keyboard.Controller().release(keyboard.Key.num_lock)

# ---- classes ----
class KeyGame:
    def __enter__(self):
        # store key states
        # HACK: expects the indicies of caps lock, scroll lock and num lock keys not to move in const array (that might be subject to change, especially if the key options can change with an increase in difficulty)
        self.caps_was_on = is_caps_on()
        self.scrlk_was_on = is_scrlk_on()
        self.numlk_was_on = is_numlk_on()
        # print(
        #     "caps:",
        #     self.caps_was_on,
        #     "scroll:",
        #     self.scrlk_was_on,
        #     "num:",
        #     self.numlk_was_on
        # )

        # turn off lights if they were already on
        toggle_all_lock_keys_off()

    def __exit__(self, exc_type, exc_val, exc_tb):
        # restore key states
        if self.caps_was_on != is_caps_on():
            # print("caps needs toggling back")
            keyboard.Controller().press(keyboard.Key.caps_lock)
            keyboard.Controller().release(keyboard.Key.caps_lock)
        if self.scrlk_was_on != is_scrlk_on():
            # print("scroll needs toggling back")
            keyboard.Controller().press(keyboard.Key.scroll_lock)
            keyboard.Controller().release(keyboard.Key.scroll_lock)
        if self.numlk_was_on != is_numlk_on():
            # print("num needs toggling back")
            keyboard.Controller().press(keyboard.Key.num_lock)
            keyboard.Controller().release(keyboard.Key.num_lock)

def blink_all(wait_after_on, wait_after_off):
    # print("blinking on!")
    keyboard.Controller().press(keyboard.Key.caps_lock)
    keyboard.Controller().release(keyboard.Key.caps_lock)
    keyboard.Controller().press(keyboard.Key.scroll_lock)
    keyboard.Controller().release(keyboard.Key.scroll_lock)
    keyboard.Controller().press(keyboard.Key.num_lock)
    keyboard.Controller().release(keyboard.Key.num_lock)
    sleep(wait_after_on)
    # print("blinking off!")
    keyboard.Controller().press(keyboard.Key.caps_lock)
    keyboard.Controller().release(keyboard.Key.caps_lock)
    keyboard.Controller().press(keyboard.Key.scroll_lock)
    keyboard.Controller().release(keyboard.Key.scroll_lock)
    keyboard.Controller().press(keyboard.Key.num_lock)
    keyboard.Controller().release(keyboard.Key.num_lock)
    sleep(wait_after_off)

# ---- SCRIPT ----
# Call the function
with KeyGame():
    print("Ready?")
    print("3...")
    sleep(1)
    print("2...")
    sleep(1)
    print("1...")
    sleep(1)
    print("Go!")

    time_to_wait = INITIAL_SECONDS_TO_WAIT
    key_pressed = True
    rounds = 0

    while key_pressed == True:
        wait_for_keypress(time_to_wait)
        toggle_all_lock_keys_off()
        sleep(0.2)
        time_to_wait = time_to_wait*0.9
        rounds += 1

    print("rounds:", rounds)


    blink_all(0.2, 0.2)
    blink_all(0.2, 0.2)
    blink_all(0.2, 0.2)
