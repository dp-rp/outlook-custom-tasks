from time import sleep
from random import random
import pyautogui

# GAME IDEA:
# - Randomly show a prompt on screen relating to something that needs to be performed related to the lock keys
#   - toggle the keys off from left to right
#   - whack a mole (after a random number of seconds a key lights up and you need to hit it again to whack it)


from pynput import keyboard

def on_press(key):
    try:
        print('alphanumeric key {0} pressed'.format(
            key.char))
    except AttributeError:
        print('special key {0} pressed'.format(
            key))
        if key == keyboard.Key.scroll_lock:
            print("scroll lock!")

def on_release(key):
    print('{0} released'.format(
        key))
    if key == keyboard.Key.esc:
        # Stop listener
        return False

print("start listening")
listener = keyboard.Listener(
    on_press=on_press,
    on_release=on_release
)
print("now listening")
listener.start()

print("hello")
sleep(2)

listener.join()
print("stop listening")


def whack_a_mole(ctx):
    sleep(random()*2)
    key = random()*3
    if key < 1:
        pyautogui.press('capslock')
    elif key < 2:
        pyautogui.press('scrolllock')
    elif key < 3:
        pyautogui.press('numlock')
    # TODO: IF KEY PRESSED WITHIN N SECONDS, MOLE WHACKED
    # TODO: DO WHATEVER HAPPENS WHEN YOU WHACK A MOLE

whack_a_mole(None)
whack_a_mole(None)
whack_a_mole(None)
whack_a_mole(None)
whack_a_mole(None)

# while True:
    # pyautogui.press('numlock')
    # pyautogui.press('scrolllock')
    # pyautogui.press('capslock')
    # sleep(0.1)
    # pyautogui.press('numlock')
    # pyautogui.press('scrolllock')
    # pyautogui.press('capslock')
    # sleep(0.1)

# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
# sleep(0.1)
# pyautogui.press('numlock')
# pyautogui.press('scrolllock')
