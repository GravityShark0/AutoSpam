# Import
from pynput.keyboard import Key, Controller
import pywintypes
import win32gui
import win32con
import pickle
import time
import os


# Clear Function
def cls():
    os.system('cls')


# Fake Startup
print("Opening AutoSpam.")
time.sleep(1)
cls()
print("Opening AutoSpam..")
time.sleep(1)
cls()
print("Opening AutoSpam...")
time.sleep(1)
cls()
print("Welcome to AutoSpam")
time.sleep(1)
print("Follow the instructions to setup the program")
time.sleep(1)
cls()

# Globals

# Variables/Settings
keyboard = Controller()


def AutoSpam():
    global timer, word, amount, inf
    cls()
    saves = str(
        input("Use saved settings? [y/n]\n[if you don't have saved settings it will automatically use the default "
              "settings]\n:")).lower()
    if saves == "y":
        word, amount, timer = pickle.load(open("settings.dat", "rb"))
        cls()
        print("Settings Loaded")
        time.sleep(1)
    elif saves == "n":
        cls()
        word = str(input("1.Type the word you want to spam.\n:"))
        cls()
        amount = input("2.Type how many times to repeat the word.\n[Type inf for infinity]\n:")
        cls()
        timer = int(
            input("3.Type how long for word to be sent\n(Measured in Seconds)\n:"))
        cls()
        save = str(input("Save settings? [y/n]\n:")).lower()
        if save == "y":
            pickle.dump([word, amount, timer], open("settings.dat", "wb"))
            print("Settings Saved.")
    cls()
    correction = str(input("Are you sure these are the correct settings? [y/n]\n1." + str(word) + "\n2." + str(amount) +
                           "\n3." + str(timer) + "\n:")).lower()
    if correction == "n":
        AutoSpam()

    # Tes
    ass = "notepad"

    # Inf Check
    if amount == "inf":
        amount = 1
        inf = True

    # Countdown Loop
    time.sleep(1)
    cls()
    input("Press Enter to Start\n")
    cls()
    print("AutoSpam is starting in")
    s = 5
    while s != 0:
        print(s)
        s -= 1
        time.sleep(1)

    # Message
    cls()
    print("AutoSpam is Running...")
    if inf:
        print("inf is enabled exit the program to stop it")

    def windowEnumerationHandler(hwnd, top_windows):
        global amount
        top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

    # While Loop
    while int(amount) > 0:
        if __name__ == "__main__":
            results = []
            top_windows = []
            win32gui.EnumWindows(windowEnumerationHandler, top_windows)
            for i in top_windows:
                if ass in i[1].lower():
                    win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
                    win32gui.SetForegroundWindow(i[0])
                    for c in word:
                        keyboard.press(c)
                        keyboard.release(c)
                        time.sleep(0.1)
                    keyboard.press(Key.enter)
                    keyboard.release(Key.enter)
                    win32gui.ShowWindow(i[0], win32con.SW_MINIMIZE)
                    time.sleep(timer)
                    if not inf:
                        amount -= 1

    # Restart
    continues = str(input("Do you want to redo? [y/n]\n:"))
    if continues == "y":
        AutoSpam()


AutoSpam()

# Quit
cls()
print("AutoSpam finished action")
time.sleep(1)
cls()
print("Closing AutoSpam.")
time.sleep(3)
quit()
