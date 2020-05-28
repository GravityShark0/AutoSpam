# Import
from pynput.keyboard import Key, Controller
import win32com.client
import pywintypes
import win32api
import win32con
import win32gui
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

# Variables/Settings
keyboard = Controller()
amount1 = 0
key = "|"


def AutoSpam():
    cls()
    global timer, word, inf, app_back, app, amount1, amount2, number, chose, key
    # Menu
    print("Welcome to AutoSpamv1.1")
    chose = int(input("Main Menu\n1 = Keyboard\n2 = Other\n:"))
    if chose == 1:
        cls()
        saves = str(input("Use saved settings? [y/n]\n[if you don't have saved settings it will use the default "
                          "settings]\n:")).lower()
        if saves == "y":
            word, number, amount1, amount2, timer, app_back, app = pickle.load(open("settings.dat", "rb"))
            cls()
            print("Settings Loaded")
            time.sleep(0.5)
        elif saves == "n":
            cls()
            word = str(input("1.Type what you want to spam.\n[Type the Keyword to add an newline]\n[Current Keyword: "
                             "" + key + "]\n:"))
            cls()
            number = str(input("2.Enable Increasing Number? [y/n]\n:")).lower()
            if number == "y":
                amount1 = int(input("2a.Type on what number the increasing number starts\n:"))
                amount2 = int(input("2b.Type on what number the increasing number starts\n:"))
            elif number == "n":
                amount2 = input("2b.Type how many times to repeat the word.\n[Type inf for infinity]\n:")
            cls()
            timer = float(
                input("3.Type how long for word to be sent.\n(Measured in Seconds)\n:"))
            cls()
            app_back = str(input("4.Enable instant tabbing and un tabbing to application? [y/n]\n[If you disable this "
                                 "you have to manually select the application]\n:")).lower()
            if app_back == "y":
                app = str(input("4a.Type the name of your application\n:")).lower()
            elif app_back == "n":
                app = "None"
            cls()
            save = str(input("Save settings? [y/n]\n:")).lower()
            if save == "y":
                pickle.dump([word, number, amount1, amount2, timer, app_back, app], open("settings.dat", "wb"))
                print("Settings Saved.")
        cls()
        correction = str(
            input("Are you sure these are the correct settings? [y/n]\n1." + str(word) + " (Word)\n2." + str(
                number) + " (Increasing Number)\n2a." + str(amount1) + " (Starting Number)\n2b." + str(
                amount2) + " (Stopping Number)\n3." + str(timer) + " (Time between)\n4." + str(app_back) + " (Instant "
                "Select)\n4a." + str(app) + " (Selected Application)\n:")).lower()
        if correction == "n":
            AutoSpam()

        # Accurate Number Fix
        amount2 += 1

        # Inf Check
        inf = False
        if amount2 == "inf":
            amount2 = 1
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
            global amount2
            top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

        # While Loop
        while amount1 != amount2:
            if app_back == "y":
                if __name__ == "__main__":
                    results = []
                    top_windows = []
                    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
                    for i in top_windows:
                        if app in i[1].lower():
                            win32gui.ShowWindow(i[0], win32con.SW_MAXIMIZE)
                            shell = win32com.client.Dispatch("WScript.Shell")
                            shell.SendKeys('%')
                            win32gui.SetForegroundWindow(i[0])
                            if number == "y":
                                keyboard.type(str(amount1))
                            for c in word:
                                if c == key:
                                    keyboard.press(Key.shift_l)
                                    keyboard.press(Key.enter)
                                    keyboard.release(Key.shift_l)
                                    keyboard.release(Key.enter)
                                elif c != key:
                                    keyboard.press(c)
                                    keyboard.release(c)
                            keyboard.press(Key.enter)
                            keyboard.release(Key.enter)
                            win32gui.ShowWindow(i[0], win32con.SW_MINIMIZE)
                            time.sleep(timer)
                            if not inf:
                                amount1 += 1
            elif app_back == "n":
                if number == "y":
                    keyboard.type(str(amount1))
                for c in word:
                    if c == key:
                        keyboard.press(Key.shift_l)
                        keyboard.press(Key.enter)
                        keyboard.release(Key.shift_l)
                        keyboard.release(Key.enter)
                    elif c != key:
                        keyboard.press(c)
                        keyboard.release(c)
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)
                time.sleep(timer)
                if not inf:
                    amount1 += 1

    elif chose == 2:
        cls()
        other = int(input("Other\n1 = Keyword\n2 = Back\n:"))
        if other == 1:
            cls()
            key = str(input("Type a letter when used will make an newline.\n[If you make a keyword longer than 1 "
                            "letter it will use the first letter instead]\n[Current Keyword: " + key + "]\n:"))
            key = key[0:1]
            cls()
            AutoSpam()

        elif other == "yeetus betus deletus":
            print("yeet")
    # Restart
    continues = str(input("Do you want to go back to the main menu? [y/n]\n:"))
    if continues == "y":
        cls()
        AutoSpam()


AutoSpam()

# Quit
cls()
print("Closing AutoSpam.")
time.sleep(1)
cls()
print("Closing AutoSpam..")
time.sleep(1)
cls()
print("Closing AutoSpam...")
time.sleep(1)
quit()
