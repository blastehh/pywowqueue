# Requirements:
# python -m pip install pywin32 pyautogui pytesseract opencv-python pillow pynput
# https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.0-alpha.20200223.exe
import cv2
import pyautogui
import pytesseract
import numpy as np
import win32gui
import time
import win32con
from PIL import Image
from pynput.mouse import Button, Controller as MouseController
import sys,os, datetime
mouse = MouseController()

pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'

def findWindow(window_title, window_class=None):
    return win32gui.FindWindow(window_class, window_title)

def setFGW(hwnd):
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetWindowPos(hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
        win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
        win32gui.SetWindowPos(hwnd,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)
    except:
        # Handle exception, no one cares
        pass

def screenshot(window_class=None, window_title=None, setFG=False):
    if window_title:
        hwnd = findWindow(window_title, window_class)
        if hwnd:
            if setFG:
                setFGW(hwnd)
                time.sleep(1)

            im = pyautogui.screenshot()
            return im
        else:
            pout(f"{window_title} window not found!")
    else:
        im = pyautogui.screenshot()
        return im

def closeWindow(window_title=None):
    if window_title:
        hwnd = findWindow(window_title)
        if hwnd:
            win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
        else:
            pout(f"Can't find {window_title} window to close")

def launchWow():
    hwnd = findWindow("World of Warcraft", "GxWindowClass")
    if not hwnd:
        bnet = screenshot("Qt5QWindowOwnDCIcon", "Blizzard Battle.net", True)
        if isinstance(bnet, Image.Image):
            img = np.array(bnet)
            img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            res = cv2.matchTemplate(img_gray, playbtn, cv2.TM_CCOEFF_NORMED)
            matches = np.where(res >= 0.9)
            if np.shape(matches)[1] >= 1:
                x = matches[1][0]
                y = matches[0][0]
                pout("Launching WoW...")
                mouse.position = (int(x)+15, int(y)+5)
                mouse.click(Button.left)
            else:
                pout("Failed to find Battle.net window. Is it on the screen?")
        else:
            pout("Is Battle.net Launcher running?")

def pout(*args):
    curDT = datetime.datetime.now()
    dtString = curDT.strftime("%H:%M:%S") + " - "
    for arg in args:
        dtString += arg
    print(dtString)

if not os.path.isfile(pytesseract.pytesseract.tesseract_cmd):
    pout("Tesseract not found! Check you have the correct path in the script.")
    input()
    exit()

playpath = 'play.png'
if not os.path.isfile(playpath):
    playpath = os.path.dirname(sys.argv[0])+'/play.png'
if not os.path.isfile(playpath):
    pout(f"play.png not found at: {playpath}")
    input()
    exit()

playbtn = cv2.imread(playpath, 0)
queue = 0
wait = ""

if findWindow("World of Warcraft", "GxWindowClass"):
    pout("WoW is running!")
else:
    pout("WoW isn't running!")
    launchWow()
    time.sleep(10)

while True:
    ss = screenshot("GxWindowClass", "World of Warcraft", False)
    if isinstance(ss, Image.Image):
        img = np.array(ss)
        # Define area of the screen where we're looking for the queue/disconnect message
        startx = int(round(img.shape[1]/3.2))
        endx = int(round(img.shape[1]-startx))
        starty = int(round(img.shape[0]/2.5))
        endy = int(round(img.shape[0]-starty))

        cropped = img[starty:endy, startx:endx]
        hsv = cv2.cvtColor(cropped, cv2.COLOR_RGB2HSV)

        # Define the HSV colour range of the status text
        lower = np.array([5, 130, 150])
        upper = np.array([70, 255, 255])

        mask = cv2.inRange(hsv, lower, upper)

        #cropped_BGR = cv2.cvtColor(cropped, cv2.COLOR_RGB2BGR)
        #masked = cv2.bitwise_and(cropped_BGR,cropped_BGR, mask=mask)
        #cv2.imwrite('mask.png',mask)
        #cv2.imwrite('masked.png',res)
        #cv2.imwrite('raw.png',cropped_BGR)

        text = pytesseract.image_to_string(mask)
        inqueue = False
        if "You have been disconnected from the server." in text:
            pout("Disconnected! Closing WoW...")
            closeWindow('World of Warcraft')
            time.sleep(12)
        else:
            lines = text.splitlines()
            for i, val in enumerate(lines):
                if "Position in queue" in val:
                    inqueue = True
                    queue = val.split("queue:",1)[1].strip()
                elif "Estimated time" in val:
                    inqueue = True
                    wait = val.split("time:",1)[1].strip()
                
            #print(text)
        if inqueue:
            pout(f"Queue position: {queue}, Wait time: {wait}")
    else:
        launchWow()
    time.sleep(10)