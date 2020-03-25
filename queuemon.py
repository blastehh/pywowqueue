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
mouse = MouseController()

pytesseract.pytesseract.tesseract_cmd = r'D:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def screenshot(window_title=None):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            #win32gui.SetForegroundWindow(hwnd)
            
            im = pyautogui.screenshot()

            return im
        else:
            print('Window not found!')
    else:
        im = pyautogui.screenshot()
        return im

def closeWindow(window_title=None):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
        else:
            print('Window not found!')

def launchWow():
    while True:
        hwnd = win32gui.FindWindow(None, "World of Warcraft")
        if hwnd:
            break
        else:
            bnet = screenshot("Blizzard Battle.net")
            if isinstance(bnet, Image.Image):
                img = np.array(bnet)
                img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
                res = cv2.matchTemplate(img_gray, playbtn, cv2.TM_CCOEFF_NORMED)
                matches = np.where(res >= 0.9)
                if np.shape(matches)[1] >= 1:
                    x = matches[1][0]
                    y = matches[0][0]
                    print("Launching WoW...")
                    mouse.position = (int(x)+15, int(y)+5)
                    mouse.click(Button.left)
                else:
                    hwnd = win32gui.FindWindow(None, "Blizzard Battle.net")
                    if hwnd:
                        win32gui.SetForegroundWindow(hwnd)
        time.sleep(5)

playbtn = cv2.imread('play.png', 0)

while True:
    ss = screenshot('World of Warcraft')
    if isinstance(ss, Image.Image):

        img = np.array(ss)

        startx = int(round(img.shape[1]/3.2))
        endx = int(round(img.shape[1]-startx))
        starty = int(round(img.shape[0]/2.5))
        endy = int(round(img.shape[0]-starty))

        cropped_1 = img[starty:endy, startx:endx]
        cropped = cv2.cvtColor(cropped_1, cv2.COLOR_RGB2BGR)
        hsv = cv2.cvtColor(cropped, cv2.COLOR_BGR2HSV)

        lower = np.array([5, 130, 150])
        upper = np.array([70, 255, 255])
        mask = cv2.inRange(hsv, lower, upper)
        res = cv2.bitwise_and(cropped,cropped, mask=mask)

        #cv2.imwrite('mask.png',mask)
        #cv2.imwrite('res.png',res)
        #cv2.imwrite('raw.png',cropped)

        text = pytesseract.image_to_string(mask)
        if "You have been disconnected from the server." in text:
            print("Disconnected! Closing WoW...")
            closeWindow('World of Warcraft')
            time.sleep(12)
            
    launchWow()
    time.sleep(10)