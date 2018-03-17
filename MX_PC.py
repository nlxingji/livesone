import os
import random
import time
import webbrowser
import csv

import win32api
import win32gui
import win32con
from ctypes import *

def find_idxSubHandle(pHandle, winClass, index=0):
        """
        已知子窗口的窗体类名
        寻找第index号个同类型的兄弟窗口
        """
        assert type(index) == int and index >= 0
        handle = win32gui.FindWindowEx(pHandle, 0, winClass, None)
        while index > 0:
            handle = win32gui.FindWindowEx(pHandle, handle, winClass, None)
            index -= 1
        return handle



def find_subHandle(pHandle, winClassList):
        """
        递归寻找子窗口的句柄
        pHandle是祖父窗口的句柄
        winClassList是各个子窗口的class列表，父辈的list-index小于子辈
        """
        assert type(winClassList) == list
        if len(winClassList) == 1:
            return find_idxSubHandle(pHandle, winClassList[0][0], winClassList[0][1])
        else:
            pHandle = find_idxSubHandle(pHandle, winClassList[0][0], winClassList[0][1])
            return find_subHandle(pHandle, winClassList[1:])
            print(find_subHandle(pHandle, winClassList[1:]))

def up():
    win32gui.SetForegroundWindow(Mhandle)
    rnum=random.randint(150,350)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, rnum,win32con.WHEEL_DELTA)
    time.sleep(random.random()*5)
    print('up')

def down():
    win32gui.SetForegroundWindow(Mhandle)
    rnum= -1 * random.randint(150,350)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, rnum,win32con.WHEEL_DELTA)
    time.sleep(random.random()*5)
    print('down')


def see_wy():  # 网页浏览操作
    i = 50
    while i > 1:
        if i < 15:
            down()
        else:
            rd = random.randint(1, 10)
            print(rd)
            if rd > 4:
                up()
            else:
                down()
            
            if rd>8:
                print('点击页面')
                click_page()
    
        i=i-1



def close_page():
    win32api.keybd_event(VK_CODE['ctrl'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['F4'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['F4'], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.random()*10)
    print('关闭当前页面')

def close_otherpage():
    win32api.keybd_event(VK_CODE['ctrl'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['shift'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['w'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['w'], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.random()*10)
    print('关闭其他页面')


def quanping():
    win32api.keybd_event(VK_CODE['F11'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['F11'], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.random()*10)
    print('进入全屏模式')

def clickLeftCur():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN|win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def moveCurPos(x,y):
    windll.user32.SetCursorPos(x, y)

def getCurPos():
    return win32gui.GetCursorPos()


def click_page():
    xy=getCurPos()
    left, top, right, bottom = win32gui.GetWindowRect(Mhandle)
    dx=right-left
    dy=bottom-top
    x=random.randint(int(0.4*dx),int(0.7*dx))+left
    y=random.randint(int(0.4*dy),int(0.7*dy))+top
    #num=random.randint(1,4)
    #if num==1:
        #moveCurPos(xy[0]+100, xy[1])
    moveCurPos(x, y)
    time.sleep(5)
    clickLeftCur()
    time.sleep(random.randint(10,30))
    num=random.randint(7,15)
    while num>1:
        down()
        num-=1
    time.sleep(random.randint(3,10))

     



VK_CODE = {'backspace': 0x08,
           'tab': 0x09,
           'clear': 0x0C,
           'enter': 0x0D,
           'shift': 0x10,
           'ctrl': 0x11,
           'alt': 0x12,
           'pause': 0x13,
           'caps_lock': 0x14,
           'esc': 0x1B,
           'spacebar': 0x20,
           'page_up': 0x21,
           'page_down': 0x22, 'end': 0x23,
           'home': 0x24,
           'left_arrow': 0x25,
           'up_arrow': 0x26,
           'right_arrow': 0x27,
           'down_arrow': 0x28,
           'select': 0x29,
           'print': 0x2A, 'execute': 0x2B, 'print_screen': 0x2C, 'ins': 0x2D, 'del': 0x2E, 'help': 0x2F, '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
           'a': 0x41, 'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'numpad_0': 0x60,
           'numpad_1': 0x61,
           'numpad_2': 0x62,
           'numpad_3': 0x63,
           'numpad_4': 0x64,
           'numpad_5': 0x65,
           'numpad_6': 0x66,
           'numpad_7': 0x67,
           'numpad_8': 0x68,
           'numpad_9': 0x69,
           'multiply_key': 0x6A,
           'add_key': 0x6B,
           'separator_key': 0x6C,
           'subtract_key': 0x6D,
           'decimal_key': 0x6E,
           'divide_key': 0x6F,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87,
           'num_lock': 0x90,
           'scroll_lock': 0x91,
           'left_shift': 0xA0,
           'right_shift ': 0xA1,
           'left_control': 0xA2,
           'right_control': 0xA3,
           'left_menu': 0xA4,
           'right_menu': 0xA5,
           'browser_back': 0xA6,
           'browser_forward': 0xA7,
           'browser_refresh': 0xA8,
           'browser_stop': 0xA9,
           'browser_search': 0xAA,
           'browser_favorites': 0xAB,
           'browser_start_and_home': 0xAC,
           'volume_mute': 0xAD,
           'volume_Down': 0xAE,
           'volume_up': 0xAF,
           'next_track': 0xB0,
           'previous_track': 0xB1,
           'stop_media': 0xB2,
           'play/pause_media': 0xB3,
           'start_mail': 0xB4,
           'select_media': 0xB5,
           'start_application_1': 0xB6,
           'start_application_2': 0xB7,
           'attn_key': 0xF6,
           'crsel_key': 0xF7,
           'exsel_key': 0xF8,
           'play_key': 0xFA,
           'zoom_key': 0xFB,
           'clear_key': 0xFE,
           '+': 0xBB,
           ',': 0xBC,
           '-': 0xBD,
           '.': 0xBE,
           '/': 0xBF,
           '`': 0xC0,
           ';': 0xBA,
           '[': 0xDB,
           '\\': 0xDC,
           ']': 0xDD,
           "'": 0xDE,
           '`': 0xC0}



       

if __name__ == '__main__':
    url0 = 'https://www.2345.com/?kq101020'
    webbrowser.open(url0,new=0,autoraise=True)
    print('正在打开网址'+ url0)
    time.sleep(2)
    #quanping()
    global Mhandle
    Mhandle = win32gui.FindWindow("Maxthon3Cls_MainFrm", None)
    if Mhandle > 0:
        print('open browser sucess')
    else:
        print('failure')
        exit
    
    time.sleep(5)
    try:
        see_wy()
        click_page()
        close_page()
    except:
        pass


filename = "url.csv"
url = []
if filename != '':
    with open(filename) as csvfile:
        rows = csv.reader(csvfile)
        for row in rows:
            url.append(row[0])
else:
    print("You may not set the options correctly,please try again")

counter = 0
while True:
    counter += 1
    urlx=random.choice(url)
    print('正在打开网址'+ urlx)
    webbrowser.open(urlx,new=0,autoraise=True)
    time.sleep(20)
    try:
        win32gui.SetForegroundWindow(Mhandle)
        see_wy()
        time.sleep(random.randint(20,50))
    except:
        pass

    if counter ==5:
        try:
            win32gui.SetForegroundWindow(Mhandle)
            close_otherpage()
            print('关闭所有页面')
        except:
            pass
        counter=0
