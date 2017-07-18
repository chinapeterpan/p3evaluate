# -*- coding: utf-8 -*-
import csv

import webbrowser
import keyboard
import time
import os
import win32gui
import win32con
import win32api
import win32clipboard


text ='亲爱的adi1949正品运动的老客户,如果觉得东西好的话,可以加我们的vx哦,我们一般在周末专柜直播并发出特价产品哦' \
      '加v:xiaozongzi520'

def openwebbrowser(wangwangid):
    url = 'https://amos.alicdn.com/getcid.aw?spm=a1z09.1.0.0.yP83sO&v=3&groupid=0&s=1&charset=utf-8&uid=' \
          + wangwangid + '&site=cntaobao&groupid=0&s=1&fromid=cntaobaoshanghai%B4%FA%B9%BA'
    return webbrowser.open_new(url)


def commandwangwang(wangwangid):
    command = 'D:\\AliWorkbench\\5.07.03N\\WWCmd aliim:sendmsg?touid=cntaobao'+wangwangid+'&site=cntaobao&status=1'
    return os.system(command)


def setclipboardww():
    """

    :rtype : object
    """
    win32clipboard.OpenClipboard()

    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
    finally:
        win32clipboard.CloseClipboard()

def replywangwang(wangwangid):
    if commandwangwang(wangwangid):
        wdname = "shanghai代购 - 接待中心"
        w1hd = win32gui.FindWindow(0, wdname)
        setclipboardww()
        win32gui.SetForegroundWindow(w1hd)
        time.sleep(2)
        win32api.keybd_event(keyboard.VK_CODE['ctrl'], 0, 0, 0)
        win32api.keybd_event(keyboard.VK_CODE['v'], 0, 0, 0)
        time.sleep(0.2)
        win32api.keybd_event(keyboard.VK_CODE['v'], 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(keyboard.VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
        keyboard.press('enter')

if __name__ == '__main__':
     with open('1.csv', encoding='GBK') as csvfile:
        reader = csv.reader(csvfile, dialect='excel')
        for row in reader:
            print(row[1])
            replywangwang(row[1])
            time.sleep(2)

