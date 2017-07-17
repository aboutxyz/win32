#coding:utf-8
import win32api
import win32con
import win32gui
import time
from constant import *
import concurrent.futures
  
def find_idxSubHandle(pHandle, winClass, index=0):
    """已知子窗口的窗体类名 寻找第index号个同类型的兄弟窗口"""
    assert type(index) == int and index >= 0
    handle = win32gui.FindWindowEx(pHandle, 0, winClass, None)
    while index > 0:
        handle = win32gui.FindWindowEx(pHandle, handle, winClass, None)
        index -= 1
    return handle

docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Document ApplicationNBZRP0')
backwindow = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.MDICLIENT.app.0.3b93019_r11_ad1',None)
bookingwindow = win32gui.FindWindowEx(backwindow,0,'WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Accept E-booking')
win32gui.SetForegroundWindow(docuwindow)


def createbll():
    confirmwindow = find_idxSubHandle(bookingwindow, 'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',0)
    createblno = find_idxSubHandle(confirmwindow, 'WindowsForms10.Window.b.app.0.3b93019_r11_ad1',0)
    win32gui.PostMessage(createblno,win32con.WM_LBUTTONDOWN,None,None)
    win32gui.PostMessage(createblno,win32con.WM_LBUTTONUP,None,None)
    time.sleep(1)
    clientchose = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Create B/L No.')
    clientchose1 = find_idxSubHandle(clientchose,'WindowsForms10.Window.b.app.0.3b93019_r11_ad1',1)
    win32gui.PostMessage(clientchose1,win32con.WM_LBUTTONDOWN,None,None)
    win32gui.PostMessage(clientchose1,win32con.WM_LBUTTONUP,None,None)
    time.sleep(1)
    acceptblno = win32gui.FindWindowEx(confirmwindow,0,'WindowsForms10.Window.b.app.0.3b93019_r11_ad1','Accept')
    win32gui.PostMessage(acceptblno,win32con.WM_LBUTTONDOWN,None,None)
    win32gui.PostMessage(acceptblno,win32con.WM_LBUTTONUP,None,None)
    time.sleep(1)
    aa = win32gui.FindWindow('#32770','Warning')
    shi = find_idxSubHandle(aa,'Button',1)
    win32gui.PostMessage(shi, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.PostMessage(shi, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    time.sleep(2)
    queren = win32gui.FindWindow('#32770',None)
    shi2 = find_idxSubHandle(queren,'Button',0)
    win32gui.PostMessage(shi2, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.PostMessage(shi2, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    time.sleep(1)



# n = 15
# while n>0:
    # n = n-1
    # createbll()
    # time.sleep(2)
    
    
with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
    for i in range(12):
        aa=executor.submit(createbll,)
        aa.result()