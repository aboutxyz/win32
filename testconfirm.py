#coding:utf-8
import win32api
import win32con
import win32gui
import time
from constant import *
  
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
confirmwindow = find_idxSubHandle(bookingwindow, 'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',0)
print hex(confirmwindow)
for i in range(10):
    print hex(find_idxSubHandle(bookingwindow, 'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',i))

