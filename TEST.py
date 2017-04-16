#coding:utf-8
import win32api
import win32con
import win32gui
import time
from concurrent import futures
from threading import current_thread

def find_idxSubHandle(pHandle, winClass, index=0):
    """已知子窗口的窗体类名 寻找第index号个同类型的兄弟窗口"""
    assert type(index) == int and index >= 0
    handle = win32gui.FindWindowEx(pHandle, 0, winClass, None)
    while index > 0:
        handle = win32gui.FindWindowEx(pHandle, handle, winClass, None)
        index -= 1
    return handle

docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Document ApplicationNBZRP0')
backwindow = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.MDICLIENT.app.0.2a2cc74_r9_ad1',None)
bookingwindow = win32gui.FindWindowEx(backwindow,0,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Accept E-booking')
backwindow = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.MDICLIENT.app.0.2a2cc74_r9_ad1',None)
bookingwindow = win32gui.FindWindowEx(backwindow,0,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Accept E-booking')
bookingtitle = find_idxSubHandle(bookingwindow,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1',4)
retrieve = win32gui.FindWindowEx(bookingtitle,0,'WindowsForms10.BUTTON.app.0.2a2cc74_r9_ad1','Retrieve')
vslvoy = find_idxSubHandle(bookingtitle,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1',0)
vslname = find_idxSubHandle(vslvoy,'WindowsForms10.Window.b.app.0.2a2cc74_r9_ad1',1)
voyname = find_idxSubHandle(vslvoy,'WindowsForms10.Window.b.app.0.2a2cc74_r9_ad1',0)
vslinput = find_idxSubHandle(vslname,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
voyinput = find_idxSubHandle(voyname,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
win32gui.SendMessage(vslinput, win32con.WM_SETTEXT,None,"SITSISU")
time.sleep(.1)
win32gui.SendMessage(voyinput, win32con.WM_SETTEXT,None,"1707N")
# win32gui.SendMessage(retrieve,win32con.BM_CLICK,None,None)
# time.sleep(2)
# aa = win32gui.FindWindow('#32770','')
# win32gui.PostMessage(aa, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
# win32gui.PostMessage(aa, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
try:
    win32gui.PostMessage(retrieve,win32con.BM_CLICK,None,None)
    time.sleep(.8)  
finally:
    aa = win32gui.FindWindow('#32770',None)
    win32gui.PostMessage(aa, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.PostMessage(aa, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


win32gui.ShowWindow(bookingwindow, win32con.SW_MINIMIZE)
