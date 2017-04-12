#coding:utf-8
import win32api
import win32con
import win32gui
import time
'''
def runApp():
    import win32api
    # 最后一个参数表示是窗口属性，0表示不显示，1表示正常显示，2表示最小化，3表示最大化
    res = win32api.ShellExecute(0, 'open', 'C:\\Windows\\System32\\notepad.exe', '', '', 3)
    
 # 关闭软件
win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
# 软件最大化
win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_MAXIMIZE, 0)
# 将软件窗口置于最前
win32gui.SetForegroundWindow(hwnd)  
'''  
def find_idxSubHandle(pHandle, winClass, index=0):
    """已知子窗口的窗体类名 寻找第index号个同类型的兄弟窗口"""
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

'''
#进入单证模块,选择ebooking
docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Document ApplicationNBZRP0')
document = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','menuStrip1')
win32gui.SetForegroundWindow(document)
win32api.keybd_event(win32con.VK_RIGHT,0,0,0)
win32api.keybd_event(win32con.VK_RIGHT,0,win32con.KEYEVENTF_KEYUP,0)
time.sleep(.2)
win32gui.PostMessage(document, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(document, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
for i in range(10):
    win32api.keybd_event(win32con.VK_DOWN,0,0,0)
    time.sleep(.3)
win32gui.PostMessage(document, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(document, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
time.sleep(.2)
'''
#e-booking
docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Document ApplicationNBZRP0')
backwindow = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.MDICLIENT.app.0.3b93019_r11_ad1',None)
bookingwindow = win32gui.FindWindowEx(backwindow,0,'WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Accept E-booking')
for i in range(10):
    print find_idxSubHandle(bookingwindow,'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',i)
bookingtitle = find_idxSubHandle(bookingwindow,'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',4)
print bookingtitle

# voyageinput = find_idxSubHandle(bookingwindow,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
# win32gui.SendMessage(vesselinput, win32con.WM_SETTEXT,None,'hu')
# time.sleep(.2)
# win32gui.SendMessage(voyageinput, win32con.WM_SETTEXT,None,'ko')
