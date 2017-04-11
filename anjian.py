#coding:utf-8
import win32api
import win32con
import win32gui
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
#窗体句柄
loginwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Login')
#控件句柄
userinput = find_idxSubHandle(loginwindow,'WindowsForms10.EDIT.app.0.3b93019_r11_ad1',1)
passinput = find_idxSubHandle(loginwindow,'WindowsForms10.EDIT.app.0.3b93019_r11_ad1',0)
#常量
USERNAME = 'NBZRP'
PASSWORD = '900502'
#输入
win32gui.SendMessage(userinput, win32con.WM_SETTEXT,None,USERNAME)
win32gui.SendMessage(passinput, win32con.WM_SETTEXT,None,PASSWORD)
#回车
win32gui.PostMessage(passinput, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(passinput, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#进入模块选择
# modulewindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Application Project')
# document = find_idxSubHandle(modulewindow,'WindowsForms10.BUTTON.app.0.3b93019_r11_ad1',11)
# win32gui.PostMessage(document, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)

#进入单证模块
docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Document ApplicationNBZRP0')
document = find_idxSubHandle(docuwindow,'WindowsForms10.Window.8.app.0.3b93019_r11_ad1',0)
win32gui.SetForegroundWindow(document)
# win32api.keybd_event(18,0,0,0)
# win32api.keybd_event(69,0,0,0)     # E
# win32api.keybd_event(69,win32con.KEYEVENTF_KEYUP,0)  #释放按键
win32api.keybd_event(win32con.VK_RIGHT,0,0,0)
win32api.keybd_event(win32con.VK_RIGHT,0,win32con.KEYEVENTF_KEYUP,0)
win32gui.PostMessage(document, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(document, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#e-booking
# bookingwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.3b93019_r11_ad1','Accept E-booking')
