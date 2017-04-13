#窗体句柄
loginwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Login')
#控件句柄
userinput = find_idxSubHandle(loginwindow,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',1)
passinput = find_idxSubHandle(loginwindow,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
#常量
USERNAME = 'NBZRP'
PASSWORD = '900502'
#输入
win32gui.SendMessage(userinput, win32con.WM_SETTEXT,None,USERNAME)
win32gui.SendMessage(passinput, win32con.WM_SETTEXT,None,PASSWORD)
#回车
win32gui.PostMessage(userinput, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(userinput, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#进入模块选择
modulewindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Application Project')
time.sleep(1)
docuchoice = find_idxSubHandle(modulewindow,'WindowsForms10.BUTTON.app.0.2a2cc74_r9_ad1',10)
win32gui.SendMessage(docuchoice,win32con.BM_CLICK,None,None)
time.sleep(3)

#按键选择 alt+E
win32api.keybd_event(18,0,0,0)
win32api.keybd_event(69,0,0,0)     # E
win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)
win32api.keybd_event(69,win32con.KEYEVENTF_KEYUP,0)  #释放按键

#前台点击按钮
left, top, right, bottom = win32gui.GetWindowRect(retrieve)
win32api.SetCursorPos((left+15,top+15))
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0, 0, 0)  
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0,0,0, 0)

#进入单证模块,选择ebooking
documodule = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Document ApplicationNBZRP0')
document = win32gui.FindWindowEx(documodule,0,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','menuStrip1')
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
time.sleep(2)


#e-booking
docuwindow = win32gui.FindWindow('WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Document ApplicationNBZRP0')
backwindow = win32gui.FindWindowEx(docuwindow,0,'WindowsForms10.MDICLIENT.app.0.2a2cc74_r9_ad1',None)
bookingwindow = win32gui.FindWindowEx(backwindow,0,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1','Accept E-booking')
bookingtitle = find_idxSubHandle(bookingwindow,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1',4)
retrieve = win32gui.FindWindowEx(bookingtitle,0,'WindowsForms10.BUTTON.app.0.2a2cc74_r9_ad1','Retrieve')
vslvoy = find_idxSubHandle(bookingtitle,'WindowsForms10.Window.8.app.0.2a2cc74_r9_ad1',0)
vslname = find_idxSubHandle(vslvoy,'WindowsForms10.Window.b.app.0.2a2cc74_r9_ad1',1)
voyname = find_idxSubHandle(vslvoy,'WindowsForms10.Window.b.app.0.2a2cc74_r9_ad1',0)
vslinput = find_idxSubHandle(vslname,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
voyinput = find_idxSubHandle(voyname,'WindowsForms10.EDIT.app.0.2a2cc74_r9_ad1',0)
win32gui.SendMessage(vslinput, win32con.WM_SETTEXT,None,'SNLDFFG')
time.sleep(.1)
win32gui.SendMessage(voyinput, win32con.WM_SETTEXT,None,'1708S')
# 按钮win32gui.PostMessage(retrieve,win32con.BM_CLICK,None,None)
win32gui.SendMessage(retrieve,win32con.BM_CLICK,None,None)
