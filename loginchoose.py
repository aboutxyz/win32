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
documodule = find_idxSubHandle(modulewindow,'WindowsForms10.BUTTON.app.0.2a2cc74_r9_ad1',10)
win32gui.PostMessage(documodule, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
win32gui.PostMessage(documodule, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


#按键选择 alt+E
win32api.keybd_event(18,0,0,0)
win32api.keybd_event(69,0,0,0)     # E
win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)
win32api.keybd_event(69,win32con.KEYEVENTF_KEYUP,0)  #释放按键
