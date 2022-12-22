import  win32con,win32api,win32gui
import win32clipboard as wt
import time
import random
import pandas as pd
import re



class QQhandle():
    def __init__(self):
        self.handle = 0
        self.group_name = None
        self.results =None

    def Get_hwnd(self,name="QQ"):
        """
        根据窗口名查找句柄号
        :param name: 窗口名
        :return: 句柄号
        """
        hwnd_title = dict()
        def get_all_hwnd(hwnd, mouse):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd):
              hwnd_title.update({win32gui.GetWindowText(hwnd):hwnd})
        win32gui.EnumWindows(get_all_hwnd, 0)

        return hwnd_title.get(name)

    def send_message(self,hwnd,grouptext):
        setText(grouptext)  # 将好友名称复制到剪切
        Open_win(hwnd)
        win32gui.SendMessage(hwnd, 258, 22, 2080193)  # SendMessage的的参数。具体的还是详细查msdn
        win32gui.SendMessage(hwnd, 770, 0, 0)
        time.sleep(1.5)
        winRect = Get_win_size(hwnd)
        mouse_click(winRect[0] + 125, winRect[1] + 150)  # 移动到QQ搜索窗口并点击
        mouse_click(winRect[0] + 125, winRect[1] + 150)  # 移动到QQ搜索窗口并点击


    def Get_hwnd_blurry(self,name="2022"):
        """ 根据窗口名模糊查找包含的群名，主要用于针对在一个窗口中同时打开了多个群存在多个会话的情况
        :param name:
        :return:
        """
        hwnd_title = dict()
        self.group_name = name
        def get_all_hwnd(hwnd, mouse):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                hwnd_title.update({win32gui.GetWindowText(hwnd): hwnd})
                if name in win32gui.GetWindowText(hwnd):
                   self.handle=  win32gui.GetWindowText(hwnd)


        win32gui.EnumWindows(get_all_hwnd, 0)

        if self.handle:
            return hwnd_title.get(self.handle)
        else:
            print(f"未运行对应 {name} 进程")
            print(f"正在运行的进程为：\n {hwnd_title}")

    def uppage(self,nowphwnd,count=1):
        nowwinRect = Get_win_size(nowphwnd)
        mouse_click(nowwinRect[0] + 300, nowwinRect[1] + 380)
        Ctrl_Home()
        time.sleep(1)  # 刷新等待1秒
        for _ in range(count):
            # mouse_click(nowwinRect[0] + 270, nowwinRect[1] + 110)  # 刷新窗口
            time.sleep(0.5)
            win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, 200)  # 向上滑动100个单位


    def crawl(self,grouptext):

        qq.send_message(qq.Get_hwnd(), grouptext[0])
        time.sleep(1.5)
        phwnd = self.Get_hwnd_blurry(grouptext[0][:10])  # 只取前10个字符
        if phwnd != None:
            Open_win(phwnd)
            winRect = Get_win_size(phwnd)
            """依次查询获取的群名，获取群消息记录"""
            for name in grouptext:
                time.sleep(1)
                mouse_click(winRect[0] + 120, winRect[1] + 125)  # 移动到QQ搜索窗口并点击
                Random_Sleep()


                if  nowphwnd := self.Get_hwnd_blurry(name):
                    print(f"{name}：窗口打开成功！")
                else:
                    print(f"未找到对应{name}：窗口")
                    return
                Open_win(nowphwnd)
                self.uppage(nowphwnd)

                Ctrl_A()
                Ctrl_C()

                """处理后的数据存储"""
                self.save_file()

    def check_info(self,info):
        data= info.split()
        if not data:
            return

        elif len(data)==1:
            self.info.append(data)
        else:
            time =None
            name =None
            if len(data)==3:
                name,year,time = data
                time = year+" "+time

            if len(data)==2:             # 没有年月
                name,time =data

            self.info_time.append(time)
            self.name.append(name)






    def save_file(self):
        data = GetText()
        information = re.split("\n|\r",data)
        f = open("results.txt", "w", encoding="utf-8")
        f.write(data)
        f.close()

        # print(len(information))
        self.name = []
        self.info_time=[]
        self.info =[]
        for info in information:
            self.check_info(info)
        length_time,length_info = len(self.info_time),len(self.info)
        # print(length_time,length_info)

        if length_time!= length_info:
            self.name =self.name[-length_info:]
            self.info_time = self.info_time[-length_info:]

        # print(len(self.info_time),len(self.info))

        data = pd.DataFrame.from_dict({"name":self.name,"time":self.info_time,"info":self.info})
        data.to_excel("results.xls")
        data.insert(0,'groupname',self.group_name)
        self.results =data.values.tolist()


def setText(info):  #将文本复制进剪切板
    wt.OpenClipboard()
    wt.EmptyClipboard()
    wt.SetClipboardData(win32con.CF_UNICODETEXT, info)
    wt.CloseClipboard()

def GetText():
    """
    获取剪切板文本
    :return:
    """
    wt.OpenClipboard()
    d = wt.GetClipboardData(win32con.CF_UNICODETEXT)
    wt.CloseClipboard()
    return d


def Random_Sleep():
    """
    随机休息一定时间
    import random
    print( random.randint(1,10) )        # 产生 1 到 10 的一个整数型随机数
    print( random.random() )             # 产生 0 到 1 之间的随机浮点数
    print( random.uniform(1.1,5.4) )     # 产生  1.1 到 5.4 之间的随机浮点数，区间可以不是整数
    print( random.choice('tomorrow') )   # 从序列中随机选取一个元素
    print( random.randrange(1,100,2) )   # 生成从1到100的间隔为2的随机整数
    :return:
    """
    time.sleep(random.uniform(0.05, 0.1))


def Ctrl_A():
    """
    全选
    :return:
    """
    win32api.keybd_event(17, 0, 0, 0)  # Ctrl
    win32api.keybd_event(65, 0, 0, 0)  # A
    win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.uniform(0.1, 0.2))


def Ctrl_C():
    """
    粘贴
    :return:
    """
    win32api.keybd_event(17, 0, 0, 0)  # Ctrl
    win32api.keybd_event(67, 0, 0, 0)  # A
    win32api.keybd_event(67, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.uniform(0.1, 0.2))

def Ctrl_Home():
    """
    向上翻页
    :return:
    """
    win32api.keybd_event(17, 0, 0, 0)  # Ctrl
    win32api.keybd_event(36, 0, 0, 0)  # Home
    win32api.keybd_event(36, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(random.uniform(0.1, 0.2))


def Open_win(phwnd):
    """
    将当前句柄窗口置顶显示
    :param phwnd:
    :return:
    """
    win32gui.ShowWindow(phwnd, win32con.SW_SHOWNORMAL)
    win32gui.SetForegroundWindow(phwnd)


def mouse_click(x, y, type=1):
    """
    鼠标移动到目标位置并点击
    :param x: x
    :param y: y
    :param type:点击次数
    :return:
    """
    if type == 2:
        win32api.SetCursorPos([x, y])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    else:
        win32api.SetCursorPos([x, y])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def Get_win_size(phwnd):
    """
    获取窗口大小
    :param phwnd:
    :return: (0,0,0,0)返回的是窗口左上角坐标和窗口右下角坐标
    """
    # winRect = win32gui.GetWindowRect(phwnd)
    return win32gui.GetWindowRect(phwnd)




if __name__ =="__main__":

    grouptext = ["2022Fall_数据库设计_小组"]        #  需要爬取的群组 确保提前打开
    qq =QQhandle()



    qq.crawl(grouptext)
    result = qq.results  # 最后的列表结果
    print(result[-10:])




