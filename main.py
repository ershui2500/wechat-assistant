from wx import *
from tkinter.filedialog import askopenfilename
from fnmatch import fnmatch
from ast import literal_eval
import ttkbootstrap as ttk
import PIL.ImageGrab
import tkinter as tk
import win32api
import datetime
import csv
import nt
import re


class WxAssistant:
    def __init__(self):
        self.csv_reader = None
        self.l_times = None
        self.wxaddress = None
        self.s1 = None
        self.s2 = None
        self.listbox = None
        self.left_stext = None
        self.right_stext = None
        self.send_f_right_main = None
        self.send_f_right_timing = None
        self.send_f_left_main = None
        self.send_f_left_update = None
        self.send_nb_right = None
        self.send_nb_left = None
        self.start_nb = None
        self.start_f_start = None
        self.start_f_send = None
        self.start_f_about = None
        self.x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.x = self.x + 500 if self.x > 1920 else self.x
        self.y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        self.y = self.y + 150 if self.y > 1080 else self.y

        # 实例化创建应用程序窗口
        self.root = ttk.Window(
            title="微信助手v1.0",  # 设置窗口的标题
            themename="solar",  # 设置主题
            size=(int(self.x * 0.5), int(self.y * 0.63)),  # 窗口的大小
            minsize=(0, 0),  # 窗口的最小宽高
            maxsize=(self.x, self.y),  # 窗口的最大宽高
            resizable=None,  # 设置窗口是否可以更改大小
            alpha=1.0,  # 设置窗口的透明度(0.0完全透明）
        )
        self.style = ttk.Style()
        self.style.configure('my.TNotebook', tabposition='wn')
        self.style.configure('TNotebook.Tab', font=('宋体', 10))

        self.root.iconbitmap(r'cat.ico')  # 窗口图标
        self.root.place_window_center()  # 让显现出的窗口居中
        self.root.resizable(True, True)  # 让窗口不可更改大小
        # self.root.wm_attributes('-topmost', 1)  # 让窗口位置其它窗口之上

        self.frame_top = tk.Frame(self.root)
        self.frame_top.pack(fill='x')
        self.frame_bottom = tk.Frame(self.root)
        self.frame_bottom.pack(fill=tk.BOTH, expand=1)

    # 主题
    def theme(self):
        theme_names = self.style.theme_names()  # 以列表的形式返回多个主题名
        theme_cbo = ttk.Combobox(master=self.frame_top, font=("微软雅黑", 10), values=theme_names)
        theme_cbo.pack(padx=10, pady=50, side='right', anchor='n')
        tk.Label(master=self.frame_top, text='主题：').pack(pady=50, side='right', anchor='n')
        self.l_times = tk.Label(master=self.frame_top, text='', font=("宋体", 20))
        self.l_times.pack(side='top', pady=50)

        def change_theme(event):
            self.style.theme_use(theme_cbo.get())
            theme_cbo.selection_clear()
            self.style.configure('my.TNotebook', tabposition='wn')
            self.style.configure('TNotebook.Tab', font=('宋体', 10))

        theme_cbo.bind('<<ComboboxSelected>>', change_theme)

    # 循环窗口
    def mainloop(self):
        self.root.mainloop()

    # 多线程开启实时显示
    def time_display(self):
        week_list = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        list_f = []
        try:
            self.l_times['text'] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
            if time.strftime('%S', time.localtime()) == '00':
                with open(f'timed.csv', mode='r', encoding='utf-8-sig', newline='') as f:
                    self.csv_reader = csv.reader(f)
                    for i in self.csv_reader:
                        list_f.append(i)
                        for j in i[1].split('--'):
                            if week_list[datetime.datetime.now().weekday()] == j or \
                                    str(datetime.datetime.now().day) == j or \
                                    '每天' == j:
                                for k in i[2].split('--'):
                                    if time.strftime('%H:%M', time.localtime()) == k or \
                                            time.strftime('%H：%M', time.localtime()) == k:
                                        self.left_stext.delete('1.0', tk.END)
                                        self.right_stext.delete('1.0', tk.END)
                                        self.left_stext.insert('1.0', i[3])
                                        self.right_stext.insert('1.0', i[4])
                                        self.right_send()
                                        if i[5] == '单次':
                                            list_f.pop()
                with open(f'timed.csv', mode='w', encoding='utf-8-sig', newline='') as f:
                    f_writer = csv.writer(f)
                    for i in list_f:
                        f_writer.writerow(i)
        except:
            pass
        self.root.after(1000, self.time_display)

    # 开始面板转操作面板
    def start_panel(self):
        def bu1():
            a = True
            basedir = ['DEFCGHIJK']
            with open('@AutomationLog.txt', mode='w', encoding='utf-8') as f:
                f.close()
            with open('init.ini', mode='r', encoding='utf-8') as f:
                find = re.search('wx_path=.*?.exe', f.read())
            if find is not None:
                self.wxaddress = find.group().split("=")[1]
                os.system(f'start {self.wxaddress}')
                a = False
            for i in basedir[0]:
                if a:
                    for root, dirs, files in os.walk(i + ':\\'):
                        for file in files:
                            path = os.path.join(root, file)
                            if fnmatch(path, "*WeChat.exe*") and a:
                                a = False
                                with open('init.ini', mode='a', encoding='utf-8') as f:
                                    f.write(f'wx_path={path}\n')
                                self.wxaddress = path
                                os.system(f'start {path}')
                                break
            self.start_nb.forget(self.start_f_start)
            self.start_nb.add(self.start_f_send, text="\n   群 发 送   \n")
            self.start_nb.add(self.start_f_about, text="\n   帮    助   \n")

        self.start_nb = ttk.Notebook(self.frame_bottom, style='my.TNotebook')
        self.start_nb.pack(fill=ttk.BOTH, expand=1)

        self.start_f_start = tk.Frame(self.start_nb)
        self.start_f_send = tk.Frame(self.start_nb)
        self.start_f_about = tk.Frame(self.start_nb)

        self.start_nb.add(self.start_f_start, text="  开 始 使 用  ")
        tk.Label(self.start_f_start, text='''点击下面的按钮打开微信
第一次打开会有点久哦 ಥ_ಥ''', font=('宋体', 20)).pack(pady=100)
        tk.Button(self.start_f_start, text='q(≧▽≦q)', font=('宋体', 20), command=bu1).pack()

    # 微信发送面板
    def send_panel(self):
        self.send_nb_left = ttk.Notebook(self.start_f_send)
        self.send_nb_left.pack(anchor='n', side='left', fill='both', expand=1)
        self.send_nb_right = ttk.Notebook(self.start_f_send)
        self.send_nb_right.pack(anchor='n', side='right', fill='both', expand=1)

        self.send_f_left_main = tk.Frame(self.send_nb_left)
        self.send_f_left_main.pack(anchor='n', fill='both', expand=1)
        self.send_f_left_update = tk.Frame(self.send_nb_left)
        self.send_f_left_update.pack(anchor='n', fill='both', expand=1)
        self.send_f_right_main = tk.Frame(self.send_nb_right)
        self.send_f_right_main.pack(anchor='n', fill='both', expand=1)
        self.send_f_right_timing = tk.Frame(self.send_nb_right)
        self.send_f_right_timing.pack(anchor='n', fill='both', expand=1)

        self.send_nb_left.add(self.send_f_left_main, text="  发 送 好 友  ")
        self.send_nb_left.add(self.send_f_left_update, text="  更 新 名 单  ")
        self.send_nb_right.add(self.send_f_right_main, text=" 发 送 信 息 ")
        self.send_nb_right.add(self.send_f_right_timing, text=" 定 时 发 送 ")

    def send_left_panel(self):
        def send_left_group():
            HaveBecomeFriends = []
            os.system(f'start {self.wxaddress}')
            time.sleep(1)
            wechat = WeChat()
            listgroupname = self.left_stext.get('1.0', tk.END).strip()
            for groupname in listgroupname.split('\n'):
                with open('init.ini', mode='r', encoding='utf-8') as f:
                    find = re.search(f'{groupname}.*?=.*?={groupname}', f.read())
                if find is not None:
                    for i in find.group().split('=[')[1].split(']=')[0].split(', '):
                        HaveBecomeFriends.append(i.split("'")[1])
                else:
                    HaveBecomeFriends, NotFriended, name = wechat.GetGroupMemberName(groupname)
                    with open('init.ini', mode='a', encoding='utf-8') as f:
                        f.write(f"{name}={HaveBecomeFriends}={name}--{NotFriended}\n")
            wechat.UiaAPI.GetWindowPattern().Close()
            ListBox()
            for i in HaveBecomeFriends:
                self.listbox.insert(tk.END, i)

        def send_left_label():
            os.system(f'start {self.wxaddress}')
            time.sleep(1)
            wechat = WeChat()
            with open('init.ini', mode='r', encoding='utf-8') as f:
                find = re.search(r'PeopleName=.*?=PeopleName', f.read())
            if find is not None:
                dicts = literal_eval(find.group().split('PeopleName=')[1].split('=PeopleName')[0])
            else:
                dicts = wechat.GetFriendListName()
                with open('init.ini', mode='a', encoding='utf-8') as f:
                    f.write(f"PeopleName={dicts}=PeopleName\n")
            wechat.UiaAPI.GetWindowPattern().Close()
            ListBox()
            for i in dicts:
                for j in dicts[i]:
                    self.listbox.insert(tk.END, f"{i} : {j}")

        def send_left_groupname():
            groupname = []
            os.system(f'start {self.wxaddress}')
            time.sleep(1)
            wechat = WeChat()
            with open('init.ini', mode='r', encoding='utf-8') as f:
                find = re.search(r'GroupName=.*?=GroupName', f.read())
            if find is not None:
                for i in find.group().split('GroupName=[')[1].split(']=GroupName')[0].split(', '):
                    groupname.append(i.split("'")[1])
            else:
                groupname = wechat.GetGroupListName()
                with open('init.ini', mode='a', encoding='utf-8') as f:
                    f.write(f"GroupName={groupname}=GroupName\n")
            wechat.UiaAPI.GetWindowPattern().Close()
            ListBox()
            for i in groupname:
                self.listbox.insert(tk.END, i)

        def ListBox():
            self.s1 = ttk.Scrollbar(self.left_stext, orient=tk.HORIZONTAL, style="round-success", cursor='arrow')
            self.s1.pack(side=tk.BOTTOM, fill=tk.X)
            self.s2 = ttk.Scrollbar(self.left_stext, orient=tk.VERTICAL, style="round-success", cursor='arrow')
            self.s2.pack(side=tk.RIGHT, fill=tk.Y)
            self.listbox = tk.Listbox(self.left_stext, listvariable=tk.StringVar(), selectmode=tk.EXTENDED,
                                      font=('宋体', 13),
                                      xscrollcommand=self.s1.set, yscrollcommand=self.s2.set)
            self.s1.config(command=self.listbox.xview)
            self.s2.config(command=self.listbox.yview)
            self.listbox.pack(side='top', anchor='w', fill='both', expand=1)

        def remove_component():
            self.s1.destroy()
            self.s2.destroy()
            self.listbox.destroy()

        def get_listbox_value():
            self.left_stext.delete('1.0', tk.END)
            try:
                for i in self.listbox.curselection():
                    self.left_stext.insert(tk.END, self.listbox.get(i) + '\n')
                remove_component()
            except:
                pass

        left_lf = ttk.Labelframe(self.send_f_left_main, text=" 获 取 好 友 名 称 方 式 ", style=ttk.PRIMARY)
        left_lf.pack(anchor='n', pady=20)
        ttk.Button(left_lf, text=' 群 聊 ', style="success-link", command=send_left_group).pack(side='left', padx=10,
                                                                                              pady=10)
        ttk.Button(left_lf, text=' 标 签 ', style="success-link", command=send_left_label).pack(side='right', padx=10,
                                                                                              pady=10)
        ttk.Button(left_lf, text=' 群 聊 名 称 ', style="success-link", command=send_left_groupname).pack(side='top',
                                                                                                      pady=10)
        scroll = ttk.Scrollbar(self.send_f_left_main, style="round-success")
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.left_stext = ttk.Text(self.send_f_left_main, font=('宋体', 12), spacing1=5, spacing2=5, spacing3=5)
        self.left_stext.pack(fill='both', expand=1)
        left_frame = ttk.Text(self.left_stext)
        left_frame.pack(side='bottom')
        ttk.Button(left_frame, text=' 取 消 ', style=ttk.SUCCESS, command=remove_component, cursor='arrow') \
            .pack(anchor='s', side='left', padx=15)
        ttk.Button(left_frame, text=' 确 认 ', style=ttk.SUCCESS, command=get_listbox_value, cursor='arrow') \
            .pack(anchor='s', side='right', padx=15)
        scroll.config(command=self.left_stext.yview)
        self.left_stext.config(yscrollcommand=scroll.set)

        update_f = tk.Frame(self.send_f_left_update)
        update_f.pack(fill='both', expand=1)
        tk.Label(update_f, text='''更新的名单有：
您的所有好友并按照标签分好
最近的群聊、最近群聊的好友
更新的时间需要较长时间
并且不能碰鼠标和键盘,请见谅''', font=('宋体', 17)).pack(pady=50)
        tk.Button(update_f, text=' 开  始 ', font=('宋体', 20), command=self.all_update).pack()

    def right_send(self):
        string = ''
        send_list = []
        name_list = []
        try:
            os.system(f'start {self.wxaddress}')
            time.sleep(1)
            wechat = WeChat()
            self.right_stext.insert(tk.END, '\n完成')
            for i in self.right_stext.get('1.0', tk.END).strip().split('\n'):
                string += i
                if i == '':
                    if string != '':
                        send_list.append(string)
                    string = ''
                elif i == '完成':
                    string = string[:-2]
                    if string != '':
                        send_list.append(string)
            for i in self.left_stext.get('1.0', tk.END).strip().split('\n'):
                if i != '':
                    name_list.append(i)
            for name in name_list:
                wechat.Search(name)
                for send in send_list:
                    if 'filepath(文件路径)' in send:
                        wechat.SendFiles(send.split('=')[-1])
                    elif 'Screenshot(屏幕截图)' in send:
                        wechat.SendFiles(send.split('=')[-1])
                    else:
                        wechat.SendMsg(send)
            time.sleep(1)
            wechat.UiaAPI.GetWindowPattern().Close()
        except:
            time.sleep(1)
            wechat.UiaAPI.GetWindowPattern().Close()

    def send_right_panel(self):
        def send_right_file():
            path = askopenfilename()
            if path:
                self.right_stext.insert(tk.END, f"\n\nfilepath(文件路径)={path}\n\n")

        def send_right_image():
            try:
                time_s = time.strftime(".\\images\\%Y年%m月%d日-%H点%M分%S秒.png", time.localtime())
                img = PIL.ImageGrab.grabclipboard()
                img.save(time_s, "png")
                nt.system("color 27")
            except AttributeError:
                nt.system("color 47")
            self.right_stext.insert(tk.END, f"\n\nScreenshot(屏幕截图)={time_s}\n\n")

        def chat_record_receive():
            name_list = []
            try:
                os.system(f'start {self.wxaddress}')
                time.sleep(1)
                wechat = WeChat()
                for i in self.left_stext.get('1.0', tk.END).strip().split('\n'):
                    name_list.append(i)
                for i in name_list:
                    wechat.Search(i)
                    msgs = wechat.GetAllMessage
                    with open(f'./chat-record/和“{i}”的聊天记录.csv', mode='w', encoding='utf-8-sig', newline='') as f:
                        f_writer = csv.writer(f)
                        for msg in msgs:
                            f_writer.writerow([msg[0], msg[1]])
                        self.right_stext.insert(tk.END, f'与{i}的聊天记录接收完成')
            except:
                pass
            wechat.UiaAPI.GetWindowPattern().Close()

        right_lf = ttk.Labelframe(self.send_f_right_main, text=" 发 送 格 式 以 及 聊 天 记 录 接 收", style=ttk.PRIMARY)
        right_lf.pack(anchor='n', pady=20)
        ttk.Button(right_lf, text=' 聊 天 记 录 ', style="success-link", command=chat_record_receive) \
            .pack(side='right', padx=10, pady=10)
        ttk.Button(right_lf, text=' 屏 幕 截 图 ', style="success-link", command=send_right_image).pack(side='left',
                                                                                                    padx=10, pady=10)
        ttk.Button(right_lf, text=' 文 件 ', style="success-link", command=send_right_file).pack(side='top', pady=10)

        scroll = ttk.Scrollbar(self.send_f_right_main, style="round-success")
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.right_stext = tk.Text(self.send_f_right_main, font=('宋体', 12), spacing1=5, spacing2=5, spacing3=5)
        self.right_stext.pack(fill='both', expand=1)
        ttk.Button(self.right_stext, text=' 发 送 ', style=ttk.SUCCESS, command=self.right_send, cursor='arrow').pack(
            anchor='s',
            side='bottom')
        scroll.config(command=self.right_stext.yview)
        self.right_stext.config(yscrollcommand=scroll.set)

    def timed_send(self):
        def new_timed():
            def cancel():
                entry.destroy()
                entry2.destroy()
                entry3.destroy()
                entry4.destroy()
                f_temp.destroy()
                self.left_stext.delete('1.0', tk.END)
                self.right_stext.delete('1.0', tk.END)

            def determine():
                l_w = []
                msgs = {'任务名称': entry.get(), '发送频率': entry2.get(), '发送时间': entry3.get(),
                        '接收人': self.left_stext.get('1.0', tk.END), '发送内容': self.right_stext.get('1.0', tk.END),
                        '任务模式': entry4.get()}
                with open(f'timed.csv', mode='a', encoding='utf-8-sig', newline='') as f:
                    f_writer = csv.writer(f)
                    for i in msgs:
                        l_w.append(msgs[i])
                    f_writer.writerow(l_w)
                cancel()

            self.left_stext.delete('1.0', tk.END)
            self.right_stext.delete('1.0', tk.END)
            entry = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry.insert('0', "任务名称")
            entry.pack(pady=20, fill='x')
            entry2 = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry2.insert('0', "发送频率,有多项时用“--”隔开，如：星期天--10--3")
            entry2.pack(pady=20, fill='x')
            entry3 = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry3.insert('0', "发送时间,有多项时用“--”隔开，如：09:00--20:00")
            entry3.pack(pady=20, fill='x')
            entry4 = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry4.insert('0', "循环还是单次")
            entry4.pack(pady=20, fill='x')
            f_temp = tk.Frame(frame_bottom)
            f_temp.pack(pady=20)
            ttk.Button(f_temp, text=' 取 消 ', style=ttk.SUCCESS, command=cancel).pack(side='left', padx=20)
            ttk.Button(f_temp, text=' 确 认 ', style=ttk.SUCCESS, command=determine).pack(side='right', padx=20)

        def delete_timed():
            def cancel():
                entry.destroy()
                f_temp.destroy()
                self.left_stext.delete('1.0', tk.END)
                self.right_stext.delete('1.0', tk.END)

            def determine():
                list_f = []
                with open(f'timed.csv', mode='r', encoding='utf-8-sig', newline='') as f:
                    f_reader = csv.reader(f)
                    for i in f_reader:
                        if i[0] != entry.get().strip():
                            list_f.append(i)
                with open(f'timed.csv', mode='w', encoding='utf-8-sig', newline='') as f:
                    f_writer = csv.writer(f)
                    for i in list_f:
                        f_writer.writerow(i)
                cancel()

            self.left_stext.delete('1.0', tk.END)
            self.right_stext.delete('1.0', tk.END)
            entry = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry.insert('0', "任务名称")
            entry.pack(pady=20, fill='x')
            f_temp = tk.Frame(frame_bottom)
            f_temp.pack(pady=20)
            ttk.Button(f_temp, text=' 取 消 ', style=ttk.SUCCESS, command=cancel).pack(side='left', padx=20)
            ttk.Button(f_temp, text=' 确 认 ', style=ttk.SUCCESS, command=determine).pack(side='right', padx=20)

        def see_timed():
            self.left_stext.delete('1.0', tk.END)
            self.right_stext.delete('1.0', tk.END)
            with open(f'timed.csv', mode='r', encoding='utf-8-sig', newline='') as f:
                f_reader = csv.reader(f)
                for i in f_reader:
                    self.left_stext.insert(tk.END, f'{i}\n\n')

        def modify_timed():
            def confirm():
                list_f = []
                for i in self.left_stext.get('1.0', tk.END).split('\n\n'):
                    list_f.append(i)
                with open(f'timed.csv', mode='a', encoding='utf-8-sig', newline='') as f:
                    f_writer = csv.writer(f)
                    f_writer.writerow(list_f)
                cancel()

            def cancel():
                entry.destroy()
                f_temp.destroy()
                self.left_stext.delete('1.0', tk.END)
                self.right_stext.delete('1.0', tk.END)

            def determine():
                list_f = []
                ttk.Button(f_temp, text=' 确 认 修 改 ', style=ttk.SUCCESS, command=confirm).pack(side='right', padx=20)
                with open(f'timed.csv', mode='r', encoding='utf-8-sig', newline='') as f:
                    f_reader = csv.reader(f)
                    for i in f_reader:
                        if i[0] == entry.get().strip():
                            for j in i:
                                self.left_stext.insert(tk.END, f'{j}\n\n')
                        else:
                            list_f.append(i)
                print(list_f)
                with open(f'timed.csv', mode='w', encoding='utf-8-sig', newline='') as f:
                    f_writer = csv.writer(f)
                    for i in list_f:
                        f_writer.writerow(i)

            self.left_stext.delete('1.0', tk.END)
            self.right_stext.delete('1.0', tk.END)
            entry = ttk.Entry(frame_bottom, textvariable=ttk.StringVar())
            entry.insert('0', "任务名称")
            entry.pack(pady=20, fill='x')
            f_temp = tk.Frame(frame_bottom)
            f_temp.pack(pady=20)
            ttk.Button(f_temp, text=' 取 消 ', style=ttk.SUCCESS, command=cancel).pack(side='left', padx=20)
            ttk.Button(f_temp, text=' 确 认 ', style=ttk.SUCCESS, command=determine).pack(side='right', padx=20)

        frame = tk.Frame(self.send_f_right_timing)
        frame.pack(expand=1, fill='both')
        labelframe = ttk.Labelframe(frame, text=" 定 时 任 务 ", style=ttk.PRIMARY)
        labelframe.pack(side='top', pady=20)
        frame_bottom = tk.Frame(frame)
        frame_bottom.pack(side='top', expand=1, fill='both')
        ttk.Button(labelframe, text=' 新 建 ', style="success-link", command=new_timed).pack(side='left', padx=10, pady=5)
        ttk.Button(labelframe, text=' 删 除 ', style="success-link", command=delete_timed).pack(side='left', padx=10,
                                                                                              pady=5)
        ttk.Button(labelframe, text=' 查 看 ', style="success-link", command=see_timed).pack(side='left', padx=10, pady=5)
        ttk.Button(labelframe, text=' 修 改 ', style="success-link", command=modify_timed).pack(side='left', padx=10,
                                                                                              pady=5)

    def all_update(self):
        a = True
        with open('init.ini', mode='w', encoding='utf-8') as f:
            f.close()
        for i in 'DEFCGHIJK':
            if a:
                for root, dirs, files in os.walk(i + ':\\'):
                    for file in files:
                        path = os.path.join(root, file)
                        if fnmatch(path, "*WeChat.exe*") and a:
                            a = False
                            with open('init.ini', mode='a', encoding='utf-8') as f:
                                f.write(f'wx_path={path}\n')
                            self.wxaddress = path
                            os.system(f'start {path}')
                            break
        wechat = WeChat()
        os.system(f'start {self.wxaddress}')
        time.sleep(1)
        groupname = wechat.GetGroupListName()
        with open('init.ini', mode='a', encoding='utf-8') as f:
            f.write(f"GroupName={groupname}=GroupName\n")
        time.sleep(1)
        wechat.UiaAPI.GetWindowPattern().Close()
        os.system(f'start {self.wxaddress}')
        time.sleep(1)
        dicts = wechat.GetFriendListName()
        with open('init.ini', mode='a', encoding='utf-8') as f:
            f.write(f"PeopleName={dicts}=PeopleName\n")
        wechat.UiaAPI.GetWindowPattern().Close()
        os.system(f'start {self.wxaddress}')
        time.sleep(1)
        for i in groupname:
            HaveBecomeFriends, NotFriended, name = wechat.GetGroupMemberName(i)
            with open('init.ini', mode='a', encoding='utf-8') as f:
                f.write(f"{name}={HaveBecomeFriends}={name}--{NotFriended}\n")
        wechat.UiaAPI.GetWindowPattern().Close()

    def program_help(self):
        scroll = ttk.Scrollbar(self.start_f_about, style="round-success")
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        a_text = tk.Text(self.start_f_about, spacing1=5, spacing2=5, spacing3=5, font=('宋体', 15))
        a_text.pack(expand=1, fill='both', side='left')
        scroll.config(command=a_text.yview)
        a_text.config(yscrollcommand=scroll.set)
        a_text.insert(tk.END, "作者箴言：此软件完成免费,为作者个人学习项目,禁止用来商用,如遇到BUG（程序出问题）请联系作者\n")
        a_text.insert(tk.END, "作者微信：ershui2500       作者电话：13672773441      作者QQ：3294161978\n\n")
        a_text.insert(tk.END, "一、更新名单：在第一次使用此程序时建议先更新一下名单,这样以后的操作就比较方便,快捷\n\n")
        a_text.insert(tk.END, "二、发送好友：\n")
        a_text.insert(tk.END, "1.群聊：在群聊下面的输入框里输入要获取的群聊名称,他会自动获取此群聊里您已加的好友,并用列表形式显示出来,在列表选择好需要发送的好友后,"
                              "点击底下的确认后就可以把选择的好友打印到输入框里\n\n")
        a_text.insert(tk.END, "2.群聊名称：在通讯录管理里面获取您微信最近使用的十几个群聊名称,并用列表形式显示出来,在列表选择好需要获取好友的群聊后,"
                              "点击底下的确认后就可以把选择的群聊打印到输入框里\n\n")
        a_text.insert(tk.END, "3.标签：在通讯录管理里面获取您微信所有的好友,并按照您给他们的标签排序好,并用列表形式显示出来,在列表选择好需要需要发送的好友后,"
                              "点击底下的确认后就可以把选择的好友打印到输入框里\n\n")
        a_text.insert(tk.END, "三、发送信息：\n")
        a_text.insert(tk.END, "1.屏幕截图：在点击这个按钮前先要截图,然后点击按钮就可以了,这样下面的输入框会出现截图的时间,不要删除他打印出来的东西\n\n")
        a_text.insert(tk.END, "2.文件：选择发送的文件,之后他会在输入框里显示文件的路径,不要删除他打印出来的东西\n\n")
        a_text.insert(tk.END, "3.聊天记录：在发送好友的输入框里输入好友名字,点击聊天记录按钮,就可以获取该好友的聊天记录了,聊天记录保存在和这程序同一文件夹的chat-record文件夹里\n\n")
        a_text.insert(tk.END, "四、定时发送：在用这个时注意新建时要发送人和发送的信息都输入到对应的输入框里,之后点击定时发送里面的新建的确认按钮完成,后面的其他三个也一样,"
                              "不要点击发现好友输入框里面的确认按钮\n\n")
        a_text.config(state=tk.DISABLED)


if __name__ == '__main__':
    wa = WxAssistant()
    wa.theme()
    wa.time_display()
    wa.start_panel()
    wa.send_panel()
    wa.send_left_panel()
    wa.send_right_panel()
    wa.timed_send()
    wa.program_help()
    wa.mainloop()
