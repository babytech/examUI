#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
本程序设计为图形化程序
模拟学多多APP
"""
import base64
import json
from pathlib import Path
import os
import pickle
import threading
import time
import tkinter as tk
import tkinter.messagebox
from tkinter.simpledialog import *
from tkinter import *
from tkinter import filedialog, dialog, ttk
import random
import pandas as pd
import pygal as pygal
from pygal.style import Style
import requests
from urllib.request import urlopen
from tqdm import tqdm
from PIL import Image, ImageTk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import pyttsx3
from aip import AipSpeech
from playsound import playsound

"""
版本历史信息--->
 v0.0.1: 支持图形化显示、支持单选题
 v0.0.2: 重构部分代码，界面增加功能按钮，支持开始测验、提交答案
 v0.0.3: 增加登录页面、支持多选题
 v0.0.4: 增加JSON转EXCEL、EXCEL转JSON功能，可用于导入题库
 v0.0.5: 增加帮助按钮，获得提示信息，增加答题倒计时功能
 v0.0.6: 增加只看错题功能，支持多用户登录
 v0.0.7: 增加生成试卷功能，支持设置考试时长，支持选择测试课程
 v0.0.8: 增加历史统计功能，支持多用户按成绩排名
 v0.0.9: 增加通过邮件找回密码功能，支持成绩按单科排名
 v0.1.0: 支持成绩按科目总分排名，支持填空题
 v0.1.1: 支持家长版、学生版配置不同界面，供登录时进行选择
 v0.1.2: 增加成绩统计分布图形化显示功能
 v0.1.3: 增加密钥加密和解密功能，可用来访问百度文字转语音请求
 v0.1.4: 增加播放录音按钮，支持英语听力题
"""
AUTHOR_INFO = "宝宝科技"
CONTACT_INFO = "babytech@126.com"
WECHAT_INFO = "microbabytech"
VERSION_INFO = "鸡娃程序 学多多 演示版本 v0.1.4"
AUTHORIZATION_FILE = "auth.key"
DEFAULT_JSON_FILE = "exam.json"
EXAM_TIME = 30 * 60
SUB_DIR = "resource"
UPLOAD_SUFFIX = "/upload"
SMTP_SERVER_ADDRESS = 'smtp.126.com'
WELCOME_INFO = '''
                           _   _ ___
   _____  ____ _ _ __ ___ | | | |_ _|
  / _ \ \/ / _` | '_ ` _ \| | | || |
 |  __/>  < (_| | | | | | | |_| || |
  \___/_/\_\__,_|_| |_| |_|\___/|___|

'''


class LoginWindow:
    user = None
    mode = None
    smtp_token = ""

    def __init__(self):
        self.frame = None
        self.login = True
        print(VERSION_INFO)
        print(WELCOME_INFO)
        print(AUTHOR_INFO, CONTACT_INFO)

    def login_window(self, master):
        def on_closing():
            if tk.messagebox.askokcancel("退出", "要退出登录吗?"):
                master.destroy()
                self.login = False

        # 登录函数
        def usr_log_in():
            def send_mail():
                # 创建一个生成6位数的随机验证码
                def create_random():
                    get_random = ''
                    for i in range(6):
                        one_number = str(random.randint(0, 9))
                        get_random += one_number
                    return get_random

                def usr_password_reset():
                    nn = new_name.get()
                    if id_code != id_pwd.get():
                        tk.messagebox.showerror(message='验证码错误，重置密码失败')
                        return
                    if new_pwd.get() != new_pwd_confirm.get():
                        tk.messagebox.showerror('错误', '新密码前后不一致')
                        return
                    try:
                        with open('usr_info.pickle', 'rb') as _usr_file:
                            exist_usr_info = pickle.load(_usr_file)
                    except FileNotFoundError:
                        exist_usr_info = {}
                    if nn in exist_usr_info:
                        exist_usr_info[nn] = new_pwd.get()
                        with open('usr_info.pickle', 'wb') as _usr_file:
                            pickle.dump(exist_usr_info, _usr_file)
                        tk.messagebox.showinfo('重置密码', '重置密码成功')
                        reset_password_window.destroy()

                top_window.destroy()
                # 收件人邮箱
                receiver = mail_str.get()
                print("收件人邮箱: ", receiver)
                try:
                    with open('usr_contact.pickle', 'rb') as usr_contact_file:
                        exist_usr_contact = pickle.load(usr_contact_file)
                except FileNotFoundError:
                    exist_usr_contact = {}
                if usr_name in exist_usr_contact:
                    if receiver != exist_usr_contact[usr_name]:
                        error_str = '用户名: ' + usr_name + '\n'
                        error_str += '注册邮箱: ' + exist_usr_contact[usr_name] + '\n'
                        error_str += '找回密码邮箱: ' + receiver + '\n'
                        error_str += '请使用注册邮箱找回密码' + '\n'
                        tk.messagebox.showerror('找回密码邮箱与注册邮箱不符', error_str)
                        return
                else:
                    error_str = '用户名: ' + usr_name + '未注册\n'
                    tk.messagebox.showerror('用户名未注册', error_str)
                    return
                # 验证码
                id_code = create_random()
                print("随机验证码: ", id_code)
                smtp_server = SMTP_SERVER_ADDRESS
                username = CONTACT_INFO
                password = LoginWindow.get_smtp_token().strip("\n")
                # sender一般要与username一样
                sender = username
                subject = VERSION_INFO + '找回登录密码验证邮件(请勿回复)'
                subject = Header(subject, 'utf-8').encode()
                # 构造邮件对象MIMEMultipart对象
                # 主题，发件人，收件人等显示在邮件页面上的。
                msg = MIMEMultipart('mixed')
                msg['Subject'] = subject
                msg['From'] = AUTHOR_INFO
                msg['To'] = receiver
                # 构造文字内容
                receiver_user, sep, receiver_server = receiver.partition('@')
                text = 'Hi!' + receiver_user + '\n' + '这是你的验证码：' + id_code + '\n'
                print(text)
                text_plain = MIMEText(text, 'plain', 'utf-8')
                msg.attach(text_plain)
                # 发送邮件
                try:
                    smtp = smtplib.SMTP()
                    smtp.connect(smtp_server)
                    # 用set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
                    # smtp.set_debuglevel(1)
                    smtp.login(username, password)
                    smtp.sendmail(sender, receiver, msg.as_string())
                    smtp.quit()
                except smtplib.socket.error:
                    error_str = '连接邮件服务器: ' + smtp_server + '失败\n'
                    tk.messagebox.showerror(title='错误', message=error_str)
                    return
                # 重设密码界面
                reset_password_window = tk.Toplevel(login_frame)
                window_position_center(reset_password_window, 350, 210)
                reset_password_window.title('忘记密码')
                new_name = tk.StringVar()
                tk.Label(reset_password_window, text='用户名：').place(x=10, y=10)
                tk.Entry(reset_password_window, textvariable=new_name).place(x=150, y=10)
                id_pwd = tk.StringVar()
                tk.Label(reset_password_window, text='请输入验证码：').place(x=10, y=50)
                tk.Entry(reset_password_window, textvariable=id_pwd).place(x=150, y=50)
                new_pwd = tk.StringVar()
                tk.Label(reset_password_window, text='请输入新密码：').place(x=10, y=90)
                tk.Entry(reset_password_window, textvariable=new_pwd, show='*').place(x=150, y=90)
                new_pwd_confirm = tk.StringVar()
                tk.Label(reset_password_window, text='再次输入密码：').place(x=10, y=130)
                tk.Entry(reset_password_window, textvariable=new_pwd_confirm, show='*').place(x=150, y=130)
                tk.Button(reset_password_window, text='重设密码', command=usr_password_reset).place(x=150, y=170)

            # 输入框获取用户名密码
            usr_name = var_usr_name.get().strip()
            usr_pwd = var_usr_pwd.get()
            usr_config = var_usr_config.get()
            LoginWindow.set_mode(usr_config)
            # 从本地字典获取用户信息，如果没有则新建本地数据库
            try:
                with open('usr_info.pickle', 'rb') as usr_file:
                    usr_info = pickle.load(usr_file)
            except FileNotFoundError:
                with open('usr_info.pickle', 'wb') as usr_file:
                    usr_info = {'admin': 'admin'}
                    pickle.dump(usr_info, usr_file)
            # 判断用户名和密码是否匹配
            if usr_name in usr_info:
                if usr_pwd == usr_info[usr_name]:
                    lw = LoginWindow()
                    lw.set_user(usr_name)
                    lw.validate_mode(master)
                else:
                    tk.messagebox.showerror(message='密码错误')
                    find_back = tk.messagebox.askyesno('登录失败', '密码输入错误，是否找回密码？')
                    if find_back:
                        top_window = tk.Toplevel()
                        window_position_center(top_window, 260, 100)
                        top_window.title('找回密码')
                        top_window.resizable(height=False, width=False)
                        mail_str = tk.StringVar()
                        tk.Label(top_window, text='邮箱：').place(x=10, y=10)
                        tk.Entry(top_window, textvariable=mail_str).place(x=60, y=10)
                        tk.Button(top_window, text='发送邮件获取验证码', command=send_mail).place(x=60, y=40)
            # 用户名密码不能为空
            elif usr_name == '' or usr_pwd == '':
                tk.messagebox.showerror(message='用户名或密码为空')
            # 不在数据库中弹出是否注册的框
            else:
                is_signup = tk.messagebox.askyesno('欢迎', '您还没有注册，是否现在注册')
                if is_signup:
                    usr_sign_up()

        # 注册函数
        def usr_sign_up():
            # 确认注册时的相应函数
            def usr_sign_to():
                # 获取输入框内的内容
                nn = new_name.get().strip()
                np = new_pwd.get()
                npf = new_pwd_confirm.get()
                nm = new_mail.get()
                # 本地加载已有用户信息,如果没有则已有用户信息为空
                try:
                    with open('usr_info.pickle', 'rb') as usr_info_file:
                        exist_usr_info = pickle.load(usr_info_file)
                except FileNotFoundError:
                    exist_usr_info = {}
                # 检查用户名存在、密码为空、密码前后不一致
                if nn in exist_usr_info:
                    tk.messagebox.showerror('错误', '用户名已存在')
                elif np == '' or nn == '':
                    tk.messagebox.showerror('错误', '用户名或密码为空')
                elif np != npf:
                    tk.messagebox.showerror('错误', '密码前后不一致')
                # 注册信息没有问题则将用户名密码写入数据库
                else:
                    exist_usr_info[nn] = np
                    with open('usr_info.pickle', 'wb') as usr_info_file:
                        pickle.dump(exist_usr_info, usr_info_file)
                    try:
                        with open('usr_contact.pickle', 'rb') as usr_contact_file:
                            exist_usr_contact = pickle.load(usr_contact_file)
                    except FileNotFoundError:
                        exist_usr_contact = {}
                    if nn in exist_usr_contact:
                        tk.messagebox.showerror('错误', '用户名已存在')
                    else:
                        exist_usr_contact[nn] = nm
                        with open('usr_contact.pickle', 'wb') as usr_contact_file:
                            pickle.dump(exist_usr_contact, usr_contact_file)
                    tk.messagebox.showinfo('欢迎', '注册成功')
                    # 注册成功关闭注册框
                    window_sign_up.destroy()

            # 新建注册界面
            window_sign_up = tk.Toplevel(login_frame)
            window_position_center(window_sign_up, 350, 220)
            window_sign_up.resizable(width=False, height=False)
            window_sign_up.title('注册')
            # 用户名变量及标签、输入框
            new_name = tk.StringVar()
            tk.Label(window_sign_up, text='用户名：').place(x=10, y=10)
            tk.Entry(window_sign_up, textvariable=new_name).place(x=150, y=10)
            # 密码变量及标签、输入框
            new_pwd = tk.StringVar()
            tk.Label(window_sign_up, text='请输入密码：').place(x=10, y=50)
            tk.Entry(window_sign_up, textvariable=new_pwd, show='*').place(x=150, y=50)
            # 重复密码变量及标签、输入框
            new_pwd_confirm = tk.StringVar()
            tk.Label(window_sign_up, text='请再次输入密码：').place(x=10, y=90)
            tk.Entry(window_sign_up, textvariable=new_pwd_confirm, show='*').place(x=150, y=90)
            # 注册邮箱变量及标签、输入框
            new_mail = tk.StringVar()
            tk.Label(window_sign_up, text='注册邮箱：').place(x=10, y=130)
            tk.Entry(window_sign_up, textvariable=new_mail).place(x=150, y=130)
            # 确认注册按钮及位置
            bt_confirm_sign_up = tk.Button(window_sign_up, text='确认注册', command=usr_sign_to)
            bt_confirm_sign_up.place(x=150, y=170)

        master.title("请登录鸡娃程序")
        master.protocol("WM_DELETE_WINDOW", on_closing)
        window_position_center(master, 260, 130)
        login_frame = tk.Frame(master)
        login_frame.grid(padx=15, pady=15)
        # 标签 用户名密码
        ttk.Label(login_frame, text='用户名').grid(column=1, row=1, columnspan=2)
        ttk.Label(login_frame, text='密码').grid(column=1, row=2, columnspan=2)
        # 用户名输入框
        var_usr_name = tk.StringVar()
        entry_usr_name = ttk.Entry(login_frame, textvariable=var_usr_name)
        entry_usr_name.grid(column=3, row=1, columnspan=3)
        # 密码输入框
        var_usr_pwd = tk.StringVar()
        entry_usr_pwd = ttk.Entry(login_frame, textvariable=var_usr_pwd, show='*')
        entry_usr_pwd.grid(column=3, row=2, columnspan=3)
        # 配置输入框
        var_usr_config = tk.StringVar()
        var_usr_config.set("学生版")
        entry_usr_student = ttk.Radiobutton(login_frame, text="学生版", variable=var_usr_config, value="学生版")
        entry_usr_student.grid(column=2, row=3, columnspan=3)
        entry_usr_parent = ttk.Radiobutton(login_frame, text="家长版", variable=var_usr_config, value="家长版")
        entry_usr_parent.grid(column=5, row=3, columnspan=4)
        ttk.Button(login_frame, text='注册', command=usr_sign_up).grid(column=3, row=4, columnspan=1, pady=5)
        ttk.Button(login_frame, text='登录', command=usr_log_in).grid(column=5, row=4, columnspan=1, pady=5)
        self.frame = login_frame
        return login_frame

    @classmethod
    def get_smtp_token(cls):
        decrypt_dict = {}
        try:
            with open(file=AUTHORIZATION_FILE, mode='r+', encoding='utf-8') as key_file:
                raw_text_lines = key_file.readlines()
            for raw_text in raw_text_lines:
                key_prefix, key_content = raw_text.split(":")
                key = key_content.strip()
                if key_prefix == 'SMTP_TOKEN':
                    decrypt_dict['SMTP_TOKEN'] = base64.b64decode(key).decode('utf-8')
            return decrypt_dict['SMTP_TOKEN']
        except FileNotFoundError:
            tk.messagebox.showerror(title='错误', message='找不到密钥文件: ' + AUTHORIZATION_FILE)
            return ""

    @classmethod
    def set_user(cls, usr_name):
        cls.user = usr_name

    @classmethod
    def get_user(cls):
        return cls.user

    @classmethod
    def set_mode(cls, usr_mode):
        cls.mode = usr_mode

    @classmethod
    def get_mode(cls):
        return cls.mode

    def validate_mode(self, master):
        def do_validate():
            def is_valid_idiom(sample_chars):
                for entry_char in entry.get():
                    if entry_char not in sample_chars:
                        return False
                return True

            right_counts = 0
            window_validation.destroy()
            for entry in idiom_input_list:
                if entry.get() in idiom_words and is_valid_idiom(sample_char_list) is True:
                    right_counts += 1
            if right_counts < 2:
                tk.messagebox.showerror(title='验证失败', message='您填入的成语不符合要求\n以学生版登录')
                LoginWindow.set_mode("学生版")
            else:
                tk.messagebox.showinfo(title='欢迎', message='欢迎您：' + self.get_user() + '\n登录专业版')
            master.destroy()
            return

        if self.get_mode() == "学生版":
            tk.messagebox.showinfo(title='欢迎', message='欢迎您：' + self.get_user() + '\n登录学生版')
            master.destroy()
            return
        # 新建专业版验证界面
        window_validation = tk.Toplevel(self.frame)
        window_validation.wm_attributes('-topmost', True)
        window_position_center(window_validation, 310, 220)
        window_validation.resizable(width=False, height=False)
        window_validation.title('专业版验证')
        welcome_str = "Hi, " + self.user + "请选择下面合适的字组成2个成语"
        validate_label_frame = LabelFrame(window_validation, text=welcome_str, padx=5, pady=5)
        validate_label_frame.place(x=10, y=10)
        char_list = []
        idiom_words = ["黔驴技穷", "高深莫测", "学海无涯", "呕心沥血", "专心致志", "志同道合", "聚精会神", "博大精深"]
        for idiom in idiom_words:
            for char in idiom:
                char_list.append(char)
        sort_char_list = sorted(set(char_list), key=char_list.index)
        sample_char_list = random.sample(sort_char_list, 24)
        sample_char_str = "\t"
        for i in range(len(sample_char_list)):
            sample_char_str += sample_char_list[i] + " "
            if (i + 1) % 8 == 0:
                sample_char_str += '\n\t'
        validate_label = ttk.Label(validate_label_frame, text=sample_char_str)
        validate_label.place(x=30, y=30)
        validate_label.pack(anchor=W)
        idiom_input_list = []
        for i in range(2):
            idiom_str = tk.StringVar()
            validate_entry = ttk.Entry(window_validation, textvariable=idiom_str, width=10)
            validate_entry.place(x=20, y=130 + i * 40)
            idiom_input_list.append(idiom_str)
        ttk.Button(window_validation, text='确认', command=do_validate).place(x=200, y=150)
        return True


class WindowShell:
    def __init__(self, _window, axes):
        self.file_dir = Path(SUB_DIR)
        self.timer_start = time.time()
        self.timer_running = True
        self.start_exam = False
        self.finish_exam = False
        self.load_exam = False
        self.look_wrong = False
        self.run_threads = True
        self.audio_mode = "local"
        self.window = _window
        self.child = self
        self.exam_time = EXAM_TIME
        self.author_info = AUTHOR_INFO
        self.contact_info = CONTACT_INFO
        self.wechat_info = WECHAT_INFO
        self.version_info = VERSION_INFO
        self.window.title(VERSION_INFO)
        self.window.geometry('800x600+600+400')
        self.frame = tk.Frame(self.window, padx=5)
        self.frame.grid(row=axes[0], column=axes[1])
        # 标签
        self.input_label = Label(self.window, text="试题显示")
        self.input_label.place(x=330, y=45)
        self.input_exam_label = Label(self.window, text="")
        self.input_exam_label.place(x=390, y=45)
        self.output_label = Label(self.window, text="结果显示")
        self.output_label.place(x=330, y=370)
        self.output_exam_label = Label(self.window, text="")
        self.output_exam_label.place(x=390, y=370)
        self.timer_label = Label(self.window, text="考试总时长为: " + str(self.exam_time // 60) + "分钟")
        self.timer_label.place(x=120, y=30)

        # 处理结果展示
        self.input_Text = Text(self.window, width=60, height=17)
        self.input_Text.place(x=120, y=70)
        self.input_Text.configure(state="disabled")
        self.output_Text = Text(self.window, width=60, height=10)
        self.output_Text.place(x=120, y=400)
        self.output_Text.configure(state="disabled")

        # 右侧按钮
        administrator, _ = CONTACT_INFO.rsplit("@")
        self.play_audio_button = Button(self.window, text="播放录音", bg="lightblue",
                                        width=10, command=self.play_audio)
        self.play_audio_button.place(x=500, y=30)
        if LoginWindow.get_user() == administrator:
            self.encrypt_button = Button(self.window, text="生成密钥", bg="lightblue", width=10, command=self.run_encrypt)
            self.encrypt_button.place(x=650, y=0)
            self.decrypt_button = Button(self.window, text="查看密钥", bg="lightblue", width=10, command=self.run_decrypt)
            self.decrypt_button.place(x=650, y=30)
        self.load_button = Button(self.window, text="导入题库", bg="lightblue", width=10, command=self.run_load)
        self.load_button.place(x=650, y=60)
        self.exam_button = Button(self.window, text="开始测验", bg="lightblue", width=10, command=self.run_exam)
        self.exam_button.place(x=650, y=90)
        self.submit_button = Button(self.window, text="提交答案", bg="lightblue", width=10, command=self.run_submit)
        self.submit_button.place(x=650, y=120)
        self.help_button = Button(self.window, text="帮助", bg="lightblue", width=10, command=self.run_help)
        self.help_button.place(x=650, y=150)
        self.history_button = Button(self.window, text="历史统计", bg="lightblue", width=10, command=self.run_history)
        self.history_button.place(x=650, y=180)

        # 左侧按钮
        self.time_set_button = Button(self.window, text="设置考试时长", bg="lightblue",
                                      width=10, command=self.set_time)
        self.time_set_button.place(x=5, y=70)
        self.json2excel_button = Button(self.window, text="JSON转EXCEL", bg="lightblue",
                                        width=10, command=self.json2excel)
        self.json2excel_button.place(x=5, y=160)
        self.excel2json_button = Button(self.window, text="EXCEL转JSON", bg="lightblue",
                                        width=10, command=self.excel2json)
        self.excel2json_button.place(x=5, y=190)
        if LoginWindow.get_mode() == "家长版":
            self.gen_paper_button = Button(self.window, text="生成试卷", bg="lightblue",
                                           width=10, command=self.gen_paper)
            self.gen_paper_button.place(x=5, y=100)
            self.gen_answer_button = Button(self.window, text="参考答案", bg="lightblue",
                                            width=10, command=self.gen_answer)
            self.gen_answer_button.place(x=5, y=130)
            self.open_button = Button(self.window, text="打开文件", bg="lightblue",
                                      width=10, command=self.open_file)
            self.open_button.place(x=5, y=220)
            self.save_button = Button(self.window, text="保存文件", bg="lightblue",
                                      width=10, command=self.save_file)
            self.save_button.place(x=5, y=250)
            self.upload_button = Button(self.window, text="上传文件", bg="lightblue",
                                        width=10, command=self.upload_file)
            self.upload_button.place(x=5, y=280)
            self.download_button = Button(self.window, text="下载文件", bg="lightblue",
                                          width=10, command=self.download_file)
            self.download_button.place(x=5, y=310)
        self.clear_button = Button(self.window, text="清除文本框", bg="lightblue", width=10, command=self.clear_text)
        self.clear_button.place(x=5, y=340)
        self.wrong_button = Button(self.window, text="只看错题", bg="lightblue", width=10, command=self.run_wrong)
        self.wrong_button.place(x=5, y=400)
        self.first_button = Button(self.window, text="第一道题", bg="lightblue", width=10, command=self.run_first)
        self.first_button.place(x=5, y=440)
        self.last_button = Button(self.window, text="最后一题", bg="lightblue", width=10, command=self.run_last)
        self.last_button.place(x=5, y=470)
        self.previous_button = Button(self.window, text="上一道题", bg="lightblue", width=10, command=self.run_previous)
        self.previous_button.place(x=5, y=510)
        self.next_button = Button(self.window, text="下一道题", bg="lightblue", width=10, command=self.run_next)
        self.next_button.place(x=5, y=540)

        # Window
        self.history_windows = []
        self.paper_windows = []
        self.answer_windows = []

        # Frame
        self.label_frame = LabelFrame(self.window, text="你选择的答案是？", padx=5, pady=5)
        self.label_frame.place(x=620, y=230)
        self.score_frame = None
        self.image_frame = []
        self.image_threads = []

        # Combo Box
        number = tk.StringVar()
        self.course_combobox = ttk.Combobox(self.window, width=12, textvariable=number, postcommand=self.run_choose)
        self.course_combobox['values'] = ("英语", "数学", "语文", "综合")
        self.course_combobox['state'] = 'readonly'
        self.course_combobox.place(x=5, y=30)
        self.course_combobox.current(0)

        # Data
        self.index = 0
        self.params = []
        self.questions = []
        self.questions_name = {}
        self.questions_dict = {}
        self.questions_type = {}
        self.questions_audio = {}
        self.audios_dict = {}
        self.answers_right = {}
        self.answers_yours = {}
        self.answers_result = {}
        self.choices_list = []
        self.choices_str = []
        self.fills_dict = {}
        self.radio_button = []
        self.check_button = []
        self.check_value = []
        self.fill_entry = []
        self.fill_label = []
        self.fill_value = []
        self.wrong_questions = []
        self.wrong_choices_str = []
        self.wrong_choices_list = []
        self.wrong_fills_dict = {}

        # register on_closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

        # init json data
        self.json_file = os.path.abspath(DEFAULT_JSON_FILE)
        try:
            f = open(self.json_file)
            f.close()
            self.json_data = self.read_from_json(self.json_file)
        except IOError:
            error_str = "缺省题库文件: " + self.json_file + "不存在" + "\n"
            self.insert_text(self.output_Text, error_str)
            error_info = "请生成" + self.json_file + "文件，然后重新“开始测验”，或者点击“导入题库”使用自定义题库" + "\n"
            self.insert_text(self.output_Text, error_info)

    @staticmethod
    def insert_text(text, string):
        text.configure(state="normal")
        text.insert(END, string)
        text.configure(state="disabled")

    @staticmethod
    def delete_text(text):
        text.configure(state="normal")
        text.delete(1.0, END)
        text.configure(state="disabled")

    @staticmethod
    def read_from_json(json_file):
        with open(json_file, 'r', encoding='utf8') as file:
            json_data = json.load(file)
        return json_data

    @staticmethod
    def read_from_excel(excel_file):
        excel_data = pd.read_excel(excel_file)
        return excel_data

    @staticmethod
    def parse_data(content, parameter):
        question, question_name, question_dict, question_type, question_audio, audio_dict, \
            answer, choice_str, choice_list, fill_dict = parameter
        strings = []
        lists = []
        for key, value in content.items():
            if key == 'Questions':
                for sub_content in value:
                    for sub_key, sub_value in sub_content.items():
                        if sub_key == 'answer':
                            answer[sub_content['name']] = sub_value
                        if sub_key == 'type':
                            question_type[sub_content['name']] = sub_value
                        if sub_key == 'question':
                            question_title = sub_content['name'] + ":\n" + sub_value + '\n\n'
                            question.append(question_title)
                            question_name[sub_content['name']] = sub_content['name']
                            question_dict[sub_content['name']] = sub_value
                        if sub_key == 'audio':
                            audio_dict[sub_content['name']] = sub_value
                        if sub_key == 'choices':
                            for d in sub_value:
                                lists.append(list(d.values()))
                                d_str = '. '.join(str(i) for i in list(d.values())) + '\n'
                                strings.append(d_str)
                            choice_str.append(strings)
                            choice_list.append(lists)
                            strings = []
                            lists = []
                        if sub_key == 'fills':
                            for d in sub_value:
                                lists.append(list(d.values()))
                            fill_dict[sub_content['name']] = lists
                            lists = []
        for key, value in content.items():
            if key == 'Questions':
                for sub_content in value:
                    for sub_key, sub_value in sub_content.items():
                        if sub_key == 'audio':
                            element = sub_content['name']
                            if sub_value == "yes" and element in answer.keys():
                                question_audio[element] = question_dict[element].replace("___", answer[element])

    @staticmethod
    def check_score_level(score):
        if score == 100:
            image_file = 'smile.png'
            image_level = 'A.png'
            level = "A☆"
        elif 95 <= score < 100:
            image_file = 'smile.png'
            image_level = 'A.png'
            level = "A+"
        elif 90 <= score < 95:
            image_file = 'smile.png'
            image_level = 'A.png'
            level = "A"
        elif 85 <= score < 90:
            image_file = 'smile.png'
            image_level = 'A.png'
            level = "A-"
        elif 80 <= score < 85:
            image_file = 'smile.png'
            image_level = 'B.png'
            level = "B"
        elif 75 <= score < 80:
            image_file = 'cry.png'
            image_level = 'B.png'
            level = "B-"
        elif 60 <= score < 75:
            image_file = 'cry.png'
            image_level = 'C.png'
            level = "C"
        else:
            image_file = 'cry.png'
            image_level = 'D.png'
            level = "D"
        return image_file, image_level, level

    def run_help(self):
        help_info = self.version_info + "\n\n"
        banner_info = '请点击“下拉菜单”选择科目\n点击“导入题库”进行自定义测试\n' \
                      '点击“开始测验”进行作答\n' \
                      '然后点击“下一道题”依次作答\n直到所有题目均完成作答\n' \
                      '答题完毕后，请点击“提交答案”\n' \
                      '测试结束后，成绩和等级会在右侧显示\n' \
                      '点击“历史成绩”查看测试记录\n'
        help_info += banner_info + "\n"
        help_info += "程序作者: " + self.author_info + "\n"
        help_info += "联系方式: " + self.contact_info + "\n"
        help_info += "微信号: " + self.wechat_info + "\n"
        tk.messagebox.showinfo("帮助", help_info)

    def run_choose(self):
        def course_case(value):
            courses = {
                "英语": "english.json",
                "数学": "math.json",
                "语文": "chinese.json",
            }
            return courses.get(value, 'exam.json')

        course_str = str(self.course_combobox.get())
        message_str = "选择测试课程: " + course_str + "\n"
        tk.messagebox.showinfo("测试课程", message=message_str)
        self.delete_text(self.input_Text)
        self.delete_text(self.output_Text)
        self.insert_text(self.output_Text, message_str)
        json_file = course_case(course_str)
        question_bank_info = "题库文件: " + os.path.abspath(json_file) + "\n"
        self.insert_text(self.output_Text, "使用" + question_bank_info)
        if os.path.exists(json_file) is False:
            self.insert_text(self.output_Text, '题库文件{}不存在'.format(json_file))
            tk.messagebox.showwarning('警告', '题库文件{}不存在'.format(json_file))
        else:
            self.insert_text(self.output_Text, '题库文件{}存在'.format(json_file))
            self.json_file = json_file

    def gen_paper(self):
        def save_paper():
            file_path = filedialog.asksaveasfilename(title=u'保存文件')
            if isinstance(file_path, tuple) or file_path == "":
                return None
            file_text = text.get(1.0, END)
            if file_path is not None:
                with open(file=file_path, mode='a+', encoding='utf-8') as file:
                    file.write(file_text)
                self.delete_text(self.output_Text)
                info_str = "已保存试卷:\n" + file_path + "\n"
                self.insert_text(self.output_Text, info_str)
                dialog.Dialog(None, {'title': '文件修改', 'text': '保存完成',
                                     'bitmap': 'warning', 'default': 0,
                                     'strings': ('确定', '取消')})

        self.questions = []
        self.choices_str = []
        for window in self.paper_windows:
            if window is not None:
                window.destroy()
        self.window.wm_attributes('-topmost', False)
        top_window = tk.Toplevel()
        window_position_center(top_window, 800, 600)
        top_window.title('生成试卷')
        top_window.resizable(height=False, width=False)
        label_v = Label(top_window, text="", padx=5, pady=5)
        label_v.pack(anchor=W)
        scrollbar_v = Scrollbar(top_window)
        scrollbar_v.pack(side=RIGHT, fill=Y)
        text = Text(top_window, width=450, height=150)
        text.config(yscrollcommand=scrollbar_v.set)
        text.pack(expand=YES, fill=BOTH)
        self.params = [self.questions, self.questions_name, self.questions_dict, self.questions_type,
                       self.questions_audio, self.audios_dict,
                       self.answers_right, self.choices_str, self.choices_list, self.fills_dict]
        self.parse_data(self.json_data, self.params)
        text.delete(1.0, END)
        for idx in range(len(self.questions)):
            question_type = self.questions_type[str(idx + 1)]
            question_suffix = ""
            if question_type == "radio":
                question_suffix += ".单选题"
            elif question_type == "check":
                question_suffix += ".多选题"
            elif question_type == "fill":
                question_suffix += ".填空题"
            question_item, question_content = self.questions[idx].split(":")
            question_str = question_item + question_suffix + question_content
            text.insert(END, question_str)
            for choice_str in self.choices_str[idx]:
                text.insert(END, choice_str)
            text.insert(END, "\n")
        scrollbar_v.config(command=text.yview)
        save_file_button = Button(top_window, text="保存试卷", bg="lightblue",
                                  width=10, command=save_paper)
        save_file_button.place(x=550, y=0)
        upload_file_button = Button(top_window, text="上传试卷", bg="lightblue",
                                    width=10, command=self.upload_file)
        upload_file_button.place(x=650, y=0)
        self.paper_windows.append(top_window)

    def gen_answer(self):
        self.window.wm_attributes('-topmost', False)
        answer_str = ""
        if self.start_exam is False:
            for window in self.answer_windows:
                if window is not None:
                    window.destroy()
            top_window = tk.Toplevel()
            window_position_center(top_window, 600, 200)
            top_window.title('参考答案')
            top_window.resizable(height=False, width=False)
            scrollbar_v = Scrollbar(top_window)
            scrollbar_v.pack(side=RIGHT, fill=Y)
            text = Text(top_window, width=450, height=150)
            text.config(yscrollcommand=scrollbar_v.set)
            text.pack(expand=YES, fill=BOTH)
            text.delete(1.0, END)
            self.answers_right = {}
            self.params = [self.questions, self.questions_name, self.questions_dict, self.questions_type,
                           self.questions_audio, self.audios_dict,
                           self.answers_right, self.choices_str, self.choices_list, self.fills_dict]
            self.parse_data(self.json_data, self.params)
            answer_str += "正确答案: " + "\n" + json.dumps(self.answers_right) + "\n\n"
            text.insert(END, answer_str)
            scrollbar_v.config(command=text.yview)
            self.answer_windows.append(top_window)
        elif self.start_exam is True and self.finish_exam is False:
            self.delete_text(self.output_Text)
            answer_idx = str(self.index + 1)
            answer_str += "第" + answer_idx + "题"
            answer_str += "正确答案: " + "\n" + self.answers_right[answer_idx] + "\n\n"
            self.insert_text(self.output_Text, answer_str)

    def run_history(self):
        def exam_history():
            def exam_rank():
                def score_statistics():
                    rank_window.destroy()
                    stat_window = tk.Toplevel()
                    window_position_center(stat_window, 600, 500)
                    stat_window.title('成绩统计')
                    stat_window.resizable(height=False, width=False)
                    if listbox_list[item_index] != '综合':
                        level_tuple = ("A☆", "A+", "A", "A-", "B", "B-", "C", "D")
                        level_dict = {}
                        for level in level_tuple:
                            level_dict[level] = 0
                        for score_record in score_list:
                            level_str = score_record[3]
                            level_dict[level_str] += 1
                        # style = Style(font_family='Yahei')
                        style = Style(font_family='SimHei')
                        course_pie_chart = pygal.Pie(style=style, value_formatter=lambda x: "{}".format(x))
                        course_pie_chart.title = '科目成绩统计'
                        for level_key, level_value in level_dict.items():
                            course_pie_chart.add(level_key, level_value)
                        pie_file = str(self.file_dir / 'score_level.png')
                        course_pie_chart.render_to_png(pie_file)
                        course_level_thread = self.start_image_thread(pie_file, (550, 400),
                                                                      stat_window, self.image_frame)
                        self.image_threads.append(course_level_thread)
                    else:
                        # style = Style(font_family='Yahei')
                        style = Style(font_family='SimHei')
                        course_bar_chart = pygal.Bar(style=style)
                        course_bar_chart.title = '总成绩分数统计'
                        for user, score in totals_dict.items():
                            course_bar_chart.add(user, score)
                        bar_file = str(self.file_dir / 'score_total.png')
                        course_bar_chart.render_to_png(bar_file)
                        course_level_thread = self.start_image_thread(bar_file, (550, 400),
                                                                      stat_window, self.image_frame)
                        self.image_threads.append(course_level_thread)

                top_window.destroy()
                rank_window = tk.Toplevel()
                window_position_center(rank_window, 600, 500)
                rank_window.title('成绩排名')
                rank_window.resizable(height=False, width=False)
                rank_score_label = Label(rank_window, text=title_str, padx=5, pady=5)
                rank_score_label.pack(anchor=W)
                rank_scrollbar_v = Scrollbar(rank_window)
                rank_scrollbar_v.pack(side=RIGHT, fill=Y)
                rank_text = Text(rank_window, width=600, height=500)
                rank_text.place(x=0, y=50)
                rank_text.config(yscrollcommand=rank_scrollbar_v.set)
                rank_text.pack(expand=YES, fill=BOTH)
                rank_text.insert(1.0, score_str)
                rank_scrollbar_v.config(command=rank_text.yview)
                tk.Button(rank_window, text='成绩统计', command=score_statistics).place(x=500, y=0)

            def course_score(course_records, course_dict):
                user_name = course_records['用户']
                user_score = course_records['得分']
                course_dict[user_name] = user_score

            try:
                with open('history_score.pickle', 'rb') as history_file:
                    score_history_info = pickle.load(history_file)
            except FileNotFoundError:
                score_history_info = {}
            title_str = ""
            history_str = ""
            score_str = ""
            score_list = []
            current_selection = lb.curselection()
            if current_selection:
                item_index = current_selection[0]
                if listbox_list[item_index] != '综合':
                    title_tuple = ('用户', '科目', '得分', '成绩', '测验时间')
                    for title_element in title_tuple:
                        title_str += "{:^20}".format(title_element)
                    for timestamp, records in score_history_info.items():
                        if records['科目'] == listbox_list[item_index]:
                            records_str = ""
                            for record in records.values():
                                history_str += "{:^10}".format(record)
                                records_str += "{:^10}".format(record)
                            history_str += '\n'
                            score_list.append(records_str.split())
                    score_list.sort(key=lambda x: float(x[2]), reverse=True)
                    for score_element in score_list:
                        for element in score_element:
                            score_str += "{:^10}".format(element)
                        score_str += '\n'
                else:
                    title_tuple = ('用户', '英语', '数学', '语文', '总成绩')
                    for title_element in title_tuple:
                        title_str += "{:^20}".format(title_element)
                    english_dict = {}
                    math_dict = {}
                    chinese_dict = {}
                    for timestamp, records in score_history_info.items():
                        if records['科目'] == title_tuple[1]:
                            course_score(records, english_dict)
                        elif records['科目'] == title_tuple[2]:
                            course_score(records, math_dict)
                        elif records['科目'] == title_tuple[3]:
                            course_score(records, chinese_dict)
                    totals = {}
                    totals_list = []
                    totals_dict = {}
                    for name in english_dict.keys() | math_dict.keys() | chinese_dict.keys():
                        totals[title_tuple[0]] = name
                        if name in english_dict.keys():
                            totals[title_tuple[1]] = english_dict[name]
                        else:
                            totals[title_tuple[1]] = "-"
                        if name in math_dict.keys():
                            totals[title_tuple[2]] = math_dict[name]
                        else:
                            totals[title_tuple[2]] = "-"
                        if name in chinese_dict.keys():
                            totals[title_tuple[3]] = chinese_dict[name]
                        else:
                            totals[title_tuple[3]] = "-"
                        for i in range(1, 4):
                            totals_list.append(totals[title_tuple[i]])
                        totals_score = [int(x) for x in totals_list if x != "-"]
                        totals[title_tuple[4]] = str(sum(totals_score))
                        totals_dict[name] = sum(totals_score)
                        records_str = ""
                        for record in totals.values():
                            history_str += "{:^10}".format(record)
                            records_str += "{:^10}".format(record)
                        history_str += '\n'
                        score_list.append(records_str.split())
                        totals_list.clear()
                    score_list.sort(key=lambda x: float(x[4]), reverse=True)
                    for score_element in score_list:
                        for element in score_element:
                            score_str += "{:^10}".format(element)
                        score_str += '\n'
                top_window = tk.Toplevel()
                window_position_center(top_window, 600, 500)
                top_window.title('历史统计')
                top_window.resizable(height=False, width=False)
                history_score_label = Label(top_window, text=title_str, padx=5, pady=5)
                history_score_label.pack(anchor=W)
                scrollbar_v = Scrollbar(top_window)
                scrollbar_v.pack(side=RIGHT, fill=Y)
                text = Text(top_window, width=600, height=500)
                text.place(x=0, y=50)
                text.config(yscrollcommand=scrollbar_v.set)
                text.pack(expand=YES, fill=BOTH)
                text.insert(1.0, history_str)
                scrollbar_v.config(command=text.yview)
                self.history_windows.append(top_window)
                tk.Button(top_window, text='成绩排名', command=exam_rank).place(x=500, y=0)
            else:
                tk.messagebox.showwarning('警告', '您未选择科目')

        for window in self.history_windows:
            if window is not None:
                window.destroy()
        self.window.wm_attributes('-topmost', False)
        listbox_window = tk.Toplevel()
        window_position_center(listbox_window, 280, 260)
        listbox_window.title('历史统计')
        listbox_window.resizable(height=False, width=False)
        listbox_list = list(self.course_combobox['values'])
        listbox_frame = Frame(listbox_window)
        listbox_frame.pack(padx=3, pady=1)
        lb = Listbox(listbox_frame, width=30, selectmode=EXTENDED)
        lb.pack(fill=BOTH)
        for key in listbox_list:
            lb.insert(END, key)
        b = Button(listbox_frame, text="选择科目", command=exam_history)
        b.pack(side=RIGHT)
        self.history_windows.append(listbox_window)

    def timer_refresh(self):
        def count_down():
            time_aid = self.timer_start + self.exam_time
            dead_line = int(time_aid - time.time())
            dead_hours = dead_line // (60 * 60) % 24 % 30
            dead_minutes = dead_line // 60 % 60
            dead_seconds = dead_line % 60
            if dead_line <= 0:
                time_info = "倒计时结束" + "\n"
                tk.messagebox.showinfo("答题完毕", time_info)
                self.run_submit()
            time_str = '剩余时间:%02d小时%02d分钟%02d秒' % (dead_hours, dead_minutes, dead_seconds)
            return time_str

        try:
            if self.timer_running is True:
                self.timer_label["text"] = count_down()
                self.window.after(1000, self.timer_refresh)
        except ValueError as e:
            print(e)

    def on_closing(self):
        if tk.messagebox.askokcancel("退出", "要退出登录吗?"):
            self.run_threads = False
            self.window.destroy()

    def json2excel(self):
        file_path, json_data = self.open_file_with_ext(".json")
        if json_data is None:
            return
        questions = []
        questions_name = {}
        questions_dict = {}
        questions_type = {}
        questions_audio = {}
        audios_dict = {}
        answers_right = {}
        choices_list = []
        choices_str = []
        fills_dict = {}
        params = [questions, questions_name, questions_dict, questions_type,
                  questions_audio, audios_dict,
                  answers_right, choices_str, choices_list, fills_dict]
        self.parse_data(json_data, params)
        # df1
        df1 = pd.DataFrame()
        df1['name'] = pd.Series(questions_name)
        df1['name'] = pd.to_numeric(df1['name'], errors='coerce')
        df1['question'] = pd.Series(questions_dict)
        df1['audio'] = pd.Series(audios_dict, dtype=pd.StringDtype())
        df1['type'] = pd.Series(questions_type)
        df1['answer'] = pd.Series(answers_right)
        # df2
        df2 = pd.DataFrame()
        idx = 0
        for list_e in choices_list:
            df2.at[idx, 'name'] = idx + 1
            for _e in list_e:
                if _e[0] == 'A':
                    df2.at[idx, 'A'] = _e[1]
                if _e[0] == 'B':
                    df2.at[idx, 'B'] = _e[1]
                if _e[0] == 'C':
                    df2.at[idx, 'C'] = _e[1]
                if _e[0] == 'D':
                    df2.at[idx, 'D'] = _e[1]
            idx += 1
        # df3
        df3 = pd.DataFrame()
        for idx in range(1, len(questions) + 1):
            df3.at[idx, 'name'] = idx
        for fills_question, fills_list in fills_dict.items():
            idx = int(fills_question)
            for list_e in fills_list:
                df3.at[idx, list_e[0]] = list_e[1]
        # merge
        df1_2 = pd.merge(df1, df2, how='outer', on='name')
        df = pd.merge(df1_2, df3, how='outer', on='name')
        file_prefix, file_ext = os.path.splitext(file_path)
        excel_file = file_prefix + ".xlsx"
        df.to_excel(excel_file, sheet_name='summary', startcol=0, index=False)
        self.insert_text(self.output_Text, "生成XLSX文件:" + excel_file + "\n")

    def excel2json(self):
        file_path, excel_data = self.open_file_with_ext(".xlsx")
        if excel_data is None:
            return
        data_prefix_list = []
        data = []
        data_dict = {}
        for i in excel_data.index.values:
            data_prefix_dict = excel_data.loc[i, ['name', 'question', 'audio', 'type', 'answer']].to_dict()
            data_prefix_dict['name'] = str(data_prefix_dict['name'])
            data_prefix_list.append(data_prefix_dict)
            # choice
            data_middle_dict = excel_data.loc[i, ['A', 'B', 'C', 'D']].to_dict()
            row_choice_dict = {}
            for key, value in data_middle_dict.items():
                if isinstance(value, str):
                    row_choice_dict[key] = value
            choice_dict = {}
            choice_list = []
            choices_dict = {}
            for key, value in row_choice_dict.items():
                choice_dict['item'] = key
                choice_dict['content'] = value
                choice_list.append(choice_dict)
                choice_dict = {}
            choices_dict['choices'] = choice_list
            # fills
            header_data = excel_data.columns
            header_list = []
            for idx in range(len(header_data)):
                header_list.append(header_data[idx])
            data_suffix_dict = excel_data.loc[i, header_list[9:]].to_dict()
            row_fill_dict = {}
            for key, value in data_suffix_dict.items():
                if isinstance(value, str):
                    row_fill_dict[key] = value
            fill_dict = {}
            fill_list = []
            fills_dict = {}
            for key, value in row_fill_dict.items():
                fill_dict['item'] = key
                fill_dict['content'] = value
                fill_list.append(fill_dict)
                fill_dict = {}
            fills_dict['fills'] = fill_list
            # merge
            question_dict = {}
            for question_dict_prefix in data_prefix_list:
                question_dict_middle = question_dict_prefix.copy()
                question_dict_middle.update(choices_dict)
                question_dict = question_dict_middle.copy()
                question_dict.update(fills_dict)
            data.append(question_dict)
        data_dict['Questions'] = data
        json_dict = json.dumps(data_dict, ensure_ascii=False, indent=4)
        file_prefix, file_ext = os.path.splitext(file_path)
        json_file = file_prefix + ".json"
        self.insert_text(self.output_Text, "生成JSON文件:" + json_file + "\n")
        with open(json_file, 'w') as fd:
            fd.write(json_dict)

    def run_first(self):
        self.destroy_button()
        self.index = 0
        self.delete_text(self.output_Text)
        self.show_data()
        tk.messagebox.showinfo('提示', '回到第一题')

    def run_last(self):
        self.destroy_button()
        self.index = len(self.questions) - 1
        self.delete_text(self.output_Text)
        self.show_data()
        tk.messagebox.showinfo('提示', '回到最后一题')

    def run_previous(self):
        self.index -= 1
        self.delete_text(self.output_Text)
        if self.index < 0:
            _str_ = '到了第一道题'
            tk.messagebox.showwarning('警告', _str_)
            return
        self.destroy_button()
        self.show_data()

    def run_next(self):
        self.index += 1
        self.delete_text(self.output_Text)
        if self.index >= len(self.questions):
            _str_ = '已经是最后一道题'
            tk.messagebox.showwarning('警告', _str_)
            return
        self.destroy_button()
        self.show_data()

    def destroy_button(self):
        for r in self.radio_button:
            r.destroy()
        for c in self.check_button:
            c.destroy()
        for label in self.fill_label:
            label.destroy()
        for entry in self.fill_entry:
            entry.destroy()
        for i in self.image_frame:
            if i is not None:
                i.destroy()

    def show_data(self):
        def event_radio():
            _str = '你选择的答案是:{}'.format(v.get())
            print(_str)
            index_num = self.index + 1
            index_str = str(index_num)
            self.answers_yours[index_str] = v.get()
            self.label_frame.config(text=_str)
            self.delete_text(self.output_Text)
            self.insert_text(self.output_Text, _str)

        def event_check():
            str_list = []
            for cv in self.check_value:
                str_list.append(cv.get())
            _str = ''.join(str(sl) for sl in str_list)
            print(str_list)
            answer_str = '你选择的答案是:{}'.format(_str)
            print(answer_str)
            index_num = self.index + 1
            index_str = str(index_num)
            self.answers_yours[index_str] = _str
            self.label_frame.config(text=answer_str)
            self.delete_text(self.output_Text)
            self.insert_text(self.output_Text, answer_str)
            str_list.clear()

        def event_fill(event):
            str_list = []
            for fe in self.fill_value:
                str_list.append(fe.get())
            _str = ' '.join(str(sl) for sl in str_list)
            answer_str = '你填入的答案是:{}'.format(_str)
            index_num = self.index + 1
            index_str = str(index_num)
            self.answers_yours[index_str] = _str
            self.delete_text(self.output_Text)
            current_str = "当前填入的答案是" + event.widget.get() + "\n"
            self.insert_text(self.output_Text, current_str)
            self.insert_text(self.output_Text, answer_str)
            str_list.clear()

        self.check_value = []
        self.fill_value = []
        questions, questions_name, questions_dict, questions_type, questions_audio, audios_dict, \
            answers_right, choices_str, choices_list, fills_dict = self.params
        if self.index < 0 or self.index >= len(questions):
            return
        self.delete_text(self.input_Text)
        question_type = questions_type[str(self.index + 1)]
        question_suffix = ""
        if question_type == "radio":
            question_suffix += ".单选题"
            self.label_frame.config(text="你选择的答案是？")
            if not self.look_wrong:
                choices = choices_list[self.index]
                v = tk.StringVar()
                for choice in choices:
                    choice_string = '. '.join(str(i) for i in choice)
                    r = Radiobutton(self.label_frame, text=choice_string,
                                    variable=v, value=choice[0], command=event_radio)
                    r.pack(anchor=W)
                    self.radio_button.append(r)
        elif question_type == "check":
            question_suffix += ".多选题"
            self.label_frame.config(text="你选择的答案是？")
            if not self.look_wrong:
                choices = choices_list[self.index]
                for choice in choices:
                    v = tk.StringVar()
                    choice_string = '. '.join(str(i) for i in choice)
                    c = Checkbutton(self.label_frame, text=choice_string,
                                    variable=v, onvalue=choice[0], offvalue="", command=event_check)
                    c.pack(anchor=W)
                    self.check_value.append(v)
                    self.check_button.append(c)
        elif question_type == "fill":
            question_suffix += ".填空题"
            self.label_frame.config(text="请填空作答")
            if not self.look_wrong:
                choices = fills_dict[str(self.index + 1)]
                for choice in choices:
                    v = tk.StringVar()
                    choice_string = '. '.join(str(i) for i in choice)
                    label = ttk.Label(self.label_frame, text=choice_string)
                    label.pack(anchor=W)
                    entry = ttk.Entry(self.label_frame, textvariable=v)
                    entry.pack(anchor=W)
                    self.fill_label.append(label)
                    self.fill_entry.append(entry)
                    self.fill_value.append(v)
                    entry.bind('<Key-Return>', event_fill)
        # show question
        question_item, question_content = questions[self.index].split(":")
        question_str = question_item + question_suffix + question_content + "\n"
        self.insert_text(self.input_Text, question_str)
        for choice_str in choices_str[self.index]:
            self.insert_text(self.input_Text, choice_str)
        if self.look_wrong is True:
            tips_str = "正确答案: " + self.answers_right[question_item] + "\n"
            if question_item in self.answers_yours.keys():
                tips_str += "你的答案: " + self.answers_yours[question_item] + "\n"
            else:
                tips_str += "你的答案: " + "未作答" + "\n"
            self.insert_text(self.output_Text, tips_str)
        self.show_current_exam_info()
        # play audio
        element = str(self.index + 1)
        if element in audios_dict.keys():
            if audios_dict[element] == "yes":
                audio_str = questions_audio[element]
                file_path, _ = os.path.splitext(self.json_file)
                file_prefix = file_path.replace(os.path.dirname(file_path) + '/', "")
                file_name = file_prefix + '_' + element + '.mp3'
                audio_file = str(self.file_dir / file_name)
                info_str = "音频文件:" + audio_file + "\n"
                self.insert_text(self.output_Text, info_str)
                self.gen_audio(audio_str, audio_file)

    def set_time(self):
        if self.start_exam is True:
            if self.finish_exam is False:
                info_str = "正在答题中，请“提交答案”后再试"
                tk.messagebox.showinfo('提示', info_str)
                return
        self.window.wm_attributes('-topmost', False)
        time_str = askstring('考试时长', '设置考试时长（分钟）？')
        if time_str is None:
            time_str_set = '考试时长未设置, 使用默认时长: ' + str(EXAM_TIME // 60) + '分钟\n'
        else:
            time_str_set = '考试时长已设置: {}'.format(time_str + '分钟\n')
            if int(time_str) > 0:
                self.exam_time = int(time_str) * 60
                self.timer_label.config(text="考试总时长为: " + str(self.exam_time // 60) + "分钟")
        tk.messagebox.showinfo('考试时长', time_str_set)
        self.delete_text(self.output_Text)
        self.insert_text(self.output_Text, time_str_set)

    def run_decrypt(self):
        decrypt_dict = {}
        try:
            with open(file=AUTHORIZATION_FILE, mode='r+', encoding='utf-8') as key_file:
                raw_text_lines = key_file.readlines()
            for raw_text in raw_text_lines:
                key_prefix, key_content = raw_text.split(":")
                key = key_content.strip()
                if key_prefix == 'SMTP_TOKEN':
                    decrypt_dict['SMTP_TOKEN'] = base64.b64decode(key).decode('utf-8')
                elif key_prefix == 'APP_ID':
                    decrypt_dict['APP_ID'] = base64.b64decode(key).decode('utf-8')
                elif key_prefix == 'API_KEY':
                    decrypt_dict['API_KEY'] = base64.b64decode(key).decode('utf-8')
                elif key_prefix == 'SECRET_KEY':
                    decrypt_dict['SECRET_KEY'] = base64.b64decode(key).decode('utf-8')
            self.delete_text(self.output_Text)
            output_str = "SMTP_TOKEN = " + decrypt_dict['SMTP_TOKEN'] + "\n"
            output_str += "APP_ID = " + decrypt_dict['APP_ID'] + "\n"
            output_str += "API_KEY = " + decrypt_dict['API_KEY'] + "\n"
            output_str += "SECRET_KEY = " + decrypt_dict['SECRET_KEY'] + "\n"
            self.insert_text(self.output_Text, output_str)
            return decrypt_dict
        except FileNotFoundError:
            tk.messagebox.showerror(title='错误', message='找不到密钥文件: ' + AUTHORIZATION_FILE)
            return

    def run_encrypt(self):
        def do_encrypt():
            smtp_token_str = smtp_token.get().strip()
            app_id_str = app_id.get().strip()
            api_key_str = api_key.get().strip()
            secret_key_str = secret_key.get().strip()
            if smtp_token_str == '' or app_id_str == '' or api_key_str == '' or secret_key_str == '':
                tk.messagebox.showerror(title='错误', message='无法生成密钥')
                return
            encrypt_smtp_token = base64.b64encode(smtp_token_str.encode('utf-8'))
            key = "SMTP_TOKEN: " + encrypt_smtp_token.decode('utf-8') + '\n'
            encrypt_app_id = base64.b64encode(app_id_str.encode('utf-8'))
            key += "APP_ID: " + encrypt_app_id.decode('utf-8') + '\n'
            encrypt_api_key = base64.b64encode(api_key_str.encode('utf-8'))
            key += "API_KEY: " + encrypt_api_key.decode('utf-8') + '\n'
            encrypt_secret_key = base64.b64encode(secret_key_str.encode('utf-8'))
            key += "SECRET_KEY: " + encrypt_secret_key.decode('utf-8') + '\n'
            with open(file=AUTHORIZATION_FILE, mode='w+', encoding='utf-8') as key_file:
                key_file.write(key)
            window_encrypt.destroy()
            if os.path.exists(AUTHORIZATION_FILE) is True:
                self.delete_text(self.output_Text)
                output_str = "密钥已写入文件: " + AUTHORIZATION_FILE + "\n"
                self.insert_text(self.output_Text, output_str)

        self.window.wm_attributes('-topmost', False)
        window_encrypt = tk.Toplevel(self.window)
        window_position_center(window_encrypt, 350, 220)
        window_encrypt.resizable(width=False, height=False)
        window_encrypt.title('生成密钥')
        smtp_token = tk.StringVar()
        tk.Label(window_encrypt, text='SMTP_TOKEN：').place(x=10, y=10)
        tk.Entry(window_encrypt, textvariable=smtp_token).place(x=150, y=10)
        app_id = tk.StringVar()
        tk.Label(window_encrypt, text='APP_ID：').place(x=10, y=50)
        tk.Entry(window_encrypt, textvariable=app_id).place(x=150, y=50)
        api_key = tk.StringVar()
        tk.Label(window_encrypt, text='API_KEY：').place(x=10, y=90)
        tk.Entry(window_encrypt, textvariable=api_key).place(x=150, y=90)
        secret_key = tk.StringVar()
        tk.Label(window_encrypt, text='SECRET_KEY：').place(x=10, y=130)
        tk.Entry(window_encrypt, textvariable=secret_key, show='*').place(x=150, y=130)
        gen_encrypt = tk.Button(window_encrypt, text='生成密钥', command=do_encrypt)
        gen_encrypt.place(x=150, y=180)

    def run_load(self):
        if self.start_exam is True:
            tk.messagebox.showinfo('提示', '已经开始测试，请作答')
            return
        self.delete_text(self.output_Text)
        load_info = "请导入题库文件，该文件必须为JSON格式" + "\n"
        self.insert_text(self.output_Text, load_info)
        self.json_file, self.json_data = self.open_file_with_ext(".json")
        if self.json_data is None:
            load_error = "导入题库出错" + "\n"
            self.insert_text(self.output_Text, load_error)
            self.load_exam = False
            return
        else:
            question_bank_info = "题库文件: " + self.json_file + "\n"
            self.insert_text(self.output_Text, "导入" + question_bank_info)
            self.insert_text(self.output_Text, "当前" + question_bank_info)
            load_ok = "导入题库成功" + "\n"
            self.insert_text(self.output_Text, load_ok)
            self.load_exam = True
            return

    def run_exam(self):
        if os.path.exists(self.json_file) is False:
            tk.messagebox.showwarning('警告', '题库文件{}不存在，请点击“导入题库”，然后“开始测验”'.format(self.json_file))
            return
        else:
            self.json_data = self.read_from_json(self.json_file)

        for image_frame in self.image_frame:
            if image_frame is not None:
                image_frame.destroy()
        if self.start_exam is True:
            tk.messagebox.showinfo('提示', '已经开始测试，请作答')
            return
        # init
        self.look_wrong = False
        self.destroy_button()
        if self.score_frame:
            self.score_frame.destroy()
        frame_str = "你选择的答案是？"
        self.label_frame.config(text=frame_str)
        self.delete_text(self.input_Text)
        self.delete_text(self.output_Text)

        # start timer count down
        self.timer_running = True
        self.timer_start = time.time()
        self.timer_refresh()

        # re-init
        self.index = 0
        self.questions = []
        self.questions_name = {}
        self.questions_dict = {}
        self.questions_type = {}
        self.answers_right = {}
        self.answers_yours = {}
        self.answers_result = {}
        self.choices_list = []
        self.choices_str = []
        self.check_value = []
        self.check_button = []
        self.fill_entry = []
        self.fill_label = []
        self.params = [self.questions, self.questions_name, self.questions_dict, self.questions_type,
                       self.questions_audio, self.audios_dict,
                       self.answers_right, self.choices_str, self.choices_list, self.fills_dict]
        self.parse_data(self.json_data, self.params)
        print("正确答案是: ", self.answers_right)
        print("题目类型: ", self.questions_type)
        for key in self.answers_right.keys():
            self.answers_yours[key] = '未作答'
        self.show_data()
        self.start_exam = True
        self.finish_exam = False
        question_bank_info = "当前题库文件: " + os.path.abspath(self.json_file) + "\n"
        self.insert_text(self.output_Text, question_bank_info)

    def show_current_exam_info(self):
        exam_str_input = str(self.index + 1) + "/" + str(len(self.questions))
        exam_str_output = "({})".format(exam_str_input)
        self.input_exam_label.config(text=exam_str_output)
        self.output_exam_label.config(text=exam_str_output)

    def run_submit(self):
        def course_from_json(value):
            courses = {
                "english.json": "英语",
                "math.json": "数学",
                "chinese.json": "语文",
            }
            return courses.get(value, '英语')

        if self.start_exam is False:
            info_str = ""
            if self.finish_exam is True:
                info_str += "测试结束，作答已提交\n"
            info_str += "请点击“开始测验”进行作答"
            print(info_str)
            tk.messagebox.showinfo('提示', info_str)
            return
        # init
        self.destroy_button()
        self.index = 0
        self.timer_running = False
        self.start_exam = False
        self.finish_exam = True
        self.look_wrong = False
        # self.questions = []
        frame_str = "作答已提交"
        self.label_frame.config(text=frame_str)
        self.delete_text(self.input_Text)
        self.delete_text(self.output_Text)
        record_str = "试卷作答完毕" + "\n\n"
        record_str += "正确答案是: " + "\n" + json.dumps(self.answers_right) + "\n\n"
        print("你的答案是: ", self.answers_yours)
        record_str += "你的答案是: " + "\n" + json.dumps(self.answers_yours, ensure_ascii=False) + "\n\n"
        score_of_each_question = round(float(100 / len(self.answers_right)), 2)
        score_of_each_question_str = "每题分值: " + str(score_of_each_question) + "分" + "\n"
        print(score_of_each_question_str)
        score_sum = 0
        for index in self.answers_right.keys():
            if self.questions_type[index] != "fill":
                if index in self.answers_yours.keys():
                    if self.answers_yours[index] == self.answers_right[index]:
                        self.answers_result[index] = '对'
                        score_sum += score_of_each_question
                    else:
                        self.answers_result[index] = '错'
            else:
                right_num = 0
                if index in self.answers_yours.keys():
                    answers_yours = self.answers_yours[index].split(' ')
                    answers_right = self.answers_right[index].split(' ')
                    for idx in range(len(answers_yours)):
                        if answers_yours[idx] == answers_right[idx]:
                            right_num += 1
                    if right_num == len(answers_right):
                        self.answers_result[index] = '对'
                    elif right_num == 0:
                        self.answers_result[index] = '错'
                    else:
                        self.answers_result[index] = '半对'
                    score_of_fill_question = float(right_num / len(answers_right)) * score_of_each_question
                    score_of_fill_question_str = "填空题得分: " + str(round(score_of_fill_question, 2)) + "分" + "\n"
                    print(score_of_fill_question_str)
                    score_sum += round(score_of_fill_question, 1)
        print("成绩统计结果: ", self.answers_result)
        result_str = "成绩统计结果: " + "\n" + json.dumps(self.answers_result, ensure_ascii=False) + "\n"
        if score_sum > 100:
            score = 100
        else:
            score = round(score_sum)
        result_str += "\n" + "你的成绩: " + str(score) + "分" + "\n"
        self.insert_text(self.input_Text, record_str)
        self.insert_text(self.output_Text, result_str)
        score_str = "用户: " + LoginWindow.get_user() + "\n"
        score_str += "得分: " + str(score) + "分" + "\n"
        self.score_frame = LabelFrame(self.window, text="你的成绩", padx=5, pady=5)
        self.score_frame.place(x=620, y=390)
        self.score_frame.config(text=score_str)
        image_filename, image_level, score_level = self.check_score_level(score)
        t1 = self.start_image_thread(self.file_dir / image_filename, (100, 100), self.score_frame, self.image_frame)
        self.image_threads.append(t1)
        t2 = self.start_image_thread(self.file_dir / image_level, (100, 100), self.label_frame, self.image_frame)
        self.image_threads.append(t2)
        score_str += "成绩: " + score_level
        tk.messagebox.showinfo('测试成绩', score_str)
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        score_user = {'用户': LoginWindow.get_user(),
                      '科目': course_from_json(self.json_file),
                      '得分': str(round(score, 1)),
                      '成绩': score_level,
                      '时间': timestamp}
        try:
            with open('history_score.pickle', 'rb') as history_file:
                score_history_info = pickle.load(history_file)
        except FileNotFoundError:
            score_history_info = {}
        if timestamp in score_history_info:
            tk.messagebox.showerror('错误', '记录已存在')
        else:
            score_history_info[timestamp] = score_user
            with open('history_score.pickle', 'wb') as history_file:
                pickle.dump(score_history_info, history_file)

    def run_wrong(self):
        self.wrong_questions = []
        self.wrong_choices_str = []
        self.wrong_choices_list = []
        self.wrong_fills_dict = {}
        if self.look_wrong is True:
            info_str = "当前已进入“只看错题”状态\n请“开始测验”-“提交答案”后再试"
            tk.messagebox.showinfo('提示', info_str)
            return
        if self.finish_exam is False:
            if self.start_exam is False:
                info_str = "当前状态-未开始测验\n请“开始测验”-“提交答案”后再试"
                tk.messagebox.showinfo('提示', info_str)
                return
            info_str = "正在答题中，请“提交答案”后再试"
            tk.messagebox.showinfo('提示', info_str)
            return
        self.look_wrong = True
        self.delete_text(self.input_Text)
        self.delete_text(self.output_Text)
        for key, value in self.answers_result.items():
            if value == '错' or value == '半对':
                idx = int(key) - 1
                self.wrong_questions.append(self.questions[idx])
                self.wrong_choices_str.append(self.choices_str[idx])
                self.wrong_choices_list.append(self.choices_list[idx])
        self.params = [self.wrong_questions, self.questions_name, self.questions_dict, self.questions_type,
                       self.questions_audio, self.audios_dict,
                       self.answers_right, self.wrong_choices_str, self.wrong_choices_list, self.wrong_fills_dict]
        self.questions = self.wrong_questions
        self.choices_str = self.wrong_choices_str
        self.choices_list = self.wrong_choices_list
        self.fills_dict = self.wrong_fills_dict
        self.show_data()

    def pack_image(self, image_file, image_size, parent_frame, image_frame_list):
        pil_image = Image.open(image_file)
        pil_image_small = pil_image.resize(image_size, Image.ANTIALIAS)
        img = ImageTk.PhotoImage(pil_image_small)
        child_frame = Label(parent_frame, image=img)
        child_frame.pack()
        image_frame_list.append(child_frame)
        while self.run_threads is True:
            time.sleep(1)

    def start_image_thread(self, image_file, image_size, parent_frame, image_frame_list):
        t = threading.Thread(target=self.pack_image, args=(image_file, image_size, parent_frame, image_frame_list))
        t.start()
        return t

    def upload_file(self):
        def do_upload():
            top_window.destroy()
            self.delete_text(self.output_Text)
            select_file_name = filedialog.askopenfilename(title='上传文件')
            if isinstance(select_file_name, tuple) or select_file_name == "":
                return None
            url_str = "http://" + ip_str.get() + ":" + port_str.get() + UPLOAD_SUFFIX
            url = url_str.strip('\n')
            requests.post(url, files={'file': open(select_file_name, 'rb')})
            file_info = '上传文件: ' + select_file_name + '\n' + '至服务器URL==> ' + url + '\n'
            self.insert_text(self.output_Text, file_info)

        self.window.wm_attributes('-topmost', False)
        top_window = tk.Toplevel()
        window_position_center(top_window, 240, 130)
        top_window.title('服务器地址')
        top_window.resizable(height=False, width=False)
        ip_str = tk.StringVar()
        tk.Label(top_window, text='IP：').place(x=10, y=10)
        tk.Entry(top_window, textvariable=ip_str).place(x=60, y=10)
        port_str = tk.StringVar()
        tk.Label(top_window, text='Port：').place(x=10, y=50)
        tk.Entry(top_window, textvariable=port_str).place(x=60, y=50)
        tk.Button(top_window, text='确认', command=do_upload).place(x=60, y=80)

    def download_file(self):
        def download_from_url(url, dst):
            """
            @param: url to download file
            @param: dst place to put the file
            :return: bool
            """
            # 获取文件长度
            try:
                file_size = int(urlopen(url).info().get('Content-Length', -1))
            except Exception as e:
                print(e)
                print("错误，访问url: %s 异常" % url)
                return False
            # 判断本地文件存在时
            if os.path.exists(dst):
                # 获取文件大小
                first_byte = os.path.getsize(dst)
            else:
                # 初始大小为0
                first_byte = 0
            # 判断大小一致，表示本地文件存在
            if first_byte >= file_size:
                print("文件已经存在,无需下载")
                return file_size
            header = {"Range": "bytes=%s-%s" % (first_byte, file_size)}
            p_bar = tqdm(
                total=file_size, initial=first_byte,
                unit='B', unit_scale=True, desc=url.split('/')[-1])
            # 访问url进行下载
            req = requests.get(url, headers=header, stream=True)
            try:
                with(open(dst, 'ab')) as f:
                    for chunk in req.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                            p_bar.update(1024)
            except Exception as e:
                print(e)
                return False
            p_bar.close()
            return True

        def do_download():
            top_window.destroy()
            url_str = "http://" + ip_str.get() + ":" + port_str.get() + "/" + file_str.get()
            url = url_str.strip('\n')
            file_info = '从服务器URL: ' + url + ' 开始下载' + '\n'
            self.insert_text(self.output_Text, file_info)
            file_path = filedialog.asksaveasfilename(title=u'下载文件')
            if isinstance(file_path, tuple) or file_path == "":
                err_info = "未选择下载文件，已取消下载"
                self.insert_text(self.output_Text, err_info)
                return None
            if bar_check.get() == 1:
                download_from_url(url, file_path)
                file_info = '下载至本地文件: ' + file_path + ' 已完成' + '\n'
            else:
                files = requests.get(url)
                files.raise_for_status()
                with open(file_path, 'wb') as f:
                    f.write(files.content)
                file_info = '下载至本地文件: ' + f.name + ' 已完成' + '\n'
            self.insert_text(self.output_Text, file_info)

        self.window.wm_attributes('-topmost', False)
        top_window = tk.Toplevel()
        window_position_center(top_window, 240, 200)
        top_window.title('服务器地址')
        top_window.resizable(height=False, width=False)
        ip_str = tk.StringVar()
        tk.Label(top_window, text='IP地址：').place(x=10, y=10)
        tk.Entry(top_window, textvariable=ip_str).place(x=60, y=10)
        port_str = tk.StringVar()
        tk.Label(top_window, text='端口：').place(x=10, y=50)
        tk.Entry(top_window, textvariable=port_str).place(x=60, y=50)
        file_str = tk.StringVar()
        tk.Label(top_window, text='文件名：').place(x=10, y=90)
        tk.Entry(top_window, textvariable=file_str).place(x=60, y=90)
        bar_check = tk.IntVar()
        bar_str = "带进度条？"
        cb = Checkbutton(top_window, text=bar_str, variable=bar_check, onvalue=1, offvalue=0)
        cb.place(x=60, y=120)
        tk.Button(top_window, text='确认', command=do_download).place(x=60, y=150)

    def open_file_with_ext(self, ext):
        file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('./')))
        # 未选择文件返回元组
        if isinstance(file_path, tuple) or file_path == "":
            return "", None
        self.delete_text(self.output_Text)
        if file_path is not None:
            file_prefix, file_ext = os.path.splitext(file_path)
            if file_ext == ext:
                self.insert_text(self.output_Text,
                                 "{ext_name}文件为: ".format(ext_name=ext.lstrip('.').upper()) + file_path + "\n")
                if ext == ".json":
                    data = self.read_from_json(file_path)
                    return file_path, data
                elif ext == ".xlsx":
                    data = self.read_from_excel(file_path)
                    return file_path, data
            else:
                error_str = "该文件不是{ext_name}格式".format(ext_name=ext.lstrip('.').upper())
                self.insert_text(self.output_Text, "文件: " + file_path + "\n" + error_str)
                tk.messagebox.showerror('文件格式出错', error_str)
                return None, None
        else:
            self.insert_text(self.output_Text, "没找到文件\n")
            return None, None

    def open_file(self):
        file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('./')))
        if isinstance(file_path, tuple) or file_path == "":
            return None
        self.input_Text.get(1.0, END)
        if file_path is not None:
            with open(file=file_path, mode='r+', encoding='utf-8') as file:
                file_text = file.read()
            self.insert_text(self.input_Text, file_text)

    def save_file(self):
        file_path = filedialog.asksaveasfilename(title=u'保存文件')
        if isinstance(file_path, tuple) or file_path == "":
            return None
        file_text = self.input_Text.get(1.0, END)
        if file_path is not None:
            with open(file=file_path, mode='a+', encoding='utf-8') as file:
                file.write(file_text)
            self.delete_text(self.input_Text)
            dialog.Dialog(None, {'title': '文件修改', 'text': '保存完成',
                                 'bitmap': 'warning', 'default': 0,
                                 'strings': ('确定', '取消')})

    def clear_text(self):
        self.delete_text(self.output_Text)
        self.delete_text(self.input_Text)
        self.destroy_button()
        if self.label_frame is not None:
            self.label_frame.config(text='')
        if self.score_frame is not None:
            self.score_frame.config(text='')
        for image_frame in self.image_frame:
            image_frame.destroy()

    def gen_audio(self, text_data, audio_file):
        def run_decrypt():
            decrypt_dict = {}
            try:
                with open(file=AUTHORIZATION_FILE, mode='r+', encoding='utf-8') as key_file:
                    raw_text_lines = key_file.readlines()
                for raw_text in raw_text_lines:
                    key_prefix, key_content = raw_text.split(":")
                    key = key_content.strip()
                    if key_prefix == 'SMTP_TOKEN':
                        decrypt_dict['SMTP_TOKEN'] = base64.b64decode(key).decode('utf-8')
                    elif key_prefix == 'APP_ID':
                        decrypt_dict['APP_ID'] = base64.b64decode(key).decode('utf-8')
                    elif key_prefix == 'API_KEY':
                        decrypt_dict['API_KEY'] = base64.b64decode(key).decode('utf-8')
                    elif key_prefix == 'SECRET_KEY':
                        decrypt_dict['SECRET_KEY'] = base64.b64decode(key).decode('utf-8')
                return decrypt_dict
            except FileNotFoundError:
                tk.messagebox.showerror(title='错误', message='找不到密钥文件: ' + AUTHORIZATION_FILE)
                return

        self.delete_text(self.output_Text)
        if self.audio_mode == "local":
            engine = pyttsx3.init()
            engine.say(text_data)
            engine.runAndWait()
        elif self.audio_mode == "remote":
            # 检查本地音频文件是否存在
            # 若存在，优先用本地文件
            # 若不存在，访问百度云
            if os.path.exists(audio_file) is False:
                info_str = "音频文件:" + audio_file + "不存在，访问百度云"
                self.insert_text(self.output_Text, info_str)
                app_dict = run_decrypt()
                # 对应填入百度控制台获取的三个参数
                client = AipSpeech(app_dict['APP_ID'], app_dict['API_KEY'], app_dict['SECRET_KEY'])
                result = client.synthesis(text_data, 'zh', 1, {
                    'per': 4,
                    'spd': 3,  # 速度
                    'vol': 7  # 音量
                })
                if not isinstance(result, dict):
                    with open(audio_file, 'wb') as f:
                        f.write(result)
            # 播放
            info_str = "播放音频文件:" + audio_file
            self.insert_text(self.output_Text, info_str)
            playsound(audio_file)

    def play_audio(self):
        element = str(self.index + 1)
        if element in self.audios_dict.keys():
            if self.audios_dict[element] == "yes":
                audio_str = self.questions_audio[element]
                file_path, _ = os.path.splitext(self.json_file)
                file_prefix = file_path.replace(os.path.dirname(file_path) + '/', "")
                file_name = file_prefix + '_' + element + '.mp3'
                audio_file = str(self.file_dir / file_name)
                self.gen_audio(audio_str, audio_file)


def window_position_center(window, length, width):
    sw = window.winfo_screenwidth()
    sh = window.winfo_screenheight()
    ww = length
    wh = width
    x = (sw - ww) / 2
    y = (sh - wh) / 2
    window.geometry("%dx%d+%d+%d" % (ww, wh, x, y))


def main_window():
    window = tk.Tk()
    window.wm_attributes('-topmost', True)
    banner_info = ' 欢迎' + LoginWindow.get_user() + '登录' + LoginWindow.get_mode() + '  '
    banner_info += '请点击下拉菜单选择科目，点击“开始测验”进行作答，点击“帮助”获取更多信息'
    tk.Label(window, text=banner_info).grid(row=0, column=1)
    # 初始化类为对象
    WindowShell(window, (0, 1))
    window.mainloop()


# 执行
if __name__ == '__main__':
    top = tk.Tk()
    s = LoginWindow()
    frame = s.login_window(top)
    try:
        top.wait_window(window=frame)
        if s.login:
            main_window()
    finally:
        top.mainloop()
