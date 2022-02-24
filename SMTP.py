import os
import ssl
from tkinter import *
from tkinter import ttk

from socket import *
import base64

import pickle

# TLS加密
context = ssl.create_default_context()

# 初始化窗口
root = Tk()
root.title("Send an Email via B-qqmail !!")

# 尝试读取联系人
try:
    with open('contacts.pkl', 'rb') as f:
        contacts = pickle.load(f)
except:
    contacts = ['example1@email']

# 尝试读取已发送邮件
try:
    with open('history.pkl', 'rb') as f:
        history = pickle.load(f)
except:
    history = []

account = StringVar()
password = StringVar()
receiver = StringVar()
subject = StringVar()
msgbody_content = ""

# 尝试读取之前存储的草稿
try:
    with open('draft.pkl', 'rb') as f:
        info = pickle.load(f)
        account.set(info[0])
        password.set(info[1])
        receiver.set(info[2])
        subject.set(info[3])
        msgbody_content = info[4]
except:
    pass
 
# 创建窗口
mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

# 增加文本框用于输入发件人邮箱
ttk.Label(mainframe, text="Your Email Account: ").grid(column=0, row=1, sticky=W)
account_entry = ttk.Entry(mainframe, width=30, textvariable=account)
account_entry.grid(column=4, row=1, sticky=(W, E))

# 增加文本框用于输入密码
ttk.Label(mainframe, text="Your Password: ").grid(column=0, row=2, sticky=W)
password_entry = ttk.Entry(mainframe, show="*", width=30, textvariable=password)
password_entry.grid(column=4, row=2, sticky=(W, E))


# 选择联系人，或输入新联系人
def contact():
    # 打开新窗口
    tk_contack = Toplevel()

    # 每个联系人生成一个复选框
    CheckButton = {}
    for k in contacts:
        CheckButton[k] = IntVar()
        Checkbutton(tk_contack, text = k, variable = CheckButton[k]).pack(anchor=W)

    # 新联系人通过一行文本输入，用分号 ';' 隔开，空格会被自动过滤
    Label(tk_contack, text="New Contact: ").pack(anchor=W)
    new_contactVar = StringVar()
    Entry(tk_contack, textvariable=new_contactVar).pack(anchor=W)

    # 读取外部已经设置的收件人，将其解析，同步到复选框和文本框
    receiver_list = [x.strip() for x in receiver.get().strip().split(';')]
    receiver_remain = []
    for k in receiver_list:
        if k in CheckButton:
            CheckButton[k].set(1)
        else:
            receiver_remain.append(k)
    new_contactVar.set('; '.join(receiver_remain))
    
    # 选择好联系人后，将其关闭，并同步到外部收件人文本框
    def contact_close():
        receiver_list = []
        for k, v in CheckButton.items():
            if v.get() == 1:
                receiver_list.append(k)

        for _ in new_contactVar.get().strip().split(";"):
            if _.strip() != "":
                receiver_list.append(_.strip())

        receiver.set(';\n'.join(receiver_list))

        tk_contack.destroy()

    Button(tk_contack, text="Accept", command=contact_close).pack(anchor=E)

    tk_contack.mainloop()

# 增加按钮用于选择收件人，旧联系人通过复选框选择，新联系人通过文本框输入
ttk.Label(mainframe, text="Recepient's Email Account: ").grid(column=0, row=3, sticky=W)
receiver_entry = ttk.Button(mainframe, width=30, textvariable=receiver, command=contact)
receiver_entry.grid(column=4, row=3, sticky=(W, E))

# 增加文本框用于输入标题
ttk.Label(mainframe, text="Subject: ").grid(column=0, row=6, sticky=W)
subject_entry = ttk.Entry(mainframe, width=30, textvariable=subject)
subject_entry.grid(column=4, row=6, sticky=(W, E))

# 增加多行文本框用于输入正文
ttk.Label(mainframe, text="Message Body: ").grid(column=0, row=7, sticky=W)
msgbody = Text(mainframe, width=30, height=10)
msgbody.insert("end", msgbody_content)
msgbody.grid(column=4, row=7, sticky=(W, E))


# 展示历史已发送文件，可选择其中之一并读取，或直接关闭窗口退出
def show_history():
    tk_history = Toplevel()
    max_name_length = 0
    for info in history:
        max_name_length = max(max_name_length, len(info[0]))
    max_name_length = min(max_name_length, 40)

    for info in history:
        # 对齐发件人的邮箱长度
        _account = info[0]
        if len(_account) > max_name_length:
            _account = _account[:max_name_length - 3] + '...'
        if len(_account) < max_name_length:
            _account = _account + ' ' * (max_name_length - len(_account))

        _subject = info[3]
        if len(_subject) > max_name_length:
            _subject = _subject[:max_name_length - 3] + '...'

        # 使用闭包函数生成对应的 command 处理函数
        def select_history():
            _info = info
            def f():
                account.set(_info[0])
                password.set(_info[1])
                receiver.set(_info[2])
                subject.set(_info[3])
                msgbody.delete('0.0', END)
                msgbody.insert('end', _info[4])

                tk_history.destroy()
            return f

        # 显示 发件人 和 邮件标题，根据这些信息选择历史邮件，读取其中内容，并退出
        Button(tk_history, width=30, height=1, text = "{}: {}".format(_account, _subject), command=select_history(), anchor=W, justify=LEFT).pack(anchor=W)

    tk_history.mainloop()

# 增加按钮用于选择历史已发送邮件
ttk.Button(mainframe, text="History", command=show_history).grid(column=0,row=8,sticky=W)


# 保存草稿，下次打开时如存在未发送的草稿，则自动恢复
def save_draft():
    info = [account.get(), password.get(), receiver.get(), subject.get(), msgbody.get('1.0','end')]
    with open('draft.pkl', 'wb') as f:
        pickle.dump(info, f)

    ttk.Label(mainframe, text="Draft saved successfully").grid(column=4,row=9,sticky=W)

# 增加按钮用于保存草稿
ttk.Button(mainframe, text="Save Draft", command=save_draft).grid(column=2,row=8,sticky=E)


# 使用 SMTP 发送邮件，仅支持 qq 邮箱
def sendemail():
    try:
        fromAddress = account.get()

        # 输入
        toAddress = [x.strip() for x in receiver.get().split(';')]

        subject_ = subject.get()
        msg = "\r\n" + msgbody.get('1.0','end')
        endMsg = "\r\n.\r\n"

        mailServer = "smtp.qq.com"

        # 发送方，验证信息
        username = str(base64.b64encode(fromAddress.encode()),'utf-8')  # 输入自己的用户名对应的编码
        password_ = str(base64.b64encode(password.get().encode()),'utf-8')  # 此处不是自己的密码，而是开启SMTP服务时对应的授权码

        # 创建客户端套接字并建立连接
        serverPort = 465  # SMTP SSL使用465号端口
        with create_connection((mailServer, serverPort)) as sock:
            with context.wrap_socket(sock, server_hostname = mailServer) as clientSocket:

                # 从客户套接字中接收信息
                recv = clientSocket.recv(1024).decode()
                if '220' != recv[:3]:
                    raise Exception('220 reply not received from server.\nInfo: \n{}'.format(recv))

                # 发送 HELO 命令
                # 开始与服务器的交互，服务器将返回状态码250,说明请求动作正确完成
                heloCommand = 'HELO Alice\r\n'
                clientSocket.send(heloCommand.encode())
                recv1 = clientSocket.recv(1024).decode()
                if '250' != recv1[:3]:
                    raise Exception('250 reply not received from server.\nInfo: \n{}'.format(recv1))

                # 发送"AUTH LOGIN"命令，验证身份.服务器将返回状态码334（服务器等待用户输入验证信息）
                clientSocket.sendall('AUTH LOGIN\r\n'.encode())
                recv2 = clientSocket.recv(1024).decode()
                if '334' != recv2[:3]:
                    raise Exception('334 reply not received from server.\nInfo: \n{}'.format(recv2))

                # 发送验证信息
                clientSocket.sendall((username + '\r\n').encode())
                recvName = clientSocket.recv(1024).decode()
                if '334' != recvName[:3]:
                    raise Exception('334 reply not received from server.\nInfo: \n{}'.format(recvName))

                clientSocket.sendall((password_ + '\r\n').encode())
                recvPass = clientSocket.recv(1024).decode()
                # 如果用户验证成功，服务器将返回状态码235
                if '235' != recvPass[:3]:
                    raise Exception('235 reply not received from server.\nInfo: \n{}'.format(recvPass))

                # TCP连接建立好之后，通过用户验证就可以开始发送邮件。邮件的传送从MAIL命令开始，MAIL命令后面附上发件人的地址。
                # 发送 MAIL FROM 命令，并包含发件人邮箱地址
                clientSocket.sendall(('MAIL FROM: <' + fromAddress + '>\r\n').encode())
                recvFrom = clientSocket.recv(1024).decode()
                if '250' != recvFrom[:3]:
                    raise Exception('250 reply not received from server.\nInfo: \n{}'.format(recvFrom))

                # 接着SMTP客户端发送一个或多个RCPT (收件人recipient的缩写)命令，格式为RCPT TO: <收件人地址>。
                # 发送 RCPT TO 命令，并包含收件人邮箱地址，返回状态码 250
                for to in toAddress:
                    clientSocket.sendall(('RCPT TO: <' + to + '>\r\n').encode())
                    recvTo = clientSocket.recv(1024).decode()  # 注意UDP使用sendto，recvfrom
                    if '250' != recvTo[:3]:
                        raise Exception('250 reply not received from server.\nInfo: \n{}'.format(recvTo))

                # 发送 DATA 命令，表示即将发送邮件内容。服务器将返回状态码354（开始邮件输入，以"."结束）
                clientSocket.send('DATA\r\n'.encode())
                recvData = clientSocket.recv(1024).decode()
                if '354' != recvData[:3]:
                    raise Exception('354 reply not received from server.\nInfo: \n{}'.format(recvData))

                # 编辑邮件信息，发送数据
                contentType = "text/plain"

                message = 'from:' + fromAddress + '\r\n'
                for to in toAddress:
                    message += 'to:' + to + '\r\n'
                message += 'subject:' + subject_ + '\r\n'
                message += 'Content-Type:' + contentType + '\t\n'
                message += '\r\n' + msg
                clientSocket.sendall(message.encode())

                # 以"."结束。请求成功返回 250
                clientSocket.sendall(endMsg.encode())
                recvEnd = clientSocket.recv(1024).decode()
                if '250' != recvEnd[:3]:
                    raise Exception('250 reply not received from server.\nInfo: \n{}'.format(recvEnd))

                # 发送"QUIT"命令，断开和邮件服务器的连接
                clientSocket.sendall('QUIT\r\n'.encode())

                clientSocket.close()

                ttk.Label(mainframe, text="Email sent successfully").grid(column=4,row=9,sticky=W)

        # 成功发送后，更新联系人名单并保存
        for _ in toAddress:
            if _ != "" and _ not in contacts:
                contacts.append(_)
        with open('contacts.pkl', 'wb') as f:
            pickle.dump(contacts, f)

        # 保存到已发送邮件中
        history.append([account.get(), password.get(), receiver.get(), subject.get(), msgbody.get('1.0','end')])
        with open('history.pkl', 'wb') as f:
            pickle.dump(history, f)

        # 如果当前邮件存在草稿，则将其删除
        if os.path.exists('draft.pkl'):
            os.remove('draft.pkl')

    except Exception as e:
        print(e)
        ttk.Label(mainframe, text=str(e)).grid(column=4,row=9,sticky=W)

# 增加按钮用于发送邮件
ttk.Button(mainframe, text="Send Email", command=sendemail).grid(column=4,row=8,sticky=E)

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

account_entry.focus()

root.mainloop()

