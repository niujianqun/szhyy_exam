#-*- encoding=UTF-8 -*-


__author__ = 'daniu'

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import exam.examClass as ex

'''
开启Entry对输入文本验证功能。
1、实现该功能，需要通过设置validate、validatecommand和invalidcommand三个选项。
2、启用验证的开关是validate选项，该选项可以设置以下的值：
    focus：当entry组件获得或者失去焦点的时候验证
    focusin：当entry组件获得焦点的时候验证
    focusout:当entry组件失去焦点的时候验证
    key:当输入框被编辑的时候验证
    all:当出现上面任何一种情况时候验证
    none:关闭验证功能。默认设置为该选项
3、validatecommand选项指定一个验证函数，该函数只能返回True或者False表示验证结果，一般情况下验证函数只需要知道输入框中的内容即可，
可以通过Entry组件的get()方法来获得该字符串。
4、invalidcommand选项指定的函数只有在validatecommand的返回值为False的时候才被调用。

Tkinter 为验证函数提供了一些隐藏的功能选项。
%d:操作代码-0表示删除操作，1表示插入操作，2表示获得、失去焦点或textvariable变量值被修改
%i:1、当用户尝试插入或删除操作时候，该选项表示插入或删除的位置（索引号）
    2、如果是获得、失去焦点或textvariable变量值被修改该而地啊哟用验证函数，那么该值是-1
%P: 1、当输入框的值允许改变的时候，该值有效
    2、该值为输入框的最新文本内容    
%s: 1、该值为调用验证函数前输入框的文本内容
%S：1、当插入或删除操作出发验证函数的时候，该值有效
    2、该选项表示文本被插入和删除的内容
%v: 1、该组件的validate选项的值
%V: 1、调用验证函数的原因
    2、该值是focusin，focusout，ke，或forced（textvariable选项指定的变量值被修改）中的一个
%W: 该组件的名称   

为了使用这些选项，我们可以这样修改我们的validatecommand选项：
validatecommand=(f,s1,s2,……)
其中，f是验证函数名，s1,s2等是额外的选项，这些选项会作为参数依次传给f函数，常用的为：validatecommand=(funName,'%P','%v','%W')
我们在使用隐藏的功能选项前需要冷却，这就是register()方法将验证函数包装起来。

'''

def checkNameEntry(nameValue, validOption, fieldName):
        return True


def checkNameEntryInvalid():
    print('nameEntry 不合法时执行')


def selFile():
    path_ = filedialog.askopenfilename()
    fileEntryV.set(path_)



# 采集录入信息，生成考场
def createExam():

    # 1、栏位检查
    name = name_entry.get().strip()
    file = file_entry.get().strip()
    num = num_entry.get().strip()
    wl = wl_cmo.get().strip()

    if checkAllField(name,file,num,wl):
        # 2、解析成绩单, 生成考场文件
        exam = ex.examaitonClass(name,file,num,wl)
        retcode, retmsg = exam.main()

        if retcode == 0:
            success()
        else:
            fail(retmsg)

def success():
    tk.messagebox.showinfo('提示', '已完成，请移步成绩单目录查看结果')

def fail(retmsg):
    tk.messagebox.showerror('提示', retmsg)

def checkAllField(name,file,num,wl):
    if name == '':
        messagebox.showwarning('警告', '请录入本次考试名称')
        name_entry.focus()
        return False
    if num == '' or not num.isdigit() :
        messagebox.showwarning('警告', '请正确录入每个考场人数')
        num_entry.focus()
        return False

    if int(num) <= 0 or int(num) < 10 or int(num) > 100:
        messagebox.showwarning('警告', '考场人数不合理，请输入合理人数[10-100]')
        num_entry.focus()
        return False

    if int(num)%2 != 0:
        messagebox.showwarning('警告', '考场人数不允许为单数')
        num_entry.focus()
        return False

    if wl == '' or (wl !='是' and wl !='否'):
        messagebox.showwarning('警告', '请选择文理科是否分考场')
        wl_cmo.focus()
        return False

    if file.strip() == '':
        messagebox.showwarning('警告', '请选择成绩单（excel）文件')
        file_sel_btn.focus()
        return False

    return True

############### 页面 ########################
root = tk.Tk()
root.title("学校考场分布系统")
root.geometry("800x450+500+200")
root.resizable(0, 0)

win = tk.Frame(root)
win.pack(side=tk.TOP, expand = tk.YES, fill=tk.NONE, anchor='nw')


fileEntryV = tk.StringVar()
banquanV =  tk.StringVar()
#
nameEntryCMD = win.register(checkNameEntry)

# 考试名称
name_label = tk.Label(win, text="本次考试名称 ")
name_label.grid(row=0, column=0, sticky=tk.E)

name_entry = tk.Entry(win, width=50, validate='focusout', validatecommand=(nameEntryCMD,'%P','%v','%W'),
                                                          invalidcommand=checkNameEntryInvalid)
# name_entry.insert(0, '期末考试')
name_entry.grid(row=0, column=1, sticky=tk.W)
name_entry.focus()


# 每个考场人数
num_label = tk.Label(win, text="每个考场人数 ")
num_label.grid(row=1, column=0, sticky=tk.E)

num_entry = tk.Entry(win, width=50)
num_entry.insert(0, '30')
num_entry.grid(row=1, column=1, sticky=tk.W)

# 文理科是否分考场
wl_label = tk.Label(win, text="文理科是否分考场 ")
wl_label.grid(row=2, column=0, sticky=tk.E)

wl_cmo = ttk.Combobox(win)
wl_cmo['value'] = ('否','是')
wl_cmo.current(0)
wl_cmo.grid(row=2, column=1, sticky=tk.W)

# 成绩单附件
file_label = tk.Label(win, text="成绩单（excel）")
file_label.grid(row=3, column=0, sticky=tk.W)

file_entry = tk.Entry(win, textvariable=fileEntryV, state="disabled", width=50)
file_entry.grid(row=3, column=1, sticky=tk.E)

file_sel_btn = tk.Button(win, text="选择", command=selFile)
file_sel_btn.grid(row=3, column=2)

filler_label = tk.Label(win, text="")
filler_label.grid()
# 执行
btn = tk.Button(win, text="生成考场", command=createExam)
btn.grid(row=5, column=1)

# 版权
frame = tk.Frame(root)
frame.pack(side=tk.BOTTOM, expand = tk.YES, fill=tk.NONE, anchor='sw')
xx_label = tk.Label(frame, height=15)
xx_label.pack()

banquan_label = tk.Label(frame, text="版权归大牛[niujianqun@126.com]所有，仅用于考场分布，使用须经本人同意，违者必究。", height=100, width=75)
banquan_label.pack()

win.mainloop()
