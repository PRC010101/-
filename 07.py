# -*- coding:utf-8 -*-
import tkinter
import tkinter.filedialog
import os
from PIL import ImageGrab
from time import sleep
from tkinter import *

from PIL import Image,ImageTk
from tkinter.filedialog import askopenfilename
import json
import codecs

# 创建tkinter主窗口
root = tkinter.Tk()
root.title('图片处理')

# 指定主窗口位置与大小
root.geometry('800x600+100+50')   # width x height + widthoffset + heightoffset

# 不允许改变窗口大小
root.resizable(False, False)
root.focusmodel()

class MyCapture:
    def __init__(self, png):

        # 变量X和Y用来记录鼠标左键按下的位置
        self.X = tkinter.IntVar(value=0)
        self.Y = tkinter.IntVar(value=0)
        self.selectPosition = None

        # 屏幕尺寸
        screenWidth = root.winfo_screenwidth()
        screenHeight = root.winfo_screenheight()

        # 创建顶级组件容器
        self.top = tkinter.Toplevel(root, width=screenWidth, height=screenWidth)

        # 不显示最大化、最小化按钮
        self.top.overrideredirect(True)
        self.canvas = tkinter.Canvas(self.top, bg='white', width=screenWidth, height=screenWidth)

        # 显示全屏截图，在全屏截图上进行区域截图
        self.image = tkinter.PhotoImage(file=png)
        self.canvas.create_image(screenWidth // 2, screenHeight // 2, image=self.image)

        # 鼠标左键按下的位置
        def onLeftButtonDown(event):
            self.X.set(event.x)
            self.Y.set(event.y)
            # 开始截图
            self.sel = True

        self.canvas.bind('<Button-1>', onLeftButtonDown)

        # 鼠标左键移动，显示选取的区域
        def onLeftButtonMove(event):
            if not self.sel:
                return
            global lastDraw
            try:
                # 删除刚画完的图形，要不然鼠标移动的时候是黑乎乎的一片矩形
                self.canvas.delete(lastDraw)
            except Exception as e:
                pass
            lastDraw = self.canvas.create_rectangle(self.X.get(), self.Y.get(), event.x, event.y, outline='red')

        self.canvas.bind('<B1-Motion>', onLeftButtonMove)

        # 获取鼠标左键抬起的位置，保存区域截图
        def onLeftButtonUp(event):
            self.sel = False
            try:
                self.canvas.delete(lastDraw)
            except Exception as e:
                pass
            sleep(0.1)

            # 更新主窗口位置坐标
            root.update()
            x = root.winfo_x()
            y = root.winfo_y()

            #考虑鼠标左键从右下方按下而从左上方抬起的截图
            myleft, myright = sorted([self.X.get(), event.x])
            mytop, mybottom = sorted([self.Y.get(), event.y])

            # Position ： x , y, w, h
            self.selectPosition = (myleft-x,  mytop-y-30, myright-x,mybottom-y-30)
            print (myleft, myright)
            self.top.destroy()
            print ("destroy")
        self.canvas.bind('<ButtonRelease-1>', onLeftButtonUp)
        self.canvas.pack(fill=tkinter.BOTH, expand=tkinter.YES)
        # 开始截图

#设置截取坐标全局变量， 元组类型
CapturePosition = ()

def buttonCaptureClick():
    filename = 'temp.png'
    im = ImageGrab.grab()
    im.save(filename)
    im.close()

    # 显示全屏幕截图
    w = MyCapture(filename)
    buttonCapture.wait_window(w.top)

    xy_text.set(str(w.selectPosition))
    global  CapturePosition
    CapturePosition = w.selectPosition

    root.state('normal')
    os.remove(filename)

#获取输入的标注内容
def getstr():
    if input_valueText.get() != '':
        words = input_valueText.get()
        get_label["text"] = words


def save_jsonfile():
    json_content = {}
    #获取坐标
    json_content['x'] = CapturePosition[0]
    json_content['y'] = CapturePosition[1]
    json_content['w'] = CapturePosition[2]
    json_content['h'] = CapturePosition[3]

    # 获取输入内容
    value = get_label["text"]
    json_content['value'] = value

    # 存入json文件
    final_json = json.dumps(json_content, ensure_ascii=False)
    with codecs.open ("d:\\data_json.json",'a','utf-8') as file:
          file.write(final_json)


# 定义坐标显示位置
xy_text = StringVar()


#打开图片文件并显示
def choosepic():
    path_ = askopenfilename()
    path.set(path_)
    img_open = Image.open(file_entry.get())
    img = ImageTk.PhotoImage(img_open)
    image_label.config(image=img)
    image_label.image = img  # keep a reference

path = StringVar()
tkinter.Button(root, text = '打开文件' ,command=choosepic).place(x = 100, y = 530,w = 150, h = 40)
file_entry = Entry(root, state='readonly', text=path)
file_entry.pack()
image_label = Label(root, bg = 'gray')
image_label.place(x=0, y=0,width = 500, height = 500)


#添加截图按钮功能
buttonCapture = tkinter.Button(root, text='标注', command=buttonCaptureClick)

#坐标名称及显示坐标
label_xytitle = tkinter.Label(root,text ='标注位置坐标x, y, w, h :')
label_xytitle.place(x=520, y=50)
label = tkinter.Label(root, textvariable=xy_text,fg = 'blue')
label.place(x=520, y=70)


# 输入标注内容
input_valueTitle = StringVar()
input_valueTitle.set('请输入内容：')
label_inputTitle = tkinter.Label(root, textvariable=input_valueTitle)
label_inputTitle.place(x=520, y=90)
#输入框
input_valueText = tkinter.Entry(root,width=12)
input_valueText.place(x =520 , y=120,width = 250, height=40,anchor = NW)

#获取输入内容
comand = tkinter.Button(root, text="获取内容" ,command=getstr, width=10, height=2)
comand.place(x =520 , y=150,width = 250, height=40,anchor = NW)
get_label = tkinter.Label(root, fg = 'blue')
get_label.place(x =520 , y=200)

buttonCapture.place(x=300, y=530, width=150, height=40)

#保存数据
save = tkinter.Button(root, text="保存数据", command=save_jsonfile, width=10, height=2)
save.place(x =500 , y=530,width = 150, height=40,anchor = NW)

# 启动消息主循环
root.update()
root.mainloop()
