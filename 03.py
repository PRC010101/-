from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import xlrd
from matplotlib import pyplot as plt
from tkinter import messagebox
from PIL import ImageTk, Image
import numpy


class AppUI():

    def __init__(self):
        self.root = Tk()
        self.root.title("图片识别")
        '''img = Image.open('imgs/3.jpg')
        photob = ImageTk.PhotoImage(img)
        firstlabel = rf_label(root, image=photob)
        firstlabel.pack()'''
        self.creat_leftcontent(self.root)
        self.creat_rightcontent(self.root)
        self.creat_photo()
        self.root.update()
        self.root.geometry('820x500+250+100')
        self.root.protocol("WM_DELETE_WINDOW", self.closeWindow)
        self.root.mainloop()

    def closeWindow(self):
        if messagebox.askokcancel("退出", "确定要退出吗?"):
            self.root.destroy()

    def creat_leftcontent(self, root):

        img = Image.open('imgs/3.jpg')
        photob = ImageTk.PhotoImage(img)
        lf = ttk.LabelFrame(root, text="详细信息", width= '400', height='445', labelanchor="n")
        lf.place(x=5, y=5)

        '''img = Image.open('imgs/3.jpg')
        photob = ImageTk.PhotoImage(img)
        firstlabel = rf_label(lf, image=photob)
        firstlabel.place()'''

        top_frame = Frame(lf)
        top_frame.pack(fill=X, side=TOP, pady='10')
        labl1 = ttk.Label(top_frame, text="导入空白表格")

        self.search_key = StringVar()
        global entl1
        entl1 = ttk.Entry(top_frame, textvariable=self.search_key, width='20')
        btnl1 = ttk.Button(top_frame, text="搜索", command=lambda : self.choosepic(self.search_key ,entl1, rf_label))

        labl1.pack(padx='5', side=LEFT)
        ttk.Frame(top_frame, width='12').pack(side=LEFT, fill=X)
        entl1.pack(padx='5', side=LEFT)
        btnl1.pack(padx='5', side=LEFT)



        second_frame = Frame(lf)
        second_frame.pack(fill=X, side=TOP, pady='10')


        lab02 = ttk.Label(second_frame, text="你要识别")
        global entl2
        entl2 = ttk.Entry(second_frame, width='5')
        lab12 = ttk.Label(second_frame, text="块")
        fra02 = ttk.Frame(second_frame, width='40')
        btn02 = ttk.Button(second_frame, text="选择识别区", command=lambda : self.choosearea())
        btn12 = ttk.Button(second_frame, text="清除重选")




        lab02.pack(padx='5', side=LEFT)
        entl2.pack(fill=X, side=LEFT)
        lab12.pack(padx='5', side=LEFT, fill=X)
        fra02.pack(padx='5', side=LEFT, fill=X)
        btn02.pack(padx='5', side=LEFT, fill=X)
        btn12.pack(padx='5', side=LEFT, fill=X)


        third_frame = Frame(lf)
        third_frame.pack(fill=X, side=TOP, pady='10')

        ttk.Label(third_frame, text="导入图片文件夹").pack(padx='5', side=LEFT)
        photo_key = StringVar()
        ent03 = ttk.Entry(third_frame, textvariable=photo_key, width='20')
        btn03 = ttk.Button(third_frame, text="搜索", command = lambda :self.openfile(ent03))

        ent03.pack(padx='5', side=LEFT)
        btn03.pack(padx='5', side=LEFT)


        fourth_frame = Frame(lf)
        fourth_frame.pack(fill=X, side=TOP, pady='10')

        ttk.Label(fourth_frame, text="结果保存到").pack(padx='5', side=LEFT)
        ttk.Frame(fourth_frame, width='24').pack(side=LEFT, fill=X)
        save_key = StringVar()
        ent04 = ttk.Entry(fourth_frame, textvariable=save_key, width='20')
        btn04 = ttk.Button(fourth_frame, text="搜索", command=lambda :self.openjpg(ent04))

        ent04.pack(padx='5', side=LEFT)
        btn04.pack(padx='5', side=LEFT)

        fifth_frame = Frame(lf)
        fifth_frame.pack(fill=X, side=TOP, pady='10')

        ttk.Frame(fifth_frame, width='70').pack(padx='5', side=LEFT)
        ttk.Button(fifth_frame, text="识别").pack(padx='5', side=LEFT)
        ttk.Button(fifth_frame, text='分析').pack(padx='5', side=LEFT)


    def creat_rightcontent(self,root):
        rf = ttk.LabelFrame(root, text="图片显示", width='400', height='445', labelanchor="n")
        rf.place(x=400, y=5)

        """global photo
        photo = PhotoImage(file="imgs/1.gif")
        labr11 = rf_label(rf, image=photo)
        labr11.pack()"""
        '''path = StringVar()
        img = Image.open(path)
        global photo_show
        photo_show = ImageTk.PhotoImage(img)  # 在root实例化创建，否则会报错'''
        global rf_label
        rf_label = Label(rf)
        rf_label.place(x=0,y=0)

    def choosepic(self, path, ent, lab):
        path_ = filedialog.askopenfilename(title='选择表格图片', filetypes=[('jpg', '*.jpg'), ('All Files', '*')])
        path.set(path_)
        img_open = Image.open(ent.get())
        img = ImageTk.PhotoImage(img_open)
        lab.config(image=img)
        lab.image = img  # keep a reference

    def choosearea(self):
        top  =Toplevel()
        canvas = Canvas(top,  cursor="cross")
        canvas.grid(row=0, column=0, sticky=N+S+E+W)

        frame_list = numpy.zeros((int(entl2.get()), 4))


        self.rect = None

        start_end = [0,0,0,0,0]

        canvas.bind("<ButtonPress-1>", lambda event:self.on_button_press(event, canvas,start_end))
        canvas.bind("<B1-Motion>", lambda event:self.on_move_press(event, canvas,start_end))
        canvas.bind("<ButtonRelease-1>", lambda event:self.on_button_release(event, start_end, frame_list))
        '''canvas.bind("<ButtonPress-1>", self.on_button_press)
        canvas.bind("<B1-Motion>", self.on_move_press)
        canvas.bind("<ButtonRelease-1>", self.on_button_release)'''

        self.imgopen = Image.open(entl1.get())
        canvas.config(width= self.imgopen.size[0], height= self.imgopen.size[1])
        self.tk_im = ImageTk.PhotoImage(self.imgopen)
        canvas.create_image(0, 0, anchor="nw", image=self.tk_im)
        print(frame_list)


        '''top = Toplevel()
        top.title('选择识别区')
        img_open = Image.open(ent.get())
        top.geometry('%dx%d+300+200' %( img_open.size[0], img_open.size[1]))
        top_label = Label(top)
        top_label.place(x=0, y=0)

        img = ImageTk.PhotoImage(img_open)
        top_label.config(image=img)
        top_label.image = img'''

    def openjpg(self, text):
        sfname = filedialog.askopenfilename(title='选择表格图片', filetypes=[('jpg', '*.jpg'), ('All Files', '*')])
        print(sfname)
        text.insert(INSERT, sfname)

    def openfile(self, text):
        sfname = filedialog.askdirectory(title='选择文件夹')
        print(sfname)
        text.insert(INSERT, sfname)


    def creat_photo(self):
        pass

    def pltPie(data, labels):
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示中文
        plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

        fig = plt.figure(figsize=(6, 9))
        ax = fig.add_subplot(111)
        patches, l_text, pct_text = ax.pie(data, explode=(0.1, 0), labels=labels, colors=['r', 'g'], autopct='%3.2f%%',
                                           pctdistance=0.6, shadow=True, labeldistance=1.1, startangle=90)
        plt.axis('equal')  # 设置x,y轴比例一直，这样饼状图才是圆的

        plt.legend()  # 加上图例
        plt.show()

    def on_button_press(self, event, canvas, wz_list):
        if entl2.get()=="":
            messagebox.showwarning('警告','请框选规定数量')
        elif entl1.get()=="":
            messagebox.showwarning('警告', '选择表格图片')
        elif wz_list[4]<int(entl2.get()):
            # save mouse drag start position
            wz_list[0] = event.x
            wz_list[1] = event.y

            # create rectangle if not yet exist
            #if not self.rect:
            self.rect = canvas.create_rectangle(0, 0, 1, 1, fill="")
        else:
            messagebox.showwarning('警告','请框选规定数量')



    def on_move_press(self, event, canvas, wz_list):
        curX, curY = (event.x, event.y)

        # expand rectangle as you drag the mouse
        canvas.coords(self.rect, wz_list[0], wz_list[1], curX, curY)


    def on_button_release(self, event, wz_list, save_list):
        wz_list[2] = event.x
        wz_list[3] = event.y
        j = wz_list[4]
        i=0
        for i in range(4):
            save_list[j][i] = wz_list[i]
        wz_list[4] += 1
        print(wz_list)




if __name__ == "__main__":
    AppUI()
