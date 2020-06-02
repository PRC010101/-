# coding = gbk
import cv2
import os
import glob
import matplotlib.pyplot as plt
import imutils
from imutils.perspective import four_point_transform
import numpy as np
import json
import requests
import base64
import openpyxl
from openpyxl.workbook import Workbook
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import xlrd
from matplotlib import pyplot as plt
from tkinter import messagebox
from PIL import ImageTk, Image
from collections import Counter
global img   #空白表格处理图
global point1, point2,myi  #选区坐标,标准计数器
g_rect = {0:[]}  #选取方框
my_N = 9
my_i = 0
global aim_filename,aim_dir,aim_place
global all_result  #结果存放
wb = Workbook()
ws = wb.worksheets[0]

class AppUI():

    def __init__(self):
        self.root = Tk()
        self.root.title("图片识别")
        img = Image.open('bgp.jpg')
        photob = ImageTk.PhotoImage(img)
        firstlabel = Label(self.root, image=photob)
        firstlabel.pack()
        self.creat_leftcontent(self.root)
        self.creat_rightcontent(self.root)
        self.creat_photo()
        self.root.update()
        self.root.geometry('685x500+250+100')
        self.root.protocol("WM_DELETE_WINDOW", self.closeWindow)
        self.root.mainloop()

    def closeWindow(self):
        if messagebox.askokcancel("退出", "确定要退出吗?"):
            self.root.destroy()

    def creat_leftcontent(self, root):
        global aim_filename,aim_dir,aim_place

        lf = ttk.LabelFrame(root, text="详细信息", width= '400', height='700', labelanchor="n")
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

        labl1.pack(padx='5', side=LEFT)
        ttk.Frame(top_frame, width='12').pack(side=LEFT, fill=X)
        entl1.pack(padx='5', side=LEFT, fill=X)

        top2_frame = Frame(lf)
        top2_frame.pack(fill=X, side=TOP, pady='10')
        btnl1 = ttk.Button(top2_frame, text="搜索", command=lambda: self.choosepic(self.search_key, entl1, rf_label))
        btnl1.pack(padx='5', fill=BOTH)



        second_frame = Frame(lf)
        second_frame.pack(fill=X, side=TOP, pady='10')



        fra02 = ttk.Frame(second_frame, width='20')
        btn02 = ttk.Button(second_frame, text="选择识别区", command=lambda : get_my_roi(aim_filename))
        btn12 = ttk.Button(second_frame, text="清除重选")
        fra02.pack(padx='5', side=LEFT, fill=X)
        btn02.pack(padx='5', side=LEFT, fill=X)
        btn12.pack(padx='5', side=LEFT, fill=X)


        second2_frame = Frame(lf)
        second2_frame.pack(fill=X, side=TOP, pady='10')
        lab02 = ttk.Label(second2_frame, text="（框选后请按Enter确认)")
        lab02.pack(padx='5', side=TOP)

        third_frame = Frame(lf)
        third_frame.pack(fill=X, side=TOP, pady='10')
        ttk.Label(third_frame, text="导入图片文件夹").pack(padx='5', side=LEFT)
        photo_key = StringVar()
        ent03 = ttk.Entry(third_frame, textvariable=photo_key, width='20')
        ent03.pack(padx='5', side=LEFT)


        third2_frame = Frame(lf)
        third2_frame.pack(fill=X, side=TOP, pady='10')
        btn03 = ttk.Button(third2_frame, text="搜索", command=lambda: self.openfile(ent03))
        btn03.pack(padx='5', fill=BOTH)


        fourth_frame = Frame(lf)
        fourth_frame.pack(fill=X, side=TOP, pady='10')

        ttk.Label(fourth_frame, text="结果保存到").pack(padx='5', side=LEFT)
        ttk.Frame(fourth_frame, width='24').pack(side=LEFT, fill=X)
        save_key = StringVar()
        ent04 = ttk.Entry(fourth_frame, textvariable=save_key, width='20')
        ent04.pack(padx='5', side=LEFT)


        fourth2_frame = Frame(lf)
        fourth2_frame.pack(fill=X, side=TOP, pady='10')
        btn04 = ttk.Button(fourth2_frame, text="搜索", command=lambda: self.openjpg(ent04))
        btn04.pack(padx='5', fill=BOTH)


        fifth_frame = Frame(lf)
        fifth_frame.pack(fill=X, side=TOP, pady='10')

        ttk.Frame(fifth_frame, width='20').pack(padx='5', side=LEFT)
        ttk.Button(fifth_frame, text="识别", command = lambda: get_temp_result(aim_dir)).pack(padx='5', side=LEFT)
        ttk.Button(fifth_frame, text='分析', command= lambda :self.fenxi()).pack(padx='5', side=LEFT)

        sixth_frame = Frame(lf)
        sixth_frame.pack(fill=X, side=TOP, pady='22')

    def creat_rightcontent(self,root):
        rf = ttk.LabelFrame(root, text="图片显示", width='400', height='470', labelanchor="n")
        rf.place(x=280, y=5)

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

        if entl1.get()=='':
            pre_img_open = Image.open('imgs/1.gif')
            pre_img = ImageTk.PhotoImage(pre_img_open)
            rf_label.config(image=pre_img)
            rf_label.image = pre_img        #这一步是少不了的

    def choosepic(self, path, ent, lab):
        global aim_filename
        path_ = filedialog.askopenfilename(title='选择表格图片', filetypes=[('jpg', '*.jpg'), ('All Files', '*')])
        path.set(path_)
        aim_filename = ent.get()
        print(aim_filename)
        orig_image = read_image(str(aim_filename))
        orig_shape = np.shape(orig_image)
        resize_image1 = resize_image(orig_image, resize_height=400, resize_width=None)
        cv2.imwrite('my_temp.jpg', resize_image1)
        img_open = Image.open('my_temp.jpg')
        img = ImageTk.PhotoImage(img_open)
        lab.config(image=img)
        lab.image = img  # keep a reference


    def openjpg(self, text):
        global aim_place
        sfname = filedialog.askdirectory(title='选择文件夹')
        aim_place = sfname
        print(sfname)
        text.insert(INSERT, sfname)

    def openfile(self, text):
        global aim_dir
        sfname = filedialog.askdirectory(title='选择文件夹')
        aim_dir = sfname
        print(sfname)
        text.insert(INSERT, sfname)


    def creat_photo(self):
        pass

    def pltPie(self,data, labels):
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示中文
        plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

        fig = plt.figure(figsize=(6, 9))
        ax = fig.add_subplot(111)
        patches, l_text, pct_text = ax.pie(data, explode=(0.1, 0), labels=labels, colors=['r', 'g'], autopct='%3.2f%%',
                                           pctdistance=0.6, shadow=True, labeldistance=1.1, startangle=90)
        plt.axis('equal')  # 设置x,y轴比例一直，这样饼状图才是圆的

        plt.legend()  # 加上图例
        plt.savefig('饼图.jpg')

        pie_top = Toplevel()
        pie_lab = ttk.Label(pie_top)
        pie_img_open = Image.open('饼图.jpg')
        pie_img = ImageTk.PhotoImage(pie_img_open)
        pie_lab.config(image=pie_img)
        pie_lab.image = pie_img

    def fenxi(self):
        top_fenxi = Toplevel()
        top_fenxi.geometry('150x80+500+200')

        fx_frame1 = Frame(top_fenxi)
        fx_frame1.pack(fill=X, side=TOP, pady='5')
        fx_lab1 = ttk.Label(fx_frame1, text='需要分析第')
        fx_ent1 = ttk.Entry(fx_frame1, width='5')
        fx_lab2 = ttk.Label(fx_frame1, text='列')

        fx_lab1.pack(padx='5',side=LEFT)
        fx_ent1.pack(padx='5',side=LEFT)
        fx_lab2.pack(padx='5',side=LEFT)

        fx_frame2 = Frame(top_fenxi)
        fx_frame2.pack(fill=X, side=TOP, pady='5')
        fx_btn1 = ttk.Button(fx_frame2, text='确认', command=lambda: self.pltPie(fx_num, fx_str))
        fx_btn1.pack(side=TOP)


        fx_col = int(fx_ent1.get()) + 1
        fx_data = all_result.get(fx_col)
        fx_dir = dict(Counter(fx_data))
        fx_str = fx_dir.keys()
        fx_num = fx_dir.values()


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


def show_image(title, image):
    '''
    调用matplotlib显示RGB图片
    :param title: 图像标题
    :param image: 图像的数据
    :return:
    '''
    # plt.figure("show_image")
    # print(image.dtype)
    plt.imshow(image)
    plt.axis('on')  # 关掉坐标轴为 off
    plt.title(title)  # 图像题目
    plt.show()


def cv_show_image(title, image):
    '''
    调用OpenCV显示RGB图片
    :param title: 图像标题
    :param image: 输入RGB图像
    :return:
    '''
    channels = image.shape[-1]
    if channels == 3:
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)  # 将BGR转为RGB
    cv2.imshow(title, image)
    cv2.waitKey(0)


def read_image(filename, resize_height=None, resize_width=None, normalization=False):
    '''
    读取图片数据,默认返回的是uint8,[0,255]
    :param filename:
    :param resize_height:
    :param resize_width:
    :param normalization:是否归一化到[0.,1.0]
    :return: 返回的RGB图片数据
    '''
    img = cv2.imread(filename,1)
    if img is None:
        print("Warning:不存在:{}", filename)
        return None
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 高斯滤波
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    # 自适应二值化方法
    blurred = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 51, 2)
    '''
    adaptiveThreshold函数：第一个参数src指原图像，原图像应该是灰度图。
        第二个参数x指当像素值高于（有时是小于）阈值时应该被赋予的新的像素值
        第三个参数adaptive_method 指： CV_ADAPTIVE_THRESH_MEAN_C 或 CV_ADAPTIVE_THRESH_GAUSSIAN_C
        第四个参数threshold_type  指取阈值类型：必须是下者之一  
                                     •  CV_THRESH_BINARY,
                            • CV_THRESH_BINARY_INV
         第五个参数 block_size 指用来计算阈值的象素邻域大小: 3, 5, 7, ...
        第六个参数param1    指与方法有关的参数。对方法CV_ADAPTIVE_THRESH_MEAN_C 和 CV_ADAPTIVE_THRESH_GAUSSIAN_C， 它是一个从均值或加权均值提取的常数, 尽管它可以是负数。
    '''
    # 这一步可有可无，主要是增加一圈白框，以免刚好卷子边框压线后期边缘检测无果。好的样本图就不用考虑这种问题
    blurred = cv2.copyMakeBorder(blurred, 5, 5, 5, 5, cv2.BORDER_CONSTANT, value=(255, 255, 255))
    # canny边缘检测
    edged = cv2.Canny(blurred, 10, 100)
    # 从边缘图中寻找轮廓，然后初始化表格对应的轮廓
    '''
    findContours
    image -- 要查找轮廓的原图像
    mode -- 轮廓的检索模式，它有四种模式：
         cv2.RETR_EXTERNAL  表示只检测外轮廓                                  
         cv2.RETR_LIST 检测的轮廓不建立等级关系
         cv2.RETR_CCOMP 建立两个等级的轮廓，上面的一层为外边界，里面的一层为内孔的边界信息。如果内孔内还有一个连通物体，
                  这个物体的边界也在顶层。
         cv2.RETR_TREE 建立一个等级树结构的轮廓。
    method --  轮廓的近似办法：
         cv2.CHAIN_APPROX_NONE 存储所有的轮廓点，相邻的两个点的像素位置差不超过1，即max （abs (x1 - x2), abs(y2 - y1) == 1
         cv2.CHAIN_APPROX_SIMPLE压缩水平方向，垂直方向，对角线方向的元素，只保留该方向的终点坐标，例如一个矩形轮廓只需
                           4个点来保存轮廓信息
          cv2.CHAIN_APPROX_TC89_L1，CV_CHAIN_APPROX_TC89_KCOS使用teh-Chinl chain 近似算法
    '''
    cnts = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if imutils.is_cv2() else cnts[1]
    docCnt = None
    # 确保至少有一个轮廓被找到
    if len(cnts) > 0:
        # 将轮廓按大小降序排序
        cnts = sorted(cnts, key=cv2.contourArea, reverse=True)
        # 对排序后的轮廓循环处理
        for c in cnts:
            # 获取近似的轮廓
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.02 * peri, True)
            # 如果近似轮廓有四个顶点，那么就认为找到了表格
            if len(approx) == 4:
                docCnt = approx
                break

    # 四点转换
    paper = four_point_transform(img, docCnt.reshape(4, 2))  # 原图
    rgb_image = resize_image(paper, resize_height, resize_width)
    rgb_image = np.asanyarray(rgb_image)
    if normalization:
        # 不能写成:rgb_image=rgb_image/255
        rgb_image = rgb_image / 255.0
    # show_image("src resize image",image)
    return rgb_image


def resize_image(image, resize_height, resize_width):
    '''
    :param image:
    :param resize_height:
    :param resize_width:
    :return:
    '''
    image_shape = np.shape(image)
    height = image_shape[0]
    width = image_shape[1]
    if (resize_height is None) and (resize_width is None):  # 错误写法：resize_height and resize_width is None
        return image
    if resize_height is None:
        resize_height = int(height * resize_width / width)
    elif resize_width is None:
        resize_width = int(width * resize_height / height)
    image = cv2.resize(image, dsize=(resize_width, resize_height))
    return image


def scale_image(image, scale):
    '''
    :param image:
    :param scale: (scale_w,scale_h)
    :return:
    '''
    image = cv2.resize(image, dsize=None, fx=scale[0], fy=scale[1])
    return image


def scale_rect(orig_rect, orig_shape, dest_shape):
    '''
    对图像进行缩放时，对应的rectangle也要进行缩放
    :param orig_rect: 原始图像的rect=[x,y,w,h]
    :param orig_shape: 原始图像的维度shape=[h,w]
    :param dest_shape: 缩放后图像的维度shape=[h,w]
    :return: 经过缩放后的rectangle
    '''
    new_x = int(orig_rect[0][0] * dest_shape[1] / orig_shape[1])
    new_y = int(orig_rect[0][1] * dest_shape[0] / orig_shape[0])
    new_w = int(orig_rect[0][2] * dest_shape[1] / orig_shape[1])
    new_h = int(orig_rect[0][3] * dest_shape[0] / orig_shape[0])
    dest_rect = [new_x, new_y, new_w, new_h]
    return dest_rect


def rgb_to_gray(image):
    image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    return image


def save_image(image_path, rgb_image, toUINT8=True):
    if toUINT8:
        rgb_image = np.asanyarray(rgb_image * 255, dtype=np.uint8)
    if len(rgb_image.shape) == 2:  # 若是灰度图则转为三通道
        bgr_image = cv2.cvtColor(rgb_image, cv2.COLOR_GRAY2BGR)
    else:
        bgr_image = cv2.cvtColor(rgb_image, cv2.COLOR_RGB2BGR)
    cv2.imwrite(image_path, bgr_image)


def get_words(filename):
    global all_result
    #request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/handwriting"
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"
    # 二进制方式打开图片文件
    f = open('./'+filename, 'rb')
    img1 = base64.b64encode(f.read())
    params = {"image": img1}
    access_token = '24.d4c5917951f269772bce2caa3ad6cef5.2592000.1593621893.282335-20157723'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print(response.json())
        myresult = response.json()  # 我获得的全部数据
        my_result=[]
        my_result1 = []
        if 'error_msg' in myresult.keys():
            my_result.append('error')
        else:
            for i in range(0,len(myresult['words_result'])):
                my_result.append(myresult['words_result'][i]['words'])
        my_result1="".join(my_result)
        temp_i = filename[4]
        print(temp_i)
        all_result[temp_i].append(my_result1)



def on_mouse(event, x, y, flags,param):
    global img, point1, point2,my_i
    img2 = img.copy()
    if event == cv2.EVENT_LBUTTONDOWN:  # 左键点击
        point1 = (x, y)
        cv2.circle(img2, point1, 10, (0, 255, 0), 5)
        cv2.imshow('image', img2)
    elif event == cv2.EVENT_MOUSEMOVE and (flags & cv2.EVENT_FLAG_LBUTTON):  # 按住左键拖曳
        cv2.rectangle(img2, point1, (x, y), (255, 0, 0), 5)
        cv2.imshow('image', img2)
    elif event == cv2.EVENT_LBUTTONUP:  # 左键释放
        point2 = (x, y)
        cv2.rectangle(img2, point1, point2, (0, 0, 255), 5)
        cv2.imshow('image', img2)
        min_x = min(point1[0], point2[0])
        min_y = min(point1[1], point2[1])
        width = abs(point1[0] - point2[0])
        height = abs(point1[1] - point2[1])
        g_rect[my_i] = [min_x, min_y, width, height]
        cut_img = img[min_y:min_y + height, min_x:min_x + width]
        print(g_rect[my_i])
        my_i += 1
        #cv2.imwrite('lena{}.jpg'.format(str(my_i)), cut_img)



def get_image_roi(rgb_image):
    '''
    获得用户ROI区域的rect=[x,y,w,h]
    :param rgb_image:
    :return:
    '''
    bgr_image = cv2.cvtColor(rgb_image, cv2.COLOR_RGB2BGR)
    global img
    global g_rect
    img = bgr_image
    my_i = 0
    cv2.namedWindow('image')
    cv2.setMouseCallback('image', on_mouse)
    cv2.startWindowThread()  # 加在这个位置
    while True:
        cv2.imshow('image', img)
        key = cv2.waitKey(0)
        if key == 13 or key == 32:  # 按空格和回车键退出
            break
    cv2.destroyAllWindows()
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    return g_rect


def select_user_roi(image_path):
    orig_image = read_image(image_path)
    orig_shape = np.shape(orig_image)
    resize_image1 = resize_image(orig_image, resize_height=800, resize_width=None)
    re_shape = np.shape(resize_image1)
    g_rect = get_image_roi(resize_image1)
    orgi_rect = scale_rect(g_rect, re_shape, orig_shape)
    roi_image = get_rect_image(orig_image, orgi_rect)
    #cv_show_image("RECT", roi_image)
    return roi_image


def get_my_roi(image_path):
    '''
    由于原图的分辨率较大，这里缩小后获取ROI，返回时需要重新scale对应原图
    :param image_path:
    '''
    print(image_path)
    orig_image = read_image(str(image_path))
    orig_shape = np.shape(orig_image)
    resize_image1 = resize_image(orig_image, resize_height=800, resize_width=None)
    re_shape = np.shape(resize_image1)   #至此获得标准图片
    g_rect = get_image_roi(resize_image1)   #至此获得方框区域坐标
    #orgi_rect = scale_rect(g_rect, re_shape, orig_shape)
    #roi_image = get_rect_image(orig_image, orgi_rect)

def get_result(image_path):

    orig_image = read_image(str(image_path))
    orig_shape = np.shape(orig_image)
    resize_image1 = resize_image(orig_image, resize_height=800, resize_width=None)
    print(len(g_rect))
    for i in range (0,len(g_rect)):
        my_rect=g_rect[i]
        cut_img = orig_image[my_rect[1]:my_rect[1] + my_rect[3], my_rect[0]:my_rect[0] + my_rect[2]]
        cv2.imwrite('temp{}.jpg'.format(str(i)), cut_img)
        get_words('temp{}.jpg'.format(str(i)))


def get_temp_result(filename):
    print(filename)
    global all_result,aim_place
    all_result= {'0': [], '1': [], '2': [], '3': [], '4': [], '5': [], '6': [], '7': [], '8': [], '9': [], '10': [], '11': [],
       '12': []}
    for root, dirs, files in os.walk(str(filename)):
        for d in dirs:
            print(d)  # 打印子资料夹的个数
        for file in files:
            print(file)
            # 讀入圖像
            img_path = root + '/' + file
            print(img_path)
            get_result(img_path)
    print(all_result)
    for i in range(0, len(all_result)):
        for j in range(0,len(all_result[str(i)])):
            ws.cell(j+1 , i+1).value = all_result[str(i)][j]
            print(all_result[str(i)][j])
    wb.save("{}/result.xlsx".format(aim_place))




if __name__ == '__main__':
    AppUI()


