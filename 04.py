from tkinter import *
from tkinter import filedialog
import tkinter.messagebox


def main():
    def selectExcelfile():
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xlsx'), ('All Files', '*')])
        print(sfname)
        text1.insert(INSERT, sfname)

    def closeThisWindow():
        root.destroy()

    def doProcess():
        tkinter.messagebox.showinfo('提示', '处理Excel文件的示例程序。')

    # 初始化
    root = Tk()

    # 设置窗体标题
    root.title('Python GUI Learning')

    # 设置窗口大小和位置
    root.geometry('500x300+570+200')

    label1 = Label(root, text='请选择文件:')
    text1 = Entry(root, bg='white', width=45)
    button1 = Button(root, text='浏览', width=8, command=selectExcelfile)
    button2 = Button(root, text='处理', width=8, command=doProcess)
    button3 = Button(root, text='退出', width=8, command=closeThisWindow)

    label1.pack()
    text1.pack()
    button1.pack()
    button2.pack()
    button3.pack()

    label1.place(x=30, y=30)
    text1.place(x=100, y=30)
    button1.place(x=390, y=26)
    button2.place(x=160, y=80)
    button3.place(x=260, y=80)

    root.mainloop()


if __name__ == "__main__":
    main()