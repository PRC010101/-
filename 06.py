import tkinter as tk
from tkinter import ttk
from PIL import Image,ImageTk

root = tk.Tk()
fra1 = ttk.LabelFrame(root).pack()
path = 'imgs/3.jpg'
img = Image.open(path)
photo = ImageTk.PhotoImage(img)#在root实例化创建，否则会报错
label = tk.Label(fra1,image=photo)
label.pack()
root.mainloop()