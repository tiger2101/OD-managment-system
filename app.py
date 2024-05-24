import customtkinter
from PIL import Image
import tkinter as tk
from read import readFile
from read import createFile
from openpyxl import Workbook
from openpyxl import *
from openpyxl import load_workbook
from main import HomePage
from main import ToplevelWindow
from tkinter import filedialog

customtkinter.set_appearance_mode("light")

class myFrame(customtkinter.CTkFrame):
    def __init__(self,master,**kwargs):
        super().__init__(master,**kwargs)


class button_frame(customtkinter.CTkFrame):

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title('Amrita OD Managment System')
        self.minsize(1200,600)
        self.iconbitmap('icon.ico')
        self.toplevel_window = None
        self.image = customtkinter.CTkImage(light_image=Image.open("upload.png"), size=(300, 300))
        self.image_ = customtkinter.CTkLabel(self, image=self.image,text='')
        self.image_.pack()

        self.empty_frame = tk.Label(self, text='Upload the students excel sheet',background='teal',fg='#efefef')
        self.empty_frame.pack(ipadx=10,ipady=5)

        self.buttons = button_frame(self)
        self.buttons.pack(pady=10)

        self.browse_button = customtkinter.CTkButton(
            master=self.buttons, text='Upload', command=self.upload)
        self.browse_button.grid(row=0, column=0, padx=5, ipadx=10, ipady=10)

        self.next_button = customtkinter.CTkButton(
            master=self.buttons, text='Next', command=self.next_frame)
        self.next_button.grid(row=0, column=1, padx=5, ipadx=10, ipady=10)
        self.toplevel_window = None
    def next_frame(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            # create window if its None or destroyed
            self.toplevel_window = ToplevelWindow(self)
        else:
            self.toplevel_window.focus()

    def upload(self):
       file_path = filedialog.askopenfilename(
           title="Select file", filetypes=[("Excel files", "*.xlsx;*.xls;*.csv")])
       createFile(file_path)
       self.empty_frame.config(text=file_path)
       


if __name__ == '__main__':

    try:
        file_path = readFile('path.txt')
    except:
        createFile('@')
        file_path = readFile('path.txt')

    if file_path != '@':
        home = HomePage()
        home.mainloop()
    else:
        app = App()
        app.mainloop()
