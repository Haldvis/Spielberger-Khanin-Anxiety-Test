import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from FileHandlerClass import *
import zipfile
#import FileHandlerClass as FHC

def donothing():
   x = 0
   
def showAbout():
    messagebox.showinfo('Разработчики', 'Python Guides aims at providing best practical tutorials')
    
def showGuide():
    messagebox.showinfo('Руководство пользователя', 'Руководство пользователя.')

def select_file():
    filetypes_zip = (
        ('ZIP файлы', '*.zip'),
    )

    filename = fd.askopenfilename(
        title='Загрузить файл',
        initialdir='/',
        filetypes=filetypes_zip)

    PathToContainingFolder = "/".join(filename.split("/")[0:-1])

    showinfo(
        title='Файлы извлечены в папку:',
        message=PathToContainingFolder
    )
    
    archive = zipfile.ZipFile(filename, 'r')
    archive.extractall(PathToContainingFolder)
    
    filetypes_csv = (
        ('CSV файлы', '*.csv'),
    )

    spil_data = fd.askopenfilename(
        title='Загрузить файл',
        initialdir=PathToContainingFolder,
        filetypes=filetypes_csv)

    showinfo(
        title='Вы загрузили файл:',
        message=spil_data
    )
    selected_file = FileHandler(spil_data)
    selected_file.ReadFile()
    selected_file.Handler()
    selected_file.WriteFile(PathToContainingFolder)
    
    showinfo(
        title='Работа завершена',
        message="Файлы с результатом работы программы находятся в папке "+PathToContainingFolder+'/Результат'
    )
   

# create the root window
root = tk.Tk()
root.title('ПСИХОМЕТР')
root.resizable(False, False)
root.geometry('450x225')
root.iconbitmap('brain.ico')
menubar = Menu(root)
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Руководство пользователя", command=showGuide)
helpmenu.add_command(label="О программе", command=showAbout)
menubar.add_cascade(label="Справка", menu=helpmenu)




# open button
open_button = ttk.Button(
    root,
    text='Выбрать файл',
    command=select_file
)

open_button.pack(expand=True)


# run the application
root.config(menu=menubar)
root.mainloop()
