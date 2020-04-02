from tkinter import filedialog
from tkinter import *
from main import GradeSheet


def get_folder():
    global folder_path
    folder = filedialog.askdirectory()
    folder_path.set(folder)


def get_file():
    global file_path
    filename = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"),))
    file_path.set(filename)


def create_sheets():
    global msg
    global folder_path
    global file_path
    global isStamp, isPDF, date_doc

    gs = GradeSheet()

    # получаем директорию
    gs.getDirForSave(folder_path.get())

    # получаем путь до файла
    gs.getFileCSV(file_path.get())

    if isPDF.get():
        try:
            gs.createPDF(gs._getDict(), isStamp.get(), date_doc.get())
            msg.set('Формирование оценочных ведомостей завершено успешно!')
        except:
            msg.set('Возникла непредвиденная ошибка!')
    else:
        try:
            gs.createDocs(gs._getDict(), isStamp.get(), date_doc.get())
            msg.set('Формирование оценочных ведомостей завершено успешно!')
        except:
            msg.set('Возникла непредвиденная ошибка!')


root = Tk()
root.title('Elvira - generator docs')
root.geometry(f'500x300+500+200')  # ширина=500, высота=400, x=300, y=200
root.resizable(False, False)  # размер окна может быть изменён только по горизонтали

folder_path = StringVar()
file_path = StringVar()
folder_path.set('Выберите папку для сохранения оценочных ведомостей')
file_path.set('Выберите scv файл с данными о волонтерах')

lbl1 = Label(master=root, width=50, height=3, textvariable=folder_path)
lbl1.grid(row=0, column=1)

buttonB1 = Button(text="Выберите папку", command=get_folder)
buttonB1.grid(row=0, column=2)

lbl2 = Label(master=root, width=50, height=3, textvariable=file_path)
lbl2.grid(row=1, column=1)

buttonB2 = Button(text="Выберите файл", command=get_file)
buttonB2.grid(row=1, column=2)

isStamp = BooleanVar()
check1 = Checkbutton(text='Выполнить простановку печати', variable=isStamp, onvalue=True, offvalue=False, height=3)
check1.grid(row=2, column=1)

isPDF = BooleanVar()
check2 = Checkbutton(text='Верстать в PDF', variable=isPDF, onvalue=True, offvalue=False, height=3)
check2.grid(row=2, column=2)

lbl_date = Label(master=root, width=50, height=3, text='Введите дату в удобном для вас формате: ')
lbl_date.grid(row=3, column=1)
date_doc = StringVar()
date_entry = Entry(textvariable=date_doc)
date_entry.grid(row=3, column=2)

buttonB3 = Button(text="Сгенерировать ведомости", command=create_sheets)
buttonB3.grid(row=4, column=1)

msg = StringVar()

lbl4 = Label(master=root, width=50, height=5, fg='red', textvariable=msg)
lbl4.grid(row=5, column=1)

root.mainloop()
