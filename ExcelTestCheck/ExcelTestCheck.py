from tkinter import Tk, Button, Label
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showerror
from PIL import ImageTk, Image

import pyexcel
import xlwt


def calculation(test_path, ans_path):
    my_array = pyexcel.get_array(file_name=test_path)
    ans = pyexcel.get_array(file_name=ans_path)
    book = xlwt.Workbook(encoding="utf-8")
    results = book.add_sheet("Результаты")
    iq = book.add_sheet("IQ")
    for row in range(1, len(my_array)):
        total = 0
        for i in range(1, len(my_array[row])):
            if str(my_array[row][i]).lower().replace(" ", '') == str(ans[i][0]):
                total += 1
        results.write(row - 1, 0, my_array[row][0])
        results.write(row - 1, 1, total)
        iq.write(row - 1, 0, my_array[row][0])
        if total == 0:
            iq.write(row, 1, "<75")
        else:
            iq.write(row, 1, 75 + 2.5 * total)

    book.save("results.xls")


def initialization():
    def open_test():
        root.test_path = askopenfile(parent=root)
        if root.test_path is not None:
            root.checkmark_l1.place(anchor='center', rely=0.4, relx=0.8)

    def open_ans():
        root.ans_path = askopenfile(parent=root)
        if root.ans_path is not None:
            root.checkmark_l2.place(anchor='center', rely=0.5, relx=0.8)

    def calculate():
        if root.test_path is not None and root.ans_path is not None:
            try:
                calculation(root.test_path.name, root.ans_path.name)
                Label(text="Готово. Провертье папку с программой.").place(anchor='center', rely=0.8, relx=0.5)
            except Exception:
                Label(text="Неверный тип входных данных.").place(anchor='center', rely=0.8, relx=0.5)
        else:
            showerror(title="Ошибка", message="Не выбраны исходные файлы")

    root.button_test = Button(root, text='Файл с тестовыми данными',
                              command=open_test).place(anchor="center", rely=0.4, relx=0.5)
    root.button_ans = Button(root, text='Файл с ответами',
                             command=open_ans).place(anchor='center', rely=0.5, relx=0.5)
    root.button_ans = Button(root, text='Подсчитать', command=calculate).place(anchor='center', rely=0.6, relx=0.5)


root = Tk()
root.title("EysenckTest")
root.geometry("360x360")
root.test_path = None
root.ans_path = None

checkmark_img = ImageTk.PhotoImage(Image.open("./checkmark.png").resize((25, 25), Image.ANTIALIAS))
root.checkmark_l1 = Label(image=checkmark_img)
root.checkmark_l2 = Label(image=checkmark_img)

initialization()
root.mainloop()
