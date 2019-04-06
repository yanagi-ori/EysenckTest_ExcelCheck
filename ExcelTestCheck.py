import pyexcel
import xlwt

input("Для работы программы необходимы таблицы test.xlsx и ans.xlsx в папке с исполнительным файлом"
      "\n(Нажмите Enter)")

my_array = pyexcel.get_array(file_name="test.xlsx")
ans = pyexcel.get_array(file_name="ans.xlsx")
book = xlwt.Workbook(encoding="utf-8")
results = book.add_sheet("Результаты")
iq = book.add_sheet("IQ")

for row in range(1, len(my_array)):
    total = 0
    for i in range(1, len(my_array[row])):
        if str(my_array[row][i]).lower().replace(" ", '') == str(ans[i][0]):
            total += 1
    results.write(row, 0, my_array[row][0])
    results.write(row, 1, total)
    iq.write(row, 0, my_array[row][0])
    iq.write(row, 1, 75 + 2.5 * total)

book.save("results.xls")
print("Результат программы сохранен в файл results.xls")
input()
