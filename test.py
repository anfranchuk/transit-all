import win32com.client as win32
import datetime
def SaveToExcel():
    # создание нового экземпляра Excel
    objExcel = win32.Dispatch("Excel.Application")
    objExcel.Visible = True

    # создание новой рабочей книги
    objWorkbook = objExcel.Workbooks.Add()
    objWorksheet1 = objWorkbook.Sheets.Add()
    objWorksheet2 = objWorkbook.Sheets.Add()

    # наименование листов
    objWorksheet1.Name = "Входящие"
    objWorksheet2.Name = "Исходящие"

    # настройка заголовков столбцов
    objWorksheet1.Cells(1, 1).Value = "Отправитель"
    objWorksheet1.Cells(1, 2).Value = "Кому"
    objWorksheet1.Cells(1, 3).Value = "Тема"
    objWorksheet1.Cells(1, 4).Value = "Получено"
    objWorksheet1.Cells(1, 5).Value = "Размер"
    objWorksheet1.Cells(1, 6).Value = "Категория"

    objWorksheet2.Cells(1, 1).Value = "Отправитель"
    objWorksheet2.Cells(1, 2).Value = "Кому"
    objWorksheet2.Cells(1, 3).Value = "Тема"
    objWorksheet2.Cells(1, 4).Value = "Получено"
    objWorksheet2.Cells(1, 5).Value = "Размер"
    objWorksheet2.Cells(1, 6).Value = "Категория"

    # выбор папки с письмами
    outlook = win32.Dispatch("Outlook.Application")
    objFolder = outlook.GetNamespace("MAPI").Folders("Технический Андеррайтер").Folders("Входящие").Folders("Вход")
    objItems = objFolder.Items

    row = 2

    # сохранение входящих писем
    for objMail in objItems:
        if objMail.Class == 43 and objMail.ReceivedTime.date() == datetime.date.today():
            objWorksheet1.Cells(row, 1).Value = objMail.SenderName
            objWorksheet1.Cells(row, 2).Value = objMail.To
            objWorksheet1.Cells(row, 3).Value = objMail.Subject
            objWorksheet1.Cells(row, 4).Value = objMail.ReceivedTime
            objWorksheet1.Cells(row, 5).Value = objMail.Size
            objWorksheet1.Cells(row, 6).Value = objMail.Categories
            row += 1

    # выбор папки с исходящими письмами
    objFolder = outlook.GetNamespace("MAPI").Folders("Технический Андеррайтер").Folders("Входящие").Folders("Выход")
    objItems = objFolder.Items

    row = 2

    # сохранение исходящих писем
    for objMail in objItems:
        if objMail.Class == 43 and objMail.ReceivedTime.date() == datetime.date.today():
            objWorksheet2.Cells(row, 1).Value = objMail.SenderName
            objWorksheet2.Cells(row, 2).Value = objMail.To
            objWorksheet2.Cells(row, 3).Value = objMail.Subject
            objWorksheet2.Cells(row, 4).Value = objMail.ReceivedTime
            objWorksheet2.Cells(row, 5).Value = objMail.Size
            objWorksheet2.Cells(row, 6).Value = objMail.Categories
            row += 1

    # создание имени файла в формате ЧЧММГГГГ.xlsx
    fileName = datetime.datetime.now().strftime("%H%M%Y") + ".xlsx"

    # сохранение книги
    objWorkbook.SaveAs(r"C:\temp" + fileName)

    # освобождение ресурсов
    objWorkbook.Close()
    objExcel.Quit()
    del objWorkbook
    del objExcel

    print("Письма сохранены в файл Excel")
SaveToExcel()