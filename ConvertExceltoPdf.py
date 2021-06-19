import win32com.client
import os


def Exportexcel(file_name):
    o = win32com.client.Dispatch("Excel.Application")

    o.Visible = False

    my_path = os.path.abspath(file_name)
    # print(type(my_path))
    # print(my_path)

    re_path = my_path.replace('\n', '\n')
    # print(re_path)

    # wb_path = r'C:\Users\mobil\Downloads\Pink w4x barcode 340 unit.xlsx'

    wb = o.Workbooks.Open(re_path)  # Purple W4x 400 unit(1-20)_GEN.xlsx

    y = []
    x = o.Worksheets.Count

    for i in range(1, x+1):
        y.append(i)

    path_pdf = str(input("กรอกชื่อไฟล.pdf = "))

    pdf_path = os.path.abspath(path_pdf)
    re_pdf_path = pdf_path.replace('\n', '\n')
    # print(pdf_path)

    # path_to_pdf = r'C:\Users\mobil\Downloads\hi.pdf'

    # print_area = 'A1:B25'

    user_print = str(input("กรอก column:row (ตัวอย่างเช่น A1:B30): "))

    for index in y:

        ws = wb.Worksheets[index - 1]

        ws.PageSetup.Zoom = False

        ws.PageSetup.FitToPagesTall = 1

        ws.PageSetup.FitToPagesWide = 1

        ws.PageSetup.PrintArea = user_print

    wb.WorkSheets(y).Select()

    wb.ActiveSheet.ExportAsFixedFormat(0, re_pdf_path)


def main():
    file_name = str(input("กรอกชื่อไฟล.xlsx = "))
    Exportexcel(file_name)


if __name__ == "__main__":
    main()
