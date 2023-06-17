import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl.reader.excel import load_workbook

def doZoom(xls_file_path, output_file_path):
    data_frame = pd.read_excel(xls_file_path)
    duration = data_frame['Тривалість']
    Name = data_frame["Ім'я(справжнє)"]
    output_wb = load_workbook(output_file_path)
    output_ws = output_wb.active
    column_duration = "Відвідуваність"
    column_name = "Ім'я(справжнє)"
    target_column_attendance = None
    target_column_name = None
    for column in output_ws.iter_cols(min_row=1, max_row=1):
        if column[0].value == column_duration:
            target_column_attendance = column[0].column
        if column[0].value == column_name:
            target_column_name = column[0].column
        if target_column_attendance != None and target_column_name != None:
            break

    data = zip(duration, Name)

    for i, (attendance, name) in enumerate(data, start=2):
        status = "absent"
        if attendance >= 46:
            status = "present"
        output_ws.cell(row=i, column=target_column_attendance, value=status)
        output_ws.cell(row=i, column=target_column_name, value=name)

    output_wb.save(output_file_path)
    print("Success")

def doMoodle(xls_file_path, output_file_path):
    data_frame = pd.read_excel(xls_file_path)
    Grades = data_frame['Загальне за курс (Бали)']
    Surname = data_frame['Прізвище']
    Name = data_frame["Ім'я"]
    output_wb = load_workbook(output_file_path)
    output_ws = output_wb.active
    column_grade = "Бали"
    column_surname = "Прізвище"
    column_name = "Ім'я"
    target_column_grade = None
    target_column_name = None
    target_column_surname = None
    for column in output_ws.iter_cols(min_row=1, max_row=1):
        if column[0].value == column_grade:
            target_column_grade = column[0].column
        if column[0].value == column_surname:
            target_column_surname = column[0].column
        if column[0].value == column_name:
            target_column_name = column[0].column
        if target_column_surname!=None and target_column_name!=None and target_column_grade!=None:
            break

    data = zip(Grades, Surname, Name)
    for i, (grade, surname, name) in enumerate(data, start=2):
        output_ws.cell(row=i, column=target_column_grade, value=grade)
        output_ws.cell(row=i, column=target_column_surname, value=surname)
        output_ws.cell(row=i, column=target_column_name, value=name)

    output_wb.save(output_file_path)
    print("Success")

def button_click():
    input_file_path = filedialog.askopenfilename(title="Select input file")
    output_file_path = filedialog.asksaveasfilename(title="Select output file")

    if input_file_path and output_file_path:
        if "Zoom" in input_file_path:
            doZoom(input_file_path, output_file_path)
        if "Moodle" in input_file_path:
            doMoodle(input_file_path, output_file_path)

root = tk.Tk()
root.title("Fill table")

button = tk.Button(root, text="Submit", command=button_click)
button.pack(pady=20)

root.mainloop()
