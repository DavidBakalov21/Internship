import tkinter as tk

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
    input1_text = input1.get()
    input2_text = input2.get()
    if "Zoom" in input1_text:
        doZoom(input1_text, input2_text)
    if "Moodle" in input1_text:
        doMoodle(input1_text, input2_text)

def clear_input(entry):
    entry.delete(0, tk.END)

root = tk.Tk()
root.title("Fill table")

# Create labels
label1 = tk.Label(root, text="Path to excel file:")
label1.grid(row=0, column=0, sticky=tk.E)

# Create input field and clear button
input1 = tk.Entry(root)
input1.grid(row=0, column=1)

clear_button1 = tk.Button(root, text="Clear",  command=lambda: clear_input(input1))
clear_button1.grid(row=0, column=2)

label2 = tk.Label(root, text="Path to result file:")
label2.grid(row=1, column=0, sticky=tk.E)

# Create input field and clear button
input2 = tk.Entry(root)
input2.grid(row=1, column=1)

clear_button2 = tk.Button(root, text="Clear", command=lambda: clear_input(input2))
clear_button2.grid(row=1, column=2)

# Create button
button = tk.Button(root, text="Submit", command=button_click)
button.grid(row=2, columnspan=3)

# Center-align the widgets
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)

root.mainloop()