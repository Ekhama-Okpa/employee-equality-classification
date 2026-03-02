
import openpyxl

wb = openpyxl.load_workbook("Task 5 Equality Table (1).xlsx")
working_sheet = wb["Sheet1"]
num = 2
employee_number = working_sheet.max_row

while num < employee_number:
    equality_score_cell = working_sheet[f"C{num}"]
    equality_class_cell = working_sheet[f"D{num}"]
    if equality_score_cell.value > 20 or equality_score_cell.value < -10:
        equality_class_cell.value = "Highly Discriminative"
    elif equality_score_cell.value > 10 or -10 < equality_score_cell.value < 0:
        equality_class_cell.value = "Unfair"
    elif 0 <= equality_score_cell.value <= 10:
        equality_class_cell.value = "Fair"
    num += 1
    print(equality_score_cell.value, equality_class_cell.value)

wb.save("Task 5 Updated Equality Table (1).xlsx")
