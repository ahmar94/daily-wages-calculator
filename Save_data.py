import xlwings as xw
from tkinter import Tk, filedialog

workbook_path = r"F:\Daily Wages Calculator\Daily Wages Calculator.xlsm"

# Ranges
Main_sheet = ["B12:AB33"]
Total_wages = ["C7:AA29"]

# ---------------------------------------------------
# SAVE FUNCTION (for Excel button)
# ---------------------------------------------------

def save_sheet1_data():
    Tk().withdraw()
    wb = xw.Book(workbook_path)
    sheet1 = wb.sheets[0]
    sheet3 = wb.sheets[2]
    
    file_path = filedialog.asksaveasfilename(
        initialdir=r"F:\Daily Wages Calculator",
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        title="Save Excel Data As Text File"
    )
    if not file_path:
        sheet1.range("I36").value = "❌ Save cancelled."
        return

    data_lines = []

    # --- MAIN SHEET ---
    for item in ",".join(Main_sheet).split(","):
        for cell in sheet1.range(item):
            if cell.value is None and not cell.formula:
                continue
            address = cell.address
            if cell.formula:
                data_lines.append(f"Main\tFormula\t{address}\t{cell.formula}")
            else:
                data_lines.append(f"Main\tValue\t{address}\t{cell.value}")

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(data_lines))

    sheet1.range("I36").value = f"✅ Data saved to:\n{file_path}" 

# ---------------------------------------------------
# LOAD FUNCTION (for Excel button)
# ---------------------------------------------------
def load_sheet1_data():
    Tk().withdraw()
    wb = xw.Book(workbook_path)
    sheet1 = wb.sheets[0]
    sheet3 = wb.sheets[2]

    sheet1.range("E14:E31,G14:H31,J14:K31,N14:N31,Q14:Q31,S14:S31,U14:U31,AA14:AA31").value = None


    txt_file_path = filedialog.askopenfilename(
        initialdir=r"F:\Daily Wages Calculator",
        title="Select Text File to Load",
        filetypes=[("Text Files", "*.txt")]
    )
    if not txt_file_path:
        sheet1.range("I36").value = "❌ No file selected."
        return

    with open(txt_file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        parts = line.strip().split("\t")
        if len(parts) < 4:
            continue
        sheet_name, data_type, address, value = parts[0], parts[1], parts[2], "\t".join(parts[3:])
        if value.strip().lower() == "none":
            value = ""

        if sheet_name.lower() == "main":
            target_sheet = sheet1 
        else:
            continue

        if data_type == "Formula":
            target_sheet.range(address).formula = value
        else:
            target_sheet.range(address).value = value

    sheet1.range("I36").value = f"✅ Data loaded from:\n{txt_file_path}"

# ---------------------------------------------------
# SAVE FUNCTION (for Excel button)
# ---------------------------------------------------

def save_sheet2_data():
    Tk().withdraw()
    wb = xw.Book(workbook_path)
    sheet3 = wb.sheets[2]
    
    file_path = filedialog.asksaveasfilename(
        initialdir=r"F:\Daily Wages Calculator",
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        title="Save Excel Data As Text File"
    )
    if not file_path:
        sheet3.range("J36").value = "❌ Save cancelled."
        return

    data_lines = []

    # --- MAIN SHEET ---
    for item in ",".join(Total_wages).split(","):
        for cell in sheet3.range(item):
            if cell.value is None and not cell.formula:
                continue
            address = cell.address
            if cell.formula:
                data_lines.append(f"Total\tFormula\t{address}\t{cell.formula}")
            else:
                data_lines.append(f"Total\tValue\t{address}\t{cell.value}")

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(data_lines))

    sheet3.range("J36").value = f"✅ Data saved to:\n{file_path}"

# ---------------------------------------------------
# LOAD FUNCTION (for Excel button)
# ---------------------------------------------------
def load_sheet2_data():
    Tk().withdraw()
    wb = xw.Book(workbook_path)
    sheet3 = wb.sheets[2]

    sheet3.range("C7:AA29").value = None


    txt_file_path = filedialog.askopenfilename(
        initialdir=r"F:\Daily Wages Calculator",
        title="Select Text File to Load",
        filetypes=[("Text Files", "*.txt")]
    )
    if not txt_file_path:
        sheet3.range("J36").value = "❌ No file selected."
        return

    with open(txt_file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        parts = line.strip().split("\t")
        if len(parts) < 4:
            continue
        sheet_name, data_type, address, value = parts[0], parts[1], parts[2], "\t".join(parts[3:])
        if value.strip().lower() == "none":
            value = ""

        if sheet_name.lower() == "total":
            target_sheet = sheet3 
        else:
            continue

        if data_type == "Formula":
            target_sheet.range(address).formula = value
        else:
            target_sheet.range(address).value = value

    sheet3.range("J36").value = f"✅ Data loaded from:\n{txt_file_path}"