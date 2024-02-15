
import openpyxl

def extract_data_from_excel(file_path):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    
    # Prepare a list to store the extracted data
    extracted_data = []
    
    # Iterate through each row in the sheet
    for row in sheet.iter_rows(min_row=2): # assuming the first row is the header
        if row[14].value == "Yes": # column O is the 15th column, index starts at 0
            g_value = row[6].value # column G is the 7th column
            f_value = row[5].value # column F is the 6th column
            extracted_data.append((f_value, g_value))
    
    # Return the extracted data
    extracted_data.sort()
    return extracted_data

file_path = 'MASTER FLO MARCH 2024.xlsx'
data = extract_data_from_excel(file_path)

def convert(x:tuple[str]):
    out = " ".join(x[::-1])
    while "  " in out:
        out = out.replace("  "," ")
    a = "(()=>{navigator.clipboard.writeText(this.textContent);this.style=\"background-color: red;\"})()"
    return f"{out.split(' ')[-1][0]}: <button onclick='{a}'>{out}</button><br/>"

print("\n".join(convert(i) for i in data))
