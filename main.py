import openpyxl as xl
import datetime

xlsx = xl.load_workbook('format.xlsx')
sheet = xlsx.worksheets[0]

today = datetime.date.today()
today = today.strftime('%d/%m/%Y')
print(today)

company = input("Enter Company Name: ")
company_address = input("Enter Address of Company: ")
company_gst = input("Enter Company's GST: ")

bill_supply = input("Are Billing & Shipping Address same? (y/n)")

if bill_supply == 'y':
    sheet['F11'] = company
    sheet['F12'] = company_address
    sheet['F14'] = f'GSTIN: {company_gst}'
else:
    company1 = input("Enter Company Name: ")
    company_address1 = input("Enter Address of Company: ")
    company_gst1 = input("Enter Company's GST: ")
    
    sheet['F11'] = company1
    sheet['F12'] = company_address1
    sheet['F14'] = f'GSTIN: {company_gst1}'

date = input("Date (press 'y' for today): ")

if date == 'y':
    sheet['H3'] = today
else:
    today = input('Enter Date in DD/MM/YYYY format: ')
    sheet['H3'] = today

invoice_no = int(input("Invoice No: "))
frieght_charges = int(input("Freight Charges: "))

item_num = int(input("How many items do you want to add? "))

for j in range(18, (item_num+18)):
    item_code = int(input("Item Code: "))
    des = input("Description: ")
    hsn = input("Enter HSN/SAC: ")
    qty = int(input("Enter Quantity: "))
    rate = int(input("Rate: "))
    sheet[f'B{j}'] = item_code
    sheet[f'C{j}'] = des
    sheet[f'E{j}'] = hsn
    sheet[f'F{j}'] = qty
    sheet[f'G{j}'] = rate
    sheet[f'H{j}'] = rate * qty
    



sheet['H2'] = invoice_no

sheet['A11'] = company
sheet['A12'] = company_address
sheet['A14'] = f'GSTIN: {company_gst}'
sheet['H32'] = frieght_charges


xlsx.save(f'{company}.xlsx')
