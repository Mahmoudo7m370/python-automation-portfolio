from openpyxl import Workbook
from openpyxl.styles import Font
wb= Workbook()
ws= wb.active
ws.title="Clients"
header=["Name","Email","Project","Amount"]
ws.append(header)
ws.append(["Ahmed","ahmed@gmail.com","Excel Automation" ,500])
ws.append(["sara","sara@gmail.com","Data Analysis",300])
ws.append(["omar","omar@gmail.com","Web Scraping",750])
ws.append(["Mona","mona@gmail.com","Report Generation",450])
ws.append(["Ali","ali@gmail.com","Dashboard",600])
for cell in ws[1]:
    cell.font=Font(bold=True)
wb.save("clients.xlsx")