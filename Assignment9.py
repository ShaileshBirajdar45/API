import xml.etree.ElementTree as ET
from openpyxl import Workbook

tree = ET.parse('Input.xml')
root = tree.getroot()

workbook = Workbook()
worksheet = workbook.active

headers = []

for voucher in root.iter('VOUCHER'):
    voucher_type = voucher.find('VOUCHERTYPENAME').text
    if voucher_type == "Receipt":
        transaction_data = []
        for element in voucher:
            element_name = element.tag
            element_text = element.text if element.text else''
            transaction_data.append((element_name, element_text))

        if not headers:
            headers = [header[0] for header in transaction_data]
            worksheet.append(headers)

        values = [item[1] for item in transaction_data]
        worksheet.append(values)

workbook.save('receipt_transactions.xlsx')