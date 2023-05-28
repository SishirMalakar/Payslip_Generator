import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

df = pd.read_excel('excel.xlsx', header=None)

name_column = 1
esi_number_column = 2
epf_number_column = 3
epf_uan_number_column = 4
days_worked_column = 7
holidays_enjoyed_column = 8
total_days_worked_column = 9
basic_wages_amount_column = 13
nh_column = 14
total_column = 16
pf_percentage_column = 17
esi_percentage_column = 18
total_deduction_column = 21
net_amount_paid_column = 22

pdf_file_name = 'all_payslips.pdf'
pdf = canvas.Canvas(pdf_file_name, pagesize=letter)

font_size = 8
x_pos = 50
y_pos = 1000
line_height = font_size + 5
available_space = 1000

for index, row in df.iloc[5:-1].iterrows():
    employee_name = row[name_column]
    esi_number = row[esi_number_column]
    epf_number = row[epf_number_column]
    epf_uan_number = row[epf_uan_number_column]
    days_worked = row[days_worked_column]
    holidays_enjoyed = row[holidays_enjoyed_column]
    total_days_worked = row[total_days_worked_column]
    basic_wages_amount = row[basic_wages_amount_column]
    nh = row[nh_column]
    total = row[total_column]
    pf_percentage = row[pf_percentage_column]
    esi_percentage = row[esi_percentage_column]
    total_deduction = row[total_deduction_column]
    net_amount_paid = row[net_amount_paid_column]

    payslip_content = f"""
    Employee Name: {employee_name}
    ESI Number: {esi_number}
    EPF Number: {epf_number}
    EPF UAN Number: {epf_uan_number}
    Days Worked: {days_worked}
    Holidays Enjoyed: {holidays_enjoyed}
    Total Days Worked: {total_days_worked}
    Basic Wages Amount: {basic_wages_amount}
    N.H.: {nh}
    Total: {total}
    P.F. @12%: {pf_percentage}
    E.S.I. @.75%: {esi_percentage}
    Total Deduction: {total_deduction}
    Net Amount Paid: {net_amount_paid}
    ----------------------------------
    """

    lines = payslip_content.split('\n')

    if y_pos - (len(lines) * line_height) < 50:
        # Create a new page
        pdf.showPage()
        y_pos = 1000

    pdf.setFont("Helvetica", font_size)
    for line in lines:
        pdf.drawString(x_pos, y_pos, line.strip())
        y_pos -= line_height

    y_pos -= line_height
pdf.save()

print(f'All payslips generated: {pdf_file_name}')
