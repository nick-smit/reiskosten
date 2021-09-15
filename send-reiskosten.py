from genericpath import isfile
import json
from os import unlink
import xlsxwriter
import config
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib, ssl
from os.path import basename
from sys import argv

def writeSheet(filename: str):
    workbook = xlsxwriter.Workbook(filename)
    sheet = workbook.add_worksheet()

    sheet.set_column('A:A', width=30)
    sheet.set_column('B:C', width=15)
    sheet.set_column('D:D', width=20)
    sheet.set_column('E:F', width=15)

    cell_format_header = workbook.add_format()
    cell_format_header.set_bold(True)
    cell_format_header.set_font('Calibri')
    cell_format_header.set_font_size(16)
    sheet.write('A5', 'Reiskostendeclaratie', cell_format_header)

    cell_format_name = workbook.add_format()
    cell_format_name.set_bold(True)
    cell_format_name.set_font('Calibri')
    cell_format_name.set_top()
    
    cell_format_border_right = workbook.add_format()
    cell_format_border_right.set_top()
    cell_format_border_right.set_right()

    for col in range(0, 6):
        for row in range(6, 9):
            sheet.write(row, col, '', cell_format_name)
            if col >= 2:
                sheet.write(row, col, '', cell_format_border_right)

    sheet.write('A7', f"Naam: {config.your_name}", cell_format_name)



    cell_format_table_heading = workbook.add_format()
    cell_format_table_heading.set_bold(True)
    cell_format_table_heading.set_font('Calibri')
    cell_format_table_heading.set_align('center')
    cell_format_table_heading.set_border()

    sheet.write('A10', 'Datum', cell_format_table_heading)
    sheet.write('B10', 'Van PC+Nr', cell_format_table_heading)
    sheet.write('C10', 'Naar PC+Nr', cell_format_table_heading)
    sheet.write('D10', 'Omschrijving', cell_format_table_heading)
    sheet.write('E10', 'KM\'s',cell_format_table_heading)
    sheet.write('F10', 'Bedrag', cell_format_table_heading)

    cell_format_centered = workbook.add_format()
    cell_format_centered.set_font('Calibri')
    cell_format_centered.set_align('center')
    cell_format_centered.set_border()

    cell_format_base = workbook.add_format()
    cell_format_base.set_font('Calibri')
    cell_format_base.set_border()

    today = datetime.date.today()
    curMonthFile = f"{config.base_dir}data/{today.year}-{today.month}.json"
    officeDays = []
    with open(curMonthFile, 'r') as file:
        officeDays = json.loads(file.read())

    row = 10
    total_compensation = 0
    total_kms = 0
    for officeDay in officeDays:
        sheet.write(row, 0, datetime.datetime.strptime(officeDay, '%Y-%m-%d').strftime('%d-%m-%Y'), cell_format_centered)
        sheet.write(row, 1, config.from_pht, cell_format_centered)
        sheet.write(row, 2, config.to_pht, cell_format_centered)
        sheet.write(row, 3, config.description, cell_format_base)
        sheet.write_number(row, 4, config.total_kms, cell_format_base)

        compensation_for_day = round(config.total_kms * config.compensation_per_km, 2)
        sheet.write_number(row, 5, compensation_for_day, cell_format_base)

        total_kms += config.total_kms
        total_compensation += compensation_for_day
        row += 1


    cell_format_bold = workbook.add_format()
    cell_format_bold.set_font('Calibri')
    cell_format_bold.set_bold(True)

    for col in range(0, 6):
        sheet.write(row, col, '', cell_format_base)

    sheet.write(row + 1, 3, 'Totalen:', cell_format_bold)
    sheet.write_number(row + 1, 4, total_kms, cell_format_base)
    sheet.write_number(row + 1, 5, round(total_compensation, 2), cell_format_base)

    workbook.close()

def sendMail(subject, html_body, attachments):
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(config.smtp_host, config.smtp_port, context=context, timeout=60) as server:
        server.login(config.smtp_username, config.smtp_password)

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['To'] = config.notification_email_to
        msg['From'] = config.notification_email_from
        msg.attach(MIMEText(html_body, 'html'))

        for attachment in attachments:
            with open(attachment, "rb") as file:
                part = MIMEApplication(
                    file.read(),
                    Name=basename(attachment)
                )

            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(attachment)
            msg.attach(part)

        body = msg.as_string()

        server.sendmail(config.notification_email_from, config.notification_email_to, body)
        server.quit()


if __name__ == '__main__':
    today = datetime.date.today()
    file = f"{config.base_dir}data/{today.year}-{today.month}.xlsx"
    if '--force' not in argv:
        if isfile(file):
            print("Already executed for this month")
            exit(0)
        
        if today.day < config.send_on_day_of_month:
            print("The day to send on has not come by")
            exit(0)

    try:
        writeSheet(file)

        body = f"<html><head></head><body>Beste, <br/><br/>"
        body += f"Hierbij de reiskostendeclaratie voor {today.month}-{today.year}.<br/><br />"
        body += f"Vergeet niet om deze door te sturen naar de administratie!</body></html>"

        sendMail(f"Reiskosten declaratie {today.month}-{today.year}", body, [file])
        print("Success!")
    except Exception as e:
        if isfile(file):
            unlink(file)

        raise e
