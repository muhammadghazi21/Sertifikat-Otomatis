from docxtpl import DocxTemplate
from docx2pdf import convert
import pandas as pd
import os, openpyxl, datetime, time
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Tugas akhir MK Scripting Language
# Oleh Muhammad Ghazi Alghifari dan kawan2

isUpdated = False

def generateSertifikat(context, templatefile, activityName):
    doc = DocxTemplate(templatefile)
    doc.render(context)
    
    if not os.path.exists("files_docx"):
        os.makedirs("files_docx")
    doc.save("files_docx/(halaman depan) Sertifikat {} {}.docx".format(activityName, context["Nama"]))

def generateWithSchaduling(datafile, templatefile, isSendEmail, SenderEmail, SenderPassword, activityName, emailMessage):
    excel = openpyxl.load_workbook(datafile)
    sheet = excel.active
    row = sheet.max_row
    col = sheet.max_column
    
    if sheet.cell(row=1, column=col).value == "Status":
        for i in range(2, row+1):
            #cek apa sudah pernah generate sertifikat
            if sheet.cell(row=i, column=col).value is None and sheet.cell(row=i, column=1).value is not None:
                context = {
                    "No": sheet.cell(row=i, column=1).value,
                    "Sebagai": sheet.cell(row=i, column=2).value,
                    "Nama": sheet.cell(row=i, column=3).value,
                    "Email": sheet.cell(row=i, column=4).value,
                }
                print(context)
                
                generateSertifikat(context, templatefile, activityName)
            else:
                pass

        if not os.path.exists("files_pdf"):
            os.makedirs("files_pdf")
        convert("files_docx", "files_pdf")

    
        for i in range(2, row+1):
            if sheet.cell(row=i, column=col).value is None and sheet.cell(row=i, column=1).value is not None:
                if isSendEmail == "y":
                    sendEmail(SenderEmail, SenderPassword, context["Email"], "Sertifikat {}".format(activityName), emailMessage, "files_pdf/(halaman depan) Sertifikat {} {}.pdf".format(activityName, context["Nama"]))
                    sheet.cell(row=i, column=col).value = "Sudah pada {}".format(datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S"))
                elif isSendEmail == "n":
                    pass

                excel.save(datafile)
            else:
                pass

    else:
        sheet.cell(row=1, column=col+1).value = "Status"
        excel.save(datafile)
        generateWithSchaduling(datafile, templatefile, isSendEmail, SenderEmail, SenderPassword, activityName, emailMessage)

def sendEmail(fromEmail, password, toEmail, subject, message, filename):
    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'plain'))

    attachment = open(filename, 'rb')

    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(p)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(fromEmail, password)
    text = msg.as_string()
    s.sendmail(fromEmail, toEmail, text)
    s.quit()

def checkUpdate(datafile):
    excel = openpyxl.load_workbook(datafile)
    sheet = excel.active
    row = sheet.max_row
    col = sheet.max_column
    
    if sheet.cell(row=1, column=col).value == "Status":
        for i in range(2, row+1):
            if sheet.cell(row=i, column=col).value is None and sheet.cell(row=i, column=1).value is not None:
                isUpdated = False
                break
            else:
                isUpdated = True

        if isUpdated:
            print("\n(( Data sudah paling update ))")
            pilihanLanjut = input("Masih ingin melanjutkan? [y/n]: ")
            if pilihanLanjut == "y":
                pass
            elif pilihanLanjut == "n":
                exit()
            else:
                print("Pilihan tidak tersedia")
                main()
        else:
            print("\n((...Data dalam proses update...))")
    else:
        pass

def inputData():
    datafile = input("\n1. masukkan nama file data peserta dengan format .xlsx\n>>> ")
    datafile = datafile + ".xlsx"
    datafile = str(datafile)
    checkUpdate(datafile)
    if isUpdated:
        print("Data sudah terupdate")
        exit()
    else:
        pass

    templatefile = input("2. masukkan nama file template sertifikat dengan format .docx\n>>> ")
    templatefile = templatefile + ".docx"
    templatefile = str(templatefile)

    SenderEmail = ""
    SenderPassword = ""

    def everyOneHour():
        isEveryOneHour = input("Apakah ingin mengulanginya tiap 1 jam? ([y/n])\n>>> ")
        if isEveryOneHour == "y":
            while True:
                generateWithSchaduling(datafile, templatefile, isSendEmail, SenderEmail, SenderPassword, activityName, emailMessage)
                print("Sertifikat berhasil di generate dan dikirim ke email peserta")
                print("\n  > tekan ctrl + c untuk menghentikan program  <\n")
                time.sleep(10)
        elif isEveryOneHour == "n":
            generateWithSchaduling(datafile, templatefile, isSendEmail, SenderEmail, SenderPassword, activityName, emailMessage)
        else:
            print("input tidak valid")
            everyOneHour()

    isSendEmail = input("3. Apakah sekalian ingin mengirimkan sertifikat melalui email? ([y/n])\n>>> ")
    if isSendEmail == "y":
        SenderEmail = input("4. masukkan email pengirim\n>>> ")
        SenderPassword = input("5. masukkan password email pengirim\n>>> ")
        SenderPassword = "makkrnketqostqcx"
        activityName = input("6. masukkan nama kegiatan\n>>> ")
        emailMessage = input("7. masukkan isi pesan email\n>>> ")    
        everyOneHour()
        print("Sertifikat berhasil di generate dan dikirim ke email peserta")

    elif isSendEmail == "n":
        everyOneHour()
        print("Sertifikat berhasil di generate")
        main()
    else:
        print("Pilihan tidak tersedia")
        main()


def main():
    print("\n------------------------------------------------------------------------------------------------------")
    print("Selamat datang di program generate sertifikat")
    print("------------------------------------------------------------------------------------------------------")
    print("??_Anda dapat menggunakan program ini jika semua berkas dan keperluan setup yang dibutuhkan sudah sesuai (Baca info.html)_??")
    menu = input("Menu:\n1. Generate Sertifikat\n2. Exit\nPilih menu: ")

    if menu == "1":
        inputData()

    elif menu == "2":
        exit()
    else:
        print("Pilihan tidak tersedia")
        main()

if __name__ == "__main__":
    main()