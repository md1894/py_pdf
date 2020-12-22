from sys import argv
from fpdf import FPDF
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
from email import encoders 
import smtplib

# python pdf_creater.py 0_AccountStatement.LOG  :: --> to run through script

# if you created the executable file then
# ./pdf_creater 0_AccountStatement.LOG

# command to create executable
# pyinstaller --onefile -w 'filename.py'


'''
Name: fpdf
Version: 1.7.2

Name: PyPDF2
Version: 1.26.0

Name: pyinstaller
Version: 4.1
Summary: PyInstaller bundles a Python application and all its dependencies into a single package.
'''

bank_name = ''
branch_name = ''
branch_address1 = ''
branch_address2 = ''
branch_address3 = ''
branch_ph_no = ''
logo_path = ''
from_mail_id = ''
from_mail_id_pass = ''
host = ''
save_file_path = ''
is_send_pdf = ''
branch_code = ''

subject = ''
file_name = ''
from_date = ''
to_date = ''
long_name = ''
account_no = ''
mobile_no = ''
address1 = ''
address2 = ''
address3 = ''
city = ''
pin_code = ''
pan_no = ''
aadhar_no = ''
ifsc_code = ''
micr_code = ''
schm_name = ''
cata_code = ''
to_mail_id = ''
open_bal = ''
open_bal_dc = ''

tran_part = ''
tran_date = ''
chq_no = ''
debit_amount = ''
credit_amount = ''
balance = ''
tran_bal_dc = ''

tran_group = []

def get_app_info(lines):
    global logo_path
    global from_mail_id
    global from_mail_id_pass
    global host
    global save_file_path
    global is_send_pdf
    logo_path = lines[1].split('~')[1].strip()
    if len(lines[2].strip()) != 0: 
        from_mail_id = lines[2].split('~')[1].strip()
    if len(lines[3].strip()) != 0:
        from_mail_id_pass = lines[3].split('~')[1].strip()
    if len(lines[4].strip()) != 0:
        host = lines[4].split('~')[1].strip()
    save_file_path = lines[5].split('~')[1].strip()
    is_send_pdf = lines[6].split('~')[1].strip()


def get_input_params(lines):
    global account_no
    global from_date
    global to_date
    global subject
    global file_name
    account_no = lines[10].split('~')[1].strip()
    from_date = lines[11].split('~')[1].strip()
    to_date = lines[12].split('~')[1].strip()
    subject = lines[13].split('~')[1].strip()
    file_name = lines[14].split('~')[1].strip()


def get_branch_para(lines):
    global branch_code
    global bank_name
    global branch_name
    global branch_address1
    global branch_address2
    global branch_address3
    global city
    global pin_code
    global ifsc_code
    global micr_code
    global branch_ph_no
    branch_code = lines[17].split('~')[1].strip()
    bank_name = lines[18].split('~')[1].strip()
    branch_name = lines[19].split('~')[1].strip()
    branch_address1 = lines[20].split('~')[1].strip()
    branch_address2 = lines[21].split('~')[1].strip()
    branch_address3 = lines[22].split('~')[1].strip()
    city = lines[23].split('~')[1].strip()
    pin_code = lines[24].split('~')[1].strip()
    ifsc_code = lines[25].split('~')[1].strip()
    micr_code = lines[26].split('~')[1].strip()
    branch_ph_no = lines[27].split('~')[1].strip()


def get_actyps(lines):
    global schm_name
    schm_name = lines[30].split('~')[1].strip()

def get_master(lines):
    global account_no
    global long_name
    global cata_code
    global to_mail_id
    global address1
    global address2
    global address3
    global city
    global pin_code
    global mobile_no
    global pan_no
    global aadhar_no
    global open_bal
    global open_bal_dc
    account_no = lines[33].split('~')[1].strip()
    long_name = lines[34].split('~')[1].strip()
    cata_code = lines[35].split('~')[1].strip()
    to_mail_id = lines[36].split('~')[1].strip()
    address1 = lines[37].split('~')[1].strip()
    address2 = lines[38].split('~')[1].strip()
    address3 = lines[39].split('~')[1].strip()
    city = lines[40].split('~')[1].strip()
    pin_code = lines[41].split('~')[1].strip()
    mobile_no = lines[42].split('~')[1].strip()
    pan_no = lines[43].split('~')[1].strip()
    aadhar_no = lines[44].split('~')[1].strip()
    open_bal = lines[45].split('~')[1].strip()
    open_bal_dc = lines[46].split('~')[1].strip()


def get_all_tran_details(lines):
    cnt = 48

    while lines[cnt] != '</TRAN>':
        cnt = cnt + 1
        if lines[cnt] == '<DET>':
            cnt = cnt + 1
            record_list = []
            tran_part = lines[cnt].split('~')[1].strip()
            record_list.append(tran_part)
            cnt = cnt + 1
            tran_date = lines[cnt].split('~')[1].strip()
            record_list.append(tran_date)
            cnt = cnt + 1
            chq_no = lines[cnt].split('~')[1].strip()
            record_list.append(chq_no)
            cnt = cnt + 1
            debit_amount = lines[cnt].split('~')[1].strip()
            record_list.append(debit_amount)
            cnt = cnt + 1
            credit_amount = lines[cnt].split('~')[1].strip()
            record_list.append(credit_amount)
            cnt = cnt + 1
            balance = lines[cnt].split('~')[1].strip()
            record_list.append(balance)
            cnt = cnt + 1
            tran_bal_dc = lines[cnt].split('~')[1].strip()
            record_list.append(tran_bal_dc)
            tran_group.append(record_list)
            cnt = cnt + 1
            if lines[cnt+1] == '</TRAN>':
                break

filtered_lines = []

with open(argv[1],'r') as fd:
    lines = fd.readlines()
    for line in lines:
        filtered_lines.append(line.replace("\n",""))
    
    get_app_info(filtered_lines)
    get_input_params(filtered_lines)
    get_branch_para(filtered_lines)
    get_actyps(filtered_lines)
    get_master(filtered_lines)
    get_all_tran_details(filtered_lines)


def add_encryption(input_pdf, output_pdf, password):
    pdf_writer = PdfFileWriter()
    pdf_reader = PdfFileReader(input_pdf)

    for page in range(pdf_reader.getNumPages()):
        pdf_writer.addPage(pdf_reader.getPage(page))

    pdf_writer.encrypt(user_pwd=password, owner_pwd=None, use_128bit=True)
    with open(output_pdf, 'wb') as fh:
        pdf_writer.write(fh)

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

class PDF(FPDF):

    def main_header(self): 
        self.set_font('Arial', 'B', 12)
        pdf.set_text_color(153,0,76)
        self.cell(150, 5, bank_name)
        pdf.set_text_color(0,0,0)
        self.set_draw_color(153,0,76)
        self.set_line_width(0.5)
        self.line(5,36,210,36)
        self.ln(2)
        self.set_font('Arial', '', 9)
        second_line = '%s,%s,%s,%s,%s %s'%(branch_address1,branch_address2,branch_address3,city,pin_code,branch_ph_no)
        self.cell(140,10, second_line, align='L')
        if len(logo_path) != 0:
            self.image(logo_path, 180, 4, w=27,h=27)
        self.ln(20)
    
    def header(self):
        if self.page_no() != 1:
            self.set_font('Arial','B', 9)
            self.cell(150, 10, bank_name)
            self.ln(5)
            self.cell(150, 10, '%s : %s'%(long_name,account_no))
            self.set_font('Arial','B', 9)
            self.ln(15)


    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')
    
    def acc_details(self):
        self.set_font('Arial','B', 9)
        self.ln(5)
        self.cell(150, 10, long_name)
        self.ln(5)
        self.set_font('Arial','', 9)
        self.cell(90, 10, 'A/C NO. : %s'%(account_no))
        self.cell(90, 10, 'MOBILE NO : %s'%(mobile_no))
        self.ln(5)
        self.cell(90, 10, address1)
        self.cell(90, 10, 'PAN NO : %s'%(pan_no))
        self.ln(5)
        self.cell(90, 10, address2)
        self.cell(90, 10, 'AADHAR NO : %s'%(aadhar_no))
        self.ln(5)
        self.cell(90, 10, address3)
        self.cell(90, 10, 'IFSC CODE : %s'%(ifsc_code))
        self.ln(5)
        self.cell(90, 10, '%s %s'%(city,pin_code))
        self.cell(90, 10, 'MICR CODE : %s'%(micr_code))
        self.ln(5)
        self.cell(90, 10, 'A/C TYPE : %s'%(schm_name))
        self.cell(90, 10, 'CATEGORY : %s'%(cata_code))

pdf = PDF(format='letter')
pdf.alias_nb_pages()
pdf.add_page()
pdf.main_header()
 
epw = pdf.w - 2*pdf.l_margin
col_width = epw/4

pdf.acc_details()
pdf.ln(10)
pdf.line(5,77,210,77)
th = pdf.font_size
pdf.set_font('Arial', 'B', 11)
pdf.set_text_color(153,0,76)
# account_no[-4:] # if you want to show only last 4 characters
stmt_line = 'Detailed Statement for %s between %s to %s'%(account_no, from_date, to_date)
pdf.cell(80, 10, stmt_line)
pdf.ln(10)
cnt = 0
pdf.set_font('Arial', '', 10)
pdf.set_draw_color(0,0,0)
pdf.set_text_color(0,0,0)
pdf.set_line_width(0.2)
for row in tran_group:
    if cnt == 0 or cnt == 1 or cnt == len(tran_group)-1:
        pdf.set_font('Arial', 'B', 9)
    else:
        pdf.set_font('Arial', '', 8)
    
    if cnt == 0:
        pdf.set_fill_color(211,211,211)
    elif cnt == 1 or cnt == len(tran_group)-1:
        pdf.set_fill_color(221, 221, 221)
    else:
        pdf.set_fill_color(255, 255, 255)

    pdf.cell(20, 2*th, str(tran_group[cnt][1]), border=1, fill = 1, align='C')
    pdf.cell(70, 2*th, str(tran_group[cnt][0]), border=1, align = 'L' if cnt != 0 else 'C', fill = 1)
    pdf.cell(18, 2*th, str(tran_group[cnt][2]), border=1, fill = 1)
    pdf.cell(28, 2*th, str(tran_group[cnt][3]), border=1, align ='R' if cnt != 0 else 'C', fill = 1)
    pdf.cell(28, 2*th, str(tran_group[cnt][4]), border=1, align ='R' if cnt != 0 else 'C', fill = 1)
    pdf.cell(28, 2*th, str(tran_group[cnt][5]), border=1, align ='R' if cnt != 0 else 'C', fill = 1)
    pdf.cell(8, 2*th, str(tran_group[cnt][6]), border=1, fill = 1)
    pdf.ln(2*th)
    cnt = cnt + 1

pdf.ln(10)
pdf.set_font('Arial', 'B', 10)
pdf.cell(80, 10, 'Disclaimer')
pdf.ln(5)
pdf.set_font('Arial', '', 10)
pdf.cell(150, 10,'This is a system generated statement. please contact your nearest branch for further details/any queries')

fname = '%s.pdf'%(file_name)
# CHANDWAD_0002052204634_2020/12/16_2020/12/16.pdf --> cant create pdf file due to '/'
# CHANDWAD_0002052204634_20201216_20201216.pdf   --> correct file name
qualified_fname = '%s_%s'%(save_file_path,fname) # this is a temporary file
qualified_fname_ = '%s%s'%(save_file_path,fname)
pdf.output(qualified_fname, 'F')
passwd = long_name.strip()[:4] + account_no[-4:]
add_encryption(input_pdf=qualified_fname, output_pdf=qualified_fname_, password=passwd)
os.system('rm -rf %s'%(qualified_fname)) # remove temporary file

# gmail_password ==>> 1234*abcd
# is_send_pdf = 'TRUE'
if is_send_pdf == 'TRUE':
    MY_ADDRESS = from_mail_id
    PASSWORD = from_mail_id_pass
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)
    msg = MIMEMultipart()
    msg['From']=MY_ADDRESS
    msg['To'] = 'mehul.dubey@kimayainfotech.com'
    msg['Subject'] = subject
    attachment = open(qualified_fname_, "rb")
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    attachment.close()
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % fname)
    msg.attach(p)
    message_template = read_template('template_file.txt')
    message = message_template.substitute(LONG_NAME=long_name, SUBJECT=subject)
    msg.attach(MIMEText(message, 'plain'))
    s.send_message(msg)
    del msg
    s.quit()