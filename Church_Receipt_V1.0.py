# This is a sample Python script.


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
import sys
import pandas as pd
from num2words import num2words
from mailmerge import MailMerge
from datetime import datetime
from docx2pdf import convert
from PyPDF2 import PdfMerger, PdfReader, PdfWriter

from reportlab.pdfgen import canvas
import wx
import json
import datetime


class mainApp(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(mainApp, self).__init__(*args, **kwargs)
        self.pdf_folder = None
        self.Json_file_path = None
        self.Json_file_path2 = None
        self.InitUI()

    def init_steps(self):

        # Construct the path to the file located in the same folder as the executable
        self.Json_file_path = os.path.join(self.exe_folder_path, 'data_requirements\\dbhistory.txt')
        self.Json_file_path2 = os.path.join(self.exe_folder_path, 'data_requirements\\lastbillno.txt')
        self.pdf_folder = self.exe_folder_path + '\\Output\\'
        print(self.Json_file_path)
        print(self.Json_file_path2)
        print(self.pdf_folder)

    def on_submit(self, event):
        self.label.SetLabel("Processing Started...   Please wait for the update")
        self.init_steps()

        # calling a method
        mergePdf(self)
        printDoc(self)
        self.label.SetLabel("Processing Completed...   Please verify the results")

    def OnQuit(self, e):
        self.Close()

    def get_max_billno_from_json(self, event, month, year):

        # read JSON file into dictionary
        with open(self.Json_file_path2, 'r') as f:
            bill_numbers = json.load(f)

        # get month and year for the entry

        year = year
        month = month

        # get last bill number for current month and year
        if year in bill_numbers and month in bill_numbers[year]:
            last_bill_no = bill_numbers[year][month]['last_bill_no']
        else:
            # start new series with starting bill number of 0
            last_bill_no = int(year[2:] + month + '00')

        # increment last bill number and save updated dictionary to JSON file
        self.new_bill_no = last_bill_no + 1
        bill_numbers[year] = bill_numbers.get(year, {})
        bill_numbers[year][month] = {'last_bill_no': self.new_bill_no}
        with open(self.Json_file_path2, 'w') as f:
            json.dump(bill_numbers, f)

    def create_new_data(self, billno, billdate, title, name, address, rs, mode_of_pymt, ch_no, ch_date, approver,
                        checked_by):
        # Get current datetime in ISO 8601 format
        current_time = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + 'Z'

        # Create new data object
        new_data = {
            "billno": billno,
            "BillDate": billdate,
            "Title": title,
            "Name": name,
            "Address": address,
            "Rs": rs,
            "ModeOfPymt": mode_of_pymt,
            "ChNo": ch_no,
            "Chdate": ch_date,
            "Approver": approver,
            "CheckedBy": checked_by,
            "created_on": current_time,
            "last_updated_on": current_time,
            "updated_by": ""
        }

        return new_data

    def update_or_insert_json_data(self, new_data):
        with open(self.Json_file_path, 'r+') as f:
            data = json.load(f)
            for item in data:
                if item['billno'] == new_data['billno']:
                    # update existing item
                    item.update(new_data)
                    item['last_updated_on'] = datetime.datetime.utcnow().isoformat() + 'Z'
                    break
            else:
                # insert new item
                new_data['created_on'] = datetime.datetime.utcnow().isoformat() + 'Z'
                new_data['last_updated_on'] = new_data['created_on']
                data.append(new_data)
            # move file pointer to the beginning of the file
            f.seek(0)
            # write the updated/inserted data back to the file
            json.dump(data, f, indent=4)
            # truncate the remaining content if any
            f.truncate()

    def InitUI(self):
        #   self.SetBackgroundColour(wx.BLUE)
        self.SetBackgroundColour("#625AB6")
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
        menubar.Append(fileMenu, '&File')
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.OnQuit, fileItem)

        font = wx.Font(pointSize=16, family=wx.FONTFAMILY_DEFAULT, style=wx.FONTSTYLE_NORMAL,
                       weight=wx.FONTWEIGHT_NORMAL)

        # Get the path to the executable file
        exe_file_path = sys.argv[0]

        # Get the path to the directory containing the executable file
        self.exe_folder_path, _ = os.path.split(os.path.abspath(exe_file_path))
        _, exe_ext = os.path.splitext(exe_file_path)

        if not exe_ext:
            # If the file does not have an extension, use the directory component as the executable filename
            exe_file_path = self.exe_folder_path
            self.exe_folder_path, _ = os.path.split(exe_folder_path)

            print(self.exe_folder_path)

        # Get the path to the directory containing the executable file
        #   self.exe_folder_path = os.path.dirname(os.path.abspath(exe_file_path))

        # Get the path to the folder containing the executable file
        #     self.exe_folder_path = os.path.dirname(os.path.abspath(__file__))

        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        vboxh11 = wx.BoxSizer(wx.VERTICAL)
        vboxh12 = wx.BoxSizer(wx.VERTICAL)
        Address_label1 = wx.StaticText(self, label='VAILANKANNI ANNAI CHURCH')
        Address_label1.SetFont(font)
        Address_label2 = wx.StaticText(self, label='RENOVATION / RECONSTRUCTION COMMITTEE')
        Address_label2.SetFont(font)
        Address_label3 = wx.StaticText(self, label='Annai Vailankanni Church,')
        Address_label3.SetFont(font)
        Address_label4 = wx.StaticText(self, label='State Bank Colony, Tuticorin – 628002')
        Address_label4.SetFont(font)
        Address_label5 = wx.StaticText(self, label='Donation Receipt form ')
        Address_label5.SetFont(font)
        vboxh11.Add(Address_label1, flag=wx.LEFT | wx.RIGHT, border=5)
        vboxh11.Add((-1, 20))
        vboxh11.Add(Address_label2, flag=wx.LEFT | wx.RIGHT, border=5)
        vboxh11.Add((-1, 20))
        vboxh11.Add(Address_label3, flag=wx.LEFT | wx.RIGHT, border=5)
        vboxh11.Add((-1, 5))
        vboxh11.Add(Address_label4, flag=wx.LEFT | wx.RIGHT, border=10)
        vboxh11.Add((-1, 5))
        vboxh11.Add(Address_label5, flag=wx.LEFT | wx.RIGHT, border=10)
        vboxh11.Add((-1, 20))
        vboxh11.Add((-1, 20))

        image = wx.StaticBitmap(self, wx.ID_ANY,
                                wx.Bitmap(
                                    os.path.join(self.exe_folder_path, 'data_requirements\\veilankanni_pic_small.jpg'),
                                    wx.BITMAP_TYPE_ANY))
        image.SetMinSize(wx.Size(180, 350))  # Set the minimum size to 20 x 10 pixels

        vboxh12.Add(image, proportion=0, flag=wx.ALL | wx.ALIGN_LEFT, border=5)  # Use ALIGN_RIGHT and proportion=0

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.file1 = wx.FilePickerCtrl(self, style=wx.FLP_USE_TEXTCTRL)
        self.file1.SetFont(font)
        hbox2.Add(self.file1, proportion=1, flag=wx.EXPAND)
        vboxh11.Add(hbox2, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        vboxh11.Add((-1, 20))
        vboxh11.Add((-1, 20))

        hbox7 = wx.BoxSizer(wx.HORIZONTAL)
        submit_button = wx.Button(self, label='Submit')
        hbox7.Add(submit_button, proportion=1, flag=wx.EXPAND)
        vboxh11.Add(hbox7, flag=wx.LEFT | wx.BOTTOM, border=10)

        self.Bind(wx.EVT_BUTTON, self.on_submit, submit_button)

        hbox1.Add(vboxh11)
        hbox1.Add(vboxh12)

        vbox.Add(hbox1, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        vbox.Add((-1, 20))

        font = wx.Font(10, wx.DECORATIVE, wx.ITALIC, wx.NORMAL)

        hbox8 = wx.BoxSizer(wx.HORIZONTAL)
        self.label = wx.StaticText(self, label="Welcome, Annai Veilankanni church receipt V1.0")
        self.label.SetMinSize((400, 30))  # set the minimum size of the text box
        self.label.SetFont(font)
        #     self.label.SetBackgroundColour("#FFFFFF")

        hbox8.Add(self.label, proportion=1, flag=wx.EXPAND)
        vbox.Add(hbox8, flag=wx.LEFT | wx.BOTTOM, border=10)

        # vbox.AddSpacer(100)  # Add some padding to the top

        self.gauge = wx.Gauge(self, range=100, size=(-1, 25), style=wx.GA_HORIZONTAL)
        self.gauge.SetForegroundColour(wx.Colour(0, 255, 0))  # Green color
        vbox.Add(self.gauge, proportion=0, flag=wx.EXPAND)

        self.SetSizer(vbox)

        #        self.update_progress(0)

        self.SetSize((700, 500))
        self.SetTitle('Annai Veilankanni receipt Program V1.0')
        self.Centre()

        self.Show()


def mergePdf(self):
    self.label.SetLabel("Merging started...   Please wait for the update")
    # Load Excel data into a pandas dataframe
    data = pd.read_excel(self.file1.GetPath())
    # data = pd.read_excel('C:\\Users\\Johnson\\Desktop\\receipt_data.xlsx')

    # Open the Word document template

    template = os.path.join(self.exe_folder_path, 'data_requirements\\Church_Receipt_template.docx')
    document = MailMerge(template)

    # Merge data into each document and save as a PDF
    pdf_files = []

    # Create a PdfMerger object
    merger = PdfMerger()

    # Perform mail merge
    for i, row in data.iterrows():
        document = MailMerge(template)
        self.label.SetLabel(f"Processing receipt {i}...   Please wait for the update")

        # Parse the bill date string to a datetime object
        bill_date = datetime.datetime.strptime((str(row['BillDate'])), '%Y-%m-%d %H:%M:%S')

        # Convert the datetime object to the desired output format
        bill_date_str = bill_date.strftime('%d-%b-%Y')
        bill_date_year = bill_date.strftime('%Y')
        bill_date_month = bill_date.strftime('%m')

        cheq_date_str = ''
        modeOfPymt = str(row['ModeOfPymt'])

        if modeOfPymt == 'Cheque':
            # Parse the cheque date string to a datetime object
            cheq_date = datetime.datetime.strptime((str(row['Chdate'])), '%Y-%m-%d %H:%M:%S')

            # Convert the datetime object to the desired output format
            cheq_date_str = "Cheque Date: " + cheq_date.strftime('%d-%b-%Y')

            cheq_no = "Cheque No: " + str(row['ChNo'])
        else:
            cheq_date_str = ''
            cheq_no = ""

        title_st = str(row['Title'])
        name_st = str(row['Name'])
        address_st = str(row['Address'])

        # format the rupees to two decimal.
        num_str = "{:.2f}".format(row['Rs'])

        rs_full = num_str.split('.')

        amtwords_st = f"Rupees {num2words(rs_full[0], lang='en_IN', to='cardinal')} "
        amtwords_st = amtwords_st.replace(',', '')

        if (rs_full[1]) != "00":
            amtwords_st = amtwords_st + f" and {num2words(rs_full[1], lang='en_IN', to='cardinal')} Paise"
        else:
            amtwords_st = amtwords_st + ""

        rsinnum_st = num_str
        modeofpymt_st = str(row['ModeOfPymt'])
        approver_st = str(row['Approver'])
        checkedby_st = str(row['CheckedBy'])

        self.get_max_billno_from_json(self, bill_date_month, bill_date_year)

        billno_st = str(self.new_bill_no)
        formatted_billno = billno_st[:4] + '/' + billno_st[4:]

        document.merge(BillNo=formatted_billno, BillDate=bill_date_str, Title=title_st,
                       Name=name_st, Address=address_st, amtinwords=amtwords_st, RsinNum=rsinnum_st,
                       ModeOfPymt=modeofpymt_st, ChNo=cheq_no, chadate=cheq_date_str,
                       ApprovedBy=approver_st, CheckedBy=checkedby_st)

        # Save the merged document as a PDF
        pdf_file = self.pdf_folder + f"\\receipt_{i + 1}.docx"

        document.write(pdf_file)
        convert(pdf_file)
        pdf_files.append(pdf_file)

        # create the Json file update call
        data_json = self.create_new_data(formatted_billno, bill_date_str, title_st, name_st, address_st,
                                         rsinnum_st, modeofpymt_st, cheq_no, cheq_date_str, approver_st, checkedby_st)

        self.update_or_insert_json_data(data_json)

        # Delete the Word document
        os.remove(pdf_file.replace(".pdf", ".docx"))


def printDoc(self):
    # Open the input PDF files and add their pages to the merger object
    input_files = []

    self.label.SetLabel(f"Processing print file...   Please wait for the update")

    for filename in os.listdir(self.pdf_folder):
        if filename.endswith('.pdf'):
            full_path = os.path.join(self.pdf_folder, filename)
            input_files.append(full_path)

    # Define the output PDF file path and name
    output_path = os.path.join(self.exe_folder_path, 'Output\\print\\Consolidated_print.pdf')

    # Create a PDF file merger object
    merger = PdfMerger()

    for input_file in input_files:
        with open(input_file, 'rb') as f:
            reader = PdfReader(f)
            merger.append(reader)

    # Create a new PDF file with A4 size for printing
    output = PdfWriter()

    c = canvas.Canvas("temp.pdf", pagesize=(595.27, 841.89))
    c.save()
    with open("temp.pdf", 'rb') as f:
        reader = PdfReader(f)
        if len(reader.pages) >= 1:
            output.addPage(reader.pages[0])

    # Merge the PDF files into the output PDF file
    with open(output_path, 'wb') as f:
        merger.write(f)

    # Remove the temporary PDF file
    os.remove("temp.pdf")


def main():
    app = wx.App()
    ex = mainApp(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
