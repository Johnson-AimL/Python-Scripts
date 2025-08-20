# This is a software which Identifies the duplicate in a question and remove it.
# Designed as per Customer Needs
# Licenced for KK Associates
# Developed by WJ Tech Solutions

import os
import wx
import pandas as pd
import datetime


class mainApp(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(mainApp, self).__init__(*args, **kwargs)
        self.InitUI()

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

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add((-1, 40))

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        stock_label = wx.StaticText(self, label='Select Input folder')
        stock_label.SetFont(font)
        hbox1.Add(stock_label, flag=wx.LEFT | wx.RIGHT, border=10)
        vbox.Add(hbox1, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.file1 = wx.FilePickerCtrl(self, style=wx.FLP_USE_TEXTCTRL)
        self.file1.SetFont(font)
        hbox2.Add(self.file1, proportion=1, flag=wx.EXPAND)
        vbox.Add(hbox2, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        vbox.Add((-1, 200))

        hbox7 = wx.BoxSizer(wx.HORIZONTAL)
        submit_button = wx.Button(self, label='Submit')
        hbox7.Add(submit_button, proportion=1, flag=wx.EXPAND)
        vbox.Add(hbox7, flag=wx.LEFT | wx.BOTTOM, border=10)

        self.Bind(wx.EVT_BUTTON, self.on_submit, submit_button)

        vbox.Add((-1, 20))

        font = wx.Font(10, wx.DECORATIVE, wx.ITALIC, wx.NORMAL)

        hbox8 = wx.BoxSizer(wx.HORIZONTAL)
        self.label = wx.StaticText(self, label="Welcome  - Duplicate removing Program V1.0")
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

        self.update_progress(0)

        self.SetSize((700, 500))
        self.SetTitle('Duplicate removing Program V2.1')
        self.Centre()

        self.Show()

    def load_questions(self, file_path):
        questions = []
        with open(file_path, 'r') as file:
            lines = file.readlines()

        i = 0
        while i < len(lines):
            if lines[i].strip() == "":
                i += 1
                continue

            question = lines[i].strip()
            options = []
            for j in range(1, 5):
                option_line = lines[i + j].strip()
                if option_line.startswith("A)") or option_line.startswith("A."):
                    options.append(option_line[:].strip())
                elif option_line.startswith("B)") or option_line.startswith("B."):
                    options.append(option_line[:].strip())
                elif option_line.startswith("C)") or option_line.startswith("C."):
                    options.append(option_line[:].strip())
                elif option_line.startswith("D)") or option_line.startswith("D."):
                    options.append(option_line[:].strip())
                else:
                    raise ValueError(f"Invalid option format in line {i + j + 1}")

            answer_line = lines[i + 5].strip()
            if answer_line.startswith("ANSWER:"):
                answer = answer_line[:].strip()
            else:
                raise ValueError(f"Invalid answer format in line {i + 6}")

            question_data = {
                'question': question,
                'options': options,
                'answer': answer
            }
            questions.append(question_data)
            i += 6

        return questions

    def remove_duplicates(self, questions):
        unique_questions = []
        seen_questions = set()

        for question in questions:
            question_text = question['question']
            if question_text not in seen_questions:
                seen_questions.add(question_text)
                unique_questions.append(question)

        return unique_questions

    def write_questions_to_file(self, questions, file_path):
        with open(file_path, 'w') as file:
            for question in questions:
                file.write(question['question'] + '\n')
                for option in question['options']:
                    file.write(option + '\n')
                file.write(question['answer'] + '\n\n')

    def OnQuit(self, e):
        self.Close()

    def prefix_files_in_folder(self, folder_path, prefix):
        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        for file in files:
            if file.endswith(".txt"):
                original_file_path = os.path.join(folder_path, file)
                new_file_name = f"{prefix}_{file}"
                new_file_path = os.path.join(folder_path, new_file_name)
                loaded_questions = self.load_questions(original_file_path)
                unique_questions = self.remove_duplicates(loaded_questions)
                self.write_questions_to_file(unique_questions, new_file_path)

                for idx, question_data in enumerate(loaded_questions):
                    print(f"Question {idx + 1}: {question_data['question']}")
                    print("Options:")
                    for option in question_data['options']:
                        print(f"- {option}")
                    print(f"Answer: {question_data['answer']}\n")

        question_text_list = [question['question'] for question in loaded_questions]
        duplicate_questions = [question for question in question_text_list if question_text_list.count(question) > 1]

        df_duplicates = pd.DataFrame({'Question': duplicate_questions})
        print(df_duplicates)

        #           with open(original_file_path, 'r') as original_file:

    #               content = original_file.read()

    #          with open(new_file_path, 'w') as new_file:
    #                   new_file.write(content)
    #
    #              print(f"Created file: {new_file_path}")

    def on_submit(self, event):
        self.label.SetLabel("Processing Started...   Please wait for the update")
        self.inputFile_path = self.file1.GetPath()

        self.directory = os.path.dirname(self.inputFile_path)
        print("Directory:", self.directory)

        self.update_progress(20)
        self.prefix_files_in_folder(self.directory, "asa")

        self.label.SetLabel("writing output, Please wait.")
        self.update_progress(51)
        self.Layout()

        #  self.update_progress(90)
        self.Layout()

        print("Button clicked")
        self.update_progress(100)

        self.label.SetLabel("Processing Completed.Result files are available in " + self.directory)

    def update_progress(self, value):
        self.gauge.SetValue(value)


def main():
    app = wx.App()
    ex = mainApp(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()