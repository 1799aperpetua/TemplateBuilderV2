import customtkinter
from tkinter import filedialog
from tkinter.filedialog import asksaveasfilename
import os
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
import math
from version2 import captureTemplateData, buildTemplate

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("850x450")
        self.title("Template Builder")
        self.minsize(800, 300)

        # Create a 1x5 grid
        #self.grid_rowconfigure(0, weight = 1)
        self.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight = 1)

        # Textbox: Display text to the user
        self.table_space = customtkinter.CTkTextbox(master=self, bg_color="white")
        self.table_space.grid(row = 0, column = 2, columnspan = 7, rowspan = 10, padx = 20, pady = (20, 0), sticky = "nsew")

        # Header label
        self.header_label = customtkinter.CTkLabel(master=self, text="Select Desired Project")
        self.header_label.grid(row = 0, column = 0, columnspan = 2, padx = 20, pady = (20, 0), sticky="w")

        # Project variable
        self.sport = customtkinter.StringVar()
        self.sport.set("FWB")

        # Buttons: Radiobuttons for each of our projects
        self.project_button_a = customtkinter.CTkRadioButton(master = self, text = "FWB", variable = self.sport, value = "FWB")
        self.project_button_b = customtkinter.CTkRadioButton(master = self, text = "CMP", variable = self.sport, value = "CMP")
        self.project_button_c = customtkinter.CTkRadioButton(master = self, text = "OE", variable = self.sport, value = "OE")
        self.project_button_d = customtkinter.CTkRadioButton(master = self, text = "SS", variable = self.sport, value = "SS")
        self.project_button_e = customtkinter.CTkRadioButton(master = self, text = "QH", variable = self.sport, value = "QH")

        self.project_button_a.grid(row = 1, column = 0, padx = 20, pady = (10, 0), sticky="ew")
        self.project_button_b.grid(row = 1, column = 1, padx = 20, pady = (10, 0), sticky="ew")
        self.project_button_c.grid(row = 2, column = 0, padx = 20, pady = (10, 0), sticky="ew")
        self.project_button_d.grid(row = 2, column = 1, padx = 20, pady = (10, 0), sticky="ew")
        self.project_button_e.grid(row = 3, column = 0, padx = 20, pady = (10, 0), sticky="ew")

        # Header: Download a fresh template
        self.data_label = customtkinter.CTkLabel(master=self, text="Download a fresh meeting data template")
        self.data_label.grid(row = 4, column = 0, padx = 20, pady = (20, 0), sticky="w")

        # Button: Download template
        self.master_file_button = customtkinter.CTkButton(master = self, text = "Download Template", command = self.pullTemplate)
        self.master_file_button.grid(row = 5, column = 0, rowspan=2, padx = 20, pady = 10, sticky="ew")

        # Header: Select data file
        self.data_label = customtkinter.CTkLabel(master=self, text="Select Data File")
        self.data_label.grid(row = 4, column = 1, padx = 20, pady = (20, 0), sticky="w")

        # Button: Select data (Excel) file (contains your meeting data)
        self.master_file_button = customtkinter.CTkButton(master = self, text = "Choose File", command = self.openFile)
        self.master_file_button.grid(row = 5, column = 1, rowspan=2, padx = 20, pady = 10, sticky="ew")

        # Variable: Mtg Notes in filename boolean 
        self.note_button_var = customtkinter.IntVar()

        # Header: Mtg Notes label
        self.mtg_notes_checkbox_label = customtkinter.CTkLabel(master = self, text = 'Include "_Mtg Notes_" in the Filename')
        self.mtg_notes_checkbox_label.grid(row = 7, column = 0, padx = 20, pady = (20, 0), sticky="w")

        # Button:  Mtg Notes Checkbox (includes/excludes "_Mtg Notes_" from the filename)
        self.note_checkbox = customtkinter.CTkCheckBox(master = self, text="OFF", variable=self.note_button_var, command = self.toggle_button)
        self.note_checkbox.grid(row = 8, column = 0, padx = 20, pady = (10, 0), sticky="ew")

        # Button:  Submit - Create your templates (must have selected a data file and project)
        self.submit_button = customtkinter.CTkButton(master = self, text = "Build Templates", command = self.Submit)
        self.submit_button.grid(row = 9, column = 1, padx = 20, pady = (30, 10), sticky="e")
    
    # Update our checkbox to display on/off depending on which is shown
    def toggle_button(self):
        if self.note_button_var.get() == 1:
            self.note_checkbox.configure(text="ON")
        else:
            self.note_checkbox.configure(text="OFF")
    
    # Method: Execute - Build meeting templates based on what's provided
    def Submit(self):
        # Execute the function that builds templates
        templates = {
        'FWB' : 'Base_Templates/fwb-template.docx',
        'OE' : 'Base_Templates/oe-template.docx', 
        'SS' : 'Base_Templates/ss-template.docx',
        'CMP' : 'Base_Templates/cmp-template.docx',
        'QH' : 'Base_Templates/qh-template.docx',
        }

        if self.note_button_var.get() == 1:
            note_toggle = True
        else:
            note_toggle = False

        sport = self.sport.get()
        project_template = templates.get(sport)
        #print(project_template)

        buildTemplate(template = project_template, xl_file = self.master_file_name, note_bool = note_toggle)

    # Need to incorporate functionality where it determines whether the 

    # Method: Download a fresh template
    def pullTemplate(self):
        filepath = asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        wb = load_workbook('Base_Templates/yymmdd_project_meeting builder.xlsx')
        wb.save(filepath)

    # Method: Ask the user which file they want to use to create meeting templates
    def openFile(self):
        
        file = filedialog.askopenfile(mode = 'r', filetypes = [("XLSX Files", "*.xlsx")])

        def captureFilename(string):
            '''
            Helper function that strips the location of the file off and just returns the filename
            '''
            i = -1
            ltrs = []
            while string[i] != '/': # add the letters of our file name (in reverse) to a list and stop when we reach the slash indicating that we've reached a folder
                ltrs.append(string[i])
                i -= 1
            
            s = []
            for i in range(len(ltrs)):
                s.append(ltrs.pop())

            return "".join(s)

        if file:
            filepath = os.path.abspath(file.name)
            #self.current_master_file = str(filepath)
            master_file_full_name = str(filepath)
            self.master_file_name = captureFilename(master_file_full_name)
            self.master_file_text = customtkinter.CTkLabel(master = self, text = "Selected File: " + self.master_file_name)
            self.master_file_text.grid(row=10, column = 0, columnspan=8, padx = 10, pady = (20, 0))
            self.open_file_bool = True
            #self.master_file_text.text = str(filepath)

if __name__ == "__main__":
    app = App()
    app.mainloop()