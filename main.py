import customtkinter
from tkinter import filedialog
import os

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("1000x500")
        self.title("Template Builder")
        self.minsize(900, 400)

        # Create a 1x5 grid
        #self.grid_rowconfigure(0, weight = 1)
        self.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight = 1)

        
        # Table to display text to the user
        self.table_space = customtkinter.CTkTextbox(master=self, bg_color="white")
        self.table_space.grid(row = 0, column = 2, columnspan = 7, rowspan = 10, padx = 20, pady = (20, 0), sticky = "nsew")

        # Header label
        self.header_label = customtkinter.CTkLabel(master=self, text="Select Desired Project")
        self.header_label.grid(row = 0, column = 0, columnspan = 2, padx = 20, pady = (20, 0), sticky="w")

        # Project variable
        sport = customtkinter.StringVar()
        sport.set("FWB")

        # Project radiobuttons
        self.project_button_a = customtkinter.CTkRadioButton(master = self, text = "FWB", variable = sport, value = "FWB")
        self.project_button_b = customtkinter.CTkRadioButton(master = self, text = "CMP", variable = sport, value = "CMP")
        self.project_button_c = customtkinter.CTkRadioButton(master = self, text = "OE", variable = sport, value = "OE")
        self.project_button_d = customtkinter.CTkRadioButton(master = self, text = "SS", variable = sport, value = "SS")
        self.project_button_e = customtkinter.CTkRadioButton(master = self, text = "QH", variable = sport, value = "QH")

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

        # Button: Select Excel file (contains your meeting data)
        self.master_file_button = customtkinter.CTkButton(master = self, text = "Choose File", command = self.openFile)
        self.master_file_button.grid(row = 5, column = 1, rowspan=2, padx = 20, pady = 10, sticky="ew")

        # Variable: Mtg Notes in filename boolean 
        self.note_button_var = customtkinter.IntVar()

        # Mtg Notes label
        self.mtg_notes_checkbox_label = customtkinter.CTkLabel(master = self, text = 'Include "_Mtg Notes_" in the Filename')
        self.mtg_notes_checkbox_label.grid(row = 7, column = 0, padx = 20, pady = (20, 0), sticky="w")

        # Button:  Mtg Notes Checkbox (includes/excludes "_Mtg Notes_" from the filename)
        self.note_checkbox = customtkinter.CTkCheckBox(master = self, text="OFF", variable=self.note_button_var, command = self.toggle_button)
        self.note_checkbox.grid(row = 8, column = 0, padx = 20, pady = (10, 0), sticky="ew")

        # Button:  Submit - Create your templates (must have selected a data file and project)
        self.submit_button = customtkinter.CTkButton(master = self, text = "Build Templates", command = self.Submit)
        self.submit_button.grid(row = 9, column = 1, padx = 20, pady = (30, 10), sticky="e")
    
    def toggle_button(self):
        if self.note_button_var.get() == 1:
            self.note_checkbox.config(text="ON")
        else:
            self.note_checkbox.config(text="OFF")
    
    def Submit(self):
        input_key = self.api_entry.get() # Do this but instead of grabbing the api entry, we grab the stuff we care about
        # Execute the function that builds templates
        print(input_key)

    # Need to incorporate functionality where it determines whether the 

    def pullTemplate(self):
        pass

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
            self.current_master_file = str(filepath)
            master_file_full_name = str(filepath)
            master_file_name = captureFilename(master_file_full_name)
            self.master_file_text = customtkinter.CTkLabel(master = self, text = "Selected File: " + master_file_name)
            self.master_file_text.grid(row=10, column = 0, columnspan=8, padx = 10, pady = (20, 0))
            self.open_file_bool = True
            #self.master_file_text.text = str(filepath)

if __name__ == "__main__":
    app = App()
    app.mainloop()