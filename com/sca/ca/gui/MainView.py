import os
import tkinter as tk
import tkinter.filedialog
import PIL.Image
from PIL import ImageTk
from functools import partial

from tkinter import messagebox
import traceback

from com.sca.ca.FacilityProximityAssessment import FacilityProximityAssessment
from com.sca.ca.gui.EntryWithPlaceHolder import EntryWithPlaceholder
from com.sca.ca.gui.Styles import *
from com.sca.ca.model.ACSDataset import ACSDataset
from com.sca.ca.model.ACSdefaults import ACSdefaults
from com.sca.ca.model.CensusDataset import CensusDataset
from com.sca.ca.model.FacilityList import FacilityList


class MainView(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        self.censusblks_df = None
        self.acs_df = None
        self.acsCountyTract_df = None

        tk.Frame.__init__(self, master=master, *args, **kwargs)

        # set mainframe background color
        self.output_file = None
        self.facility_list_file = None
        self.home = master
        self.container = tk.Frame(self, width=400, height=300, bg=MAIN_COLOR)
        self.container.pack(fill="both", expand=True)

        self.folder_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')
        self.faclist_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')
        self.radius_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')

        self.faclist_frame.grid(row=3, columnspan=5, sticky="nsew")
        self.radius_frame.grid(row=4, columnspan=5, sticky="nsew")
        self.folder_frame.grid(row=5, columnspan=5, sticky="nsew")

        self.header = tk.Label(self.container, font=TEXT_FONT, bg='white', width=48,
                               text="Proximity Analysis Tool v1.0")
        self.header.grid(row=2, column=1, sticky='WE', pady=10)

        # First step - choose a facilities list file
        self.step1 = tk.Label(self.faclist_frame,
                              text='1. Browse to select the input file of latitude/longitude locations around which the tool will analyze demographics. \n    There is a Sample Template file in the “Inputs” folder to use to format your input file.', font=SMALL_TEXT_FONT, bg='white', anchor="w", justify="left")
        self.step1.grid(pady=10, row=3, column=0)

        self.inputs1_frame = tk.Frame(self.faclist_frame, height=32, padx=20, bg='white')
        self.inputs1_frame.grid(row=4, columnspan=1, sticky="nsew")

        fu = PIL.Image.open('images\icons8-document-48.png').resize((30, 30))
        ficon = self.add_margin(fu, 5, 0, 5, 0)
        fileicon = ImageTk.PhotoImage(ficon)
        self.fileLabel = tk.Label(self.inputs1_frame, image=fileicon, bg='white')
        self.fileLabel.image = fileicon
        self.fileLabel.grid(row=4, column=0)

        self.step1_instructions = tk.Label(self.inputs1_frame, text="Select file",
                                           font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step1_instructions.grid(row=4, column=1)
        self.fileLabel.bind("<Button-1>", partial(self.browse_file, self.step1_instructions))
        self.step1_instructions.bind("<Button-1>", partial(self.browse_file, self.step1_instructions))

        # Second step - choose a radius
        self.step2 = tk.Label(self.radius_frame,
                              text="2. Enter the radius (km) to be analyzed around each location.", font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step2.grid(pady=10, row=3, column=0)

        self.inputs2_frame = tk.Frame(self.radius_frame, height=32, padx=20, bg='white')
        self.inputs2_frame.grid(row=4, columnspan=1, sticky="nsew")

        self.radius_num = EntryWithPlaceholder(
            self.inputs2_frame, placeholder="Enter a radius ≤ 50 km", name="radius")
        self.radius_num["width"] = 24
        self.radius_num.grid(row=3, column=1, pady=10)

        # Third step - choose an output file location
        self.step3 = tk.Label(self.folder_frame,
                              text="3.  Enter the name (without extension) to be given to the output file of demographic results, which will be located in the “output” folder.",
                              font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step3.grid(pady=10, row=3, column=0)

        self.inputs3_frame = tk.Frame(self.folder_frame, height=32, padx=20, bg='white')
        self.inputs3_frame.grid(row=4, columnspan=1, sticky="nsew")

        self.output_file_name = EntryWithPlaceholder(
            self.inputs3_frame, placeholder="Enter a file name", name="filename")
        self.output_file_name["width"] = 24
        self.output_file_name.grid(row=4, column=0, pady=10)

        self.run_button = tk.Label(self.container, text="Run", font=TEXT_FONT,
                                   bg='lightgrey', relief='solid', borderwidth=2, width=8)
        self.run_button.grid(row=7, column=1, padx=20, pady=20, sticky='S')

        self.run_button.bind("<Button-1>", self.run_reports)

    def run_reports(self, event):

        # Make sure a list of facilities with  locations was entered
        if self.facility_list_file == None:
            messagebox.showinfo("Input error!", "Please select an input facility file.")
            return
 

        # Make sure radius is an integer
        if not self.radius_num.get_text_value().isnumeric():
            messagebox.showinfo("Input error!", "Radius must be an integer (no decimals). Please re-enter.")
            return

        # And make sure radius is not more than 50km        
        if int(self.radius_num.get_text_value()) > 50:
            messagebox.showinfo("Input error!", "Radius cannot exceed 50km. Please re-enter.")
            return

        
        print("Creating a proximity analysis!")

        self.show_running()

        # Load auxiliary files
        # Create censusblks dataframe
        try:
            censusblks = CensusDataset(path="resources/us_blocks_2020.csv")
        except BaseException as e:
            messagebox.showinfo("Error", "An error has occurred while trying to read the census block input file (csv format). \n" +
                                "The error says: \n\n" + str(e))
            self.reset_run_button()
            return
            
        self.censusblks_df = censusblks.dataframe
        print("Loaded census blocks")

        # Create acs dataframe
        try:
            acs = ACSDataset(path="resources/acs.csv")
        except BaseException as e:
            messagebox.showinfo("Error", "An error has occurred while trying to read the ACS data input file (csv format). \n" +
                                "The error says: \n\n" + str(e))
            self.reset_run_button()
            return
            
        self.acs_df = acs.dataframe
        print("Loaded ACS data")

        # Create ACS default dataframe
        try:
            acsDefault = ACSdefaults(path="resources/acs_defaults.csv")
        except BaseException as e:
            messagebox.showinfo("Error", "An error has occurred while trying to read the ACS default input file (csv format). \n" +
                                "The error says: \n\n" + str(e))
            self.reset_run_button()
            return
            
        self.acsDefault_df = acsDefault.dataframe
        print("Loaded ACS County/Tract data")

        # Create faclist dataframe
        try:
            faclist = FacilityList(path=self.facility_list_file)
        except BaseException as e:
            messagebox.showinfo("Error", "An error has occurred while trying to read the facility input file. \n" +
                                "The error says: \n\n" + str(e))
            self.reset_run_button()
            return
            
        faclist_df = faclist.dataframe
        print("Loaded facility data")

        name_only = self.output_file_name.get_text_value()
        folder = "output"

        try:
            if not os.path.exists(folder):
                os.makedirs(folder)
        except OSError:
            print ('Error: Creating directory. ' + folder)

        assessment = FacilityProximityAssessment(filename_entry=name_only,
                                                 output_dir=folder,
                                                 faclist_df=faclist_df,
                                                 radius=self.radius_num.get_text_value(),
                                                 census_df=self.censusblks_df,
                                                 acs_df=self.acs_df,
                                                 acsDefault_df=self.acsDefault_df)
    
        while True:
            try:        
                assessment.create()
                messagebox.showinfo("Complete", "Run complete. Check your output folder for results.")
                self.reset_gui()
                break
            except Exception as e:
                # messagebox.showinfo("Error", "An error has been generated while trying to run the assessment Create function. \n" +
                #                     "The error says: \n" + str(e))
                messagebox.showinfo("Error", "An error has been generated while trying to run the assessment Create function. \n" +
                                    "See traceback")
                print(traceback.format_exc())
                self.reset_run_button()
                break
                

    def show_running(self):
        self.run_button["text"] = "Running..."
        self.run_button["state"] = "disabled"
        self.home.update_idletasks()
        
    def reset_run_button(self):
        self.run_button["text"] = "Run"
        self.run_button["state"] = "normal"
        self.home.update_idletasks()

    def reset_gui(self):
        self.output_file = None
        self.facility_list_file = None

        self.run_button["text"] = "Run"
        self.run_button["state"] = "normal"
        self.step1_instructions["text"] = "Select file"
        self.radius_num.put_placeholder()
        self.output_file_name.put_placeholder()
        self.home.update_idletasks()

    # The input file browse handler.
    def browse_file(self, icon, event):
        self.facility_list_file = tkinter.filedialog.askopenfilename()
        if not self.facility_list_file:
            return

        icon["text"] = self.facility_list_file.split("/")[-1]

    def add_margin(self, pil_img, top, right, bottom, left):
        width, height = pil_img.size
        new_width = width + right + left
        new_height = height + top + bottom
        result = PIL.Image.new(pil_img.mode, (new_width, new_height))
        result.paste(pil_img, (left, top))
        return result
