import tkinter as tk
import tkinter.filedialog
import PIL.Image
from PIL import ImageTk
from functools import partial
import webbrowser

from com.sca.ca.CommunityAssessment import CommunityAssessment
from com.sca.ca.gui.EntryWithPlaceHolder import EntryWithPlaceholder
from com.sca.ca.gui.Styles import *


class MainView(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master=master, *args, **kwargs)

        # set mainframe background color
        self.output_dir = None
        self.facility_list_file = None

        self.home = master
        self.container = tk.Frame(self, width=400, height=300, bg=MAIN_COLOR)
        self.container.pack(fill="both", expand=True)

        self.folder_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')
        self.faclist_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')
        self.radius_frame = tk.Frame(self.container, height=120, pady=1, padx=5, bg='white')

        self.folder_frame.grid(row=3, columnspan=5, sticky="nsew")
        self.faclist_frame.grid(row=4, columnspan=5, sticky="nsew")
        self.radius_frame.grid(row=5, columnspan=5, sticky="nsew")

        self.header = tk.Label(self.container, font=TEXT_FONT, bg='white', width=48,
                               text="Community Assessment Stand-alone")
        self.header.grid(row=2, column=1, sticky='WE')

        # First step - choose an output folder
        self.step1 = tk.Label(self.folder_frame,
                              text="1.", font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step1.grid(pady=10, row=3, column=0)

        fu = PIL.Image.open('images\icons8-folder-48.png').resize((30,30))
        ficon = self.add_margin(fu, 5, 0, 5, 0)
        fileicon = ImageTk.PhotoImage(ficon)
        self.fileLabel = tk.Label(self.folder_frame, image=fileicon, bg='white')
        self.fileLabel.image = fileicon
        self.fileLabel.grid(row=3, column=1)

        self.step1_instructions = tk.Label(self.folder_frame, text="Select output folder",
                                           font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step1_instructions.grid(row=3, column=2)
        self.fileLabel.bind("<Button-1>", partial(self.browse, self.step1_instructions))
        self.step1_instructions.bind("<Button-1>", partial(self.browse, self.step1_instructions))

        # Second step - choose a facilities list file
        self.step2 = tk.Label(self.faclist_frame,
                              text="2.", font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step2.grid(pady=10, row=3, column=0)

        fu = PIL.Image.open('images\icons8-document-48.png').resize((30,30))
        ficon = self.add_margin(fu, 5, 0, 5, 0)
        fileicon = ImageTk.PhotoImage(ficon)
        self.fileLabel = tk.Label(self.faclist_frame, image=fileicon, bg='white')
        self.fileLabel.image = fileicon
        self.fileLabel.grid(row=3, column=1)

        self.step2_instructions = tk.Label(self.faclist_frame, text="Select facility list file",
                                           font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step2_instructions.grid(row=3, column=2)
        self.fileLabel.bind("<Button-1>", partial(self.browse_file, self.step2_instructions))
        self.step2_instructions.bind("<Button-1>", partial(self.browse_file, self.step2_instructions))

        # Third step - choose a radius
        self.step3 = tk.Label(self.radius_frame,
                              text="3.", font=SMALL_TEXT_FONT, bg='white', anchor="w")
        self.step3.grid(pady=10, row=3, column=0)

        self.radius_num = EntryWithPlaceholder(
            self.radius_frame, placeholder="Enter a radius <= 50 km", name="radius")
        self.radius_num["width"] = 24
        self.radius_num.grid(row=3, column=1, pady=10)

        run_button = tk.Label(self.container, text="Run", font=TEXT_FONT, bg='white')
        run_button.grid(row=5, column=2, padx=20, pady=20, sticky='E')

        run_button.bind("<Button-1>", self.run_reports)

    def run_reports(self, event):

        print("Creating a community assessment!")

        ca = CommunityAssessment(self.output_dir, self.facility_list_file,
                                 self.radius_num.get_text_value())

        ca.calculate_distances()
        ca.create_workbook()

    # The folder browse handler.
    def browse(self, icon, event):
        self.output_dir = tkinter.filedialog.askdirectory()
        if not self.output_dir:
            return

        icon["text"] = self.output_dir.split("/")[-1]

    # The folder browse handler.
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
