from enum import Enum
from typing import List
import tkinter as tk
from tkinter import messagebox
from tkinter import * 
from tkinter import filedialog
import json
from pathlib import Path

# TODO:
# TODO:
# https://www.geeksforgeeks.org/python/file-explorer-in-python-using-tkinter/

class FileType(Enum):
    FOLDER = 1
    FILE = 2

class UIContext:
    class ContextPersistData(Enum):
        Input_Folder = 0
        Output_Folder = 1

    def __init__(self):
        self._callback_create_file = lambda: None
        self.__initialize_file__()
        self._input_file_name = None
        self._output_file_name = None

    def __initialize_file__(self):
        self._persist_file = Path(__file__).parent.resolve() / Path("persist_data.json")
        self._input_folder = "/"
        self._output_folder = "/"

        if not(self._persist_file.exists()):
            with open(self._persist_file, "w") as f:
                json.dump({}, f)
        else:
            with open(self._persist_file, "r") as f:
                json_file = json.load(f)
                self._input_folder = json_file[UIContext.ContextPersistData.Input_Folder.name]
                self._output_folder = json_file[UIContext.ContextPersistData.Output_Folder.name]

    @property
    def input_file_name(self):
        return self._input_file_name
    
    @input_file_name.setter
    def input_file_name(self, value):
        if not isinstance(value, str) and value is not None:
            raise TypeError("Filename must be a string")
        self._input_file_name = value
        self.__store_file_name___(UIContext.ContextPersistData.Input_Folder, value)

    @property
    def output_file_name(self):
        return self._output_file_name
    
    @output_file_name.setter
    def output_file_name(self, value):
        if not isinstance(value, str) and value is not None:
            raise TypeError("Filename must be a string")
        self._output_file_name = value
        self.__store_file_name___(UIContext.ContextPersistData.Output_Folder, value)
    
    @property
    def input_folder(self):
        return self._input_folder
    
    @property
    def output_folder(self):
        return self._output_folder

    def __store_file_name___(self, file_type: ContextPersistData, file_name: str):
        with open(self._persist_file, "r") as output_file:
            json_file = json.load(output_file)
        
        json_file[file_type.name] = str(Path(file_name).parent.resolve()) if not(Path(file_name).is_dir()) else str(Path(file_name))

        with open(self._persist_file, "w") as output_file:
            json.dump(json_file, output_file)
    

def browseFiles(label: tk.Label, types_list: List, ui_ctx: UIContext) -> None:  
    filename = ""
    if types_list[2] == FileType.FILE:
        ui_ctx.input_file_name = filedialog.askopenfilename(initialdir = ui_ctx.input_folder, title = "Select a File", filetypes = ((types_list[0], types_list[1]),))
        filename = ui_ctx.input_file_name
    else:
        ui_ctx.output_file_name = filedialog.askdirectory(initialdir = ui_ctx.output_folder, title = "Select a Folder")
        filename = ui_ctx.output_file_name
     
    # Change label contents
    label.configure(text=label["text"] + filename)


def create_window() -> Tk:
    '''Create the window with the properties'''

    # Create the root window
    window = Tk()
    
    # Set window title
    window.title('Questions Extractor')
       
    #Set window background color
    window.config(background = "white")

    window.resizable(False, False)

    # Set window size
    window_width = 640
    window_height = 150

    # get the screen dimension
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # find the center point
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)

    # set the position of the window to the center of the screen
    window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    return window


def manage_ui(ui_ctx: UIContext):
    '''Create the window, UI and buttons'''
    
    window = create_window()
    window.columnconfigure(0, weight=1, minsize=150)
    window.columnconfigure(1, weight=1)
    
    # Create a File Explorer label
    label_select_file = Label(window, 
                                text = ui_ctx.input_folder,
                                width = 100, height = 2, 
                                fg = "blue",
                                justify=LEFT)

    button_input_data = ["PDF files", "*.pdf", FileType.FILE]

    button_select_file = Button(window, 
                            text = "Select input .pdf file",
                            command = lambda: browseFiles(label_select_file, button_input_data, ui_ctx))
    
    label_output_file = Label(window, 
                                text = ui_ctx.output_folder,
                                width = 100, height = 2, 
                                fg = "blue")
    
    button_output_data = ["all files", "*.*", FileType.FOLDER]

    button_select_output = Button(window, 
                            text = "Select output folder",
                            command = lambda: browseFiles(label_output_file, button_output_data, ui_ctx))
    
    button_create_file = Button(window, 
                            text = "Create file",
                            command = lambda: ui_ctx._callback_create_file(ui_ctx))
    
    # check_button_create_file(button_create_file, ui_ctx)
    
    button_exit = Button(window, 
                        text = "Exit",
                        command = exit) 
    
    # Grid method is chosen for placing the widgets at respective positions in a table like structure by specifying rows and columns

    label_select_file.grid(column = 1, row = 1, sticky="e", padx=10)
    
    button_select_file.grid(column = 0, row = 1, sticky="w", padx=10)

    label_output_file.grid(column = 1, row = 2, sticky="e", padx=10)
    
    button_select_output.grid(column = 0, row = 2, sticky="w", padx=10)

    button_create_file.grid(column = 0, row = 3, columnspan=2, pady=5)
    
    button_exit.grid(column = 0, row = 4, columnspan=2, pady=5)

    # label_select_file.grid(column = 0, row = 1)
    
    # button_select_file.grid(column = 0, row = 2)

    # label_output_file.grid(column = 0, row = 3)
    
    # button_select_output.grid(column = 0, row = 4)

    # button_create_file.grid(column = 0, row = 6)
    
    # button_exit.grid(column = 0, row = 7)
    
    # Let the window wait for any events
    window.mainloop()


def enable_button(curr_button: Button, status: bool) -> None:
    if status:
        curr_button.config(state="normal")
    else:
        curr_button.config(state="disabled")


def check_button_create_file(curr_button: Button, ui_ctx: UIContext) -> None:
    if ui_ctx.output_file_name == None:
        enable_button(curr_button, False)
    else:
        enable_button(curr_button, True)


