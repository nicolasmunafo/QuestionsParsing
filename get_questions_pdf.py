import pypdf as pdf
import docx
from pathlib import Path
import re
from collections import deque
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import sys
import pymupdf

import tkinter as tk
from tkinter import messagebox
from UI.ui import manage_ui
from UI.ui import UIContext

def starts_with_number(line):
    return re.match(r'[0-9]+\)', line)

def remove_number(line):
    # sub eliminates the string when it finds it
    # first with ^ it forces to check it only at the start
    # Then it can have more than one number and then parenthesis and spaces
    return(re.sub(r'^[0-9]+\)\s*', '', line,))

def is_questions_page(line):
    return line.startswith("Lernziele")

def is_solutions_page(line):
    return line.startswith("Lösungen")
    
    # return(re.split(r'\s[0-9]+\)', line, maxsplit=0))

def can_be_omitted(line):
    return ("Medizinische Grundlagen; 2026-2027" in line or
            "Barbara" in line or
            re.match(r'^[0-9]\s', line) or
            line == " " or
            "Lösungen" in line
            )

def is_new_line(line: str):
    return (re.match(r'^[0-9].\s', line) or
            line.startswith("-")
            )

def add_paragraph_from_line(output_file: docx.Document, line: str) -> None:
    added_paragraph = output_file.add_paragraph(line)
    added_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    added_paragraph.paragraph_format.space_after = 0


def is_bold(flags: int) -> bool:
    return flags & 2 ** 4

def is_italic(flags: int) -> bool:
    return flags & 2 ** 1

def is_underlined(flags: int) -> bool:
    return flags & 2 ** 0


def get_solutions(output_file: docx.Document, solutions_full_text: deque, current_question: int) -> None:

    solution_to_add = ""
    omitted_lines = 0
    line: str

    while solutions_full_text:
        line = solutions_full_text[0]
        
        print("New line" if "\n" in line else "")

        if starts_with_number(line):                      
            # Save the current solution question only if solution_to_add is empty. Otherwise, break the loop
            if solution_to_add == "":
                line = remove_number(line)           
                
                # For cases in which line is null because it was just a number, like 10)
                solution_to_add = line if line != "" else line + " "        
                omitted_lines = 0
            else:
                break            
        else:
            # If the line does not start with a number and it's relevant, concatenate it with the previous line
            if not(can_be_omitted(line)):
                # If previously a line was omitted, and it was only one, it means this is a title, so the solution is finalized
                if omitted_lines == 1:
                    break
                
                if solution_to_add != "":
                    # If it's an ordered or unordered list, add the previous line as a paragraph and add a line break
                    if is_new_line(line):
                        add_paragraph_from_line(output_file, solution_to_add)
                        solution_to_add = line
                    else:
                        solution_to_add += line
                
                omitted_lines = 0
            else:
                # If the line can be omitted, detect this in case the next line is still part of the solution
                omitted_lines += 1
        
        # Always pop the line at the end of an iteration. If the next line is a new solution, it will not add it
        solutions_full_text.popleft()


    # Add the latest solution if there is something
    if solution_to_add != "":
        add_paragraph_from_line(output_file, solution_to_add)

    # Add an additional line
    output_file.add_paragraph("")

def create_questions_file(ui_ctx: UIContext):

    # --------------------------------------------
    # Open the file to read
    # --------------------------------------------

    try:
        file_path = Path(ui_ctx._input_file_name).absolute()
    except (UnboundLocalError, TypeError):
        tk.messagebox.showerror(title="File not found", message=f"Select a .pdf file.")
        
    try:
        questions_file = pymupdf.open(file_path)
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        tk.messagebox.showerror(title="File not found", message=f"File {file_path} not found.")
        # sys.exit(1)


    # --------------------------------------------
    # Open the file to write the final results
    # --------------------------------------------

    if ui_ctx.output_file_name == None:
        tk.messagebox.showerror(title="File not found", message=f"Select an output folder.")

    output_file_name = file_path.stem
    output_file_path = Path(ui_ctx.output_file_name) / Path(output_file_name + '.docx')
    output_file = docx.Document()

    try:
        with open(output_file_path, "a"):
            pass
    except PermissionError:
        print(f"File {file_path} is locked or already open.")
        tk.messagebox.showerror(title="File locked", message=f"File {file_path} is locked or already open.")
        # sys.exit(1)

    # --------------------------------------------
    # Create a dictionary with the text of each page
    # --------------------------------------------

    text_per_page = dict()

    for page_num, page in enumerate(questions_file):
        text_per_page[page_num+1] = page.get_text("text").splitlines()

    pass

    # --------------------------------------------
    # Look for the solutions page
    # --------------------------------------------

    solutions_page = None
    for page_num, page in enumerate(questions_file):
        lines = page.get_text("text").splitlines()
        for line in lines:
            if starts_with_number(line) or is_questions_page(line):
                print(f"{page} not solutions page")
                break
            if is_solutions_page(line):
                solutions_page = page
                break
        if solutions_page is not None:
            break

    # if solutions_page is None:
    #     raise ValueError("Solutions page not found.")
    
    for page in questions_file[1:]:
        text_blocks = page.get_text("dict", flags=pymupdf.TEXTFLAGS_TEXT)["blocks"]
        for block in text_blocks:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    if is_bold(span["flags"]):
                        print("It's bold")
                    color = pymupdf.sRGB_to_rgb(span["color"])
                    print(f"Text: {text}, Color: {color}")

    # --------------------------------------------
    # Create a list with the full text of the questions and solutions
    # --------------------------------------------

    questions_full_text = deque()
    solutions_full_text = deque()
    for page in text_per_page.keys():
        if page < solutions_page:
            questions_full_text += text_per_page[page]
        elif page >= solutions_page:
            solutions_full_text += text_per_page[page]


    # --------------------------------------------
    # Start saving the contents in a new file
    # --------------------------------------------

    # Clean up and remove the first lines
    current_line = 0
    while questions_full_text:
        line = questions_full_text[current_line]
        if not(is_questions_page(line)):
            questions_full_text.popleft()
        else:
            break

    current_question = 0

    prev_line = questions_full_text.popleft()
    curr_line = ""
    is_heading = False
    is_question = False

    while questions_full_text:
        curr_line = questions_full_text.popleft()

        # If the previous line is a question
        if starts_with_number(prev_line):
            
            is_question = True
            # if the next line is not a question and is not blank, concatenate them
            if not(starts_with_number(curr_line)) and curr_line != " ":
                prev_line += curr_line
                continue
        
        else:
            if can_be_omitted(prev_line):
                prev_line = curr_line
                continue

            if prev_line != " ":
                is_heading = True
            elif curr_line == " ":      # If the previous and current lines are spaces, just finish the document
                break
            

        # In any other case, add it as a paragraph
        added_paragraph = output_file.add_paragraph(prev_line)
        
        # If is heading, set the style as a heading
        if is_heading:
            added_paragraph.style = "Heading 2"
            added_paragraph.paragraph_format.space_after = Pt(12)
            is_heading = False
        
        # If is a question, add the associated solution
        elif is_question:
            added_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in added_paragraph.runs:
                run.bold = True

            # Add the solution and an additional line
            get_solutions(output_file, solutions_full_text, current_question)

            # Reset all question variables
            is_question = False
            current_question += 1

        prev_line = curr_line

    # Change the first paragraph as a title
    output_file.paragraphs[0].style = "Title"
    output_file.paragraphs[0].style.font.size = Pt(20)

    # Chenge the font to calibri for normal
    style = output_file.styles['Normal']
    font = style.font
    font.name = 'Calibri'

    output_file.save(output_file_path)
    tk.messagebox.showinfo(title="File created", message=f"File \"{output_file_name}.docx\" successfully created!")


# TODO: Delete this later, just for testing
ui_ctx = UIContext()
ui_ctx.input_file_name = "C:\\Users\\Nicolas\\GitHub\\Automation_Py\\7. Lernziele Herz_ Kreislauf und Gefasssystem.pdf"
# ui_ctx.input_file_name = "C:\\Users\\Nicolas\\GitHub\\Automation_Py\\FilesReading\\Fragen.pdf"
# ui_ctx.output_file_name = "C:\\Users\\Nicolas\\GitHub\\Automation_Py\\FilesReading"
ui_ctx.output_file_name = "C:\\Users\\Nicolas\\GitHub\\Automation_Py"

create_questions_file(ui_ctx)

'''
https://pymupdf.readthedocs.io/en/latest/app1.html
<page>
    <text block>
        <line>
            <span>
                <char>
    <image block>
        <img>
A text page consists of blocks (= roughly paragraphs).

A block consists of either lines and their characters, or an image.

A line consists of spans.

A span consists of adjacent characters with identical font properties: name, size, flags and color.

'''