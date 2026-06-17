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
            re.match(r'^[0-9]', line) or
            line == " "
            )


def create_questions_file(ui_ctx: UIContext):

    # --------------------------------------------
    # Open the file to read
    # --------------------------------------------

    try:
        file_path = Path(ui_ctx._input_file_name).absolute()
    except (UnboundLocalError, TypeError):
        tk.messagebox.showerror(title="File not found", message=f"Select a .pdf file.")
        
    try:
        questions_file = pdf.PdfReader(file_path)
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        tk.messagebox.showerror(title="File not found", message=f"File {file_path} not found.")
        # sys.exit(1)

    pdf_doc_example = pymupdf.open(file_path)
    print(pdf_doc_example)

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

    for page_num, pages in enumerate(questions_file.pages):
        text_per_page[page_num+1] = pages.extract_text().splitlines()


    # --------------------------------------------
    # Look for the solutions page
    # --------------------------------------------

    solutions_page = None
    for page in text_per_page.keys():
        for line_num, line in enumerate(text_per_page[page]):
            if starts_with_number(line) or is_questions_page(line):
                print(f"{page} not solutions page")
                break
            if is_solutions_page(line):
                solutions_page = page
                break
        if solutions_page is not None:
            break

    if solutions_page is None:
        raise ValueError("Solutions page not found.")

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
    # Create a new solutions deque with one solution per element
    # --------------------------------------------

    solutions_list = deque()
    previous_line = ""

    while solutions_full_text:
        line = solutions_full_text.popleft()

        if starts_with_number(line) and previous_line == "":
            line = remove_number(line)
            previous_line = line
        elif previous_line != "":
            if ord(line[0]) > 32:
                previous_line += line
            else:
                solutions_list.append(previous_line)
                previous_line = ""

    if previous_line != "":
        solutions_list.append(previous_line)


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
            # output_file.paragraphs[-1].paragraph_format
            added_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in added_paragraph.runs:
                run.bold = True

            # Add the solution and an additional line
            added_paragraph = output_file.add_paragraph(solutions_list[current_question])
            added_paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            added_paragraph.paragraph_format.space_after = 0
            output_file.add_paragraph("")

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
