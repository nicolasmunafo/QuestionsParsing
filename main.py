from UI.ui import manage_ui
from UI.ui import UIContext
from get_questions_pdf import create_questions_file

def main():
    ui_ctx = UIContext()
    
    ui_ctx._callback_create_file = create_questions_file

    manage_ui(ui_ctx)

if __name__ == "__main__":
    main()

# TODO: Error with solution 21 und 24 Lernziele 1
# TODO: When the line starts with -, add \n
# TODO: PyMuPDF (fitz) to get Bold and format