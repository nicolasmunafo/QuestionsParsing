from UI.ui import manage_ui
from UI.ui import UIContext
from get_questions_pdf import create_questions_file

def main():
    ui_ctx = UIContext()
    
    ui_ctx._callback_create_file = create_questions_file

    manage_ui(ui_ctx)

if __name__ == "__main__":
    main()

# TODO: check if the file actually has questions, otherwise, return
# TODO: PyMuPDF (fitz) to get Bold and format