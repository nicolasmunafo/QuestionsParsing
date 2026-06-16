from UI.ui import manage_ui
from UI.ui import UIContext
from get_questions_pdf import create_questions_file

def main():
    ui_ctx = UIContext()
    
    ui_ctx._callback_create_file = create_questions_file

    manage_ui(ui_ctx)

if __name__ == "__main__":
    main()

# TODO: Check case when .json does not exist
# TODO: Test edge cases when one of the directories is not set
# TODO: PyMuPDF (fitz) to get Bold and format