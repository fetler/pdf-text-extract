import fitz # for PDF handling
import re # for regex
import openpyxl # for handling XLSX files
import tkinter as tk # for GUI
from tkinter import filedialog, messagebox # for file dialog and messagebox GUI
from openpyxl.utils import get_column_letter # for automatically adjusting column widths of the XLSX file (not yet implemented in script)

# GUI window to select the PDF file to use
def load_pdf():
    global pdf_path
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        lbl_pdf.config(text=f"Selected: {pdf_path}")

# Use regex to extract student IDs, page numbers, course names, accepted outcomes, board decisions, and classifications
def extract_data():
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file first.")
        return

    try:
        document = fitz.open(pdf_path)
        total_pages = document.page_count
        entries = []

        for page_num in range(total_pages):
            try:
                page = document.load_page(page_num)
                text = page.get_text("text")

                # Split and clean lines
                lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
                
                # Skip completely blank pages
                if not lines:
                    continue

                # Compile regex patterns
                accepted_outcome_pattern = re.compile(r"^Accepted Outcome:")
                eight_digit_pattern = re.compile(r"\b\d{8}\b")
                student_pattern = re.compile(r"^.+,\s.+\s\(\d{8}\)$")

                # Check if "Accepted Outcome:" is on the page
                has_accepted_outcome = any(accepted_outcome_pattern.match(line) for line in lines)

                if has_accepted_outcome:
                    # Print lines that match 8-digit IDs or "Accepted Outcome:"
                    for idx, line in enumerate(lines):
                        if eight_digit_pattern.search(line):
                            student_line = f"{line}"

                            # Search forward for "Accepted Outcome:" line
                            for j in range(idx + 1, len(lines)):
                                if accepted_outcome_pattern.match(lines[j]):
                                    outcome_line = f"{lines[j]}"
                                    print(f"{page_num + 1}${student_line}${outcome_line}")
                                    break  # Stop searching once matched
                else:
                    # Run the alternative block if "Accepted Outcome:" is not present
                    results = []

                    for i in range(len(lines) - 2):  # Ensure at least 3 lines for student, outcome, classification
                        if student_pattern.match(lines[i]):
                            name_id = lines[i]
                            outcome = lines[i + 1].strip()
                            classification = lines[i + 2].strip()

                            # Only include classification if outcome is one of the specified values
                            if outcome in ["Pass Award", "Pass Award with Compensation"]:
                                results.append((name_id, outcome, classification))
                            else:
                                results.append((name_id, outcome, classification))  # Skip classification

                    for name_id, outcome, classification in results:
                        if classification:
                            print(f"{page_num + 1}${name_id}${outcome}${classification}")
                        else:
                            print(f"{page_num + 1}${name_id}${outcome}")

            except Exception as e:
                print(f"Error processing page {page_num + 1}: {e}")

    except Exception as e:
        print(f"Failed to process PDF: {e}")

# GUI Setup
root = tk.Tk()
root.title("PDF Data Extractor")
root.geometry("500x300")

pdf_path = None
extracted_data = None

# Load PDF button
btn_load = tk.Button(root, text="Load PDF", command=load_pdf)
btn_load.pack(pady=10)

# Label to display the filepath of the selected PDF file
lbl_pdf = tk.Label(root, text="No file selected", wraplength=400)
lbl_pdf.pack()

# Button to trigger the matching and extracting of the data from the PDF using regex
btn_extract = tk.Button(root, text="Extract Data", command=extract_data)
btn_extract.pack(pady=10)

# Save button for extracting the data as an XLSX file
btn_save = tk.Button(root, text="Save to Excel", command=save_xlsx)
btn_save.pack(pady=10)

# Exit script button
btn_exit = tk.Button(root, text="Exit", command=root.destroy)
btn_exit.pack(pady=10)

root.mainloop()
