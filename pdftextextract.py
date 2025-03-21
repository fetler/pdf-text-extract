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
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Extracted Data"
        ws.append(["ID Number", "Page Number", "Course Name", "Accepted Outcome"])

        for page_num in range(total_pages):
            try:
                page = document.load_page(page_num)
                text = page.get_text("text")

                # Students who have not completed the academic year have their full academic year transcript show in the PDF. It uses different wording to those who have completed
                # the academic year, so different regex is required to match and extract the text.
                if "Classification" not in text:
                    id_matches = re.findall(r'\((\d{8})\)', text)
                    outcome_matches = re.findall(r'Accepted Outcome:\s*(.+)', text)
                    course_matches = re.findall(r'Course\s(.+?)(\(.+)', text)

                    if id_matches and outcome_matches and course_matches:
                        for i in range(min(len(id_matches), len(outcome_matches), len(course_matches))):
                            id_number = id_matches[i]
                            accepted_outcome = outcome_matches[i]
                            course_name = course_matches[i][0]
                            course_name = course_name.replace("Master of Science in", "MSc").replace(" MSc", "MSc").replace(" PG Cert", "PG Cert").replace(" PG Dip", "PG Dip").replace(" Postgraduate", "Postgraduate").replace(" DClin", "DClin")

                            ws.append([id_number, page_num + 1, course_name, accepted_outcome])

                # Students who have completed the academic year appear in a table with students from their cohort who have also completed the academic year. As it uses different wording
                # to those who have not completed the academic year, different regex is required to match and extract the text. It also loops through multiple matches per page where applicable.
                else:
                    programme_matches = re.findall(r'Programme:\s*(.+)', text)
                    student_matches = re.findall(r'([A-Za-z]+,\s[A-Za-z]+(?:\s[A-Za-z]+)*)\s*\((\d{8})\)', text)
                    board_decision_matches = re.findall(r'\)\s*([A-Z]+\s*-\s*[A-Za-z ]+)', text)
                    classification_matches = re.findall(r'([A-Za-z]+)$', text.strip())
                    print(f"{page_num + 1} - {student_matches}")

                    if student_matches:
                        for i, student_match in enumerate(student_matches):
                            student_name = student_match[0].strip()
                            student_id = student_match[1].strip()

                            programme_name = programme_matches[i].strip() if i < len(programme_matches) else "Programme Name Not Found"
                            board_decision = board_decision_matches[i].strip() if i < len(board_decision_matches) else "Board Decision Not Found"
                            classification = classification_matches[i].strip() if i < len(classification_matches) else "Classification Not Found"
                            accepted_outcome_board = f"{board_decision} - {classification}".strip()
                        
                            programme_name = programme_name.replace("Master of Science in", "MSc").replace(" MSc", "MSc").replace(" PG Cert", "PG Cert").replace(" PG Dip", "PG Dip").replace(" Postgraduate", "Postgraduate").replace(" DClin", "DClin")
                            ws.append([student_id, page_num + 1, programme_name, accepted_outcome_board])
            
            except Exception as e:
                print(f"Error processing page {page_num + 1}: {e}")
                continue

        global extracted_data
        extracted_data = wb
        messagebox.showinfo("Success", "Data extracted successfully!")
        document.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process PDF: {e}")

# GUI window to save the XLSX file containing the data found by the regex
def save_xlsx():
    if extracted_data is None:
        messagebox.showerror("Error", "No extracted data to save. Run extraction first.")
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        extracted_data.save(file_path)
        messagebox.showinfo("Success", f"File saved: {file_path}")

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
