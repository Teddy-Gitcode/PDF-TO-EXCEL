import os
import re
import openpyxl
import webbrowser
import customtkinter as ctk
from tkinter import StringVar, IntVar
from tkinter.filedialog import askopenfilename
from pdfminer.high_level import extract_text

# Define the function to extract details from PDF

def extract_details_from_pdf(pdf_file):
    text = extract_text(pdf_file)

    # Extract CU IN No.
    cu_in_pattern = re.compile(r"CU IN No\.:(.*?)\n")
    cu_in_match = cu_in_pattern.search(text)
    cu_in = cu_in_match.group(1).strip() if cu_in_match else ""

    # Extract CU SN No.
    cu_sn_pattern = re.compile(r"CU SN No\.:(.*?)\n")
    cu_sn_match = cu_sn_pattern.search(text)
    cu_sn = cu_sn_match.group(1).strip() if cu_sn_match else ""

    # Document_Type
    Document_Type_pattern = re.compile(r"Document_Type:(.*?)\n")
    Document_Type_match = Document_Type_pattern.search(text)
    Document_Type = Document_Type_match.group(1).strip() if Document_Type_match else ""

    # PIN
    PIN_pattern = re.compile(r"PIN:(.*?)\n")
    PIN_match = PIN_pattern.search(text)
    PIN = PIN_match.group(1).strip() if PIN_match else ""

    # INVOICE_NO
    INVOICE_NO_pattern = re.compile(r"INVOICE_NO\s*:\s*(.*?)\n")
    INVOICE_NO_match = INVOICE_NO_pattern.search(text)
    INVOICE_NO = INVOICE_NO_match.group(1).strip() if INVOICE_NO_match else ""

    #Invoice Date
    Invoice_Date_pattern = re.compile(r"Invoice\s+Date\s*:\s*(.*?)\n")
    Invoice_Date_match = Invoice_Date_pattern.search(text)
    Invoice_Date = Invoice_Date_match.group(1).strip() if Invoice_Date_match else ""

    #PIN_No
    PIN_No_pattern = re.compile(r"PIN_No\s*:\s*(.*?)\n")
    PIN_No_match = PIN_No_pattern.search(text)
    PIN_No = PIN_No_match.group(1).strip() if PIN_No_match else ""

    # Extract TOTAL
    total_match = re.search(r"TOTAL\s*:\s*([0-9,.]+)", text, re.IGNORECASE)
    total = total_match.group(1).strip() if total_match else ""
    if not total:
        total_match = re.search(r"([0-9,.]+)\s*TOTAL", text, re.IGNORECASE)
        total = total_match.group(1).strip() if total_match else ""

    # Extract VAT percentage
    vat_percentage_match = re.search(r"VAT\s*:\s*(\d+\.\d+)%", text, re.IGNORECASE)
    vat_percentage = vat_percentage_match.group(1).strip() if vat_percentage_match else ""

    # Extract customer using line matching
    lines = text.split('\n')

    # Find the line index containing "Customer:"
    customer_index = None
    customer_code = None
    for i, line in enumerate(lines):
        if "Customer :" in line:
            customer_index = i
            match = re.search(r"(?<=Customer : )\w+", line)  # Extract customer code (e.g., lnk123)
            if match:
                customer_code = match.group()
            break

    # Extract the customer name and address
    customer = ""
    if customer_index is not None and customer_code:
        # Find the indices of "Order_Date" and "Delivery_Note_No" lines
        order_date_index = None
        delivery_note_index = None
        for j, line in enumerate(lines[customer_index:]):
            if "Order_Date:" in line:
                order_date_index = customer_index + j
            elif "Delivery_Note_No:" in line:
                delivery_note_index = customer_index + j
                break  # Stop searching once Delivery_Note_No is found

        # Extract customer details between Order_Date and Delivery_Note_No, excluding dates
        if order_date_index is not None and delivery_note_index is not None:
            for line in lines[order_date_index + 1:delivery_note_index]:
                if line and "Order_No:" not in line and not re.match(r"\d{2}/\d{2}/\d{2}",
                                                                     line):  # Exclude lines with date format
                    customer += line.strip() + " "

    customer = f"{customer_code} {customer.strip()}"  # Add customer code and remove trailing whitespace

    return cu_in, cu_sn, Document_Type, customer, PIN, INVOICE_NO, Invoice_Date, PIN_No, total, vat_percentage, pdf_file

# Define the function to write onto excel worksheet

def write_to_excel(details_list, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["cu_in","cu_sn","Document_Type","customer","PIN","INVOICE_NO","Invoice_Date","PIN_No","total","vat_percentage", "PDF_File"])
    for details in details_list:
        sheet.append(details)
    workbook.save(output_file)

# Define the function to browse the PDF file using file dialog

def browse_files():
    file_names = ctk.CTkFileDialog.askopenfilenames(
        title="Select PDF files",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
    )
    for file_name in file_names:
        listbox.insert("end", file_name)

# Define the function to extract text from PDF files

def extract_text():
    pdf_files = list(listbox.get(0, "end"))
    details_list = []
    for pdf_file in pdf_files:
        details = extract_details_from_pdf(pdf_file)
        details_list.append(details)
    output_file = "TaxInvoice.xlsx"
    write_to_excel(details_list, output_file)
    ctk.CTkMessageBox.showinfo(
        title="Text Extraction Completed",
        message=f"Details extracted from PDFs and saved to {output_file}."
    )

# Create a customtkinter window

root = ctk.CTk()

# Set the window title

root.title("PDF Details Extractor")

# Create a customtkinter label to display the instruction

instruction_label = ctk.CTkLabel(
    root,
    text="Select PDF Files:"
)
instruction_label.pack(pady=10)

# Create a textbox to display selected PDF files

listbox = ctk.CTkTextbox(
root,
width=50,
height=10
)
listbox.pack()

# Create a button to browse for PDF files

browse_button = ctk.CTkButton(
    root,
    text="Browse",
    command=browse_files
)
browse_button.pack()

# Create radiobuttons to select details to include

selected_option = StringVar()

cu_in_radiobutton = ctk.CTkRadiobutton(
    root,
    text="CU IN",
    variable=selected_option,
    value="CU IN"
)
cu_in_radiobutton.pack()

cu_sn_radiobutton = ctk.CTkRadiobutton(
    root,
    text="CU SN",
    variable=selected_option,
    value="CU SN"
)
cu_sn_radiobutton.pack()

document_type_radiobutton = ctk.CTkRadiobutton(
    root,
    text="Document Type",
    variable=selected_option,
    value="Document Type"
)
document_type_radiobutton.pack()

customer_radiobutton = ctk.CTkRadiobutton(
    root,
    text="Customer",
    variable=selected_option,
    value="Customer"
)
customer_radiobutton.pack()

pin_radiobutton = ctk.CTkRadiobutton(
    root,
    text="PIN",
    variable=selected_option,
    value="PIN"
)
pin_radiobutton.pack()

invoice_no_radiobutton = ctk.CTkRadiobutton(
    root,
    text="Invoice No",
    variable=selected_option,
    value="Invoice No"
)
invoice_no_radiobutton.pack()

invoice_date_radiobutton = ctk.CTkRadiobutton(
    root,
    text="Invoice Date",
    variable=selected_option,
    value="Invoice Date"
)
invoice_date_radiobutton.pack()

pin_no_radiobutton = ctk.CTkRadiobutton(
    root,
    text="PIN No",
    variable=selected_option,
    value="PIN No"
)
pin_no_radiobutton.pack()

total_radiobutton = ctk.CTkRadiobutton(
    root,
    text="Total",
    variable=selected_option,
    value="Total"
)
total_radiobutton.pack()

vat_percentage_radiobutton = ctk.CTkRadiobutton(
    root,
    text="VAT Percentage",
    variable=selected_option,
    value="VAT Percentage"
)
vat_percentage_radiobutton.pack()

# Create a button to extract text from PDF files

extract_button = ctk.CTkButton(
    root,
    text="Extract Text",
    command=extract_text
)
extract_button.pack(pady=10)

# Start the main loop of the window

root.mainloop()
