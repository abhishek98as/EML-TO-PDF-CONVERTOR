import os
import email
import quopri
from bs4 import BeautifulSoup
from datetime import datetime
import pdfkit
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import threading

def decode_email_payload(payload):
    """
    Decode the email payload and return decoded content.
    """
    charset = payload.get_content_charset()
    content_transfer_encoding = payload.get('Content-Transfer-Encoding')

    if content_transfer_encoding == 'quoted-printable':
        return quopri.decodestring(payload.get_payload()).decode(charset)
    else:
        return payload.get_payload(decode=True).decode(charset)

def eml_to_html(eml_file):
    """
    Convert .EML file to HTML content and extract email details.
    """
    with open(eml_file, 'r', encoding='utf-8') as file:
        msg = email.message_from_file(file)

    sender_email = msg["From"]
    receiver_email = msg["To"]
    date_time_str = msg["Date"]
    date_time_obj = datetime.strptime(date_time_str, "%a, %d %b %Y %H:%M:%S %z")
    subject = msg["Subject"]

    # Constructing the HTML header with email details
    header_html = f"""
    <html>
    <head>
    <title>Email Details</title>
    <style>
    .header {{
        text-align: center;
        line-height: 2;
    }}
    .header p {{
        margin: 0;
    }}
    </style>
    </head>
    <body>
    <div class="header">
    <p>From: {sender_email}</p>
    <p>To: {receiver_email}</p>
    <p>Subject: {subject}</p>
    <p>Date and Time: {date_time_obj.strftime("%Y-%m-%d %H:%M:%S %Z")}</p>
    </div>
    """

    # Extract HTML content from .EML file
    html_content = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                html_content += decode_email_payload(part)
            elif part.get_content_type() == 'text/plain':
                # Convert plain text to HTML
                html_content += f"<p>{decode_email_payload(part)}</p>"
    else:
        html_content += decode_email_payload(msg)

    # Use BeautifulSoup to parse the HTML and remove unwanted headers
    soup = BeautifulSoup(html_content, 'html.parser')
    body_content = str(soup.body)

    # Return the full HTML content
    return header_html + body_content

def convert_html_to_pdf(html_content, output_pdf_file):
    """
    Convert HTML content to PDF with A4 size.
    """
    options = {
        'page-size': 'A4',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': 'UTF-8',  # Specify encoding
    }

    pdfkit.from_string(html_content, output_pdf_file, options=options)

def bulk_convert_eml_to_pdf(eml_files, output_folder):
    """
    Function to handle bulk conversion of .eml files to .pdf.
    """
    total_files = len(eml_files)
    progress_step = 100 / total_files

    for i, eml_file in enumerate(eml_files, start=1):
        try:
            # Convert .EML to HTML
            html_content = eml_to_html(eml_file)

            # Construct output PDF file path
            output_pdf_file = os.path.join(output_folder, f"BulkPDF{i}.pdf")

            # Convert HTML to PDF
            convert_html_to_pdf(html_content, output_pdf_file)

            # Update progress bar
            progress_var.set(i * progress_step)
            root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Conversion Error", f"Error converting {eml_file}: {str(e)}")

    # Show success message
    messagebox.showinfo("Bulk Conversion Successful", f"{total_files} files conversion completed!")
    # Clear selected file label
    selected_file_label.config(text="")

def bulk_convert():
    """
    Function to handle bulk conversion of .eml files to .pdf.
    """
    # Prompt user to select multiple .eml files
    eml_files = filedialog.askopenfilenames(filetypes=[("EML files", "*.eml")])
    if not eml_files:
        return

    # Update selected file label
    selected_file_label.config(text=f"Selected Files: {len(eml_files)} files")

    # Prompt user to select folder to save the PDF files
    output_folder = filedialog.askdirectory()
    if not output_folder:
        return

    # Start bulk conversion in a separate thread
    threading.Thread(target=bulk_convert_eml_to_pdf, args=(eml_files, output_folder)).start()

# Create main window
root = tk.Tk()
root.title("EML to PDF Converter")

# Create label to display selected file
selected_file_label = tk.Label(root, text="")
selected_file_label.pack(pady=5)

# Create button for bulk conversion
bulk_convert_button = tk.Button(root, text="Bulk Convert", command=bulk_convert)
bulk_convert_button.pack(pady=10)

# Create progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, length=300, mode='determinate')
progress_bar.pack(pady=10)

# Run the GUI
root.mainloop()
