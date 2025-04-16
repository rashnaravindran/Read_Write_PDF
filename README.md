# Read_Write_PDF
Read text from images from PDF file and compares data with excel and write the alternate text in PDF file

📝 Project Overview
This project automates the process of annotating Amazon order stickers with item details.

🔧 How It Works
The script reads a PDF file named DF.pdf, located inside the PDF folder.

Each page in the PDF is assumed to contain a single Amazon order sticker image.

For each page:

The image is extracted and saved in PDF folder with Page number as file name (Eg: page1_img1.png, page2_img2.png).

The Amazon Order ID is detected from the image using OCR (e.g., Tesseract).

The script then looks up the extracted Order ID in an Excel file named DF.xlsx, also located in the PDF folder.

This Excel sheet contains mappings of Order ID → SKU and Item Title.

The corresponding SKU and Item Title are then printed (annotated) directly onto the original image/page.

The updated pages are saved as a new PDF (DF_O.pdf) in PDF folder, with each page now labeled with the associated item details.

📂 Folder Structure
Read_Write_PDF/
├── PDF/
│   ├── DF.pdf        # Input PDF containing Amazon order sticker images (1 per page)
│   ├── DF.xlsx       # Excel file mapping Order ID → SKU & Item Title
├── imgfrompdf.py    # Your main processing script
├── README.md         # This file
💡 Dependencies
Python 3.x

fitz from PyMuPDF (for PDF handling)

pandas (for reading Excel)

PIL (for image annotation)

pytesseract (for OCR)

📌 Notes
Make sure the PDF/DF.pdf and PDF/DF.xlsx files exist before running the script.

You can change the paths in the script if needed.
