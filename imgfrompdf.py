import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import pandas as pd
import re

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pdf_path = "PDF/DF.pdf"
doc = fitz.open(pdf_path)

# Load Excel file
excel_path = "PDF/DF.xlsx"  # Replace with your file path
df = pd.read_excel(excel_path)  # Assumes first column contains text to search, second column contains text to add

position = (540, 320)

def extract_nine_letter_words(paragraph):
    return re.findall(r'\b\w{9}\b', paragraph)

def find_similar_word(search_word, word_list):
    for word in word_list:
        if len(word) == len(search_word):  # Ensure same length
            diff_count = sum(1 for a, b in zip(search_word, word) if a != b)
            if diff_count <= 1:  # Only one letter difference
                return word
    return None  # No match found

for page_num in range(len(doc)):
    for img_index, img in enumerate(doc[page_num].get_images(full=True)):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_ext = base_image["ext"]
        with open(f"pdf/page{page_num+1}_img{img_index+1}.{image_ext}", "wb") as img_file:
            img_file.write(image_bytes)
            imgfilepath = img_file.name
            print(imgfilepath)
            image = Image.open(imgfilepath)  # Replace with your image file
            # Extract text
            text = pytesseract.image_to_string(imgfilepath)
            # Convert extracted text into a set of words
            
            found_words = set((text.lower()).split())
            # found_words = {s.lower()
            for index,row in df.iterrows():
                matchedtext = ""
                #search_text = set(df.iloc[:, 0].astype(str)) #
                search_text = str(row[0])  # Text to search

                # Split text into lines
                max_length = 50 
                words = str(row[3]).split()
                lines = []
                line = ""
                formatted_text = ""
                line_length = 0

                for word in words:
                    if line_length + len(word) + 1 <= max_length:  # +1 for space
                        formatted_text += word + " "
                        line_length += len(word) + 1
                    else:
                        formatted_text = formatted_text.rstrip() + "\n" + word + " "
                        line_length = len(word) + 1

                insert_text = str(row[2]) + "\n" + str(formatted_text) # Text to insert if found
                #insert_text = str(insert_text)
                 # Find matches
               
                #matches = {search_text.lower()}.intersection(found_words)
                #print(found_words,search_text.lower())
                similar_words = find_similar_word(search_text.lower(), found_words)

                #for match in matches:
                #    row_index = df.index[df[0] == match].tolist()
                #if matches:
                if similar_words:
                    #print(found_words,':', matches)
                    matchedtext = search_text
                    print("Extracted Text:",page_num, "-",imgfilepath , " " ,search_text)
                    doc[page_num].insert_text(position, insert_text, fontsize=12, color=(0, 0, 1), rotate=90)  # Blue color
                    break
                                   
            if (matchedtext==""):
                print("Not Matched:", page_num, '-',search_text, " - ", extract_nine_letter_words(text[-50:]))
            

# Define text and position
text = "Confidential"
position = (20, 250)  # (x, y) coordinates

# Iterate through pages and add text
#for page in doc:
#    page.insert_text(position, text, fontsize=12, color=(1, 0, 0), rotate=90)  # Red color

# Save the new PDF
doc.save("PDF/DF_O.pdf")
doc.close()


print("Images extracted successfully!")
