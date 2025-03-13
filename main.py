import os
import win32com.client

# Function to convert Word to PDF
def convert_word_to_pdf(input_path, output_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
    doc.Close()
    word.Quit()

# Function to convert PowerPoint to PDF
def convert_ppt_to_pdf(input_path, output_path):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt.Presentations.Open(input_path, WithWindow=False)
    presentation.SaveAs(output_path, FileFormat=32)  # 32 = PDF
    presentation.Close()
    ppt.Quit()

# Main part

input_folder = r"C:\Users\YourName\Desktop\input_files"  #Change these paths to your folders
output_folder = r"C:\Users\YourName\Desktop\output_pdfs"

# Make output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Go through each file in the input folder
for file_name in os.listdir(input_folder):
    file_path = os.path.join(input_folder, file_name)
    name, ext = os.path.splitext(file_name)
    ext = ext.lower()

    output_path = os.path.join(output_folder, name + ".pdf")

    try:
        if ext in ['.doc', '.docx']:
            print("Converting Word file:", file_name)
            convert_word_to_pdf(file_path, output_path)
        elif ext in ['.ppt', '.pptx']:
            print("Converting PowerPoint file:", file_name)
            convert_ppt_to_pdf(file_path, output_path)
        else:
            print("Skipping (not Word or PowerPoint):", file_name)
    except Exception as e:
        print("Error converting", file_name, "->", e)
