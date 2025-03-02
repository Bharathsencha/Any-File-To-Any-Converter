import os
import sys
import win32com.client

def convert_word_to_pdf(input_path, output_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in the background
    try:
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 is the format for PDF
        doc.Close()
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}")
    finally:
        word.Quit()

def convert_excel_to_pdf(input_path, output_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(input_path)
        wb.ExportAsFixedFormat(0, output_path)  # 0 is for PDF format
        wb.Close(False)
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}")
    finally:
        excel.Quit()

def convert_ppt_to_pdf(input_path, output_path):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = False
    try:
        presentation = ppt.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)  # 32 is the format for PDF
        presentation.Close()
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}")
    finally:
        ppt.Quit()

def convert_office_to_pdf(input_files, output_dir):
    for file in input_files:
        ext = os.path.splitext(file)[1].lower()
        base_name = os.path.splitext(os.path.basename(file))[0]
        output_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        if ext in ['.doc', '.docx']:
            convert_word_to_pdf(file, output_path)
        elif ext in ['.xls', '.xlsx']:
            convert_excel_to_pdf(file, output_path)
        elif ext in ['.ppt', '.pptx']:
            convert_ppt_to_pdf(file, output_path)
        else:
            print(f"Unsupported file format: {file}")

def main():
    if len(sys.argv) < 3:
        print("Usage: python script.py <output_directory> <file1> <file2> ...")
        sys.exit(1)
    
    output_dir = sys.argv[1]
    input_files = sys.argv[2:]
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    convert_office_to_pdf(input_files, output_dir)
    print("Conversion complete!")

if __name__ == "__main__":
    main()
