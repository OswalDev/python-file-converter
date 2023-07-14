import os
import subprocess
from docx2pdf import convert as convert_to_pdf
from openpyxl import load_workbook
from pptx import Presentation


def convert_to_pdf(input_file, output_file):
    if input_file.lower().endswith(('.doc', '.docx')):
        try:
            convert_to_pdf(input_file, output_file=output_file)  # Corrected line
            print(f"Converted {input_file} to PDF successfully.")
        except Exception as e:
            print(f"Failed to convert {input_file} to PDF: {e}")
    elif input_file.lower().endswith(('.pptx')):
        try:
            presentation = Presentation(input_file)
            presentation.save(output_file)
            print(f"Converted {input_file} to PDF successfully.")
        except Exception as e:
            print(f"Failed to convert {input_file} to PDF: {e}")
    elif input_file.lower().endswith(('.xls', '.xlsx')):
        try:
            workbook = load_workbook(input_file)
            workbook.save(output_file)
            print(f"Converted {input_file} to PDF successfully.")
        except Exception as e:
            print(f"Failed to convert {input_file} to PDF: {e}")
    elif input_file.lower().endswith('.pdf'):
        try:
            subprocess.run(['cp', input_file, output_file])
            print(f"Converted {input_file} to {output_file} successfully.")
        except Exception as e:
            print(f"Failed to convert {input_file} to {output_file}: {e}")
    else:
        print('Unsupported file format.')


# Example usage - Convert from one format to another
input_file = './test_file.docx'
output_file = os.path.join('./convertedFiles', 'output.pdf')
convert_to_pdf(input_file, output_file)

# Example usage - Convert from PDF to another format
input_file = './ICE-card.pdf'
output_file = os.path.join('./convertedFiles', 'output.docx')
convert_to_pdf(input_file, output_file)
