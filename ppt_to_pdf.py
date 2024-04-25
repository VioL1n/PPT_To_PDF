import os
import comtypes.client

def ppt_to_pdf(ppt_path, pdf_path):
    # Load PowerPoint application
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    # Open the PPT file
    presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
    # Save as PDF
    presentation.SaveAs(pdf_path, FileFormat=32)  # 32 stands for PDF format
    # Close the presentation and quit PowerPoint
    presentation.Close()
    powerpoint.Quit()

def convert_ppts_in_directory(directory):
    # List all files in the given directory
    files = os.listdir(directory)
    # Filter PPT files
    ppt_files = [f for f in files if f.endswith(('.ppt', '.pptx'))]
    # Convert each PPT file to PDF
    for ppt in ppt_files:
        ppt_path = os.path.join(directory, ppt)
        pdf_path = os.path.join(directory, os.path.splitext(ppt)[0] + '.pdf')
        ppt_to_pdf(ppt_path, pdf_path)
        print(f'Converted {ppt} to PDF.')

# Use the function on the current directory
convert_ppts_in_directory(os.getcwd())
