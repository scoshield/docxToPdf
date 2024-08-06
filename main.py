# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# from docx2pdf import convert
import sys
from spire.doc import *
from docx2pdf import convert
import random

def print_hi(name, inputfile, outfile):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

    try:
    #     document = Document()
    #     document.LoadFromFile('Final.docx')
          name2 = f'Final {random.randint(1, 1000)}'
          convert(inputfile, outfile)
    #     # Save the document as a PDF file
    #     document.SaveToFile('Final.pdf', FileFormat.PDF)
    #     print("PDF saved successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
    #     document.Close()
        print(f'Saved successfully {outfile}')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    if len(sys.argv) != 3:
        print("Usage: python main.py <input_file> <output_file>")
        # sys.exit(1)

    print_hi('PyCharm', sys.argv[1], f'{sys.argv[2]}.pdf')

