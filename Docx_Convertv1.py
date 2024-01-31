import os
import comtypes.client
import docx
import sys


def convert_doc(file_path):
    """this picks a docx file with a name and converts it to pdf
    :param file_path: file_path of docx
    """

    # load word document using Microsoft Word
    doc = docx.Document(file_path)

    # Open word app
    word = comtypes.client.CreateObject("Word.Application")

    docx_path = os.path.abspath(file_path)

    pdf_path = os.path.abspath(file_path.replace(".docx", ".pdf"))  # because the name isn't static
    # we will have to get a function/code to change the pdf file_name

    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    print("Converting file " + file_path + ".........")
    in_file.SaveAs(pdf_path, FileFormat=pdf_format)
    in_file.Close()
    print("Conversion completed")
    os.remove(file_path)  # a logging function can come in here to show the activities if they won't be displayed in the console
    print("File deleted")
    # Quit Microsoft Word
    word.Quit()

#do a check for if file exists prnt pdf file exists and exit


if sys.argv[1][0] == ".":
    new_path = sys.argv[1][2:]
    print(new_path, " detected")
    convert_doc(new_path)
else :
    print(sys.argv[1], " detected")
    convert_doc(sys.argv[1])



# if type(sys.argv[1]) == str and os.path.isfile(sys.argv[1]):
#     # file_p = sys.argv[1]
#
#     convert_doc(sys.argv[1])

