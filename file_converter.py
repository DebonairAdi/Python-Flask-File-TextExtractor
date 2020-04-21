# package to convert docx files
import docx2txt

# package to convert pdf files
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

# input output string operations and conversions
from io import StringIO 


def pdf_to_text(filename):
    """
    this method uses the pdfminer library to 
    convert pdf file and return text in the 
    form of a string.
    """
    resource_manager = PDFResourceManager() # manages formatting
    str_implementation = StringIO() # string conversions
    codec = 'utf-8' # encoding of the file 
    laparams = LAParams() # layout parameters 
    device = TextConverter(resource_manager, str_implementation, codec=codec, laparams=laparams)
    current_file = open(filename, 'rb') 
    interpreter = PDFPageInterpreter(resource_manager, device)
    password = "";maxpages = 0;caching = True;page_num=set()
    for page in PDFPage.get_pages(current_file, page_num, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    current_file.close()
    device.close()
    converted_text = str_implementation.getvalue()
    str_implementation.close()
    print(converted_text)
    return converted_text


def docx_to_text(filename):
    """
    this method uses the docx2txt library to 
    convert docx file and return text in the 
    form of a string.
    """
    converted_text = docx2txt.process(filename)
    print(converted_text)
    return converted_text


def file_to_text_converter(filename):
    """
    this is the main function that accepts 
    filename as a parameter and passes 
    through conditions based on filename
    """

    # for docx files
    if filename[-5:] == ".docx":
        return docx_to_text(filename)
        
    # for pdf files
    elif filename[-4:] == ".pdf":
        return pdf_to_text(filename)



# maincaller
if __name__ == "__main__":
    file_to_text_converter("New Leave Policy_Communication v2.pdf")