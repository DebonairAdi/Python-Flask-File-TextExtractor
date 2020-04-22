# package to convert docx files
import docx2txt

# package to convert pdf files
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

# input output string operations and conversions
from io import StringIO 

# package to write to excel
import xlsxwriter

# system path related
import os

# list declared in which all data will be appended
data_list=[]

def path_to_folder(filename):
    """
        this method is used to add the path
        of the uploaded files folder with
        the filenames
    """
    folder_path = "uploaded_files"
    full_path = os.path.join(folder_path, filename)
    return full_path


def pdf_to_text(filename):
    """
        this method uses the pdfminer library to 
        convert pdf file and return text in the 
        form of a string.
    """
    resource_manager = PDFResourceManager() # manages formatting
    str_implementation = StringIO()         # string conversions
    codec = 'utf-8'                         # encoding of the file 
    laparams = LAParams()                   # layout parameters 
    device = TextConverter(resource_manager, str_implementation, codec=codec, laparams=laparams)
    filename_with_path=path_to_folder(filename)
    current_file = open(filename_with_path, 'rb') 
    interpreter = PDFPageInterpreter(resource_manager, device)
    password = "";maxpages = 0;caching = True;page_num=set()
    for page in PDFPage.get_pages(current_file, page_num, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    current_file.close()
    device.close()
    converted_text = str_implementation.getvalue()
    str_implementation.close()
    return data_list.extend([filename,converted_text]) 


def docx_to_text(filename):
    """
        this method uses the docx2txt library to 
        convert docx file and return text in the 
        form of a string.
    """
    filename_with_path=path_to_folder(filename)
    converted_text = docx2txt.process(filename_with_path)
    return data_list.extend([filename,converted_text])


def write_to_excel(): 
    """
        this method is used to write the
        converted text into excel file
    """
    workbook = xlsxwriter.Workbook('extracted_text.xlsx')   # excel file to be genertaed
    worksheet = workbook.add_worksheet()                    # add sheet to excel file
    row = 0                                                 # default row value
    col = 0                                                 # default col value
    i = 0                                                   # variable to auto inc.
    for data in range(0,int((len(data_list))/2)):
        worksheet.write(row, col, data_list[i])             # write to excel
        worksheet.write(row, col + 1, data_list[i+1]) 
        row += 1
        i+=2
    workbook.close()                                        # excel closed after writing data
    return print("Successfully written to excel file")


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

    # for other files rather than pdf or docx
    else:
        return print(filename+" can't convert this file, only pdf and docx files only")


def main():
    """
        this is the main method which
        will be called from the flask 
        application
    """
    try:
        uploaded_files=os.listdir("uploaded_files") # path to folder containing files
    except:
        print("folder named uploaded_files was not found")
    for file_to_convert in uploaded_files:          # to get and operate each file in folder
        file_to_text_converter(file_to_convert)     # method to convert files
    write_to_excel()                                # write data to excel


if __name__ == "__main__":
    pass