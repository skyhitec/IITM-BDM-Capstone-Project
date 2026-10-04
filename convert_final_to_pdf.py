import os
import win32com.client

def docx_to_pdf(docx_path, pdf_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        print(f"CONVERTED TO PDF: {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path}: {e}")
    finally:
        word.Quit()

if __name__ == "__main__":
    f1_docx = 'c:\\Users\\HP\\Downloads\\BDM\\23f1000204ds.study.iitm.ac.in - Final Report.docx'
    f1_pdf = 'c:\\Users\\HP\\Downloads\\BDM\\23f1000204ds.study.iitm.ac.in - Final Report.pdf'
    
    f2_docx = 'c:\\Users\\HP\\Downloads\\BDM\\Final_Report.docx'
    f2_pdf = 'c:\\Users\\HP\\Downloads\\BDM\\Final_Report.pdf'
    
    docx_to_pdf(f1_docx, f1_pdf)
    docx_to_pdf(f2_docx, f2_pdf)
