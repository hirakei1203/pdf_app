from pdfminer.pdffinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import bPDFPage
from io import StringIO

def convert_pdf_to_txt(path):
    
    rsrcmgr = PDFResource<anager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    laparams.detect_vertical = True
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    maxpages = 0
    fstr =''
    for page in PDFPage.get_pages(fp, maxpages=maxpages):
        interpreter.process_page(page)
        
        str = retstr.getvalue()
        fstr += str
        
    fp.close()
    device.close()
    retstr.close()
    return fstr