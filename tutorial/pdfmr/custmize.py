import openpyxl

def merge_excel(book, result_list, temp_file)
    sheet = book['請求書一覧']
    index = 6
    for i in range(len(result_list)):
        allText = result_list[i].split("\n\n")
        _, seikyu_no = allText[3].split()
        company_name = allText[4]
        _, bill = allText[7].split(" ")
        _, meigi = allText[24].split(":")
        kouza = allText[25]
        
        cell_b = 'B' + str(index+i)
        cell_c = 'C' + str(index+i)
        cell_d = 'D' + str(index+i)
        cell_e = 'E' + str(index+i)
        cell_f = 'F' + str(index+i)
        
        sheet[cell_b] = seikyu_no
        sheet[cell_c] = company_name
        sheet[cell_d] = bill
        sheet[cell_e] = meigi
        sheet[cell_f] = kouza
        
        book.save(temp_file)