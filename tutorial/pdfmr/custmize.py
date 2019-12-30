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
        
        