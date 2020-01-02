import openpyxl

def merge_excel(book, result_list, temp_file):
    # try:
    #     sheet = book['請求書一覧']
    #     index = 6
    #     for i in range(len(result_list)):
    #         allText = result_list[i].split("\n\n")
    #         _, seikyu_no = allText[3].split()
    #         company_name = allText[4]
    #         _, bill = allText[7].split(" ")
    #         _,meigi = allText[24].split("：")
    #         kouza = allText[25]
            
    #         cell_b = 'B' + str(index+i)
    #         cell_c = 'C' + str(index+i)
    #         cell_d = 'D' + str(index+i)
    #         cell_e = 'E' + str(index+i)
    #         cell_f = 'F' + str(index+i)
            
    #         sheet[cell_b] = seikyu_no
    #         sheet[cell_c] = company_name
    #         sheet[cell_d] = bill
    #         sheet[cell_e] = meigi
    #         sheet[cell_f] = kouza
            
    #         book.save(temp_file)
    
    # D社version
    try: 
        sheet = book['SAP処理依頼書']
        # 複数のpdfファイルの際にsheetを複製していくコードが課題
        for i in range(len(result_list)):
            allText = result_list[i].split("\n\n")
            hosoku = allText[22]
            bikou = allText[45] + ", " + allText[46] + ", " + allText[47] + ", " + allText[48]
            allText.reverse()
            price = allText[2]
            # allText.reverse()　配列が逆になるよ！for price
            
            cell_hosoku = 'D20'
            cell_bikou = 'D22'
            cell_price = 'N28'
            
            sheet[cell_hosoku] = cell_hosoku
            sheet[cell_bikou] = cell_bikou
            sheet[cell_price] = cell_price
            
            book.save(temp_file)
            
    except Exception as e:
            err_message ="Excelファイルへのデータ転記処理でエラーが発生しました。<br>\
            アップロードしたPDFファイルが正しいフォーマットか確認してください。<br>\
            エラーメッセージ：" + str(e)
            return err_message