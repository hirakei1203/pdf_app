import openpyxl
import datetime

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
        dt_now = datetime.datetime.now()
        # 複数のpdfファイルの際にsheetを複製していくコードが課題
        for i in range(len(result_list)):
            if i >= 1:
                ws = book.worksheets[0]
                book.copy_worksheet(ws)
                if i == 1:
                    sheet = book['SAP処理依頼書 Copy']
                elif i == 2:
                    sheet = book['SAP処理依頼書 Copy1']
                elif i == 3:
                    sheet = book['SAP処理依頼書 Copy2']
                elif i == 4:
                    sheet = book['SAP処理依頼書 Copy3']
            
            allText = result_list[i].split("\n")
            length = len(allText)
            year = dt_now.year
            month = dt_now.month
            day = dt_now.day
            hosoku = allText[22]
            if length > 86: 
                bikou = allText[45] + ", " + allText[46] + ", " + allText[47] + ", " + allText[48]
                allText.reverse()
                price = allText[2]
            else: 
                bikou = allText[45] + ", " + allText[46]
                allText.reverse()
                price = allText[2]

            cell_hosoku = 'D20'
            cell_bikou = 'D22'
            cell_price = 'N28'
            cell_total_price = 'N41'
            cell_year = 'N1'
            cell_month = 'P1'
            cell_day = 'R1'
            
            sheet[cell_hosoku] = hosoku
            sheet[cell_bikou] = bikou
            sheet[cell_price] = price
            sheet[cell_total_price] = price
            sheet[cell_year] = year
            sheet[cell_month] = month
            sheet[cell_day] = day
            
            # elif  i == 1:
                
            # elif  i == 2:
            #     sheet = book['SAP処理依頼書 Copy2']
            # elif  i == 3:
            #     sheet = book['SAP処理依頼書 Copy3']
            
            book.save(temp_file)
            

            
    except Exception as e:
            err_message ="Excelファイルへのデータ転記処理でエラーが発生しました。<br>\
            アップロードしたPDFファイルが正しいフォーマットか確認してください。<br>\
            エラーメッセージ：" + str(e)
            return err_message