import re
import openpyxl
import xlrd
file="word.xlsx"
#EXCEL的檔名

data=xlrd.open_workbook(file)
#打開EXCEL檔

sh=data.sheet_by_name('1')
wb = openpyxl.load_workbook("word.xlsx")
ws1 = wb.get_sheet_by_name('1')
#讀取EXCEL的工作表1
k=2
rows=sh.row_values(1)
column=['',''
        ,'C','D','E','F','G','H','I','J','K','L'
        ,'M','N','O','P','Q','R','S','T','U','V'
        ,'W','X','Y','Z','AA','AB','AC','AD','AE'#29
        ,'AF','AG','AH','AI','AJ','AK','AL'
        ,'AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV'
        ,'AW','AX','AY','AZ',
        'BA','BB','BC','BD','BE'
        ,'BF','BG','BH','BI','BJ','BK','BL'
        ,'BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV'
        ,'BW','BX','BY','BZ',
        'CA','CB','CC','CD','CE'#29
        ,'CF','CG','CH','CI','CJ','CK','CL'
        ,'CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV'
        ,'CW','CX','CY','CZ',
        'DA','DB','DC','DD','DE'#29
        ,'DF','DG','DH','DI','DJ','DK','DL'
        ,'DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV'
        ,'DW','DX','DY','DZ',
        'EA','EB','EC','ED','EE'#29
        ,'EF','EG','EH','EI','EJ','EK','EL'
        ,'EM','EN','EO','EP','EQ','ER','ES','ET','EU','EV'
        ,'EW','EX','EY','EZ',
        'FA','FB','FC','FD','FE'#29
        ,'FF','FG','FH','FI','FJ','FK','FL'
        ,'FM','FN','FO','FP','FQ','FR','FS','FT','FU','FV'
        ,'FW','FX','FY'] 
#將EXCEL的第一行關鍵字欄位建立成串列

sheet=['','']
#前兩列是服務業百大企業跟公司名稱，不用讀

for i in range(1,6):
    for j in range(2,len(column)):
        sheet.append(column[j]+str(i+2))



for i in range(1,6,1):
    i=str(i)
    f = open(i+".txt",'r',encoding = 'UTF-8')
    source = f.read()
    #打開資料夾裡面的TXT檔，依序讀入，所以檔案的命名要用數字依序排列下來
    for j in range(2,101,1):
        ws1[sheet[k]] = len(re.findall(rows[j],source))
        k=k+1
wb.save("word.xlsx")






    


    


