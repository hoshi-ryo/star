import os
import openpyxl
import sys 


print(sys.argv[1])
print(os.listdir(sys.argv[1]))
path = sys.argv[1]
fl =  os.listdir(sys.argv[1])
for f in fl:
    print(f)
    print(f[-5:])
    if f[-5:]== ".xlsx" :
        fp =    path+"/"+f
        wb = openpyxl.load_workbook(fp)
        sheets = wb.sheetnames
        for sn in sheets:
            ws = wb[sn]
            leftheader = str(ws.oddHeader.left)
            centerheader = str(ws.oddHeader.center)
            rightheader = str(ws.oddHeader.right)
            print(f+","+sn+","+leftheader+","+centerheader+","+rightheader)
    else:
     continue


