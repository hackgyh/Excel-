# Excel
## 由于最近经常处理表格的增添删改工作，所以写一个简单的程序辅助工作
```
import xlrd
import xlwt

book = xlrd.open_workbook('stu_1.xls')#打开一个excel
sheet = book.sheet_by_index(0)#根据顺序获取sheet
sheet2 = book.sheet_by_name('case1_sheet')#根据sheet页名字获取sheet
a=[]
b=[]

for m in range(sheet.nrows):#sheet.nrows是行数
    for n in range(sheet.ncols):#sheet.ncols是列数
        a.append(str(sheet.cell(m,n).value))            #构建列表
    b.append(a)
    a=[]
    
cd='0'
while cd!='q':
    c=input("姓名：")
    d=input("成绩：")

    for k in b:
        if k[3]!="分数" and k[3]!="":
            k[3]=float(k[3])
        if k[1]!="年龄":
            k[1]=float(k[1])
        if k[0]==c:
            k[3]=float(d)
    cd=input("回车继续,q退出\n")
    cb=0
        
book = xlwt.Workbook()#新建一个excel
sheet = book.add_sheet('case1_sheet')#添加一个sheet页
row = 0#控制行表格
for stu in b:
    col = 0#控制列表格
    for s in stu:                              #循环列表
        sheet.write(row,col,s)
        col+=1
    row+=1
book.save('stu_1.xls')#保存到当前目录下
```
