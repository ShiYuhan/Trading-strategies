# Trading-strategies
Trading strategies based on Directional Change Framework
import xlrd
import xlwt
import xlutils
from xlutils.copy import copy
from xlrd import xldate_as_tuple
data = xlrd.open_workbook("SH2.xlsx")
excel = copy(data)
table_wt = excel.get_sheet(0)
table = data.sheets()[0] # 通过索引顺序获取
nrows = table.nrows
ncols = table.ncols
index_list = []
date_list = []
for row in range(0, nrows):
    index_value = table.cell(row, 1).value
    index_list.append(index_value)
    
    date_value=table.cell(row,0).value
    date=xldate_as_tuple(date_value,0)
    date_list.append(date)

#--------初始化变量---------------
upturn_event = True
p_h = p_l = index_list[0]
dc_range = []
os_range = []
r=0
thresh=0.04
sum=100000
osup=0
osdown=0

#---------------寻找DCC和OS点----------
for i in range (len(index_list)):
    p_t = index_list[i]
    if upturn_event:
        if p_t <= p_h * (1 - thresh): #最高点下跌一个theta
            upturn_event = False  #更改up标志位
            p_l = p_t #低点为当前点

            dc_range.append(index_list[i])
            os_range.append(index_list[osup])
            table_wt.write(i,8,index_list[i])
            excel.save("SH2.xlsx")
            table_wt.write(i,10,index_list[osup])
            excel.save("SH2.xlsx")
            
        else:
            
            if p_h < p_t:
                p_h = p_t
                osup=i

    else: # if we're in a downturn
        if p_t >= p_l * (1 + thresh):
            upturn_event = True
            p_h = p_t
            dc_range.append(index_list[i])
            os_range.append(index_list[osdown])
            table_wt.write(i,9,index_list[i])
            excel.save("SH2.xlsx")
            table_wt.write(i,11,index_list[osdown])
            excel.save("SH2.xlsx")
        
        else:
            if p_l > p_t:
                p_l = p_t
                osdown=i


print (dc_range)

#----------profit----------------
for x in range (1,len (dc_range)-1,1):
    
    if x%2==1:
        profit=(dc_range[x+1]-dc_range[x])/dc_range[x]
        sum=sum*(1+profit)
    else:
        profit=(dc_range[x]-dc_range[x+1])/dc_range[x]
        sum=sum*(1+profit)
print (sum)



