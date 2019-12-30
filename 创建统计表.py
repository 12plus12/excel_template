import xlwings as xw
import datetime
import numpy


def get_last_month(count):
	# 获取上个月上上个月的月份 参数为-1或-2
	month_now = datetime.datetime.now().month
	if month_now == 1:
		return 13 + count
	elif month_now == 2 and count == -2:
		return 14 + count
	else:
		return month_now + count

# 记录表格的年份
year = datetime.datetime.now().year
month = datetime.datetime.now().month
if month == 1:
	year -= 1

app = xw.App(visible=False, add_book=False)
# 获取销售表中数据
workbook_sale = app.books.open('销售出库.xls')
sheet_sale = workbook_sale.sheets[0]
date_sale_list = sheet_sale.range('A2').expand('down').value
ticketnum_sale_list = sheet_sale.range('B2').expand('down').value
name_sale_list = sheet_sale.range('D2').expand('down').value
size_sale_list = sheet_sale.range('E2').expand('down').value
unit_sale_list = sheet_sale.range('F2').expand('down').value
# print(unit_sale_list)
count_sale_list = sheet_sale.range('G2').expand('down').value  
# print(count_sale_list)
customer_sale_list = sheet_sale.range('O2').expand('down').value
# print(customer_sale_list)
address_sale_list = sheet_sale.range('Q2').expand('down').value
supplier_sale_list = sheet_sale.range('R2').expand('down').value
if type(date_sale_list) is not list:
	date_sale_list = [date_sale_list]
	ticketnum_sale_list = [ticketnum_sale_list]
	name_sale_list = [name_sale_list]
	size_sale_list = [size_sale_list]
	unit_sale_list = [unit_sale_list]
	count_sale_list = [count_sale_list]
	customer_sale_list = [customer_sale_list]
	address_sale_list = [address_sale_list]
	supplier_sale_list = [supplier_sale_list]
app.kill()

# 获取上上个月盘点库存数据
app = xw.App(visible=False, add_book=False)
workbook_last = app.books.open('{}月份视云商品库存盘点统计.xlsx'.format(get_last_month(-2)))
sheet_last = workbook_last.sheets[0]
for i in range(1, 100):
	if sheet_last.range('F' + str(i)).value == '{}月实际结余库存'.format(get_last_month(-2)):
		last_list = sheet_last.range('F' + str(i+1)).expand('down').value
		break
app.kill()


# 在盘点表填入销售表数据
app = xw.App(visible=False, add_book=False)
workbook_check = app.books.open('模板.xlsx')
sheet_check = workbook_check.sheets[0]
A1 = sheet_check.range('A1').value
sheet_check.range('A1').value = A1.format(year, get_last_month(-1))  # TODO 动态输入

num = 0
count_num_list = [0, 0, 0, 0]
name_list = ['片仔癀牙火清牙膏清火炫白（臻选留兰香）', '片仔癀牙火清牙膏清火清新（白茶薄荷）', '片仔癀牙火清牙膏清火护龈（臻选留兰香）', '片仔癀牙火清牙膏清火护龈（菁萃药香）']
for i in range(len(date_sale_list)):
	sheet_check.api.Rows(3 + num).Insert()
	sheet_check.range('A' + str(3 + num)).value = date_sale_list[i]
	sheet_check.range('B' + str(3 + num)).value = ticketnum_sale_list[i]
	sheet_check.range('C' + str(3 + num)).value = name_sale_list[i]
	sheet_check.range('D' + str(3 + num)).value = size_sale_list[i]
	sheet_check.range('E' + str(3 + num)).value = unit_sale_list[i]
	sheet_check.range('F' + str(3 + num)).value = count_sale_list[i]
	sheet_check.range('G' + str(3 + num)).value = customer_sale_list[i]
	sheet_check.range('H' + str(3 + num)).value = address_sale_list[i]
	sheet_check.range('I' + str(3 + num)).value = supplier_sale_list[i]
	count_num_list[name_list.index(name_sale_list[i])] += count_sale_list[i]
	num += 1
count_list = sheet_check.range('F3').expand('down').value
if type(count_list) is not list:
	count_list = [count_list]
res = 0
for count in count_list:
	res += count
sheet_check.range('F' + str(3 + num)).value = res

# 找到第二张表的标题所在单元格 遍历A1~A100（超出范围可以改为死循环遍历）
for i in range(1, 101):
	if sheet_check.range('A' + str(i)).value == '视云片仔癀各系列产品{}年{}月份库存盘点':
		table1_row_num = i
		break

# 构建第二张表格标题数据
sheet_check.range('A' + str(table1_row_num)).value = sheet_check.range('A' + str(table1_row_num)).value.format(year, get_last_month(-1))
sheet_check.range('C' + str(table1_row_num + 1)).value = sheet_check.range('C' + str(table1_row_num + 1)).value.format(get_last_month(-2))
sheet_check.range('D' + str(table1_row_num + 1)).value = sheet_check.range('D' + str(table1_row_num + 1)).value.format(get_last_month(-1))
sheet_check.range('E' + str(table1_row_num + 1)).value = sheet_check.range('E' + str(table1_row_num + 1)).value.format(get_last_month(-1))
sheet_check.range('F' + str(table1_row_num + 1)).value = sheet_check.range('F' + str(table1_row_num + 1)).value.format(get_last_month(-1))
sheet_check.range('H' + str(table1_row_num + 1)).value = sheet_check.range('H' + str(table1_row_num + 1)).value.format(get_last_month(-1))

# 在第二张表插入数据
sheet_check.range('C' + str(table1_row_num + 2)).value = numpy.array(last_list).reshape(len(last_list),1)
sheet_check.range('E' + str(table1_row_num + 2)).value = numpy.array(count_num_list).reshape(len(count_num_list),1)
# TODO

# 第三张表标题
sheet_check.range('A' + str(table1_row_num + 10)).value = sheet_check.range('A' + str(table1_row_num + 10)).value.format(year, get_last_month(-1))
sheet_check.range('A' + str(table1_row_num + 11)).value = sheet_check.range('A' + str(table1_row_num + 11)).value.format(get_last_month(-2))
sheet_check.range('A' + str(table1_row_num + 12)).value = sheet_check.range('A' + str(table1_row_num + 12)).value.format(get_last_month(-1))
sheet_check.range('A' + str(table1_row_num + 13)).value = sheet_check.range('A' + str(table1_row_num + 13)).value.format(get_last_month(-1))
sheet_check.range('E' + str(table1_row_num + 11)).value = sheet_check.range('E' + str(table1_row_num + 11)).value.format(get_last_month(-1))
sheet_check.range('E' + str(table1_row_num + 12)).value = sheet_check.range('E' + str(table1_row_num + 12)).value.format(get_last_month(-1))
sheet_check.range('G' + str(table1_row_num + 13)).value = datetime.datetime.now().strftime('%Y/%m/%d')


workbook_check.save('{}月份视云商品库存盘点统计【test】.xlsx'.format(get_last_month(-1)))
app.kill() 