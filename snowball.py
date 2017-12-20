#-*- coding:utf-8 -*-
# Read urls from Excel file
# Parse information from itooza
# Write information to Excel file
import xlrd
import xlsxwriter
import os
from bs4 import BeautifulSoup
import urllib.request
import pickle
import getopt
import sys

def main():

	input_mode = 0
	try:
		opts, args = getopt.getopt(sys.argv[1:], "h", ["mode="])
	except getopt.GetoptError as err:
		print(err)
		sys.exit(2)
	for option, argument in opts:
		if option == "-h":
			help_msg = """
	input mode 0: Crawling 1: Pickle
			"""
			print(help_msg)
			sys.exit(2)
		elif option == "--mode":
			input_mode = int(argument)
	
	### PART I - Read Excel file
	num_stock = 2003
	#num_stock = 100
	
	input_file = "basic_20170729.xlsx"
	cur_dir = os.getcwd()
	workbook_name = input_file
	
	stock_cat_list = []
	stock_name_list = []
	stock_num_list = []
	stock_url_list = []
	
	workbook = xlrd.open_workbook(os.path.join(cur_dir, workbook_name))
	sheet_list = workbook.sheets()
	sheet1 = sheet_list[0]

	for i in range(num_stock):
	#for i in range(1600,1800):
		stock_cat_list.append(sheet1.cell(i+1,0).value)
		stock_name_list.append(sheet1.cell(i+1,1).value)
		stock_num_list.append(int(sheet1.cell(i+1,2).value))
		stock_url_list.append(sheet1.cell(i+1,3).value)

	### PART II - Read information from URLs
	eps_list = []
	bps_list = []
	dps_list = []
	roe_list = []
	avg_roe_list = []
	future_bps_list = []
	invest_price_list = []
	close_price_list = []
	expected_rate_list = []
	div_ratio_list = []

	if input_mode == 0:
		for j in range(num_stock):
			print(j, stock_name_list[j])
			if j%10 == 0: print (j)
		
			roe_sub_list = []

			url = stock_url_list[j]
			#print(url)
		
			handle = None
			while handle == None:
				try:
					handle = urllib.request.urlopen(url)
					#print(handle)
				except:
					pass

			data = handle.read()
			#print(data)
			soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
			
			# Find closing month
			table_close = soup.findAll('div', {'class':'detail-data'})

			#print(table_close[0])
			if table_close == []:
				close_month = 12
			else:
				lists = table_close[0].findAll('span')
				close_month = lists[1].text
			
			#print(type(close_month))

			# Find Price
			#table_price = soup.findAll('div', {'class':'item-detail'})
			#span = table_price[0].findAll('span')
			#stock_price = int(span[0].text.replace(',',''))

			# tr 0 : Column
			# tr 1 : EPS (connected)
			# tr 2 : EPS (individual)
			# tr 3 : PER
			# tr 4 : BPS
			# tr 5 : PBR
			# tr 6 : Dividens per share ***
			# tr 7 : Diviens ratio ***
			# tr 8 : ROE
			# tr 9 : ROS
			# tr 10 : ROA??
			
			# Find PBR
			table_quarter = soup.findAll('div', {'id':'indexTable3'})
			tr_list = table_quarter[0].findAll('tr')
			eps_tds = tr_list[1].findAll('td')
			
			if eps_tds[0].text == 'N/A':
				eps_1 = 0
			else:
				eps_1 = float(eps_tds[0].text.replace(',',''))
			if eps_tds[1].text == 'N/A':
				eps_2 = 0
			else:
				eps_2 = float(eps_tds[1].text.replace(',',''))
			if eps_tds[2].text == 'N/A':
				eps_3 = 0
			else:
				eps_3 = float(eps_tds[2].text.replace(',',''))
			if eps_tds[3].text == 'N/A':
				eps_4 = 0
			else:
				eps_4 = float(eps_tds[3].text.replace(',',''))

			eps =  eps_1 + eps_2 + eps_3 + eps_4
			
			bps_tds = tr_list[4].findAll('td')
			if bps_tds[0].text == 'N/A':
				bps = 0
			else:
				bps = float(bps_tds[0].text.replace(',',''))

			eps_list.append(eps)
			bps_list.append(bps)

			# Find DPS
			table_year = soup.findAll('div', {'id':'indexTable2'})
			tr_list = table_year[0].findAll('tr')
			td_list = tr_list[6].findAll('td')
			
			if close_month[0] == "3":
				print("March")
				if td_list[0].text == 'N/A':
					dps = 0
				else:
					dps = float(td_list[0].text.replace(',',''))
			else:
				if td_list[1].text == 'N/A':
					dps = 0
				else:
					dps = float(td_list[1].text.replace(',',''))
			
			dps_list.append(dps)
			
			roe_tds = tr_list[8].findAll('td')
			if roe_tds[5].text == 'N/A':
				#roe_sub_list.append("None")
				roe_sub_list.append(0.0)
			else:
				roe_4 = float(roe_tds[5].text.replace(',',''))
				roe_sub_list.append(float(roe_4))
			if roe_tds[4].text == 'N/A':
				#roe_sub_list.append("None")
				roe_sub_list.append(0.0)
			else:
				roe_3 = float(roe_tds[4].text.replace(',',''))
				roe_sub_list.append(float(roe_3))
			if roe_tds[3].text == 'N/A':
				#roe_sub_list.append("None")
				roe_sub_list.append(0.0)
			else:
				roe_2 = float(roe_tds[3].text.replace(',',''))
				roe_sub_list.append(float(roe_2))
			if roe_tds[2].text == 'N/A':
				#roe_sub_list.append("None")
				roe_sub_list.append(0.0)
			else:
				roe_1 = float(roe_tds[2].text.replace(',',''))
				roe_sub_list.append(float(roe_1))
			if roe_tds[1].text == 'N/A':
				#roe_sub_list.append("None")
				roe_sub_list.append(0.0)
			else:
				roe_0 = float(roe_tds[1].text.replace(',',''))
				roe_sub_list.append(float(roe_0))
			
			roe_list.append(roe_sub_list)
			
			table_price = soup.findAll('div', {'class':'item-detail'})
			span_price = table_price[0].find('span')
			close_price = int(span_price.text.replace(',',''))
			
			div_ratio = dps / close_price

			if roe_sub_list[4] > 0 and roe_sub_list[3] >0 and roe_sub_list[2] > 0:
				avg_roe = ((roe_sub_list[4] + roe_sub_list[3] + roe_sub_list[2]) / 3) - (div_ratio * 15.4)
			else:
				avg_roe = 0


			future_bps = bps * (1 + avg_roe/100)**10

			invest_price = future_bps / (1.15**10)

			expected_rate = ((future_bps / close_price)**0.1) - 1

			avg_roe_list.append(avg_roe)
			future_bps_list.append(future_bps)
			invest_price_list.append(invest_price)
			close_price_list.append(close_price)
			expected_rate_list.append(expected_rate)
			div_ratio_list.append(div_ratio)
			
			#print("dps", dps)
			
		f_pickle = open("crawling_list", "wb")
		pickle.dump(eps_list, f_pickle)
		pickle.dump(bps_list, f_pickle)
		pickle.dump(dps_list, f_pickle)
		f_pickle.close()
	
	# input mode is not 0
	else:
		f_pickle = open("crawling_list", "rb")
		eps_list = pickle.load(f_pickle)
		bps_list = pickle.load(f_pickle)
		dps_list = pickle.load(f_pickle)

	### PART III - Write information to new Excel file
	workbook_name = "snowball_value.xlsx"
	if os.path.isfile(os.path.join(cur_dir, workbook_name)):
		os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)
	worksheet_pbr = workbook.add_worksheet('result')

	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})
	percent_format = workbook.add_format({'num_format': '0.00%'})
	
	num_format = workbook.add_format({'num_format':'0.00'})
	num2_format = workbook.add_format({'num_format':'#,##0'})
	num3_format = workbook.add_format({'num_format':'#,##0.00',
									  'fg_color':'#FCE4D6'})
	
	# Write filter
	worksheet_pbr.write(0, 0, "Category", filter_format)
	worksheet_pbr.set_column('A:A', 15)
	worksheet_pbr.write(0, 1, "Name", filter_format)
	worksheet_pbr.set_column('B:B', 15)
	worksheet_pbr.write(0, 2, "Code", filter_format)
	worksheet_pbr.set_column('C:C', 10)
	worksheet_pbr.write(0, 3, "URL", filter_format)
	worksheet_pbr.set_column('D:D', 30)
	worksheet_pbr.write(0, 4,  "EPS", filter_format)
	worksheet_pbr.write(0, 5,  "BPS", filter_format)
	worksheet_pbr.write(0, 6,  "DPS", filter_format)
	worksheet_pbr.write(0, 7,  "ROE 2012", filter_format)
	worksheet_pbr.write(0, 8,  "ROE 2013", filter_format)
	worksheet_pbr.write(0, 9,  "ROE 2014", filter_format)
	worksheet_pbr.write(0, 10,  "ROE 2015", filter_format)
	worksheet_pbr.write(0, 11,  "ROE 2016", filter_format)
	#worksheet_pbr.write(0, 12,  "AVG ROE", filter_format)
	#worksheet_pbr.write(0, 13,  "Future BPS", filter_format)
	#worksheet_pbr.write(0, 14,  "close price", filter_format)
	#worksheet_pbr.write(0, 15,  "Invest price", filter_format)
	#worksheet_pbr.write(0, 16,  "Expected ratio", filter_format)
	#worksheet_pbr.write(0, 17,  "Div ratio", filter_format)
	worksheet_pbr.write(0, 12,  "3년 평균 ROE", filter_format)
	worksheet_pbr.write(0, 13,  "미래가치 BPS", filter_format)
	worksheet_pbr.write(0, 14,  "종가", filter_format)
	worksheet_pbr.write(0, 15,  "투자가능가격", filter_format)
	worksheet_pbr.write(0, 16,  "미래수익률", filter_format)
	worksheet_pbr.write(0, 17,  "시가배당률", filter_format)

	for k in range(num_stock):
		worksheet_pbr.write(1+k, 0, stock_cat_list[k])
		worksheet_pbr.write(1+k, 1, stock_name_list[k])
		worksheet_pbr.write(1+k, 2, stock_num_list[k])
		worksheet_pbr.write(1+k, 3, stock_url_list[k])
		worksheet_pbr.write(1+k, 4, eps_list[k], num2_format)
		worksheet_pbr.write(1+k, 5, bps_list[k], num2_format)
		worksheet_pbr.write(1+k, 6, dps_list[k], num2_format)
		worksheet_pbr.write(1+k, 7, roe_list[k][0])
		worksheet_pbr.write(1+k, 8, roe_list[k][1])
		worksheet_pbr.write(1+k, 9, roe_list[k][2])
		worksheet_pbr.write(1+k, 10, roe_list[k][3])
		worksheet_pbr.write(1+k, 11, roe_list[k][4])
		worksheet_pbr.write(1+k, 12, avg_roe_list[k], num_format)
		worksheet_pbr.write(1+k, 13, future_bps_list[k], num2_format)
		worksheet_pbr.write(1+k, 14, close_price_list[k], num2_format)
		worksheet_pbr.write(1+k, 15, invest_price_list[k], num2_format)
		worksheet_pbr.write(1+k, 16, expected_rate_list[k], percent_format)
		worksheet_pbr.write(1+k, 17, div_ratio_list[k], percent_format)


# Main
if __name__ == "__main__":
	main()


