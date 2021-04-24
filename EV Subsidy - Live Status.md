import requests
import openpyxl
import pandas as pd
from bs4 import BeautifulSoup

excel_file = openpyxl.Workbook()
excel_sheet = excel_file.active
excel_sheet.append(['시도','지역구분','공고대수[총]','공고대수[우선]','공고대수[법인]','공고대수[일반]','접수대수[총]','접수대수[우선]','접수대수[법인]','접수대수[일반]','출고대수[총]','출고대수[우선]','출고대수[법인]','출고대수[일반]','잔여수량[법인]','잔여수량[일반]','접수율[법인]','접수율[일반]'])
excel_sheet.column_dimensions['B'].width = 15
excel_sheet.column_dimensions['C'].width = 14
excel_sheet.column_dimensions['D'].width = 14
excel_sheet.column_dimensions['E'].width = 14
excel_sheet.column_dimensions['F'].width = 14
excel_sheet.column_dimensions['G'].width = 14
excel_sheet.column_dimensions['H'].width = 14
excel_sheet.column_dimensions['I'].width = 14
excel_sheet.column_dimensions['J'].width = 14
excel_sheet.column_dimensions['K'].width = 14
excel_sheet.column_dimensions['L'].width = 14
excel_sheet.column_dimensions['M'].width = 14
excel_sheet.column_dimensions['N'].width = 14
excel_sheet.column_dimensions['O'].width = 14
excel_sheet.column_dimensions['P'].width = 14

url = 'https://www.ev.or.kr/portal/localInfo'
res = requests.get(url)
soup = BeautifulSoup(res.content, 'html.parser')
for number in range(1,162):
    #공고대수
    totalcount = soup.select_one('#sub_cont > table > tbody > tr:nth-child(' + str(number) + ') > td:nth-child(6)')
    totalcount_clear = int(totalcount.get_text().split(' ')[0])
    primary_count = int(totalcount.get_text().split(' ')[1].replace('(','').replace(')',''))
    company_count = int(totalcount.get_text().split(' ')[2].replace('(','').replace(')',''))
    
    #접수대수
    actcount = soup.select_one('#sub_cont > table > tbody > tr:nth-child(' + str(number) + ') > td:nth-child(7)')
    actcount_clear = int(actcount.get_text().split(' ')[0])
    primary_actcount = int(actcount.get_text().split(' ')[1].replace('(','').replace(')',''))
    company_actcount = int(actcount.get_text().split(' ')[2].replace('(','').replace(')',''))
    
    #출고대수
    act_delivery = soup.select_one('#sub_cont > table > tbody > tr:nth-child(' + str(number) + ') > td:nth-child(8)')
    act_delivery_count = int(act_delivery.get_text().split(' ')[0])
    act_delivery_primary = int(act_delivery.get_text().split(' ')[1].replace('(','').replace(')',''))
    act_delivery_company = int(act_delivery.get_text().split(' ')[2].replace('(','').replace(')',''))
    
    sido = soup.select_one('#sub_cont > table > tbody > tr:nth-child(' + str(number) + ') > th:nth-child(1)').get_text()
    region = soup.select_one('#sub_cont > table > tbody > tr:nth-child(' + str(number) + ') > th:nth-child(2)').get_text()
    
    #일반공고대수(우선, 법인 제외)
    normal_count = totalcount_clear - primary_count - company_count
    
    #일반접수대수(우선, 법인 제외)
    normal_actcount = actcount_clear - primary_actcount - company_actcount
    
    #일반출고대수(우선, 법인 제외)
    normal_delivery_count = act_delivery_count - act_delivery_primary - act_delivery_company
    
    #잔여수량(법인)
    company_balance = company_count - company_actcount - act_delivery_company
    
    #잔여수량(일반)
    normal_balance = normal_count - normal_actcount - normal_delivery_count
    
    #접수율(법인)
    if company_count == 0:
        percent_company = company_actcount/1 * 100
    else:
        percent_company = company_actcount/company_count * 100
    
    #접수율(일반)
    if normal_count == 0:
        percent_normal = normal_actcount/1 * 100
    else: 
        percent_normal = normal_actcount/normal_count * 100
    
    #print(sido, region, totalcount_clear, primary_count, company_count, normal_count, actcount_clear, primary_actcount, company_actcount)
    excel_sheet.append([sido, region, totalcount_clear, primary_count, company_count, normal_count, actcount_clear, primary_actcount, company_actcount, normal_actcount, act_delivery_count, act_delivery_primary, act_delivery_company, normal_delivery_count, company_balance, normal_balance, "%.2f%%"%percent_company, "%.2f%%"%percent_normal])
    
excel_file.save('e-sesang.xlsx')
excel_file.close()
show = pd.read_excel('e-sesang.xlsx')
show
