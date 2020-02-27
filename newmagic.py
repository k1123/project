import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

from pymongo import MongoClient
client = MongoClient('localhost', 27017)
db = client.dbsparta

work_book = load_workbook('/Users/User/Desktop/sparta/py_project/excel/magic.xlsx')
work_sheet = work_book['Sheet1']

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}

comp_list = []

multiple_cells = work_sheet['A4':'A100']
for row in multiple_cells:
    for cell in row:
        code = cell.value

        data = requests.get(f'http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode={code}&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701',headers=headers)

        soup = BeautifulSoup(data.text, 'html.parser')

        # 종목코드
        symbol = soup.select_one('#compBody > div.section.ul_corpinfo > div.corp_group1 > h2')

        # 종목명
        name = soup.select_one('#giName').text

        # PBR
        PBR = soup.select_one('#corp_group2 > dl:nth-child(4) > dd')

        if PBR is None:
            PBR_val = 100
        else:
            try:
                PBR_val = float(PBR.text)
            except:
                print(f'PBR error : {PBR.text}')
                PBR_val = 100

        data = requests.get(
            f'http://comp.fnguide.com/SVO2/ASP/SVD_Finance.asp?pGB=1&gicode={code}&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701',
            headers=headers)

        soup = BeautifulSoup(data.text, 'html.parser')

        # 매출총이익
        gross_profit1 = soup.select_one('#divSonikQ > table > tbody > tr:nth-child(3) > td:nth-child(2)')
        gross_profit2 = soup.select_one('#divSonikQ > table > tbody > tr:nth-child(3) > td:nth-child(3)')
        gross_profit3 = soup.select_one('#divSonikQ > table > tbody > tr:nth-child(3) > td:nth-child(4)')
        gross_profit4 = soup.select_one('#divSonikQ > table > tbody > tr:nth-child(3) > td:nth-child(5)')

        # 총자산
        assets = soup.select_one('#divDaechaY > table > tbody > tr:nth-child(1) > td.r.cle')

        # GP/A 계산하기
        if gross_profit1 is None:
            gross_profit1_val = 0
        else:
            try:
                gross_profit1_val = float(gross_profit1.text.replace(',', ''))
            except:
                gross_profit1_val = 0

        if gross_profit2 is None:
            gross_profit2_val = 0
        else:
            try:
                gross_profit2_val = float(gross_profit2.text.replace(',', ''))
            except:
                gross_profit2_val = 0

        if gross_profit3 is None:
            gross_profit3_val = 0
        else:
            try:  # 일단 시도해보자
                gross_profit3_val = float(gross_profit3.text.replace(',', ''))
            except:  # try 안에서 문제가 생겼다면 아래를 실행시켜라(try에서 문제가 없었으면 실행되지 않음)
                gross_profit3_val = 0

        if gross_profit4 is None:
            gross_profit4_val = 0
        else:
            try:
                gross_profit4_val = float(gross_profit4.text.replace(',', ''))
            except:
                gross_profit4_val = 0

        gross_profit_val = gross_profit1_val + gross_profit2_val + gross_profit3_val + gross_profit4_val
        assets_val = float(assets.text.replace(',', ''))

        GPA = gross_profit_val / assets_val


        comp_list.append([code, name, GPA, PBR_val])


GPA_rank = sorted(comp_list, key = lambda x: x[2], reverse=True)

rank = 1
for g in GPA_rank:
    g.append(rank)
    rank += 1

PBR_rank = sorted(comp_list, key = lambda y: y[3])

rank = 1
for p in PBR_rank:
    p.append(rank)
    rank += 1

result = []
GPA_code = sorted(GPA_rank, key = lambda x: x[0])
PBR_code = sorted(PBR_rank, key = lambda y: y[0])


for g, p in zip(GPA_code, PBR_code):
    if g[0] == p[0]:
        total_rank = g[4] + g[5]
        result.append([g[0], g[1], total_rank])
    else:
        print('error')
        break

sorted_result = sorted(result, key=lambda x: x[2])

rank = 1

for r in sorted_result:
    doc = {
        'rank' : rank,
        'code' : r[0],
        'name' : r[1]
    }
    db.t20200227.insert_one(doc)
    rank += 1

