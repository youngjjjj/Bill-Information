import openpyxl
from bs4 import BeautifulSoup
from urllib.parse import unquote
from openpyxl import Workbook
import requests
import json
from people_list import people

with open("secret.json") as f:
    secret = json.loads(f.read())

# Keep secret keys in secret.json
def get_secret(setting, secret=secret):
    try:
        return secret[setting]
    except KeyError:
        error_msg = "Set the {0} environment variable".format(setting)
        raise Exception(error_msg)


excel_file = openpyxl.load_workbook(
    "국회의원 입법활동 조사 명단(대표발의).xlsx"
)

sheet = excel_file["결과"]

API_KEY = get_secret("API_KEY")

url = "http://apis.data.go.kr/9710000/BillInfoService2/getBillInfoList"

decode_API_KEY = unquote(API_KEY)

people_list = people.split("\n")

idx = 1

for person in people_list:
    params = {
        "ServiceKey": decode_API_KEY,
        "pageNo": "1",
        "numOfRows": "700",
        "mem_name_check": "G01",
        "mem_name": person,
        "start_ord": "20",
        "end_ord": "20",
        "process_num": "-",
        "start_process_num": "-",
        "end_process_num": "-",
        "propose_num": "-",
        "start_propose_num": "-",
        "end_propose_num": "-",
        "proposer_kind_cd": "F01",
        "gbn": "dae_num_name",
        "amendmentyn": "false",
        "budget": "false",
    }

    response = requests.get(url, params=params)
    soup = BeautifulSoup(response.text, "html.parser")
    content = soup.find("response")

    # 의안명
    bill_name = [billname.text for billname in content.find_all("billname")]
    # 제안일자
    date = [proposedt.text for proposedt in content.find_all("proposedt")]
    # 의결결과
    result = [generalresult.text for generalresult in content.find_all("generalresult")]
    # 주요내용
    summary = [summary.text for summary in content.find_all("summary")]
    # 심사진행상태
    process = [procstagecd.text for procstagecd in content.find_all("procstagecd")]

    for data in zip(bill_name, date, result, summary, process):
        sheet.cell(row=idx + 1, column=1).value = person
        sheet.cell(row=idx + 1, column=2).value = data[0]
        sheet.cell(row=idx + 1, column=3).value = data[1]
        sheet.cell(row=idx + 1, column=4).value = data[2]
        sheet.cell(row=idx + 1, column=5).value = data[3]
        sheet.cell(row=idx + 1, column=6).value = data[4]
        idx += 1
    idx += 1

excel_file.save(filename="국회의원 입법활동 조사 명단(대표발의).xlsx")
print("Complete!")
