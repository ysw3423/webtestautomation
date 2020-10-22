import gspread
from oauth2client.service_account import ServiceAccountCredentials

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# 테스터가 구글스프레드시트 발급받은 API key json 파일 명
JSON_FILE_NAME = "webtestautomation-293312-023523830718.json"

CREDENTIAL = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE_NAME, SCOPE)
GC = gspread.authorize(CREDENTIAL)

TESTCASE_URL = "https://docs.google.com/spreadsheets/d/1J4xlSeXzjT80ZYG2AQTwFOwv6v2e6ppazewCImUYowU/edit?usp=sharing"
