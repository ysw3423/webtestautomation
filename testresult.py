import os
import time
from selenium import webdriver
from datetime import datetime
from beautifultable import BeautifulTable

PASS = "\x1b[1;34mPass\x1b[1;m"
FAIL = "\x1b[1;31mFail\x1b[1;m"
NT = "N/T"


def screenshot(driver: webdriver, tc_name: str):
    temp = tc_name.split("0")
    file_dir = f"screenshot/{temp[0]}/"
    screenshot_name = f"{tc_name}_fail_{datetime.today().strftime('%Y%m%d%H%M')}.png"
    try:
        if not os.path.exists(file_dir):
            os.makedirs(file_dir)
    except OSError:
        print("Error: Creating directory.")
        return False
    driver.save_screenshot(file_dir + screenshot_name)


class Test:
    def __init__(self):
        self.__total = 0
        self.__no = 0
        self.__tc = {}
        self.__sum_pass = 0
        self.__sum_fail = 0
        self.__sum_nt = 0
        self.__start = time.time()
        self.__end = time.time()

    # setter
    def set_total(self, total: int):
        self.__total = total

    def set_no(self, no: int):
        self.__no = no

    def set_tc(self, tc: dict):
        self.__tc = tc

    def set_end(self):
        self.__end = time.time()

    # getter
    def get_total(self) -> int:
        return self.__total

    def get_no(self) -> int:
        return self.__no

    def get_tc(self) -> dict:
        return self.__tc

    def get_sum_pass(self) -> int:
        return self.__sum_pass

    def get_sum_fail(self) -> int:
        return self.__sum_fail

    def get_sum_nt(self) -> int:
        return self.__sum_nt

    def get_start(self) -> time:
        return self.__start

    def get_end(self) -> time:
        return self.__end

    def result_pass(self, driver: webdriver):
        self.__sum_pass += 1
        print(PASS)
        driver.quit()

    def result_fail(self, driver: webdriver, result: str):
        screenshot(driver, self.__tc["TC 명"])
        self.__sum_fail += 1
        print(f"{FAIL} - {result}")
        driver.quit()

    def result_nt(self, driver: webdriver, result: str):
        self.__sum_nt += 1
        print(f"{NT} - {result}")
        driver.quit()

    def result_total(self):
        self.__total += 1

    def result_pass_no_quit(self):
        self.__sum_pass += 1
        print(PASS)

    def result_fail_no_quit(self, driver: webdriver, result: str):
        screenshot(driver, self.__tc["TC 명"])
        self.__sum_fail += 1
        print(f"{FAIL} - {result}")

    def result_nt_no_quit(self, result: str):
        self.__sum_nt += 1
        print(f"{NT} - {result}")


def result_table(test: Test):
    table = BeautifulTable()
    table.column_headers = ["Total", "Pass", "Fail", "N/T"]
    table.append_row([test.get_total(), test.get_sum_pass(), test.get_sum_fail(), test.get_sum_nt()])
    print("------------------------테스트 결과------------------------")
    print(f"실행 날짜: {datetime.today().strftime('%Y-%m-%d')}")
    print(f"총 실행시간: {int(test.get_end() - test.get_start())}s")
    print(f"한 케이스당 평균 실행시간: {int((test.get_end() - test.get_start()) / test.get_total())}s")
    print(table)
