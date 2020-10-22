from time import sleep

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait

from gspreadAPI import GC, TESTCASE_URL
from testresult import Test, result_table

doc = GC.open_by_url(TESTCASE_URL)
tc_sheet = doc.worksheet('sheet1')


# 구글 스프레드 시트에 저장된 데이터를 한 TestCase 별로 dict 타입으로 파싱하여 반환
def tc_to_json(i: int) -> dict:

    key = tc_sheet.row_values(1)
    col = len(key)

    value = tc_sheet.row_values(i + 1)

    return {
        key[x]: value[x] for x in range(col)
    }


# tc에 저장된 입력값 선택
def select_input(driver: webdriver, cma_test: Test) -> bool:

    tc = cma_test.get_tc()

    # 월 평균 잔액 선택, 월 평균 잔액 입력 항목을 20초내에 찾지못하면 입력 페이지 로드 Timeout으로 생각하여, false를 반환하며 N/T 처리합니다.
    try:

        balance_button = WebDriverWait(driver, 20).until(
            expected_conditions.presence_of_element_located((By.XPATH, "//button[contains(text(), '" + format(int(tc["월 평균 잔액"]), ",") + "만')]"))
        )

        balance_button.click()

    except TimeoutException:
        result = "입력 페이지 로드 Timeout"
        cma_test.result_nt(driver, result)
        return False

    sleep(1)

    # 실적 선택, 실적 선택 항목을 찾지못하면 false를 반환하며 N/T 처리합니다.
    try:

        temp = tc["실적"].split(", ")
        perf = driver.find_elements(By.CLASS_NAME, "item_3aRCD")

        for x in range(0, len(temp)):
            for y in range(0, len(perf)):
                if perf[y].text == temp[x]:
                    perf[y].click()
                    break

    except NoSuchElementException:
        result = "실적 선택 항목을 찾지못함"
        cma_test.result_nt(driver, result)
        return False

    sleep(1)

    # 통장 기능 선택, 통장 기능 선택 항목을 찾지못하면 false를 반환하며 N/T 처리합니다.
    try:

        temp = tc["통장기능"].split(", ")
        perf = driver.find_elements(By.CLASS_NAME, "item_2l67q")

        for x in range(0, len(temp)):
            for y in range(0, len(perf)):
                if perf[y].text == temp[x]:
                    perf[y].click()
                    break

    except NoSuchElementException:
        result = "통장기능 선택 항목을 찾지못함"
        cma_test.result_nt(driver, result)
        return False

    return True


# 입력 페이지 우측에 입력 항목이 정상적으로 노출되는지 확인합니다.
# tc에 저장된 값대로 노출되지 않으면 false를 반환하면 Fail 처리합니다.
# 입력 페이지 우측의 '입력항목'을 찾지 못하면 false를 반환하며 N/T 처리합니다.
def confirm_input(driver: webdriver, cma_test: Test) -> int:

    tc = cma_test.get_tc()

    try:
        balance = driver.find_element(By.CLASS_NAME, "value_3k1fw").text

        # 선택한 월 평균 잔액과 입력 페이지 우측에 노출되는 월 평균 잔액 비교
        if format(int(tc["월 평균 잔액"]) * 10000, ",") + "원" != balance:
            result = f"입력 항목에 예상 월 평균잔액이 정상적으로 노출되지 않음 \n " \
                     f"기대결과: { format(int(tc['월 평균 잔액']) * 10000, ',') }원 \n" \
                     f"실제결과: { balance }"

            cma_test.result_fail(driver, result)
            return -1

    except NoSuchElementException:
        result = "입력 항목 내 예상 월 평균잔액을 찾지못함"
        cma_test.result_nt(driver, result)
        return -1

    try:
        perf = driver.find_element(By.CSS_SELECTOR, ".body_25Fqi > :nth-child(2) > :nth-child(1) > ul").text
        temp = tc["실적"].split(", ")

        # 실적 비교
        for x in range(0, len(temp)):
            if not temp[x] in perf:
                result = f"입력 항목에 선택한 수수료 혜택이 정상적으로 노출되지 않음 \n" \
                         f"기대결과: { tc['실적'] } \n" \
                         f"실제결과: { temp[x] } 미노출"

                cma_test.result_fail(driver, result)
                return -1

    except NoSuchElementException:
        result = "입력 항목 내 수수료혜택을 찾지못함"
        cma_test.result_nt(driver, result)
        return -1

    try:
        func = driver.find_element(By.CSS_SELECTOR, ".body_25Fqi > :nth-child(2) > :nth-child(2) > ul").text
        temp = tc["통장기능"].split(", ")

        # 통장기능 비교
        for x in range(0, len(temp)):
            if not temp[x] in func:
                result = f"입력 항목에 선택한 통장기능이 정상적으로 노출되지 않음 \n" \
                         f"기대결과값: { tc['통장기능'] } \n" \
                         f"실제결과값: { temp[x] } 미노출"

                cma_test.result_fail(driver, result)
                return -1

    except NoSuchElementException:
        result = "입력 항목내 통장기능을 찾지못함"
        cma_test.result_nt(driver, result)
        return -1

    wait = 0

    try:

        while True:

            if wait > 20:
                result = f"예상 연 이자금액 계산 Timeout"
                cma_test.result_fail(driver, result)
                return -1

            if "원" in driver.find_element(By.CLASS_NAME, "resultValue_rynW5").text:
                break

            wait += 1
            sleep(1)

        interest_result = int(driver.find_element(By.CLASS_NAME, "resultValue_rynW5").text.replace(",", "").replace("원", ""))
        return interest_result

    except NoSuchElementException:
        result = "예상 연 이자금액 항목을 찾지못함"
        cma_test.result_nt(driver, result)
        return -1


# 결과 리스트 로드
def load_result(driver: webdriver):

    wait = 0

    while True:

        if wait > 20:
            return False

        elements = driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_")
        num = len(elements)

        if num > 0:
            break

        sleep(1)
        wait += 1

    return True


# 한 페이지에서 로드할 수 있는 상품 모두 로드
def load_invest(driver: webdriver):

    num = 0

    while True:

        elements = driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_")
        temp = num
        num = len(elements)

        if num - temp < 10:
            break

        elements[num - 1].location_once_scrolled_into_view
        sleep(1)


# 입력 페이지에서 입력한 값들이 결과 페이지로 정상적으로 전달됐는지 확인합니다.
# 입력값과 동일하지 않은 값이 있을 때, false를 반환하며 Fail 처리합니다.
def confirm_result(driver: webdriver, cma_test: Test, interest: int) -> bool:

    tc = cma_test.get_tc()

    try:
        # 결과 페이지에 노출되는 월 평균잔액
        balance = int(driver.find_element(By.ID, "balance").get_attribute("value"))

        # 결과 페이지에 노출되는 월 평균잔액과 입력 페이지에서 선택한 월 평균잔액이 일치하는지 확인, 일치하지 않으면 Fail
        if int(tc["월 평균 잔액"]) * 10000 != balance:
            result = f"cma 분석 결과 페이지 상단에 선택한 월 평균잔액이 정상적으로 노출되지 않음 \n" \
                     f"기대결과값: { format(int(tc['월 평균 잔액']) * 10000, ',') }원 \n" \
                     f"실제결과값: { format(balance, ',') }원"

            cma_test.result_fail(driver, result)
            return False

    except NoSuchElementException:
        result = "분석 결과 페이지 상단 월 평균잔액을 찾지못함"
        cma_test.result_nt(driver, result)
        return False

    try:
        # 결과 페이지 최 상단에 노출되는 상품의 연 이자 금액과 입력 페이지에 노출되는 예상 연 이자금액이 일치하는지 확인
        first_interest = driver.find_element(By.CSS_SELECTOR, ".resultsContainer_24LR6 > :nth-child(4) > :nth-child(1) > :nth-child(2) > dl > dd > strong > span:nth-child(2)").text

        if format(interest, ",") != first_interest.replace("원", ""):
            result = f"입력 페이지에서 노출되던 예상 연 이자금액과 결과 페이지의 예상 연 이자금액이 일치하지 않음 \n" \
                     f"기대결과: { format(interest, ',') }원 \n" \
                     f"실제결과: { first_interest }"

            cma_test.result_fail(driver, result)
            return False

    except NoSuchElementException:
        result = "분석 결과 페이지의 예상 연 이자금액을 찾지못함"
        cma_test.result_nt(driver, result)
        return False

    # 결과 페이지 통장 기능 체크박스에 체크된 항목이 입력 페이지에서 선택한 항목과 일치하는지 확인
    checked_func = driver.find_elements(By.XPATH, "//div[@class='listWrap_119TX']/div[position()=2]/ul/li/ul/li/div/input[@checked]")
    not_checked_func = driver.find_elements(By.XPATH, "//div[@class='listWrap_119TX']/div[position()=2]/ul/li/ul/li/div/input[@checked=false]")

    checked_num = len(checked_func)
    not_checked_num = len(not_checked_func)

    try:
        for x in range(0, checked_num):
            checked_label = driver.find_element(By.XPATH, "//label[@for='" + checked_func[x].get_attribute("id") + "']").text

            if checked_label not in tc["통장기능"]:
                result = f"입력 페이지에서 선택한 통장기능과 결과 페이지에서 체크되어있는 통장기능이 일치하지 않음 \n" \
                         f"기대결과: { tc['통장기능'] } \n" \
                         f"실제결과: { checked_label }이 체크됨"

                cma_test.result_fail(driver, result)
                return False

        for x in range(0, not_checked_num):
            not_checked_label = driver.find_element(By.XPATH, "//label[@for='" + not_checked_func[x].get_attribute("id") + "']").text

            if not_checked_label in tc["통장기능"]:
                result = f"입력 페이지에서 선택한 통장기능과 결과 페이지에서 체크되어있는 통장기능이 일치하지 않음 \n" \
                         f"기대결과: { tc['통장기능'] } \n" \
                         f"실제결과: { not_checked_label }이 체크되지 않음"

                cma_test.result_fail(driver, result)
                return False

    except NoSuchElementException:
        result = "분석 결과 페이지의 통장기능 체크박스를 찾지못함"
        cma_test.result_nt(driver, result)
        return False

    checked_func_span = driver.find_elements(By.XPATH, "//span[@class='checkedSpecialService_3rMGy specialService_qhKGV']")
    not_checked_func_span = driver.find_elements(By.XPATH, "//span[@class='specialService_qhKGV']")

    checked_num = len(checked_func_span)
    not_checked_num = len(not_checked_func_span)

    for x in range(0, checked_num):
        if checked_func_span[x].text not in tc["통장기능"]:
            result = f"입력 페이지에서 선택한 통장기능이 결과 페이지에 노출되는 상품에 정상적으로 표시되지 않음 \n" \
                     f"기대결과: { tc['통장기능'] } \n" \
                     f"실제결과: { checked_func_span[x].text }이 표시됨"

            cma_test.result_fail(driver, result)
            return False

    for x in range(0, not_checked_num):
        if not_checked_func_span[x].text in tc["통장기능"]:
            result = f"입력 페이지에서 선택한 통장기능이 결과 페이지에 노출되는 상품에 정상적으로 표시되지 않음 \n" \
                     f"기대결과: { tc['통장기능'] } \n" \
                     f"실제결과: { not_checked_func_span[x].text }이 표시되지 않음"

            cma_test.result_fail(driver, result)
            return False

    return True


# 상품 상세보기 클릭시, 정상적인 페이지로 이동하는지 확인
# 결과 리스트로 버튼 눌렀을 때, 결과 리스트 페이지로 잘 이동하는지 확인
def confirm_detail(driver: webdriver, cma_test: Test) -> bool:

    num = len(driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_"))

    for x in range(0, num):

        if x >= 10:
            load_invest(driver)

        invest_name1 = driver.find_elements(By.CLASS_NAME, "name_-Xnbq")[x].text
        invest_detail = driver.find_element(By.CSS_SELECTOR, ".cardsContainer_2d0E9 > :nth-child(" + str(x + 1) + ") > .body_3UwsQ > .buttonGroup_1CD2Q > :nth-child(1)")

        lst_invest = driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_")

        lst_invest[x].location_once_scrolled_into_view
        invest_detail.click()

        wait_detail = 0

        while True:

            if wait_detail > 20:
                result = "상품 상세페이지 로드 Timeout"
                cma_test.result_fail(driver, result)
                return False

            invest_name2 = driver.find_element(By.CSS_SELECTOR, ".headerInfo_FgKOA > h3").text

            if invest_name2 != "":
                break

            sleep(1)
            wait_detail += 1

        if invest_name1 != invest_name2:
            result = "잘못된 상세페이지로 이동"
            cma_test.result_fail(driver, result)
            return False

        try:
            driver.find_element(By.LINK_TEXT, "결과 리스트로").click()
        except NoSuchElementException:
            result = "상품 상세보기 페이지 내 결과 리스트로 버튼을 찾지못함"
            cma_test.result_nt(driver, result)
            return False

        if not load_result(driver):
            result = "결과 리스트 로드 Timeout"
            cma_test.result_nt(driver, result)
            return False

    return True


def main():

    cma_test = Test()
    cma_test.set_total(len(tc_sheet.col_values(1)) - 1)

    # 백그라운드 옵션
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("window-size=1920x1080")
    chrome_options.add_argument("headless")

    print("------------------------테스트 시작------------------------")

    for i in range(1, cma_test.get_total() + 1):

        # 크롬드라이버를 백그라운드로 돌리고 싶을 때(또는 다른 옵션을 추가하고 싶을 때), 크롬드라이버 디렉토리 뒤에 옵션 추가
        driver = webdriver.Chrome('C:\webdriver\chromedriver.exe', options=chrome_options)

        tc = tc_to_json(i)
        print(f"{ tc['TC 명'] } Running...")

        cma_test.set_no(i)
        cma_test.set_tc(tc)

        driver.get("https://banksalad.com/cma/questions")

        sleep(1)

        if not select_input(driver, cma_test):
            continue

        sleep(1)

        interest = confirm_input(driver, cma_test)

        if interest < 0:
            driver.quit()
            continue

        try:
            driver.find_element(By.LINK_TEXT, "결과보기").click()
        except NoSuchElementException:
            result = "'결과보기' 버튼을 찾지 못함"
            cma_test.result_nt(driver, result)
            continue

        if not load_result(driver):
            result = "결과 리스트 로드되지 않음"
            cma_test.result_nt(driver, result)
            continue

        load_invest(driver)

        if not confirm_result(driver, cma_test, interest):
            continue

        if not confirm_detail(driver, cma_test):
            driver.quit()
            continue

        cma_test.result_pass(driver)

    cma_test.set_end()

    result_table(cma_test)


if __name__ == "__main__":

    main()
