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
tc_sheet = doc.worksheet('sheet2')


def tc_to_json(i: int) -> dict:

    key = tc_sheet.row_values(1)
    col = len(key)

    value = tc_sheet.row_values(i + 1)

    return {
        key[x]: value[x] for x in range(col)
    }


def select_input(driver: webdriver, cma_result_test: Test) -> bool:

    tc = cma_result_test.get_tc()

    try:
        # TC sheet에 일치하는 가입방법을 체크
        temp_signup = tc["가입 방법"].split(", ")
        for x in range(0, len(temp_signup)):
            driver.find_element(By.XPATH, "//label[contains(text(), '" + temp_signup[x] + "')]").click()

            if not load_result(driver):
                result = "가입 방법 체크 후 결과 리스트 로드 Timeout"
                cma_result_test.result_nt(driver, result)
                return False

    except NoSuchElementException:
        result = "가입 방법 체크박스를 찾지 못했습니다."
        cma_result_test.result_nt(driver, result)
        return False

    # TC sheet에 일치하는 통장기능을 체크
    try:
        temp_func = tc["통장 기능"].split(", ")
        for x in range(0, len(temp_func)):
            driver.find_element(By.XPATH, "//label[contains(text(), '" + temp_func[x] + "')]").click()

            if not load_result(driver):
                result = "통장 기능 체크 후 결과 리스트 로드 Timeout"
                cma_result_test.result_nt(driver, result)
                return False

    except NoSuchElementException:
        result = "통장 기능 체크박스를 찾지 못했습니다."
        cma_result_test.result_nt(driver, result)
        return False

    # TC sheet에 일치하는 수수료 혜택을 체크
    try:
        temp_perf = tc["수수료 혜택"].split(", ")
        for x in range(0, len(temp_perf)):
            driver.find_element(By.XPATH, "//label[contains(text(), '" + temp_perf[x] + "')]").click()

            if not load_result(driver):
                result = "수수료 혜택 체크 후 결과 리스트 로드 Timeout"
                cma_result_test.result_nt(driver, result)
                return False

    except NoSuchElementException:
        result = "수수료 혜택 체크박스를 찾지 못했습니다."
        cma_result_test.result_nt(driver, result)
        return False

    return True


# 결과 리스트 로드
def load_result(driver: webdriver) -> bool:

    wait = 0

    while True:

        if wait > 20:
            return False

        elements = driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_")
        num = len(elements)

        if num > 0:
            return True

        if is_empty(driver):
            return True

        sleep(1)
        wait += 1


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


def is_empty(driver: webdriver) -> bool:

    if "조건에 맞는 CMA가 없습니다" in driver.find_element(By.CLASS_NAME, "cardsContainer_2d0E9").text:
        return True

    return False


def confirm_result(driver: webdriver, cma_result_test: Test) -> bool:

    tc = cma_result_test.get_tc()

    try:
        elements = driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_")
        num = len(elements)

        temp_func = tc["통장 기능"].split(", ")
        temp_perf = tc["수수료 혜택"].split(", ")

        for x in range(0, num):

            # 비대면 가입만 체크한 경우, '비대면 가입' 메시지 노출 확인
            if tc == "비대면 가입(온라인/모바일)":
                if "비대면\n가입" not in elements[x].text:
                    result = "상품에 비대면 가입에대한 메시지가 노출되지 않습니다."
                    cma_result_test.result_fail(driver, result)
                    return False

            # 체크한 통장기능에 해당하는 상품이 정상 노출되는지 확인
            for y in range(0, len(temp_func)):
                if temp_func[y] not in elements[x].text:
                    result = "선택한 통장 기능이 상품에 정상적으로 노출되지 않습니다."
                    cma_result_test.result_fail(driver, result)
                    return False

            # 체크한 수수료혜택에 해당하는 상품이 정상 노출되는지 확인
            for y in range(0, len(temp_perf)):
                if temp_perf[y] not in elements[x].text:
                    result = "선택한 수수료 혜택이 상품에 정상적으로 노출되지 않습니다."
                    cma_result_test.result_fail(driver, result)
                    return False

    except NoSuchElementException:
        result = "상품요소를 찾지 못함"
        cma_result_test.result_nt(driver, result)
        return False

    return True


def confirm_kind(driver: webdriver, cma_result_test: Test) -> bool:

    num = len(driver.find_elements(By.XPATH, "//strong[contains(text(), '통장 종류')]/../div/ul/li/button"))

    for x in range(1, num):

        try:
            driver.find_element(By.CSS_SELECTOR, ".listWrap_119TX > :nth-child(1) > ul > :nth-child(4) > div").location_once_scrolled_into_view
            driver.find_element(By.CSS_SELECTOR, ".listWrap_119TX > :nth-child(1) > ul > :nth-child(4) > div").click()

            account_kind_lst = driver.find_elements(By.XPATH, "//strong[contains(text(), '통장 종류')]/../div/ul/li/button")
            account_kind = account_kind_lst[x].text

            driver.find_element(By.XPATH, "//button[contains(text(), '" + account_kind + "')]").click()

        except NoSuchElementException:
            result = "통장 종류 선택 항목을 찾지 못함"
            cma_result_test.result_nt(driver, result)
            return False

        if not load_result(driver):
            result = "통장종류 선택 결과가 로드되지 않음"
            cma_result_test.result_fail(driver, result)
            return False

        if is_empty(driver):
            continue

        load_invest(driver)

        try:
            cma_num = len(driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_"))

            for y in range(0, cma_num):

                kind_detail = driver.find_elements(By.XPATH, "//div[contains(text(), '상세보기')]")
                kind_detail[y].click()

                wait_detail = 0

                while True:

                    if wait_detail > 20:
                        result = "상세보기 페이지 로드 Timeout"
                        cma_result_test.result_fail(driver, result)
                        return False

                    name = driver.find_element(By.CLASS_NAME, "name_1K1Gf").text

                    if name != '':
                        break

                    sleep(1)
                    wait_detail += 1

                kind = driver.find_element(By.CSS_SELECTOR, ".basicInfoItem_1DlvF > table > tbody > tr > :nth-child(1)").text

                if kind != account_kind:
                    result = f"통장종류에 일치하지 않는 상품이 노출됨 \n" \
                             f"기대결과: { account_kind.text } \n" \
                             f"실제결과: { kind }"

                    cma_result_test.result_fail(driver, result)
                    return False

                driver.find_element(By.LINK_TEXT, "결과 리스트로").click()

                if not load_result(driver):
                    result = "결과 리스트 로드 Timeout"
                    cma_result_test.result_nt(driver, result)
                    return False

        except NoSuchElementException:
            result = "추천 상품 항목을 찾지못함"
            cma_result_test.result_nt(driver, result)
            return False

    driver.find_element(By.CSS_SELECTOR, ".listWrap_119TX > :nth-child(1) > ul > :nth-child(4) > div").click()
    driver.find_element(By.XPATH, "//button[contains(text(), '전체')]").click()

    if not load_result(driver):
        result = "통장종류 선택 결과 로드 Timeout"
        cma_result_test.result_nt(driver, result)
        return False

    if is_empty(driver):
        result = "다시 통장종류를 전체로 되돌렸을 때 상품이 노출되지 않음"
        cma_result_test.result_fail(driver, result)
        return False

    load_invest(driver)

    return True


# 정렬 필터 적용 했을 때 상품 비교
def confirm_filter(driver: webdriver, cma_result_test: Test) -> bool:

    num = len(driver.find_elements(By.CSS_SELECTOR, ".headSort_1LuVw > div > ul > li > button"))

    for x in range(0, num):

        try:
            driver.find_element(By.CLASS_NAME, "headSort_1LuVw").location_once_scrolled_into_view
            driver.find_element(By.CLASS_NAME, "headSort_1LuVw").click()

            sort_kind_lst = driver.find_elements(By.CSS_SELECTOR, ".headSort_1LuVw > div > ul > li > button")
            sort_kind = sort_kind_lst[x].text

            driver.find_element(By.XPATH, "//button[contains(text(), '" + sort_kind + "')]").click()
        except NoSuchElementException:
            result = "정렬 선택항목을 찾지못함"
            cma_result_test.result_nt(driver, result)
            continue

        if not load_result(driver) or is_empty(driver):
            result = f"정렬({ sort_kind }) 결과리스트가 로드되지 않음"
            cma_result_test.result_fail(driver, result)
            return False

        load_invest(driver)

        sort_num = len(driver.find_elements(By.CLASS_NAME, "cardWrap_3dP8_"))

        if sort_num == 1:
            return True

        try:

            for y in range(0, sort_num - 1):

                interest1 = 0
                interest2 = 0

                if sort_kind == "연 이자금액 높은 순":
                    interest1 = int(driver.find_elements(By.CLASS_NAME, "taxInfoBenefit_hVmDH")[y].text.replace("원", "").replace(",", ""))
                    interest2 = int(driver.find_elements(By.CLASS_NAME, "taxInfoBenefit_hVmDH")[y + 1].text.replace("원", "").replace(",", ""))

                elif sort_kind == "금리 높은 순":
                    interest1 = float(driver.find_elements(By.CLASS_NAME, "taxInfoDescription_WMdnv")[y].text.replace("적용받은 금리", "").replace("%", ""))
                    interest2 = float(driver.find_elements(By.CLASS_NAME, "taxInfoDescription_WMdnv")[y + 1].text.replace("적용받은 금리", "").replace("%", ""))

                elif sort_kind == "최대금리 순":
                    interest1 = float(driver.find_elements(By.CLASS_NAME, "maxInterest_2UmRM")[y].text.replace("최대금리", "").replace("%", ""))
                    interest2 = float(driver.find_elements(By.CLASS_NAME, "maxInterest_2UmRM")[y + 1].text.replace("최대금리", "").replace("%", ""))

                if interest1 < interest2:
                    result = f"정렬({ sort_kind })필터가 정상적으로 동작하지 않음 \n" \
                             f"기대결과: { interest2 }, { interest1 } 순으로 노출 \n" \
                             f"실제결과: { interest1 }, { interest2 } 순으로 노출"

                    cma_result_test.result_fail(driver, result)
                    return False

        except NoSuchElementException:
            result = f"정렬 후({ sort_kind }) 이자관련 항목을 찾지못함"
            cma_result_test.result_nt(driver, result)
            continue

    return True


def main():

    cma_result_test = Test()
    cma_result_test.set_total(len(tc_sheet.col_values(1)) - 1)

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("window-size=1920x1080")
    chrome_options.add_argument("headless")

    print("------------------------테스트 시작------------------------")

    for i in range(1, cma_result_test.get_total() + 1):

        # 크롬드라이버를 백그라운드로 돌리고 싶을 때(또는 다른 옵션을 추가하고 싶을 때), 크롬드라이버 디렉토리 뒤에 옵션 추가
        driver = webdriver.Chrome('C:\webdriver\chromedriver.exe')

        tc = tc_to_json(i)
        print(f"{tc['TC 명']} Running...")

        cma_result_test.set_no(i)
        cma_result_test.set_tc(tc)

        driver.get("https://banksalad.com/cma/questions")

        try:
            result_button = WebDriverWait(driver, 20).until(
                expected_conditions.presence_of_element_located((By.LINK_TEXT, "결과보기"))
            )
            result_button.click()
        except TimeoutException:
            result = "결과보기 버튼을 찾지못함"
            cma_result_test.result_nt(driver, result)
            continue

        if not load_result(driver):
            result = "결과 리스트 로드 Timeout"
            cma_result_test.result_nt(driver, result)
            continue

        if is_empty(driver):
            result = "노출되는 상품이 없음"
            cma_result_test.result_nt(driver, result)
            continue

        # 가입방법 > 비대면 가입은 디폴트로 체크되어있기때문에 초기에 체크 해제
        try:
            driver.find_element(By.XPATH, "//label[contains(text(), '비대면 가입(온라인/모바일)')]").click()
        except NoSuchElementException:
            result = "비대면 가입(온라인/모바일) 체크박스를 찾지못함"
            cma_result_test.result_nt(driver, result)
            continue

        # TC에 해당하는 값 체크
        if not select_input(driver, cma_result_test):
            continue

        if is_empty(driver):
            result = "노출되는 상품이 없음"
            cma_result_test.result_nt(driver, result)
            continue

        load_invest(driver)

        if not confirm_result(driver, cma_result_test):
            continue

        # 선택한 통장 종류에 맞지 않는 상품이 노출되는 경우 Fail 처리 & 함수 마지막에 통장종류를 전체로 바꿔줌
        if not confirm_kind(driver, cma_result_test):
            continue

        if not confirm_filter(driver, cma_result_test):
            continue

        cma_result_test.result_pass(driver)

    cma_result_test.set_end()

    result_table(cma_result_test)


if __name__ == "__main__":

    main()
