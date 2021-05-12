from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
import selenium.common
import multiprocessing as mp
import time
import random as ra
import pyperclip

# 네이버 로그인 방법: 블로그 참고
def clipboard_input(driver, user_xpath, user_input):
    pyperclip.copy(user_input)  # input을 클립보드로 복사

    driver.find_element_by_xpath(user_xpath).click()  # element focus 설정
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()  # Ctrl+V 전달
    time.sleep(1)


def auto_spinner(value_info):
    driver = webdriver.Chrome("C:/r_selenium/chromedriver")  # 직접 본인 환경에 맞게 조정해야 됨
    time.sleep(1)

    id, id_blank_x_path, pw, pw_blank_x_path = value_info

    driver.get("https://m.kin.naver.com/mobile/roulette/home.nhn")

    spin_bt = Wait(driver, 10). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="rouletteStartBtn"]')))

    spin_bt.click()

    clipboard_input(driver, id_blank_x_path, id)
    clipboard_input(driver, pw_blank_x_path, pw)

    login_bt = Wait(driver, 10). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="frmNIDLogin"]/fieldset/input')))
    login_bt.click()

    while True:
        spin_bt = Wait(driver, 10). \
            until(ec.presence_of_element_located((By.XPATH,
                                                  '//*[@id="rouletteStartBtn"]')))

        spin_bt.click()

        time.sleep(6)

        # 무언가 당첨된 경우 룰렛 결과 창을 닫는 버튼을 클릭함
        try:
            exit_bt = Wait(driver, 5). \
                until(ec.presence_of_element_located((By.XPATH,
                                                      '// *[ @ id = "dimmed"] / div[3] / div[2] / button[2]')))
            exit_bt.click()

        except selenium.common.exceptions:
            driver.close()


if __name__ == '__main__':
    id = input("아이디를 입력해 주세요:")
    id_blank_x_path = '//*[@id="id"]'
    pw = input("비밀번호를 입력해 주세요: ")
    pw_blank_x_path = '//*[@id="pw"]'
    info = ([id, id_blank_x_path, pw, pw_blank_x_path])

    pool = mp.Pool(10)
    pool.map(auto_spinner, ([info] * 10))


# id_blank = Wait(driver, 10). \
#     until(ec.presence_of_element_located((By.XPATH,
#                                           '//*[@id="id"]')))
#
# for i in id:
#     id_blank.send_keys(i)
#     time.sleep(ra.random() + 1)
#
# ps_blank = Wait(driver, 10). \
#     until(ec.presence_of_element_located((By.XPATH,
#                                           '//*[@id="pw"]')))
#
# for i in ps:
#     ps_blank.send_keys(i)
#     time.sleep(ra.random() + 1)
#
# time.sleep(ra.random() + 1)



# 자동 로그인 포기한 버전


# driver = webdriver.Chrome("C:/r_selenium/chromedriver")  # 직접 본인 환경에 맞게 조정해야 됨
# time.sleep(1)
#
# driver.get("https://m.kin.naver.com/mobile/roulette/home.nhn")
#
# spin_bt = Wait(driver, 10). \
#     until(ec.presence_of_element_located((By.XPATH,
#                                           '//*[@id="rouletteStartBtn"]')))
#
# spin_bt.click()

