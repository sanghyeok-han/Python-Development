from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as ec
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import sys
import os
import re
import pandas as pd
from numpy import mean
from tabulate import tabulate
import itertools
import string
import imgkit


pd.set_option('display.max_columns', 6)
pd.set_option('display.width', 1000)

## 크롬드라이버 여러 개로 try except 사용하기

# 1. selenium으로 크롬드라이버(같이 저장)을 통해서 페이지 들어가서 로그인 후 과목 정보 입력 받아서 보여줌 (각각의 위치 저장하기)
# 2. 수강번호를 입력하면 그것에 해당하는 강의계획서를 html과 wkhtmltoimage(같이 저장)을 통해서 이미지 파일로 제작해줌
# 3. 반복 여부 물어봄
# 4. 반복을 더 이상 하지 않는다면 로그아웃 후 종료함


# left와 top을 통해서 요소를 가져옴
def plan_element_search(left, top, all_plan_html):
    element = re.findall(
        'left: ' + str(left) + 'px; top: ' + str(top) + r'px;[\w:;.,()"\s-]+>([~&\w()/%,.?"@:;\s-]+)<', all_plan_html)

    if len(element) == 0:
        element = ""
    else:
        element = element[0]

    return element


# 해당 left에 해당하는 모든 요소를 가져옴
def plan_element_search_by_left(left, all_plan_html, output_type, print_type=1):  # output_type의 경우 1: 리스트, 2: string
    elements = re.findall('left: ' + str(left) + r'px;[\w:;.,()"\s-]+>([~&\w()/%,.?"@:;\s-]+)<', all_plan_html)

    if len(elements) == 0:
        elements = ""
    else:
        if elements[-1] == "도록 한다.":
            del elements[-1]

        if output_type == 2:
            if print_type == 2:
                elements = "<br/>&nbsp".join(elements)
                elements = "&nbsp" + elements
            else:
                elements = "<br/>".join(elements)
    return elements


# 해당 left를 가진 모든 요소의 top을 가져옴
def get_top_location_by_left(left, all_plan_html):
    top_locations = re.findall(r'left: ' + str(left) + 'px; top: (\d+)px;', all_plan_html)
    return top_locations


# 정해진 left와 첫 top을 통해서 top을 added1 또는 added2 만큼 더해나가면서 요소를 가져옴
def get_element_by_tops_sequence(fixed_left, first_top, added1, added2, all_plan_html, print_type):
    tops = [int(first_top)]
    if print_type == 2:
        elements = "&nbsp"
    else:
        elements = ""
    no_available_count = 0

    while True:
        available_element = 0
        for top in tops:
            element = plan_element_search(fixed_left, top, all_plan_html)

            if len(element) > 0:
                if no_available_count >= 1:
                    for _ in range(no_available_count):
                        elements = elements + "<br/>"
                no_available_count = 0

                if print_type == 2:
                    elements = elements + element + "<br/>&nbsp"
                else:
                    elements = elements + element + "<br/>"
                available_element += 1
        if not available_element:
            no_available_count += 1

        if no_available_count >= 5:
            break

        added_result1 = list(map(lambda x: x + added1, tops))
        added_result2 = list(map(lambda x: x + added2, tops))
        tops = list(set(added_result1 + added_result2))

    return elements


# top 정보를 기준으로 top이 근접한 정보를 하나로 합쳐주는 기능
def element_concatenate(li, li_tops, max_range):
    li_tops_changed = li_tops[:]
    for i in list(reversed(range(len(li) - 1))):
        if int(li_tops[i]) - max_range <= int(li_tops[i + 1]) <= int(li_tops[i]) + max_range:
            li[i] = str(li[i]) + "<br/>" + str(li[i + 1])
            del li[i + 1]
            li_tops_changed[i] = mean([int(li_tops_changed[i]), int(li_tops_changed[i + 1])])
            del li_tops_changed[i + 1]
    li_tops = li_tops_changed[:]

    return li, li_tops


# 각 정보를 담은 리스트를 각 정보의 top 정보를 담은 리스트의 top 정보를 기준으로 위치에 알맞게 zip 해주는 함수
def selective_zipper(li, li_tops, other_lis, other_lis_tops, including_range=17):
    all_info = []
    for i in range(len(li)):
        info = [li[i]]

        for j in range(len(other_lis)):
            is_outer_info_appended = False

            for k in range(len(other_lis[j])):
                if float(other_lis_tops[j][k]) - including_range <= float(li_tops[i]) <= float(
                        other_lis_tops[j][k]) + including_range:
                    info.append(other_lis[j][k])
                    is_outer_info_appended = True
            if not is_outer_info_appended:
                info.append("")

        all_info.append(info)

    return all_info


def access_test(file_path, options):
    driver = webdriver.Chrome(file_path, options=options)
    driver.implicitly_wait(1)
    driver.get("https://mhaksa.ajou.ac.kr:30443/index.html")  # 아주대 포탈 사이트에 접속함
    driver.implicitly_wait(3)
    return driver


def initial_access(driver, is_re):
    if not is_re:
        os.system("cls")
        print("1. 이미 아주대 포탈에 로그인된 상태인 경우 자동으로 로그아웃이 되므로 유의해")
        print("   주세요!\n")
        print("2. 강의계획서 제작시 조교 관련 정보 등 일부 빈도가 희박한 정보 및 기타 몇몇")
        print("   정보는 스크래핑 대상에서 제외하였습니다.\n")
        print("3. 강의계획서에서 기타 참고사항은 만약 해당 정보가 존재한다면 '3. 수업의 형태")
        print("   및 진행방식'에 병합되어 있습니다.\n")
        print("4. 파일 생성 위치는 해당 파일 폴더 내부입니다.\n\n")

        ajou_id = input("아주대 아이디를 입력해 주세요: ")
        ajou_ps = input("아주대 비밀번호를 입력해 주세요: ")

        os.system("cls")
        print("Accessing...")

        while True:
            # 통합 로그인 사이트에서 아이디 입력
            while True:
                try:
                    id_blank = Wait(driver, 1.5). \
                        until(ec.presence_of_element_located((By.XPATH,
                                                              '//*[@id="userId"]')))
                    id_blank.send_keys(ajou_id)  # 아이디 입력
                except:
                    time.sleep(1)
                else:
                    break

            # 통합 로그인 사이트에서 비밀번호 입력
            time.sleep(0.5)
            ps_blank = Wait(driver, 1.5). \
                until(ec.presence_of_element_located((By.XPATH,
                                                      '//*[@id="password"]')))
            ps_blank.send_keys(ajou_ps)  # 비밀번호 입력

            # 통합 로그인 사이트에서 로그인 버튼 클릭
            login_bt = Wait(driver, 1.5). \
                until(ec.presence_of_element_located((By.XPATH,
                                                      '//*[@id="loginSubmit"]')))
            driver.execute_script("arguments[0].click();", login_bt)

            # 기존 로그인 해제 경고창이 나온다면 승낙 체크
            time.sleep(0.3)

            try:
                alert = driver.switch_to.alert
                alert.accept()
            except Exception:
                pass

            # 인터넷 창 최대화
            time.sleep(0.3)
            driver.maximize_window()

            try:
                # 메인 매뉴에서 수업/비교과 버튼 클릭
                time.sleep(0.3)
                lecture_bt = Wait(driver, 1.5). \
                    until(ec.presence_of_element_located((By.XPATH,
                                                          '//*[@id="header"]/div/div[2]/haksa-top-menu/div/ul/li[2]/a')))
                driver.execute_script("arguments[0].click();", lecture_bt)
            except:
                os.system("cls")
                print("※ 아주대 사이트 로그인에 실패하였습니다!")
                time.sleep(1.5)

                os.system("cls")
                ajou_id = input("아주대 아이디를 다시 입력해 주세요: ")
                ajou_ps = input("아주대 비밀번호를 다시 입력해 주세요: ")

                os.system("cls")
                print("Accessing...")

                driver.get("https://mhaksa.ajou.ac.kr:30443/index.html")  # 아주대 포탈 사이트에 접속함
                driver.implicitly_wait(3)
            else:
                break

    if is_re:
        # 메인 매뉴에서 수업/비교과 버튼 클릭
        time.sleep(0.3)
        lecture_bt = Wait(driver, 1.5). \
            until(ec.presence_of_element_located((By.XPATH,
                                                  '//*[@id="header"]/div/div[2]/haksa-top-menu/div/ul/li[2]/a')))
        driver.execute_script("arguments[0].click();", lecture_bt)

    # 메인 매뉴에서 시간표 버튼 클릭
    time.sleep(0.3)
    schedule_bt = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="navContext"]/div/div[2]/ul/li[1]/div/div[2]/dd/ul/li[1]/a')))
    try:
        driver.execute_script("arguments[0].click();", schedule_bt)
    except:
        pass
    try:
        schedule_bt.click()
    except:
        pass

    # 팝업 창을 닫음
    time.sleep(0.3)
    check_bt = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '/html/body/div[1]/div/div[3]/div/div/button')))
    driver.execute_script("arguments[0].click();", check_bt)

    # 학년도에서 2020학년도 옵션 선택
    time.sleep(0.5)
    year = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/'
                                              'div[1]/div/select/option[6]')))
    year.click()

    # 학기에서 2학기 옵션 선택
    time.sleep(0.5)
    semester = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/'
                                              'div[2]/div/select/option[4]')))

    semester.click()


def search(driver, file_type):
    os.system("cls")

    # 교과구분을 입력 받음
    subject_type = input("교과구분을 입력해주세요(1: 전공, 2. 기초, 3. 영역별 교양, 4. 교양): ")
    # 전공 과목을 선택한 경우
    if subject_type == "1":
        subject_type_xpath = '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/div[3]/div/select/option[3]'
    # 기초 과목을 선택한 경우
    elif subject_type == "2":
        subject_type_xpath = '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/div[3]/div/select/option[5]'
    # 영역별 교양을 선택한 경우
    elif subject_type == "3":
        subject_type_xpath = '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/div[3]/div/select/option[7]'
    # 교양 과목을 선택한 경우
    else:
        subject_type_xpath = '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/div[3]/div/select/option[4]'

    time.sleep(0.5)
    subject_type_bt = Wait(driver, 3). \
        until(ec.presence_of_element_located((By.XPATH, subject_type_xpath)))
    time.sleep(0.5)
    subject_type_bt.click()

    if subject_type in ["1", "2"]:
        while True:
            try:
                majors = Wait(driver, 3). \
                    until(ec.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                               "div.nb-search-form div:nth-of-type(4) select option")))
                major_li = []
                for major in majors:
                    major_li.append(major.text)
            except:
                pass
            else:
                break

        major_df = pd.DataFrame({"학과": major_li})

        os.system("cls")
        print(tabulate(major_df, headers='keys', tablefmt='psql'))

        major_num = input("위의 리스트에서 해당 학과의 번호를 입력해주세요: ")
        while True:
            if major_num in map(str, range(len(major_df.index))):
                break
            else:
                major_num = input("\n해당 학과의 번호를 다시 입력해주세요: ")
        majors[int(major_num)].click()
    elif subject_type == "3":
        time.sleep(0.5)
        all_subjects_bt = Wait(driver, 5). \
            until(ec.presence_of_element_located((By.XPATH, '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div['
                                                            '1]/div[1]/div[4]/div/select/option[1]')))
        time.sleep(0.5)
        all_subjects_bt.click()

    subject_keyword = input("\n과목 이름의 일부 키워드를 입력해 주세요: ")
    subject_blank = Wait(driver, 1.5). \
        until(ec.presence_of_element_located(
        (By.XPATH, '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[1]/div[5]/div/input')))

    subject_blank.send_keys(subject_keyword)

    # 검색 버튼 클릭
    time.sleep(0.5)
    search_bt = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[1]/div[2]/'
                                              'div/div/button')))
    time.sleep(0.5)
    search_bt.click()

    # 전체 페이지 번호 스크래핑
    time.sleep(0.3)
    try:
        entire_page_num = Wait(driver, 5). \
            until(ec.presence_of_element_located((By.XPATH,
                                                  '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[5]/'
                                                  'sp-grid-paging/section/div/div[5]')))
    except TimeoutException:
        os.system("cls")
        print("※ 해당하는 강의가 없습니다!")
        time.sleep(2)

        return None

    entire_page_num = entire_page_num.text

    # 총 페이지 수가 한 페이지밖에 없는 경우
    if len(entire_page_num) == 0:
        entire_page_num = 1
    # 총 페이지 수가 두 페이지 이상인 경우
    else:
        entire_page_num = re.findall(r"1 / (\d*)", entire_page_num)
        entire_page_num = int("".join(entire_page_num))

    subject_info_xpath_dict = {"개설학부": '//*[(@id = "pageContainer")]//*[contains(concat('
                                       ' " ", @class, " " ), concat( " ", "ng-scope", " "'
                                       ' )) and (((count(preceding-sibling::*) + 1) = 1)'
                                       ' and parent::*)]//*[contains(concat( " ", @class,'
                                       ' " " ), concat( " ", "ng-binding", " " )) and'
                                       ' contains(concat( " ", @class, " " ), concat( " ",'
                                       ' "ng-scope", " " ))]',
                               "과목명": '//*[contains(concat( " ", @class, " " ),'
                                      ' concat( " ", "sp-grid-data-column", " "'
                                      ' )) and (((count(preceding-sibling::*) +'
                                      ' 1) = 3) and parent::*)]//*[contains('
                                      'concat( " ", @class, " " ), concat( " ",'
                                      ' "ng-scope", " " ))]',
                               "수강번호": '//*[contains(concat( " ", @class, " " ), concat( " ",'
                                       ' "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                       ' + 1) = 4) and parent::*)]//*[contains(concat( " ",'
                                       ' @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "교과구분": '//*[contains(concat( " ", @class, " " ), concat( " ",'
                                       ' "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                       ' + 1) = 5) and parent::*)]//*[contains(concat( " ",'
                                       ' @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "학점": '//*[contains(concat( " ", @class, " " ), concat('
                                     ' " ", "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                     ' + 1) = 7) and parent::*)]//*[contains(concat( " ", @class,'
                                     ' " " ), concat( " ", "ng-scope", " " ))]',
                               "주별시간": '//*[contains(concat( " ", @class, " " ),'
                                       ' concat( " ", "ng-scope", " " )) and'
                                       ' (((count(preceding-sibling::*) + 1) ='
                                       ' 8) and parent::*)]//*[contains(concat('
                                       ' " ", @class, " " ), concat( " ", "ng-scope",'
                                       ' " " ))]',
                               "교수명": '//*[contains(concat( " ", @class, " " ), concat('
                                      ' " ", "sp-grid-data-column", " " )) and ((('
                                      'count(preceding-sibling::*) + 1) = 9) and'
                                      ' parent::*)]//*[contains(concat( " ", @class,'
                                      ' " " ), concat( " ", "ng-scope", " " ))]',
                               "강의시간": '//*[contains(concat( " ", @class, " " ), concat('
                                       ' " ", "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                       ' + 1) = 10) and parent::*)]//*[contains(concat( "'
                                       ' ", @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "강의실": '//*[contains(concat( " ", @class, " " ), concat('
                                      ' " ", "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                      ' + 1) = 11) and parent::*)]//*[contains(concat('
                                      ' " ", @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "영어강의": '//*[contains(concat( " ", @class, " " ), concat( " ",'
                                       ' "sp-grid-data-column", " " )) and (((count(preceding-'
                                       'sibling::*) + 1) = 12) and parent::*)]//*[contains(concat('
                                       ' " ", @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "영어강의등급": '//*[contains(concat( " ", @class, " " ), concat('
                                         ' " ", "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                         ' + 1) = 13) and parent::*)]//*[contains(concat('
                                         ' " ", @class, " " ), concat( " ", "ng-scope", " " ))]',
                               "특기사항": '//*[contains(concat( " ", @class, " " ), concat('
                                       ' " ", "ng-scope", " " )) and (((count(preceding-sibling::*)'
                                       ' + 1) = 14) and parent::*)]//*[contains(concat('
                                       ' " ", @class, " " ), concat( " ", "ng-scope", " " ))]'}

    subject_dict = {"개설학부": [], "과목명": [], "수강번호": [],
                    "교과구분": [], "학점": [], "주별시간": [],
                    "교수명": [], "강의시간": [], "강의실": [],
                    "영어강의": [], "영어강의등급": [], "특기사항": [], "목록위치": []}

    for page in range(1, entire_page_num + 1):  # 전체 페이지 스크래핑
        time.sleep(0.5)

        # 개설학부 정보 스크래핑
        majors = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["개설학부"])))
        for index, major in enumerate(majors, 1):
            # 처음의 공백 20개를 생략해줌
            if index >= 21:
                subject_dict["개설학부"].append(major.text)

        # 과목명 정보 스크래핑
        subjects = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["과목명"])))
        for index, subject in enumerate(subjects, 1):
            # 처음의 공백 20개를 생략해줌
            if index >= 21:
                subject_dict["과목명"].append(subject.text)

        # 수강번호 정보 스크래핑
        subject_nums = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["수강번호"])))
        for index, subject_num in enumerate(subject_nums, 1):
            # 처음의 공백 20개를 생략해줌
            if index >= 13:
                subject_dict["수강번호"].append(subject_num.text)
        subject_dict["수강번호"] = [subject_num for subject_num in subject_dict["수강번호"] if len(subject_num) > 0]

        # 교과구분 정보 스크래핑
        subject_types = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["교과구분"])))
        for index, subject_type in enumerate(subject_types, 1):
            # 처음의 공백 12개를 생략해줌
            if index >= 13:
                subject_dict["교과구분"].append(subject_type.text)

        # 학점 정보 스크래핑
        grades = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["학점"])))
        for index, grade in enumerate(grades, 1):
            # 처음의 공백 6개를 생략해줌
            if index >= 7:
                subject_dict["학점"].append(grade.text)

        # 주별시간 정보 스크래핑
        weekly_times = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["주별시간"])))
        for weekly_time in weekly_times:
            subject_dict["주별시간"].append(weekly_time.text)

        # 교수명 정보 스크래핑
        professors = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["교수명"])))
        for professor in professors:
            subject_dict["교수명"].append(professor.text)

        # 강의시간 정보 스크래핑
        lecture_times = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["강의시간"])))
        for lecture_time in lecture_times:
            subject_dict["강의시간"].append(lecture_time.text)

        # 강의실 정보 스크래핑
        locations = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["강의실"])))
        for location in locations:
            subject_dict["강의실"].append(location.text)

        # 영어강의 정보 스크래핑
        is_englishs = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["영어강의"])))
        for is_english in is_englishs:
            subject_dict["영어강의"].append(is_english.text)

        # 영어강의등급 정보 스크래핑
        english_grades = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["영어강의등급"])))
        for english_grade in english_grades:
            subject_dict["영어강의등급"].append(english_grade.text)

        # 특기사항 정보 스크래핑
        special_infos = Wait(driver, 1.5). \
            until(ec.presence_of_all_elements_located((By.XPATH, subject_info_xpath_dict["특기사항"])))
        for special_info in special_infos:
            subject_dict["특기사항"].append(special_info.text)

        subject_dict["목록위치"].extend(list(zip([page] * len(special_infos), list(range(1, len(special_infos) + 1)))))

        # 마지막 페이지가 아닌 경우에 다음 페이지 버튼 클릭
        if page != entire_page_num:
            next_page_bt = Wait(driver, 1.5). \
                until(ec.presence_of_element_located((By.XPATH,
                                                      '//*[@id="pageContainer"]/dynamic-tab[2]/'
                                                      'u021301/div/div[5]/sp-grid-paging/section/div/div[3]/button')))
            driver.execute_script("arguments[0].click();", next_page_bt)
        # 총 페이지가 2페이지 이상인 상황에서 마지막 페이지인 경우에 첫 페이지로 돌아가 줌(다른 옵션을 선택한 다음 다시 검색했을 때 자동으로 1페이지로 이동하지 않음)
        elif page == entire_page_num and entire_page_num != 1:
            first_page_bt = Wait(driver, 1.5). \
                until(ec.presence_of_element_located((By.XPATH,
                                                      '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[5]/'
                                                      'sp-grid-paging/section/div/div[1]/button')))
            driver.execute_script("arguments[0].click();", first_page_bt)

    print(subject_dict)
    subject_df = pd.DataFrame(subject_dict)

    os.system("cls")
    subject_df_for_print = subject_df[["과목명", "교수명", "강의시간"]]
    print(tabulate(subject_df_for_print, headers='keys', tablefmt='psql'))

    selected_subject_num = input("\n위에서 원하는 과목의 번호를 입력해 주세요: ")
    while True:
        if selected_subject_num in map(str, range(len(subject_df.index))):
            selected_subject_num = int(selected_subject_num)
            break
        else:
            selected_subject_num = input("번호를 다시 입력해 주세요: ")

    if file_type == 1:
        return list(subject_df[["과목명", "강의시간", "강의실"]].iloc[selected_subject_num])
    else:
        subject_loca_info = subject_df.loc[subject_df.index == int(selected_subject_num), "목록위치"].values[0]
        subject_page_num = subject_loca_info[0]
        subject_list_loca = subject_loca_info[1]

        if subject_page_num >= 2:
            for page in range(1, subject_page_num):
                next_page_bt = Wait(driver, 1.5). \
                    until(ec.presence_of_element_located((By.XPATH,
                                                          '//*[@id="pageContainer"]/dynamic-tab[2]/'
                                                          'u021301/div/div[5]/sp-grid-paging/section/div/div[3]/button')))
                driver.execute_script("arguments[0].click();", next_page_bt)

        lecture_grid = Wait(driver, 1.5). \
            until(ec.presence_of_element_located((By.XPATH,
                                                  '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div['
                                                  '4]/div/div/div/div[2]/div/div[2]/div[' + str(
                                                      subject_list_loca * 3) + ']')))
        driver.execute_script("arguments[0].click();", lecture_grid)

        lecture_plan_bt = Wait(driver, 1.5). \
            until(ec.presence_of_element_located(
            (By.XPATH, '//*[@id="pageContainer"]/dynamic-tab[2]/u021301/div/div[3]/div[2]/div[1]/button')))
        driver.execute_script("arguments[0].click();", lecture_plan_bt)

        is_info_available = False

        # 처음에 과목 이름으로 강의계획서가 열린 것을 확인함
        subject_name = Wait(driver, 10). \
            until(ec.presence_of_element_located((By.XPATH, '//*[@id="m2soft-crownix-text"]/div[1]')))

        try:
            no_data_check_bt = Wait(driver, 1). \
                until(
                ec.presence_of_element_located((By.XPATH, '//*[@id="m2soft-crownix-container"]/div[5]/div/div/button')))
            driver.execute_script("arguments[0].click();", no_data_check_bt)
        except TimeoutException:
            is_info_available = True

        # 강의계획서 데이터가 있는 경우
        if is_info_available:
            all_plan_html = ''
            for i in range(6):
                time.sleep(0.5)
                plan_html = driver.page_source
                all_plan_html = all_plan_html + plan_html

                plan_next_page_bt = Wait(driver, 1). \
                    until(ec.presence_of_element_located((By.XPATH, '//*[@id="crownix-toolbar-next"]/button')))
                driver.execute_script("arguments[0].click();", plan_next_page_bt)
        else:
            os.system("cls")
            print("※ 현재 해당 과목의 강의계획서 정보가 없습니다.")
            all_plan_html = None
            time.sleep(2)
            return all_plan_html

        try:
            plan_element_search_by_left(322, all_plan_html, 1)[0]
        except:
            os.system("cls")
            print("※ 현재 해당 과목의 강의계획서 정보가 없습니다.")
            all_plan_html = None
            time.sleep(2)

        return all_plan_html


def lecture_plan_scraper(all_plan_html):
    lecture_info = {"과목이름": plan_element_search(51, 68, all_plan_html),
                    "학수구분": plan_element_search_by_left(322, all_plan_html, 1)[0],
                    "수강번호": plan_element_search(649, 98, all_plan_html),
                    "주수강대상": "<br/>".join(plan_element_search_by_left(322, all_plan_html, 1)[1:-1]),
                    "개설년도/학기": "".join(list(map(lambda x: x if len(re.findall("[가-힣]", x)) >= 1 else "",
                                                plan_element_search_by_left(649, all_plan_html, 1)))),
                    "강의시간 및 강의실": plan_element_search_by_left(322, all_plan_html, 1)[-1],
                    "영어등급": "",
                    "선수과목": "",
                    "관련 기초과목": "",
                    "동시수강 추천과목": "",
                    "관련 고급과목": "",
                    "담당교수 성명": "",
                    "담당교수 연구실": "",
                    "담당교수 구내전화": "",
                    "담당교수 e-mail": "",
                    "담당교수 상담시간": "",
                    "담당교수 홈페이지": "",
                    "교과목 개요": get_element_by_tops_sequence(62, 453, 16, 17, all_plan_html, 2),
                    "수업 목표": get_element_by_tops_sequence(60, 715, 16, 17, all_plan_html, 2),
                    "수업 형태 및 진행방식": plan_element_search_by_left(63, all_plan_html, 2, 2),
                    "4,5 column1": "",
                    "4,5 column2": "",
                    "4,5 column3": "",
                    "수강에 필요한 기초지식 및 도구능력": plan_element_search_by_left(61, all_plan_html, 2, 2),
                    "학습평가 방법": "",
                    "교재 및 참고자료": "",
                    "수업내용의 체계 및 진도계획": "",
                    "진도 계획": ""
                    }

    # 0.1 영어등급 정보 가져오기
    English_grade = re.findall("[a-zA-Z]", plan_element_search_by_left(649, all_plan_html, 1)[-1])

    if len(English_grade) >= 1:
        lecture_info["영어등급"] = English_grade
    else:
        lecture_info["영어등급"] = ""

    # 0.2 선수과목, 관련 기초과목, 동시수강 추천과목, 관련 고급과목 정보 가져오기
    other_subjects_types = plan_element_search_by_left(136, all_plan_html, 1)
    other_subjects_types_tops = get_top_location_by_left(136, all_plan_html)

    other_subjects_types = other_subjects_types[3:]
    other_subjects_types_tops = other_subjects_types_tops[3:]
    other_subjects_types_tops = list(map(int, other_subjects_types_tops))

    other_subjects = plan_element_search_by_left(325, all_plan_html, 1)
    other_subjects_tops = get_top_location_by_left(325, all_plan_html)
    other_subjects_tops = list(map(int, other_subjects_tops))

    for i in range(len(other_subjects_tops)):
        for j in range(len(other_subjects_types_tops)):
            if other_subjects_tops[i] - 1 <= other_subjects_types_tops[j] <= other_subjects_tops[i] + 1:
                lecture_info[other_subjects_types[j]] = other_subjects[i]

    # 0.3 담당교수 관련 정보 가져오기
    if len(plan_element_search_by_left(326, all_plan_html, 1)) >= 1:
        lecture_info["담당교수 성명"] = plan_element_search_by_left(326, all_plan_html, 1)[0]
    if len(plan_element_search_by_left(212, all_plan_html, 1)) >= 1:
        lecture_info["담당교수 연구실"] = plan_element_search_by_left(212, all_plan_html, 1)[0]
    if len(plan_element_search_by_left(414, all_plan_html, 1)) >= 2:
        lecture_info["담당교수 구내전화"] = plan_element_search_by_left(414, all_plan_html, 1)[0]
    if len(plan_element_search_by_left(560, all_plan_html, 1)) >= 1:
        lecture_info["담당교수 e-mail"] = plan_element_search_by_left(560, all_plan_html, 1)[0]
    if len(plan_element_search_by_left(209, all_plan_html, 1)) >= 1:
        lecture_info["담당교수 상담시간"] = plan_element_search_by_left(209, all_plan_html, 1)[0]
    if len(plan_element_search_by_left(492, all_plan_html, 1)) >= 2:
        lecture_info["담당교수 홈페이지"] = plan_element_search_by_left(492, all_plan_html, 1)[0]

    # 5. 수업운영방법, 6. 수업지원시스템 활용방법 정보 가져오기(1)
    column1_checks_tops = get_top_location_by_left(73, all_plan_html)
    column1_names_tops = get_top_location_by_left(92, all_plan_html)

    column1_checks_location = [" ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ "]
    for i in range(6):
        try:
            column1_names_tops_index = column1_names_tops.index(column1_checks_tops[i])
        except:
            pass
        else:
            column1_checks_location[column1_names_tops_index] = " V "

    lecture_info["4,5 column1"] = column1_checks_location

    # 5. 수업운영방법, 6. 수업지원시스템 활용방법 정보 가져오기(2)
    column2_checks_tops = get_top_location_by_left(287, all_plan_html)
    column2_names_tops = get_top_location_by_left(306, all_plan_html)

    column2_checks_location = [" ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ "]
    for i in range(6):
        try:
            column2_names_tops_index = column2_names_tops.index(column2_checks_tops[i])
        except:
            pass
        else:
            column2_checks_location[column2_names_tops_index] = " V "

    lecture_info["4,5 column2"] = column2_checks_location

    # 5. 수업운영방법, 6. 수업지원시스템 활용방법 정보 가져오기(3)
    column3_checks_tops = get_top_location_by_left(511, all_plan_html)
    column3_names_tops = get_top_location_by_left(530, all_plan_html)

    column3_checks_location = [" ‍  ‍  ‍ ", " ‍  ‍  ‍ ", " ‍  ‍  ‍ "]
    for i in range(6):
        try:
            column3_names_tops_index = column3_names_tops.index(column3_checks_tops[i])
        except:
            pass
        else:
            column3_checks_location[column3_names_tops_index] = " V "

    lecture_info["4,5 column3"] = column3_checks_location

    # 8. 학습평가 방법 정보 가져오기

    assessment_types = plan_element_search_by_left(51, all_plan_html, 1)

    try:
        assessment_start_index = assessment_types.index("평가항목")
    except:
        assessment_start_index = -1
    else:
        assessment_end_index = assessment_types.index("9. 교재 및 참고자료") - 1

        while True:
            if ("study" in assessment_types[assessment_end_index] or "time" in assessment_types[assessment_end_index] or
                    "주당" in assessment_types[assessment_end_index] or "시간" in assessment_types[assessment_end_index]):
                assessment_end_index -= 1
            else:
                break

    assessment_types = assessment_types[assessment_start_index + 1:assessment_end_index + 1]
    assessment_types_tops = get_top_location_by_left(51, all_plan_html)
    assessment_types_tops = assessment_types_tops[assessment_start_index + 1:assessment_end_index + 1]

    try:
        two_page_case_index = assessment_types.index("8. 학습평가 방법")
    except ValueError:
        pass
    else:
        del assessment_types[two_page_case_index: two_page_case_index + 2]
        del assessment_types_tops[two_page_case_index: two_page_case_index + 2]

    assessment_counts = plan_element_search_by_left(207, all_plan_html, 1)
    assessment_counts_tops = get_top_location_by_left(207, all_plan_html)

    assessment_ratios = plan_element_search_by_left(329, all_plan_html, 1)
    assessment_ratios_tops = get_top_location_by_left(329, all_plan_html)

    assessment_pss = plan_element_search_by_left(460, all_plan_html, 1)
    assessment_pss_tops = get_top_location_by_left(460, all_plan_html)
    assessment_pss, assessment_pss_tops = element_concatenate(assessment_pss, assessment_pss_tops, 17)

    assessment_info = selective_zipper(assessment_types, assessment_types_tops, (assessment_counts, assessment_pss),
                                       (assessment_counts_tops, assessment_pss_tops))

    assessment_located_ratios = []

    for i in range(len(assessment_types)):
        try:
            assessment_ratio = assessment_ratios[assessment_ratios_tops.index(assessment_types_tops[i])]
        except ValueError:
            assessment_ratio = ""
        assessment_located_ratios.append(assessment_ratio)

    for index, info in enumerate(assessment_info):
        info.insert(2, assessment_located_ratios[index])

    lecture_info["학습평가 방법"] = assessment_info

    # 9 교재 및 참고자료 정보 가져오기
    book_types = plan_element_search_by_left(51, all_plan_html, 1)
    book_types_tops = get_top_location_by_left(51, all_plan_html)
    book_types_tops = book_types_tops[book_types.index("9. 교재 및 참고자료") + 2:book_types.index("10. 수업내용의 체계 및 진도계획")]
    book_types = book_types[book_types.index("9. 교재 및 참고자료") + 2:book_types.index("10. 수업내용의 체계 및 진도계획")]
    book_types, book_types_tops = element_concatenate(book_types, book_types_tops, 17)

    books = plan_element_search_by_left(121, all_plan_html, 1)
    books_tops = get_top_location_by_left(121, all_plan_html)
    books, books_tops = element_concatenate(books, books_tops, 17)

    writers = plan_element_search_by_left(413, all_plan_html, 1)
    writers_tops = get_top_location_by_left(413, all_plan_html)
    writers, writers_tops = element_concatenate(writers, writers_tops, 17)

    publishers = plan_element_search_by_left(524, all_plan_html, 1)
    publishers_tops = get_top_location_by_left(524, all_plan_html)
    publishers, publishers_tops = element_concatenate(publishers, publishers_tops, 17)

    published_years = plan_element_search_by_left(673, all_plan_html, 1)
    published_years_tops = get_top_location_by_left(673, all_plan_html)
    del published_years[0]
    del published_years_tops[0]
    published_years, published_years_tops = element_concatenate(published_years, published_years_tops, 17)

    all_info = selective_zipper(books, books_tops, (book_types, writers, publishers, published_years),
                                (book_types_tops, writers_tops, publishers_tops, published_years_tops))

    for info in all_info:
        info[0], info[1] = info[1], info[0]

    lecture_info["교재 및 참고자료"] = all_info

    # 10 수업내용의 체계 및 진도계획 정보 가져오기

    tenth_info_title_top = get_top_location_by_left(51, all_plan_html)[
        plan_element_search_by_left(51, all_plan_html, 1).index("10. 수업내용의 체계 및 진도계획")]

    first_tenth_info = get_element_by_tops_sequence(62, tenth_info_title_top, 37, 38, all_plan_html, 1)
    first_tenth_info = first_tenth_info.replace("<br/>", "")

    if len(first_tenth_info) >= 1:
        first_info_top = get_top_location_by_left(62, all_plan_html)[
            plan_element_search_by_left(62, all_plan_html, 1).index(first_tenth_info)]
        tenth_info = get_element_by_tops_sequence(62, first_info_top, 16, 17, all_plan_html, 2)
    else:
        tenth_info = ""

    lecture_info["수업내용의 체계 및 진도계획"] = tenth_info

    # 10.2 진도계획 가져오기
    progress_weeks = plan_element_search_by_left(51, all_plan_html, 1)
    progress_weeks_tops = get_top_location_by_left(51, all_plan_html)
    progress_weeks_start_index = progress_weeks.index("&lt; 진도 계획 &gt;") + 2
    progress_weeks = progress_weeks[progress_weeks_start_index:]
    progress_weeks_tops = progress_weeks_tops[progress_weeks_start_index:]

    try:
        two_page_case_start_index = progress_weeks.index("&lt; 진도 계획 &gt;")
    except ValueError:
        pass
    else:
        del progress_weeks[two_page_case_start_index:two_page_case_start_index + 2]
        del progress_weeks_tops[two_page_case_start_index:two_page_case_start_index + 2]

    progress_titles = plan_element_search_by_left(94, all_plan_html, 1)
    progress_titles_tops = get_top_location_by_left(94, all_plan_html)
    progress_titles, progress_titles_tops = element_concatenate(progress_titles, progress_titles_tops, 18)

    progress_languages = plan_element_search_by_left(329, all_plan_html, 1)
    progress_languages_tops = get_top_location_by_left(329, all_plan_html)
    progress_languages, progress_languages_tops = element_concatenate(progress_languages, progress_languages_tops, 18)

    for i in reversed(range(len(progress_languages))):
        if len(re.findall("평가비율", progress_languages[i])) >= 1 or len(re.findall(r"\d", progress_languages[i])) >= 1:
            del progress_languages[i]
            del progress_languages_tops[i]

    progress_professors = plan_element_search_by_left(363, all_plan_html, 1)
    progress_professors_tops = get_top_location_by_left(363, all_plan_html)
    progress_professors, progress_professors_tops = element_concatenate(progress_professors, progress_professors_tops,
                                                                        18)

    progress_methods = plan_element_search_by_left(422, all_plan_html, 1)
    progress_methods_tops = get_top_location_by_left(422, all_plan_html)

    try:
        while True:
            progress_method_name_index = progress_methods.index("수업방법")
            del progress_methods[progress_method_name_index]
            del progress_methods_tops[progress_method_name_index]
    except ValueError:
        pass

    progress_methods, progress_methods_tops = element_concatenate(progress_methods, progress_methods_tops, 18)

    assessment_methods = plan_element_search_by_left(532, all_plan_html, 1)
    assessment_methods_tops = get_top_location_by_left(532, all_plan_html)

    try:
        while True:
            assessment_method_name_index = assessment_methods.index("평가방법")
            del assessment_methods[assessment_method_name_index]
            del assessment_methods_tops[assessment_method_name_index]

    except ValueError:
        pass

    assessment_methods, assessment_methods_tops = element_concatenate(assessment_methods, assessment_methods_tops, 18)

    lecture_info["진도 계획"] = selective_zipper(progress_weeks, progress_weeks_tops,
                                             (progress_titles, progress_languages, progress_professors, progress_methods,
                                              assessment_methods),
                                             (progress_titles_tops, progress_languages_tops, progress_professors_tops,
                                              progress_methods_tops, assessment_methods_tops), 2)

    return lecture_info


def subject_plan_maker(lecture_info, file_name):
    plan_html1 = """<!DOCTYPE html>
    <html>
    <head>
    <style>
    table, th, td {
      border: 1px solid #f2f2f2; 
    }

    th, td{
      text-align: center;
    }

    table {
      border-collapse: collapse;
      width: 900px;
      table-layout: fixed;
    }

    th{
      background-color: #C0C0C0;
      font-weight: normal;
    }

    #title_cell{
      background-color: #C0C0C0;
    }

    #outer1{
      border: 1px solid black;
      width: 900px;
      min-height: 200px
    }

    #outer2{
      border: 1px solid black;
      width: 900px;
      min-height: 100px
    }

    #inner{
      border: 1px solid black;
      display:inline;
    }

    </style>
    </head>
    <body>

    <table>
      <tr>
          <td colspan="5" class="label"><b>%s</b></td>
      </tr>
      <tr>
        <td id="title_cell" rowspan="3">교과목</td>
        <td id="title_cell">학수구분(학점/시간)</td>
        <td>%s</td>
        <td id="title_cell">수강번호</td>
        <td>%s</td>
      </tr>
      <tr>
        <td id="title_cell">주수강대상</td>
        <td>%s</td>
        <td id="title_cell">개설년도/학기</td>
        <td>%s</td>
      </tr>
      <tr>
        <td id="title_cell">강의시간 및 강의실</td>
        <td>%s</td>
        <td id="title_cell">영어등급</td>
        <td>%s</td>
      </tr>
      <tr>
        <td id="title_cell" rowspan="4">교육과정<br/>참고사항</td> 
        <td id="title_cell">선수과목</td>
        <td colspan="3" style="text-align:left">%s</td>
      </tr>
      <tr>
        <td id="title_cell">관련 기초과목</td>
        <td colspan="3" style="text-align:left">%s</td>
      </tr>
      <tr>
        <td id="title_cell">동시수강 추천과목</td>
        <td colspan="3" style="text-align:left">%s</td>
      </tr>
      <tr>
        <td id="title_cell">관련 고급과목</td>
        <td colspan="3" style="text-align:left">%s</td>
      </tr>
    </table>

    <br>

    <table>
      <tr>
        <td id="title_cell" rowspan="3">담당교수</td>
        <td id="title_cell" colspan="2">성명(직위/소속)</td>
        <td colspan="4" style="text-align:left">%s</td>
      </tr>
      <tr>
        <td id="title_cell">연구실</td>
        <td style="text-align:left">%s</td>
        <td id="title_cell">구내전화</td>
        <td>%s</td>
        <td id="title_cell">e-mail</td>
        <td style="text-align:left">%s</td>
      </tr>
       <tr>
        <td id="title_cell">상담시간</td>
        <td colspan="2">%s</td>
        <td id="title_cell">홈페이지</td>
        <td colspan="2" style="text-align:left">%s</td>
      </tr>
    </table>

    <p><b>1. 교과목 개요</b></p>
    <div id=outer1>
    %s
    </div>

    <p><b>2. 수업 목표</b></p>
    <div id=outer1>
    %s
    </div>

    <p><b>3. 수업의 형태 및 진행방식</b></p>
    <div id=outer1>
    %s
    </div>

    <p><b>4. 수업운영방법</b></p>
    <div id=outer2>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 강의
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 토론,토의
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>
    &nbsp 팀 프로젝트(발표, 사례연구 등)
    <br/>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 실험,실습(역할극 등)
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 설계,제작
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 현장학습(현장실습)
    <br/>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 기타
    <br/>
    <br/>
    </div>

    <p><b>5. 수업지원시스템 활용방법</b></p>
    <div id=outer2>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 아주 Bb
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 자동녹화시스템
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 웹과제
    <br/>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 사이버강의
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 온라인 콘텐트 활용
    <br/>
    <br/>
    &nbsp&nbsp&nbsp<div id=inner>%s</div>&nbsp 수업행동분석시스템
    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<div id=inner>%s</div>
    &nbsp 기타
    <br/>
    <br/>
    </div>

    <p><b>7. 수강에 필요한 기초지식 및 도구능력</b></p>
    <div id=outer1>
    %s
    </div>""" % (lecture_info["과목이름"],
                 lecture_info["학수구분"], lecture_info["수강번호"],
                 lecture_info["주수강대상"], lecture_info["개설년도/학기"],
                 lecture_info["강의시간 및 강의실"], lecture_info["영어등급"],
                 lecture_info["선수과목"], lecture_info["관련 기초과목"], lecture_info["동시수강 추천과목"], lecture_info["관련 고급과목"],
                 lecture_info["담당교수 성명"], lecture_info["담당교수 연구실"], lecture_info["담당교수 구내전화"],
                 lecture_info["담당교수 e-mail"], lecture_info["담당교수 상담시간"], lecture_info["담당교수 홈페이지"],
                 lecture_info["교과목 개요"],
                 lecture_info["수업 목표"],
                 lecture_info["수업 형태 및 진행방식"],
                 lecture_info["4,5 column1"][0], lecture_info["4,5 column2"][0], lecture_info["4,5 column3"][0],
                 lecture_info["4,5 column1"][1], lecture_info["4,5 column2"][1], lecture_info["4,5 column3"][1],
                 lecture_info["4,5 column1"][2],
                 lecture_info["4,5 column1"][3], lecture_info["4,5 column2"][2], lecture_info["4,5 column3"][2],
                 lecture_info["4,5 column1"][4], lecture_info["4,5 column2"][3],
                 lecture_info["4,5 column1"][5], lecture_info["4,5 column2"][4],
                 lecture_info["수강에 필요한 기초지식 및 도구능력"]
                 )

    plan_html2 = """<p><b>8. 학습평가방법</b></p>
    <table>
      <tr>
        <th>평가항목</th>
        <th>횟수</th>
        <th>평가비율</th>
        <th>비고</th>
      </tr>
    """

    assessment_method_html = ""
    for i in range(len(lecture_info["학습평가 방법"])):
        assessment_method_html_part = """
        <tr>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
        </tr>""" % (lecture_info["학습평가 방법"][i][0], lecture_info["학습평가 방법"][i][1],
                    lecture_info["학습평가 방법"][i][2], lecture_info["학습평가 방법"][i][3])
        assessment_method_html += assessment_method_html_part

    plan_html2 += assessment_method_html
    plan_html2 += "</table>"

    plan_html3 = """<p><b>9. 교재 및 참고자료</b></p>
    <table>
      <tr>
        <th>구분</th>
        <th>교재 제목</th>
        <th>저자</th>
        <th>출판사</th>
        <th>출판년도</th>
      </tr>
    """

    book_html = ""
    for i in range(len(lecture_info["교재 및 참고자료"])):
        book_html_part = """
        <tr>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
        </tr>""" % (lecture_info["교재 및 참고자료"][i][0], lecture_info["교재 및 참고자료"][i][1],
                    lecture_info["교재 및 참고자료"][i][2], lecture_info["교재 및 참고자료"][i][3],
                    lecture_info["교재 및 참고자료"][i][4])
        book_html += book_html_part

    plan_html3 += book_html
    plan_html3 += "</table>"

    plan_html4 = """<p><b>10. 수업내용의 체계 및 진도계획</b></p>
    <div id=outer1>
    %s
    </div>

    <p><b><진도 계획></b></p>
    <table>
      <tr>
        <th>주</th>
        <th>강의주제</th>
        <th>언어</th>
        <th>담당교수</th>
        <th>수업방법</th>
        <th>평가방법</th>
      </tr>
    """ % (lecture_info["수업내용의 체계 및 진도계획"])

    progress_html = ""
    for i in range(len(lecture_info["진도 계획"])):
        progress_html_part = """
        <tr>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
          <td>%s</td>
        </tr>""" % (lecture_info["진도 계획"][i][0], lecture_info["진도 계획"][i][1],
                    lecture_info["진도 계획"][i][2], lecture_info["진도 계획"][i][3],
                    lecture_info["진도 계획"][i][4], lecture_info["진도 계획"][i][5])
        progress_html += progress_html_part
    plan_html4 += progress_html
    plan_html4 += "</table>"

    plan_html5 = """</body>
    </html>"""

    plan_html = plan_html1 + plan_html2 + plan_html3 + plan_html4 + plan_html5

    with open(file_name + ".html", "w", encoding='utf8') as plan:
        plan.write(plan_html)


def timetable_maker(added_subjects, file_name):
    lecture_names = list(added_subjects["과목 이름"].values)
    lecture_times = list(added_subjects["과목 일정"].values)
    lecture_places = list(added_subjects["강의실"].values)

    # 각 요일의 강의의 순서 및 해당 강의의 이름과 장소를 저장할 dictionary
    sequence_dic = {"월": [], "화": [], "수": [], "목": [], "금": [], "토": []}

    # 강의 시간 정보를 시간표에 반영 가능한 형태로 바꿔주며, 해당 강의 이름과 강의 장소 정보를 포함하는 dictionary를 완성하는 기능
    for index, lecture_time in enumerate(lecture_times):
        # 해당 강의의 강의 시간 정보가 없는 경우
        if lecture_time == "None":
            continue
        # 해당 강의의 강의 시간 정보가 나와있는 경우
        else:
            lecture_time2 = re.sub(r"(\([\w -]+\))", "", lecture_time)
            lecture_time2 = lecture_time2.strip(" ")
            lecture_time2 = lecture_time2.split(" ")

            # 강의 시간 정보가 "화18:00~19:30" 같은 식으로 되어있는 경우
            if "~" in lecture_time2[0]:
                day = lecture_time2[0][0]  # "화"만 따로 저장하기
                lecture_time2[0] = lecture_time2[0][1:]
                lecture_time3 = lecture_time2[0].split("~")
                lecture_time3 = lecture_time3[0]  # "18:00"만 남기기
                # 시간 부분만을 시간표의 위에서부터의 순서에 반영한 값
                lecture_time_num1 = (int(re.match(r"\d+", lecture_time3).group()) - 8) * 2
                # 분 부분만을 시간표의 위에서부터의 순서에 반영한 값
                lecture_time_num2 = "".join(re.findall(r":(\d+)", lecture_time3))

                if int(lecture_time_num2) == 0:
                    lecture_time_num2 = 0
                else:
                    lecture_time_num2 = 1
                # 시간 부분과 분 부분을 모두 시간표의 위에서부터의 순서에 반영한 값
                lecture_time_num3 = lecture_time_num1 + lecture_time_num2

                time_sequence = [lecture_time_num3, lecture_time_num3 + 1, lecture_time_num3 + 2]
                sequence_dic[day].extend([{j: "{0}<br/>{1}".format(lecture_names[index], lecture_places[index])}
                                          for j in time_sequence])
            else:
                for i in lecture_time2:
                    if i[1:] == "Z":
                        time_sequence = [0, 1]
                    elif i[1:] == "A":
                        time_sequence = [2, 3, 4]
                    elif i[1:] == "B":
                        time_sequence = [5, 6, 7]
                    elif i[1:] == "C":
                        time_sequence = [8, 9, 10]
                    elif i[1:] == "D":
                        time_sequence = [11, 12, 13]
                    elif i[1:] == "E":
                        time_sequence = [14, 15, 16]
                    elif i[1:] == "F":
                        time_sequence = [17, 18, 19]
                    elif i[1:] == "G":
                        time_sequence = [20, 21, 22]
                    elif i[1:] == "H":
                        time_sequence = [23, 24, 25]
                    elif i[1:] == "I":
                        time_sequence = [26, 27, 28]
                    elif i[1:] == "J":
                        time_sequence = [29, 30, 31]
                    # 시간이 숫자(정수)로 나와있는 경우 (예: 1.5로 되어 있으면 시간표 위에서(테이블명 제외) 3번째 위치해 있음)
                    else:
                        time_sequence = [int(float(i[1:]) * 2), int(float(i[1:]) * 2) + 1]
                    sequence_dic[i[0]].extend([{j: "{0}<br/>{1}".format(lecture_names[index], lecture_places[index])}
                                               for j in time_sequence])

    is_duplicated = False  # 시간표가 겹치는 날이 있는지의 여부
    duplicated_day = []  # 시간표가 겹치는 날들

    # 시간표가 겹치는 경우를 확인해 줌
    for key, value in sequence_dic.items():
        tsfc = []  # tsfc: time sequence for check

        for i in value:
            tsfc.append(list(i.keys())[0])

        if len(set(tsfc)) < len(tsfc):
            is_duplicated = True
            duplicated_day.append(key)

    # 시간표가 겹치는 경우
    if is_duplicated:
        duplicated_day = ",".join(duplicated_day)
        print("※ {0}요일에 강의시간이 서로 겹치는 과목들이 있어 시간표가 생성되지 않았습니다!".format(duplicated_day))
        time.sleep(2)
        return None

    # 시간표에 나오는 시간들 (8:00 ~ 23:30)
    hours = list(itertools.chain.from_iterable([[str(i), str(i)] for i in range(8, 24)]))  # 같은 hour가 두 번씩 나오는 것을 반영함
    hours = list(map(int, hours))
    time_li = []  # 시간표에 나오는 시간 정보들을 담게 될 리스트
    for count, hour in enumerate(hours, 1):  # time_li에 시간 정보들을 추가해줌
        # 정각인 경우
        if count % 2 == 1:
            time_li.append("{0}:00".format(hour))
        # 30분인 경우
        else:
            time_li.append("{0}:30".format(hour))

    # Z부터 J까지 영어 알파벳으로도 시간을 표현하는 것을 반영함
    time_alphabet = list(string.ascii_uppercase[:10])
    time_alphabet = list(itertools.chain.from_iterable([[i, i, i] for i in time_alphabet]))
    time_alphabet.insert(0, "Z")
    time_alphabet.insert(0, "Z")

    # 0부터 15까지 숫자(정수)로도 시간을 표현하는 것을 반영함
    time_num = list(itertools.chain.from_iterable([[i, i] for i in range(0, 16)]))

    # 교시/요일 및 교시 열에 들어갈 정보들
    time_in_left = ["{0}<br/>{1}".format(time_alphabet[i], time_li[i]) for i in range(0, 32)]
    time_in_right = ["{0}<br/>{1}".format(time_num[i], time_li[i]) for i in range(0, 32)]

    # 데이터프레임을 만들기 위해서 먼저 만들어 주는 dictionary
    timetable_dict = {"교시/요일": time_in_left, "월": [], "화": [], "수": [], "목": [], "금": [], "토": [],
                      "교시": time_in_right}

    for day in sequence_dic.keys():
        for j in range(0, 32):
            # 해당 순번에 강의가 있는 경우
            if j in [list(k.keys())[0] for k in sequence_dic[day]]:
                timetable_dict[day].extend([list(k.values())[0] for k in sequence_dic[day] if list(k.keys())[0] == j])
            # 해당 순번에 강의가 없는 경우
            else:
                timetable_dict[day].append("")

    timetable = pd.DataFrame(timetable_dict)

    # 시간표의 css 형식 지정
    # 이미지 제작 안하는 경우에는 width: 60%;으로 하기
    timetable_css = """
    <style>
    table{
      border-collapse: collapse;
      width: 100%;
      table-layout: fixed;
    }

    th{
      border: 1px solid #ddd;
      text-align: center;
      padding: 8px;
      height: 45px;
      color: white;
      background-color: #757575;
    }

    td{
      border: 1px solid #ddd;
      text-align: center;
      padding: 8px;
      height: 70px;
    }

    th:first-child, th:last-child{
      background-color: #6E6E6E;
    }

    td:first-child, td:last-child{
      background-color: #C0C0C0;
      font-weight: bold;
    }

    tr:nth-child(even){
      background-color: #f2f2f2;
    }
    </style>
    """

    # 시간표 파일 제작
    text_file = open(file_name + ".html", "x")
    text_file.write(timetable_css)
    text_file.write(timetable.to_html(index=False, escape=False))  # escape=False는 <br/>을 사용하기 위한 목적
    text_file.close()

    return 0


def image_maker(file_name, output_file_type, file_type=2):
    output_file_type = int(output_file_type)

    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    try:
        file_path = os.path.join(base_path, 'data_files/wkhtmltopdf/bin/wkhtmltoimage.exe')
        config = imgkit.config(wkhtmltoimage=file_path)
    except Exception:
        file_path = os.path.join(base_path, 'data_files/wkhtmltopdf/wkhtmltopdf/bin/wkhtmltoimage.exe')
        config = imgkit.config(wkhtmltoimage=file_path)

    if file_type == 1:
        encoding = "EUC-KR"
    else:
        encoding = "UTF8"

    options = {'format': 'png', "encoding": encoding}

    imgkit.from_file(file_name + ".html", file_name + ".png", config=config, options=options)

    if output_file_type == 1:
        os.remove(file_name + ".html")


def logout(driver):
    logout_bt = Wait(driver, 1.5). \
        until(ec.presence_of_element_located((By.XPATH,
                                              '//*[@id="header"]/div/div[3]/div[1]/div[3]/dl/dt/button')))
    driver.execute_script("arguments[0].click();", logout_bt)


if __name__ == "__main__":
    options = Options()
    options.add_argument("--headless")

    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    try:
        file_path = os.path.join(base_path, 'data_files/chromedriver84.exe')
        driver = access_test(file_path, options)
    except:
        try:
            file_path = os.path.join(base_path, 'data_files/chromedriver78.exe')
            driver = access_test(file_path, options)
        except:
            try:
                file_path = os.path.join(base_path, 'data_files/chromedriver80.exe')
                driver = access_test(file_path, options)
            except:
                os.system("cls")
                print("※ 크롬이 설치되어 있지 않거나 78~80 버전이 아닙니다.")
                time.sleep(2)
                sys.exit()

    initial_access(driver, False)

    # 검색 및 시간표 또는 강의계획서 출력을 반복적으로 하기 위한 while문
    while True:
        os.system("cls")
        file_type = input("제작을 원하시는 파일을 입력해 주세요(시간표: 1, 강의계획서: 기타): ")
        # 시간표를 제작하는 경우
        if file_type == "1":
            os.system("cls")

            file_type = int(file_type)
            file_name = input("제작할 시간표의 이름을 입력해 주세요: ")

            added_subjects = pd.DataFrame()
            count = 0
            while True:
                os.system("cls")
                print("현재 추가된 과목들\n")
                # 추가된 과목이 있는 경우
                if len(added_subjects) >= 1:
                    subject_info_to_print = added_subjects[["과목 이름", "과목 일정"]]
                    print(tabulate(subject_info_to_print, headers='keys', tablefmt='psql'))
                print("\n\n\n")
                print("1. 시간표에 반영할 과목 추가")
                print("2. 시간표에 반영할 과목 제거")
                print("3. 시간표 제작")
                action = input("\n※ 위에서 원하는 번호를 입력해 주세요: ")
                while True:
                    if action in map(str, range(1, 4)):
                        action = int(action)
                        break
                    else:
                        action = input("\n번호를 다시 입력해주세요: ")

                if action == 1:
                    if count >= 1:
                        driver.refresh()
                        time.sleep(1.5)
                        initial_access(driver, True)
                    subject_name_and_schedule = search(driver, file_type)

                    if subject_name_and_schedule is None:
                        count += 1
                        continue

                    added_subjects = added_subjects.append(pd.DataFrame({"과목 이름": [subject_name_and_schedule[0]],
                                                                         "과목 일정": [subject_name_and_schedule[1]],
                                                                         "강의실": [subject_name_and_schedule[2]]}),
                                                           ignore_index=True)
                elif action == 2:
                    subject_to_delete_loca = input("제거할 과목의 번호를 입력해주세요: ")
                    while True:
                        if subject_to_delete_loca in map(str, range(len(added_subjects.index))):
                            subject_to_delete_loca = int(subject_to_delete_loca)
                            break
                        else:
                            subject_to_delete_loca = input("\n제거할 과목의 번호를 다시 입력해주세요: ")

                    added_subjects.drop(subject_to_delete_loca, axis=0, inplace=True)

                elif action == 3:
                    if len(added_subjects) == 0:
                        print("※ 선택한 과목이 없습니다!")
                        time.sleep(0.5)
                        continue

                    result = timetable_maker(added_subjects, file_name)
                    if result == 0:
                        image_maker(file_name, 1, file_type)

                        os.system("cls")
                        print("※ 시간표 제작이 완료되었습니다!")
                        time.sleep(1.5)
                        break
                count += 1

        # 강의계획서를 제작하는 경우
        else:
            file_type = 2
            os.system("cls")
            output_file_type = input("제작할 강의계획서 파일의 종류를 선택해주세요(.png: 1, .html: 2, 둘 다: 나머지): ")
            file_name = input("\n제작할 강의계획서 파일의 이름을 입력해 주세요: ")

            try:
                all_plan_html = search(driver, file_type)
            except Exception:
                os.system("cls")
                print("※ 오류 발생으로 인하여 프로그램을 종료합니다.")
                time.sleep(2)
                break

            if all_plan_html is not None:
                lecture_info = lecture_plan_scraper(all_plan_html)
                os.system("cls")
                print("※ 강의 계획서 파일을 제작 중입니다...")
                subject_plan_maker(lecture_info, file_name)

                if output_file_type != "2":
                    image_maker(file_name, output_file_type)

                os.system("cls")
                print("※ 강의 계획서 파일 제작이 완료되었습니다!")
                time.sleep(1.5)

        os.system("cls")
        will_make_file_again = input("다른 파일을 이어서 제작하시겠습니까?(예: 1, 종료: 나머지): ")
        os.system("cls")

        if will_make_file_again != "1":
            break
        else:
            driver.refresh()
            time.sleep(1)
            initial_access(driver, True)

    logout(driver)
    driver.close()

## activate virtual
## cd C:\Users\User\PycharmProjects\MyProject2\기타 스크래핑 필요한 프로그램
## pyinstaller --add-data "data_files/chromedriver78.exe;data_files" --add-data "data_files/chromedriver79.exe;data_files" --add-data "data_files/chromedriver80.exe;data_files" --add-data "data_files/wkhtmltopdf;data_files" "시간표 및 강의계획서 제작 프로그램.py"