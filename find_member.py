import time

from openpyxl import Workbook
from selenium import webdriver


def get_member(guild_name):
    # 멤버를 담을 공백 리스트 생성
    member_list = []

    # 크롬 드라이버 옵션 설정
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")
    # 혹은 options.add_argument("--disable-gpu")
    path = "./chromedriver"
    driver = webdriver.Chrome(path, chrome_options=options)
    # driver = webdriver.Chrome(path)

    url = "https://maple.gg/guild/elysium/" + guild_name
    driver.get(url)
    driver.implicitly_wait(5)
    # guild_data = driver.find_element_by_xpath('//*[@id="guild-content"]/section')
    try:
        # maple.gg의 guild-content 블록을 가져오도록 함
        guild_data = driver.find_element_by_id('guild-content')
        # a 태그에는 길드원의 닉네임과 공백이 담겨있음
        guild_member_name = guild_data.find_elements_by_tag_name('a')

        for i in guild_member_name:
            # 공백이 아닐때만 추가하도록 함
            if i.text != '':
                member_list.append(i.text)

        print(member_list)
        driver.quit()
    except Exception as e:
        print(e)
        driver.quit()

    return member_list


def get_members_info(member_list):
    # 엑셀 파일 경로 지정
    file_path = "./memberlist.xlsx"
    write_wb = Workbook()
    write_ws = write_wb.active

    write_ws.cell(1, 1, "닉네임")
    write_ws.cell(1, 2, "무릉도장")
    # 크롬 드라이버 옵션 설정
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")
    # 혹은 options.add_argument("--disable-gpu")
    path = "./chromedriver"
    driver = webdriver.Chrome(path, chrome_options=options)
    # driver = webdriver.Chrome(path)
    for i in member_list:
        try:
            url = "https://maple.gg/u/" + i
            driver.get(url)
            time.sleep(3)
            # 갱신버튼을 누르도록 함
            driver.find_element_by_xpath('//*[@id="btn-sync"]').click()
            time.sleep(3)
            mureungdojang = driver.find_element_by_xpath(
                '//*[@id="app"]/div[3]/div/section/div[1]/div[1]/section/div/div[1]/div/h1')
            print('{0} / {1}'.format(i, mureungdojang.text))
            write_ws.cell(member_list.index(i)+2, 1, i)
            write_ws.cell(member_list.index(i)+2, 2, mureungdojang.text)
        except Exception as e:
            print(e)

            try:
                old_mureungdojang = driver.find_element_by_xpath(
                    '//*[@id="app"]/div[3]/div/section/div[1]/div[1]/section/div/div[3]/div/b')
                print('{0} / 예전 무릉 {1}'.format(i, old_mureungdojang.text))
                write_ws.cell(member_list.index(i) + 2, 1, i)
                write_ws.cell(member_list.index(i) + 2, 2, "예전 무릉 "+old_mureungdojang.text)
            except Exception as b:
                print('{0} / 0 층'.format(i))
                write_ws.cell(member_list.index(i) + 2, 1, i)
                write_ws.cell(member_list.index(i) + 2, 2, "0 층")
                print(b)

    write_wb.save(file_path)
    driver.quit()


start = time.time()  # 시작 시간 저장

# guildname = input("길드명을 입력해주세요 : ")
guildname = "LUSH"
memberlist = get_member(guildname)
get_members_info(memberlist)

print("소요시간 :", time.time() - start)  # 현재시각 - 시작시간 = 실행 시간
