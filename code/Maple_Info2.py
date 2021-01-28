import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import WebDriverException

import pandas as pd
from urllib.request import urlopen
from PIL import Image

import os
import tkinter
import tkinter as tk
import matplotlib as mpl
import sys
sys.setrecursionlimit(10000)
mpl.use('TkAgg')


# url 읽어오기
def get_html(url):
    _html = ""
    resp = requests.get(url)
    if resp.status_code == 200:
        _html = resp.text

    return _html


class FUNCTION:
    def __init__(self):
        # data save to excel
        base_dir = "./file/"
        file_name = "DataForm.xlsx"
        file_dir = os.path.join(base_dir, file_name)
        read_equipment = pd.read_excel(file_dir, sheet_name='Sheet1')
        read_equipment = read_equipment.fillna("-")
        self.character_equipment = read_equipment.values.tolist()

        # create list
        self.rank_name = list()
        self.rank_url = list()
        self.character_stat = list()
        self.character_info_url = list()
        self.img_list = list()
        self.equipment_list = list()  # 캐릭터 장비
        self.equipment_number_list = ["3", "18", "23", "28", "24", "25", "22",  # 캐릭터 장비
                                      "8", "13", "14", "19", "30", "12",
                                      "7", "1", "6", "11", "16", "21",
                                      "15", "10", "17", "20", "5"]

    # 랭킹 검색
    def parse_maple_rank(self, name):
        url = "https://maplestory.nexon.com/Ranking/World/Total?c=" + name
        rank_html = get_html(url)
        soup = BeautifulSoup(rank_html, 'html.parser')
        maple_rank = soup.find("table", {"class": "rank_table"}).find_all("td", {"class": "left"})

        for index in maple_rank:
            info_soup = index.find("a")
            _url = info_soup["href"]
            _text = info_soup.text.split(".")
            _num = _text[0]

            self.rank_name.append(_num)
            self.rank_url.append(_url)

    # 캐릭터가 랭킹에 있는지 체크
    def check_maple_info_url(self, name):
        for index in range(0, len(self.rank_name)):
            if self.rank_name[index] == name:
                url = "https://maplestory.nexon.com" + self.rank_url[index]
                return url

    # 캐릭터 정보 URL 수집
    def parse_maple_stat_url(self, name):
        url = self.check_maple_info_url(name)
        info_html = get_html(url)
        soup = BeautifulSoup(info_html, 'html.parser')

        maple_info_url = soup.find("div", {"class": "lnb_wrap"}).find_all("a")
        for x in range(0, len(maple_info_url)):
            self.character_info_url.append(maple_info_url[x]["href"])

    # 캐릭터 정보 수집
    def parse_maple_stat(self, name):
        url = "https://maplestory.nexon.com" + self.character_info_url[1]
        character_info_html = get_html(url)
        soup = BeautifulSoup(character_info_html, 'html.parser')

        maple_level = soup.find("div", {"class": "char_info"}).find_all("dd")
        _text = maple_level[0].text
        self.character_equipment[3][4] = _text[_text.find(".") + 1:]

        maple_character_image = soup.find("div", {"class": "char_img"}).find_all("img")
        img = Image.open(urlopen(maple_character_image[0]["src"]))
        img.save("./" + name + "/img/" + name + ".png")

        maple_stat = soup.find("div", {"class": "container_wrap"}).find_all("table", {"class": "table_style01"})
        _text = maple_stat[0].text.split("\n") + maple_stat[1].text.split("\n")
        for index in range(0, len(_text)):
            self.character_stat.append(_text[index])

        _stat = self.character_stat[104].split("증가")  # 하이퍼스탯
        for index in range(0, len(_stat)):
            if _stat[index].find("공격력과 마력") >= 0:
                self.character_equipment[8][32] = _stat[index][_stat[index].find("공격력과 마력") + 7:]
                self.character_equipment[9][32] = _stat[index][_stat[index].find("공격력과 마력") + 7:]
            if _stat[index].find("힘") >= 0:
                self.character_equipment[3][32] = _stat[index][_stat[index].find("힘") + 2:]
            if _stat[index].find("민첩성") >= 0:
                self.character_equipment[4][32] = _stat[index][_stat[index].find("민첩섭") + 4:]
            if _stat[index].find("지력") >= 0:
                self.character_equipment[5][32] = _stat[index][_stat[index].find("지력") + 3:]
            if _stat[index].find("운") >= 0:
                self.character_equipment[6][32] = _stat[index][_stat[index].find("운") + 2:]

    # 캐릭터 길드 검색
    def parse_maple_guild(self):
        url = "https://maplestory.nexon.com" + self.character_info_url[10]
        character_info_html = get_html(url)
        soup = BeautifulSoup(character_info_html, 'html.parser')

        maple_guild = soup.find("div", {"class": "tab01_con_wrap"}).find_all("h1")
        _text = maple_guild[0].text
        self.character_equipment[4][2] = _text

    # 캐릭터 장비 이미지 저장
    def parse_maple_equipment_image(self, name):
        img_url = "https://maplestory.nexon.com/" + self.character_info_url[2]
        img_html = get_html(img_url)
        soup = BeautifulSoup(img_html, 'html.parser')

        maple_rank = soup.find("div", {"class": "tab01_con_wrap"}).find_all("img")

        for index in maple_rank:
            _name = index["alt"]
            _url = index["src"]
            self.img_list.append((_name, _url))

        for x in range(0, len(self.equipment_list)):
            try:
                for y in range(0, len(self.img_list)):
                    if self.equipment_list[x][0].find(")") >= 0:  # 강화, 스타포스 된 아이템 고르기
                        if self.equipment_list[x][0][:self.equipment_list[x][0].find(")") + 1] == self.img_list[y][0]:
                            img = Image.open(urlopen(self.img_list[y][1]))
                            img.save("./" + name + "/img/" + str(x) + ".gif")
                            break
                    else:  # 강화, 스타포스 안된 아이템 고르기 > 반지나 견장처럼 업횟 적은 장비는 작은 되든 말든 놔두고 스타포스만 하는 사람들 있음
                        if self.equipment_list[x][0][:self.equipment_list[x][0].find("성") - 2].strip() == self.img_list[y][0].strip():
                            img = Image.open(urlopen(self.img_list[y][1]))
                            img.save("./" + name + "/img/" + str(x) + ".gif")
                            break

                        elif self.equipment_list[x][0].strip() == self.img_list[y][0].strip():
                            img = Image.open(urlopen(self.img_list[y][1]))
                            img.save("./" + name + "/img/" + str(x) + ".gif")
                            break

                if self.equipment_list[21][1][:self.equipment_list[21][1].find(")") + 1] == self.img_list[26][0]:
                    img = Image.open(urlopen(self.img_list[26][1]))
                    img.save("./" + name + "/img/21.gif")

            except IndexError:
                img = Image.open("./img/no.gif")
                img.save("./" + name + "/img/" + str(x) + ".gif")

        if self.equipment_list[21][1][:self.equipment_list[21][1].find(")") + 1] == self.img_list[26][0]:
            img = Image.open(urlopen(self.img_list[26][1]))
            img.save("./" + name + "/img/21.gif")

    # 캐릭터 정보 df 저장 함수
    def character_info_save_tool(self, df_x, df_y, x):
        if self.character_equipment[df_x][df_y] == self.character_stat[x]:
            self.character_equipment[df_x][df_y + 1] = self.character_stat[x + 1]

    # 캐릭터 정보 df 저장
    def character_info_save(self):
        for x in range(0, len(self.character_stat)):
            self.character_info_save_tool(2, 3, x)  # world
            self.character_info_save_tool(3, 1, x)  # job
            self.character_info_save_tool(6, 1, x)  # stat offensive power
            self.character_info_save_tool(8, 1, x)  # str
            self.character_info_save_tool(8, 3, x)  # dex
            self.character_info_save_tool(9, 1, x)  # int
            self.character_info_save_tool(9, 3, x)  # luk
            self.character_info_save_tool(7, 1, x)  # hp
            self.character_info_save_tool(7, 3, x)  # mp
            self.character_info_save_tool(7, 5, x)  # critical damage
            self.character_info_save_tool(8, 5, x)  # defense rate
            self.character_info_save_tool(9, 5, x)  # boss damage
            self.character_info_save_tool(11, 1, x)  # star force
            self.character_info_save_tool(12, 1, x)  # arcane force
            self.character_info_save_tool(14, 1, x)  # ability

    # 캐릭터 장비 정보 읽어오기
    def equipment_info_search(self):
        options = webdriver.ChromeOptions()  # 크롬 드라이버 옵션설정
        options.add_argument("headless")
        options.add_argument('window-size=1920x1080')
        options.add_argument("--disable-gpu")
        browser = webdriver.Chrome("./file/chromedriver.exe", options=options)
        url = "https://maplestory.nexon.com" + self.character_info_url[2]
        browser.get(url)  # 캐릭터 장비정보 검색

        # 장비정보 등록
        for x in range(0, 24):
            xpath = """//*[@id="container"]/div[2]/div[2]/div/div[2]/div[1]/ul/li[""" + self.equipment_number_list[x] + """]/img"""
            try:
                browser.find_element_by_xpath(xpath).click()
                _equipment = browser.find_element_by_class_name("item_info").find_element_by_tag_name("div").text
                self.equipment_list.append(_equipment.split("\n"))
            except WebDriverException:
                self.equipment_list.append("")
        browser.quit()

    # 장비정보 등록
    def equipment_info_register(self):
        count_y = 2

        for x in range(0, len(self.equipment_list)):
            try:
                count_x = 0

                potential_str = 0
                potential_dex = 0
                potential_int = 0
                potential_luk = 0
                potential_all = 0
                potential_attack_damage = 0
                potential_mama_damage = 0

                additional_potential_str = 0
                additional_potential_dex = 0
                additional_potential_int = 0
                additional_potential_luk = 0
                additional_potential_all = 0
                additional_potential_attack = 0
                additional_potential_mana = 0

                # 0 아이템 옵션 / 1 잠재옵션 / 2 에디셔널 잠재옵션
                sort_equipment = []
                for z in range(0, len(self.equipment_list[x])):
                    if self.equipment_list[x][z].find("장비분류") >= 0:
                        sort_equipment.append(z)
                    if self.equipment_list[x][z].find("잠재옵션") >= 0:
                        sort_equipment.append(z)
                    if self.equipment_list[x][z].find("기타") >= 0:
                        sort_equipment.append(z)

                # 아이템 옵션 STR, DEX, INT, LUK, 공격력, 마력, 올스탯%
                for y in range(sort_equipment[0] + 1, sort_equipment[1]):
                    if self.equipment_list[x][y].find("STR") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[3][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[3][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]
                    if self.equipment_list[x][y].find("DEX") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[4][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[4][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]
                    if self.equipment_list[x][y].find("INT") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[5][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[5][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]
                    if self.equipment_list[x][y].find("LUK") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[6][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[6][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]
                    if self.equipment_list[x][y].find("올스탯") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[7][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[7][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("%")]
                    if self.equipment_list[x][y].find("공격력") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[8][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[8][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]
                    if self.equipment_list[x][y].find("마력") >= 0:
                        if self.equipment_list[x][y + 1].find("(") == -1:
                            self.character_equipment[9][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:]
                        else:
                            self.character_equipment[9][x + 8] = self.equipment_list[x][y + 1][self.equipment_list[x][y + 1].find("+") + 1:self.equipment_list[x][y + 1].find("(")]

                # 잠재옵션 STR, DEX, INT, LUK, 공격력, 마력, 올스탯%
                for y in range(sort_equipment[1] + 1, sort_equipment[2]):
                    if self.equipment_list[x][y].find("STR") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:  # % 안존재
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[12][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[12][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:  # % 존재  # % 수치를 다 합쳐야함
                            potential_str = potential_str + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[17][x + 8] = potential_str
                    if self.equipment_list[x][y].find("DEX") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[13][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[13][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_dex = potential_dex + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[18][x + 8] = potential_dex
                    if self.equipment_list[x][y].find("INT") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[14][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[14][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_int = potential_int + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[19][x + 8] = potential_int
                    if self.equipment_list[x][y].find("LUK") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[15][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[15][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_luk = potential_luk + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[20][x + 8] = potential_luk
                    if self.equipment_list[x][y].find("올스탯") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[16][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[16][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_all = potential_all + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[21][x + 8] = potential_all
                    if self.equipment_list[x][y].find("공격력") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                potential_attack_damage = potential_attack_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:])
                                self.character_equipment[22][x + 8] = potential_attack_damage
                            else:
                                potential_attack_damage = potential_attack_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")])
                                self.character_equipment[22][x + 8] = potential_attack_damage
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_attack_damage = potential_attack_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[24][x + 8] = potential_attack_damage
                            self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                            count_x = count_x + 1
                    if self.equipment_list[x][y].find("마력") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                potential_mama_damage = potential_mama_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:])
                                self.character_equipment[23][x + 8] = potential_mama_damage
                            else:
                                potential_mama_damage = potential_mama_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")])
                                self.character_equipment[23][x + 8] = potential_mama_damage
                        elif self.equipment_list[x][y].find("%") >= 0:
                            potential_mama_damage = potential_mama_damage + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[25][x + 8] = potential_mama_damage
                            self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                            count_x = count_x + 1

                    if self.equipment_list[x][y].find("보스 몬스터") >= 0:
                        self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                        count_x = count_x + 1

                    if self.equipment_list[x][y].find("몬스터 방어율") >= 0:
                        self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                        count_x = count_x + 1

                # 에디셔널 잠재옵션
                # STR, DEX, INT, LUK, 공격력, 마력, 올스탯%
                for y in range(sort_equipment[2], sort_equipment[3]):
                    if self.equipment_list[x][y].find("STR") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:  # % 안존
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[28][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[28][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:  # % 존재 %를 다 합쳐야함`# -를 0으로 바꾸고 더하는거 짜야함
                            additional_potential_str = additional_potential_str + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[33][x + 8] = additional_potential_str
                    if self.equipment_list[x][y].find("DEX") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[29][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[29][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_dex = additional_potential_dex + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[34][x + 8] = additional_potential_dex
                    if self.equipment_list[x][y].find("INT") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[30][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[30][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_int = additional_potential_int + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[35][x + 8] = additional_potential_int
                    if self.equipment_list[x][y].find("LUK") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[31][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[31][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_luk = additional_potential_luk + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[36][x + 8] = additional_potential_luk
                    if self.equipment_list[x][y].find("올스탯") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                self.character_equipment[32][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:]
                            else:
                                self.character_equipment[32][x + 8] = self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1: self.equipment_list[x][y].find("(")]
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_all = additional_potential_all + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[37][x + 8] = additional_potential_all
                    if self.equipment_list[x][y].find("공격력") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                additional_potential_attack = additional_potential_attack + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:])
                                self.character_equipment[38][x + 8] = additional_potential_attack
                            else:
                                additional_potential_attack = additional_potential_attack + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")])
                                self.character_equipment[38][x + 8] = additional_potential_attack
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_attack = additional_potential_attack + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[40][x + 8] = additional_potential_attack
                            self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                            count_x = count_x + 1

                    if self.equipment_list[x][y].find("마력") >= 0:
                        if self.equipment_list[x][y].find("%") == -1:
                            if self.equipment_list[x][y + 1].find("(") == -1:
                                additional_potential_mana = additional_potential_mana + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:])
                                self.character_equipment[39][x + 8] = additional_potential_mana
                            else:
                                additional_potential_mana = additional_potential_mana + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("(")])
                                self.character_equipment[39][x + 8] = additional_potential_mana
                        elif self.equipment_list[x][y].find("%") >= 0:
                            additional_potential_mana = additional_potential_mana + int(self.equipment_list[x][y][self.equipment_list[x][y].find("+") + 1:self.equipment_list[x][y].find("%")])
                            self.character_equipment[41][x + 8] = additional_potential_mana
                            self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                            count_x = count_x + 1

                    if self.equipment_list[x][y].find("보스 몬스터") >= 0:
                        self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                        count_x = count_x + 1

                    if self.equipment_list[x][y].find("몬스터 방어율") >= 0:
                        self.character_equipment[18 + count_x][count_y] = self.equipment_list[x][y]
                        count_x = count_x + 1

                if count_x > 0:
                    count_y = count_y + 1

                if self.equipment_list[x][0].find("강화") >= 0:
                    self.character_equipment[43][x + 8] = self.equipment_list[x][0][self.equipment_list[x][0].find("성") - 2: self.equipment_list[x][0].find("성")]
                elif self.equipment_list[x][1].find("강화") >= 0:
                    self.character_equipment[43][x + 8] = self.equipment_list[x][1][self.equipment_list[x][1].find("성") - 2: self.equipment_list[x][1].find("성")]

                if self.equipment_list[x][sort_equipment[1] + 1].find("레어") >= 0:
                    self.character_equipment[44][x + 8] = "레      어"
                elif self.equipment_list[x][sort_equipment[1] + 1].find("에픽") >= 0:
                    self.character_equipment[44][x + 8] = "에      픽"
                elif self.equipment_list[x][sort_equipment[1] + 1].find("유니크") >= 0:
                    self.character_equipment[44][x + 8] = "유 니 크"
                elif self.equipment_list[x][sort_equipment[1] + 1].find("레전드리") >= 0:
                    self.character_equipment[44][x + 8] = "레      전"

                if self.equipment_list[x][sort_equipment[2] + 1].find("레어") >= 0:
                    self.character_equipment[45][x + 8] = "레      어"
                elif self.equipment_list[x][sort_equipment[2] + 1].find("에픽") >= 0:
                    self.character_equipment[45][x + 8] = "에      픽"
                elif self.equipment_list[x][sort_equipment[2] + 1].find("유니크") >= 0:
                    self.character_equipment[45][x + 8] = "유 니 크"
                elif self.equipment_list[x][sort_equipment[2] + 1].find("레전드리") >= 0:
                    self.character_equipment[45][x + 8] = "레      전"

            except IndexError as e:
                print(x, "no equipment\n", e)

    # 데이터 저장
    def save_excel(self, name):
        new_base_dir = "./" + name + "/"
        new_file_name = name + " profile.xlsx"
        new_file_dir = os.path.join(new_base_dir, new_file_name)

        # table data ~> .xlsx
        equipment_excel = pd.DataFrame(self.character_equipment)
        equipment_excel.to_excel(new_file_dir,
                                 na_rep='NaN',
                                 float_format="%.2f",
                                 header=False,
                                 index=False,
                                 startrow=0,
                                 startcol=0)


class MapleStatUi:
    def __init__(self):
        # 제목 / 크기 / 배경 / 폰트설정
        self.main_window = tkinter.Tk()
        self.main_window.title("Maple_STAT_UI")
        self.main_window.geometry("1018x768")
        self.main_window.resizable(False, False)
        background = tk.PhotoImage(file="./img/bg.gif")
        bg_label = tk.Label(self.main_window, image=background)
        bg_label.place(x=0, y=0)

        # 프로필생성 / 종료 / 스탯확인 / 흑우등급 버튼
        _create_image = tk.PhotoImage(file="./img/_프로필생성.png")
        btn_create = tk.Button(self.main_window, image=_create_image, bg="pink", overrelief="solid", command=self.call_create_profile, repeatdelay=1000)
        btn_create.place(x=830, y=300)
        _close_image = tk.PhotoImage(file="./img/_종료.png")
        close_window = tk.Button(self.main_window, image=_close_image, bg="pink", overrelief="solid", command=self.exit, repeatdelay=1000)
        close_window.place(x=950, y=300)

        # text_index > 1.0 일때만 누를 수 있도록
        # 1.0보다 크지 않으면 프로필 생성이 안됐...나? 아닌데 ㅋ
        # 체크할 수 있는거 생각해야하는데~
        _stat_image = tk.PhotoImage(file="./img/_장비스탯확인.png")
        btn_stat = tk.Button(self.main_window, image=_stat_image, bg="pink", overrelief="solid", command=self.call_view_stat, repeatdelay=1000)
        btn_stat.place(x=30, y=230)
        _cow_image = tk.PhotoImage(file="./img/_흑우등급확인.png")
        btn_cow = tk.Button(self.main_window, image=_cow_image, bg="pink", overrelief="solid", command=self.call_view_cow, repeatdelay=1000)
        btn_cow.place(x=30, y=280)
        # 진행상황 텍스트
        self.progress = tk.Text(self.main_window, width=28, height=10)
        self.progress.configure(font=("Arial", 9, "bold"))
        self.progress.place(x=20, y=60)
        self.text_index = 1.0
        # 이름 라벨/텍스트
        _name_image = tk.PhotoImage(file="./img/_이름.png")
        self.la_name = tk.Label(self.main_window, image=_name_image)
        self.la_name.place(x=830, y=130)
        self.te_name = tk.Entry(self.main_window, width=16, textvariable=str)
        self.te_name.place(x=880, y=130)
        # 주스탯, 부스탯, 공마 라벨
        _main_stat_image = tk.PhotoImage(file="./img/_주스탯.png")
        self.la_main_stat = tk.Label(self.main_window, image=_main_stat_image)
        self.la_main_stat.place(x=878, y=170)
        _sub_stat_image = tk.PhotoImage(file="./img/_부스탯.png")
        self.la_sub_stat = tk.Label(self.main_window, image=_sub_stat_image)
        self.la_sub_stat.place(x=921, y=170)
        _om_force_image = tk.PhotoImage(file="./img/_공마.png")
        self.la_om_force = tk.Label(self.main_window, image=_om_force_image)
        self.la_om_force.place(x=963, y=170)
        # 유니온 주스탯, 부스탯, 공마 라벨/텍스트
        _union_image = tk.PhotoImage(file="./img/_유니온.png")
        self.la_union = tk.Label(self.main_window, image=_union_image)
        self.la_union.place(x=830, y=200)
        self.te_union_ms = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_union_ms.insert(0, "0")
        self.te_union_ms.place(x=878, y=200)
        self.te_union_ss = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_union_ss.insert(0, "0")
        self.te_union_ss.place(x=921, y=200)
        self.te_union_om = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_union_om.insert(0, "0")
        self.te_union_om.place(x=963, y=200)
        # 장비1, 장비2, 장비3 라벨
        _la_equipment1_image = tk.PhotoImage(file="./img/_장비1.png")
        self.la_equipment1 = tk.Label(self.main_window, image=_la_equipment1_image)
        self.la_equipment1.place(x=878, y=230)
        _la_equipment2_image = tk.PhotoImage(file="./img/_장비2.png")
        self.la_equipment2 = tk.Label(self.main_window, image=_la_equipment2_image)
        self.la_equipment2.place(x=921, y=230)
        _la_equipment3_image = tk.PhotoImage(file="./img/_장비3.png")
        self.la_equipment3 = tk.Label(self.main_window, image=_la_equipment3_image)
        self.la_equipment3.place(x=963, y=230)
        # 펫장비 이름 라벨, 장비1, 장비2, 장비3 텍스트
        _pet_image = tk.PhotoImage(file="./img/_펫장비.png")
        self.la_pet_equipment = tk.Label(self.main_window, image=_pet_image)
        self.la_pet_equipment.place(x=830, y=260)
        self.te_pet_ms = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_pet_ms.insert(0, "0")
        self.te_pet_ms.place(x=879, y=260)
        self.te_pet_ss = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_pet_ss.insert(0, "0")
        self.te_pet_ss.place(x=921, y=260)
        self.te_pet_om = tk.Entry(self.main_window, width=5, textvariable=str)
        self.te_pet_om.insert(0, "0")
        self.te_pet_om.place(x=963, y=260)

        self.main_window.mainloop()

    # 프로필생성 버튼
    def call_create_profile(self):
        if len(self.te_name.get()) > 0:
            name = self.te_name.get()
            path = "./" + name

            if not os.path.isdir(path):
                os.mkdir(path)
                os.mkdir(path + "/img")

                call = FUNCTION()
                call.character_equipment[2][2] = name

                self.check_maple_rank(call, name)
                self.check_maple_url(call, name)
                self.read_maple_stat(call, name)
                self.check_maple_guild(call)
                self.save_maple_info(call)
                self.read_equipment_info(call)
                self.sort_equipment_info(call)
                self.save_equipment_image(call, name)
                self.create_character_profile(call, name)

        else:
            print("no name")  # 이름 없다고 텍스트박스 띄우기

    # 장비스탯확인 버튼
    def call_view_stat(self):
        if len(self.te_name.get()) > 0:
            name = self.te_name.get()
            STAT(name)
        else:
            print("no name")

    # 흑우등급확인 버튼
    def call_view_cow(self):
        if len(self.te_name.get()) > 0:
            name = self.te_name.get()
            COW(name)
        else:
            print("no name")

    # 종료 버튼
    def exit(self):
        print("*********Press Exit button********")
        print("*              종료               *")
        print("**********************************\n")
        self.main_window.destroy()

    # 랭킹에 캐릭터가 있는지 확인
    def check_maple_rank(self, call, name):
        self.progress.delete("1.0", "end")
        self.progress.insert(self.text_index, "1. check_maple_ranking")
        self.text_index = self.text_index + 1.0
        call.parse_maple_rank(name)

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터 정보 url 수집
    def check_maple_url(self, call, name):
        self.progress.insert(self.text_index, "2. check_maple_url")
        self.text_index = self.text_index + 1.0
        call.parse_maple_stat_url(name)

        self.progress.insert(self.text_index, "  ...	ok  \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터 기본스탯정보 수집
    def read_maple_stat(self, call, name):
        self.progress.insert(self.text_index, "3. read_maple_stat")
        self.text_index = self.text_index + 1.0
        call.parse_maple_stat(name)

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터 길드정보 수집
    def check_maple_guild(self, call):
        self.progress.insert(self.text_index, "4. check_maple_guild")
        self.text_index = self.text_index + 1.0
        call.parse_maple_guild()

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 수집한 캐릭터 기본정보 데이터 프레임으로 정리
    def save_maple_info(self, call):
        self.progress.insert(self.text_index, "5. save_maple_info")
        self.text_index = self.text_index + 1.0
        call.character_info_save()

        self.progress.insert(self.text_index, "...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터가 장착한 장비 정보 수집
    def read_equipment_info(self, call):
        self.progress.insert(self.text_index, "6. read_equipment_info")
        self.text_index = self.text_index + 1.0
        call.equipment_info_search()

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터의 장비 정보 데이터 프레임으로 정리
    def sort_equipment_info(self, call):
        self.progress.insert(self.text_index, "7. sort_equipment_info")
        self.text_index = self.text_index + 1.0
        call.equipment_info_register()

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터 장비 이미지 저장
    def save_equipment_image(self, call, name):
        self.progress.insert(self.text_index, "8. save_equipment_image")
        self.text_index = self.text_index + 1.0
        call.parse_maple_equipment_image(name)

        self.progress.insert(self.text_index, " ...	ok \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()

    # 캐릭터 프로필 정보 엑셀 파일로 저장
    def create_character_profile(self, call, name):
        self.progress.insert(self.text_index, "9. create_character_profile")
        self.text_index = self.text_index + 1.0
        call.save_excel(name)

        self.progress.insert(self.text_index, " ... 	ok\n")
        self.text_index = self.text_index + 1.0

        self.progress.insert(self.text_index, "--- fin --- \n")
        self.text_index = self.text_index + 1.0
        self.main_window.update()


class COW:
    def __init__(self, character_name):
        self.stat_damage = 0
        self.stat = 0
        self.stat_percent = 0
        self.attack_magic = 0
        self.attack_magic_percent = 0

        self.sub_window = tkinter.Tk()
        self.sub_window.title("Maple_COW_RANK")
        self.sub_window.geometry("1096x782")
        self.sub_window.resizable(False, False)

        self.cow_rank = ["0", "0", "0", "0", "0", "0", "0", "0"]
        self.cow_rank = [0, 0, 0, 0, 0, 0, 0, 0]
        self._rank = [0, 0, 0, 0, 0, 0, 0, 0,
                      0, 0, 0, 0, 0, 0, 0, 0,
                      0, 0, 0, 0, 0, 0, 0, 0]

        # data save to excel
        base_dir = "./"
        file_name = character_name + "/" + character_name + " profile.xlsx"
        file_dir = os.path.join(base_dir, file_name)
        self.read_equipment = pd.read_excel(file_dir, sheet_name='Sheet1')

        self.calc_grade()
        self.draw_ui(character_name)

    # 캐릭터 직업분류
    def calc_grade(self):
        job_number = 0  # 1전사 2궁수 3마법사 4도적 5기타

        job = [
            "마법사", "플레임위자드", "배틀메이지", "에반", "루미너스", "키네시스", "일리움",
            "전사", "바이퍼", "캐논마스터", "미하일", "소울마스터", "스트라이커", "블래스터", "데몬슬레이어", "아란", "은월", "카이저", "아델", "아크", "제로",
            "궁수", "캡틴", "윈드브레이커", "와일드헌터", "메카닉", "메르세데스", "엔젤릭버스터",
            "도적", "나이트워커", "팬텀", "카데나", "호영",
            "제논", "데몬어벤저"]

        for x in range(0, len(job)):
            if self.read_equipment["-.1"][2].find(job[x]) >= 0:
                if 7 > x:
                    job_number = 3
                    break
                elif 7 <= x < 21:
                    job_number = 1
                    break
                elif 21 <= x < 28:
                    job_number = 2
                    break
                elif 28 <= x < 33:
                    job_number = 4
                    break
                elif x >= 33:
                    job_number = 5
                    break

        self.calc_grade_function(job_number)

    # 캐릭터 직업에 따른 아이템 옵션 분류
    def calc_grade_function(self, job):
        _len = ["-.7", "-.8", "-.9", "-.10", "-.11", "-.12",
                "-.13", "-.14", "-.15", "-.16", "-.17", "-.18",
                "-.19", "-.20", "-.21", "-.22", "-.23", "-.24",
                "-.25", "-.26", "-.27"]

        index = [16, 20, 32, 36, 37,
                 17, 20, 33, 36, 37,
                 18, 20, 34, 36, 38,
                 19, 20, 35, 36, 37]

        for x in range(0, len(_len)):
            star_cal = 0
            potential_cal = 0
            sp_cal = 0
            am_cal = 0

            # 별 개수
            try:
                star_cal = int(self.read_equipment[_len[x]][42])
            except TypeError:
                pass
            except ValueError:
                pass

            # 잠재 스탯
            try:
                print("cal :", potential_cal, job, 5 * job, 5 * job + 0, index[(5 * job) + 0], self.read_equipment[_len[x]][index[(5 * job) + 0]])
                potential_cal = potential_cal + int(self.read_equipment[_len[x]][index[(5 * (job - 1)) + 0]])
            except TypeError:
                pass
            except ValueError:
                pass

            # 잠재 올스탯
            try:
                print("cal :", potential_cal, index[(5 * job) + 1], self.read_equipment[_len[x]][index[(5 * job) + 1]])
                potential_cal = potential_cal + int(self.read_equipment[_len[x]][index[(5 * (job - 1)) + 1]])
            except TypeError:
                pass
            except ValueError:
                pass

            # 에디 스탯
            try:
                sp_cal = sp_cal + int(self.read_equipment[_len[x]][index[(5 * (job - 1)) + 2]])
            except TypeError:
                pass
            except ValueError:
                pass

            # 에디 올스탯
            try:
                sp_cal = sp_cal + int(self.read_equipment[_len[x]][index[(5 * (job - 1)) + 3]])
            except TypeError:
                pass
            except ValueError:
                pass

            # 에디 공마
            try:
                am_cal = am_cal + int(self.read_equipment[_len[x]][index[(5 * (job - 1)) + 4]])
            except TypeError:
                pass
            except ValueError:
                pass

            self.calc_star_rank(x, star_cal)
            self.calc_potential_rank(x, self.read_equipment[_len[x]][43])
            self.calc_editional_potential_rank(x, self.read_equipment[_len[x]][44])
            self.calc_potential_percent(x, potential_cal)
            self.calc_editional_potential_percent(x, float(sp_cal), float(am_cal))
            print("\n[rank : ", self._rank[x], "], ", self.read_equipment[_len[x]][1], "\n--------\n")

        self.calc_weapon_option(job)
        self.calc_cow_rank()

    # 캐릭터 아이템 옵션에 따른 등급 분류
    def calc_cow_rank(self):
        self._rank[11] = self._rank[11] + 24
        _len = [0, 1, 2, 3, 4, 5, 6,
                7, 8, 9, 10, 11, 12,
                13, 14, 15, 16, 17,
                18, 19, 20, 21, 22]

        emblem = 23

        for x in range(0, len(_len) - 2):
            # N/A
            if 0 <= self._rank[x] <= 12:
                print("n/a", self._rank[x])
                self.cow_rank[7] = int(self.cow_rank[7]) + 1
            # 1C
            elif 13 <= self._rank[x] <= 18:
                print("1c", self._rank[x])
                self.cow_rank[6] = int(self.cow_rank[6]) + 1
            # 1B
            elif 19 <= self._rank[x] <= 27:
                print("1b", self._rank[x])
                self.cow_rank[5] = int(self.cow_rank[5]) + 1
            # 1+B
            elif 28 <= self._rank[x] <= 45:
                print("1+b", self._rank[x])
                self.cow_rank[4] = int(self.cow_rank[4]) + 1
            # 1++B
            elif 46 <= self._rank[x] <= 56:
                print("1++b", self._rank[x])
                self.cow_rank[3] = int(self.cow_rank[3]) + 1
            # 1A
            elif 57 <= self._rank[x] <= 85:
                print("1a", self._rank[x])
                self.cow_rank[2] = int(self.cow_rank[2]) + 1
            # 1+A
            elif 86 <= self._rank[x] <= 135:
                print("1+a", self._rank[x])
                self.cow_rank[1] = int(self.cow_rank[1]) + 1
            # 1++A
            elif 136 <= self._rank[x]:
                print("1++a", self._rank[x])
                self.cow_rank[0] = int(self.cow_rank[0]) + 1

        for x in range(len(_len) - 2, len(_len)):
            # N/A
            if 0 <= self._rank[x] <= 10:
                print("n/a", self._rank[x])
                self.cow_rank[7] = int(self.cow_rank[7]) + 1
            # 1C
            elif 11 <= self._rank[x] <= 17:
                print("1c", self._rank[x])
                self.cow_rank[6] = int(self.cow_rank[6]) + 1
            # 1B
            elif 18 <= self._rank[x] <= 30:
                print("1b", self._rank[x])
                self.cow_rank[5] = int(self.cow_rank[5]) + 1
            # 1+B
            elif 31 <= self._rank[x] <= 45:
                print("1+b", self._rank[x])
                self.cow_rank[4] = int(self.cow_rank[4]) + 1
            # 1++B
            elif 46 <= self._rank[x] <= 54:
                print("1++b", self._rank[x])
                self.cow_rank[3] = int(self.cow_rank[3]) + 1
            # 1A
            elif 55 <= self._rank[x] <= 93:
                print("1a", self._rank[x])
                self.cow_rank[2] = int(self.cow_rank[2]) + 1
            # 1+A
            elif 94 <= self._rank[x] <= 138:
                print("1+a", self._rank[x])
                self.cow_rank[1] = int(self.cow_rank[1]) + 1
            # 1++A
            elif 139 <= self._rank[x]:
                print("1++a", self._rank[x])
                self.cow_rank[0] = int(self.cow_rank[0]) + 1

        # N/A
        if 0 <= self._rank[emblem] <= 6:
            print("n/a", self._rank[emblem])
            self.cow_rank[7] = int(self.cow_rank[7]) + 1
        # 1C
        elif 7 <= self._rank[emblem] <= 11:
            print("1c", self._rank[emblem])
            self.cow_rank[6] = int(self.cow_rank[6]) + 1
        # 1B
        elif 12 <= self._rank[emblem] <= 24:
            print("1b", self._rank[emblem])
            self.cow_rank[5] = int(self.cow_rank[5]) + 1
        # 1+B
        elif 25 <= self._rank[emblem] <= 31:
            print("1+b", self._rank[emblem])
            self.cow_rank[4] = int(self.cow_rank[4]) + 1
        # 1++B
        elif 32 <= self._rank[emblem] <= 37:
            print("1++b", self._rank[emblem])
            self.cow_rank[3] = int(self.cow_rank[3]) + 1
        # 1A
        elif 38 <= self._rank[emblem] <= 69:
            print("1a", self._rank[emblem])
            self.cow_rank[2] = int(self.cow_rank[2]) + 1
        # 1+A
        elif 70 <= self._rank[emblem] <= 103:
            print("1+a", self._rank[emblem])
            self.cow_rank[1] = int(self.cow_rank[1]) + 1
        # 1++A
        elif 104 <= self._rank[emblem]:
            print("1++a", self._rank[emblem])
            self.cow_rank[0] = int(self.cow_rank[0]) + 1

    # 캐릭터 등급표 출력
    def draw_ui(self, name):
        background = tk.PhotoImage(file="./img/bg2.gif", master=self.sub_window)
        bg_label = tk.Label(self.sub_window, image=background)
        bg_label.place(x=0, y=0)

        font = "Colossus", 12, "bold"
        font2 = "Colossus ", 9, "bold"

        _character_image = tk.PhotoImage(file="./" + name + "/img/" + name + ".png", master=self.sub_window)
        character_image = tk.Label(self.sub_window, image=_character_image, width=167, height=167)
        character_image.place(x=110, y=110)

        _hat = self.read_equipment["-.7"][42] + "성\n" + self.read_equipment["-.7"][43] + "\n" + self.read_equipment["-.7"][44]
        hat = tk.Label(self.sub_window, text=_hat, font=font2, background="white", anchor="center")
        hat.place(x=513, y=112)
        _face = self.read_equipment["-.14"][42] + "성\n" + self.read_equipment["-.14"][43] + "\n" + self.read_equipment["-.14"][44]
        face = tk.Label(self.sub_window, text=_face, font=font2, background="white")
        face.place(x=513, y=180)
        _eyes = self.read_equipment["-.15"][42] + "성\n" + self.read_equipment["-.15"][43] + "\n" + self.read_equipment["-.15"][44]
        eyes = tk.Label(self.sub_window, text=_eyes, font=font2, background="white")
        eyes.place(x=513, y=249)
        _shirt = self.read_equipment["-.8"][42] + "성\n" + self.read_equipment["-.8"][43] + "\n" + self.read_equipment["-.8"][44]
        shirt = tk.Label(self.sub_window, text=_shirt, font=font2, background="white")
        shirt.place(x=513, y=318)
        _pants = self.read_equipment["-.9"][42] + "성\n" + self.read_equipment["-.9"][43] + "\n" + self.read_equipment["-.9"][44]
        pants = tk.Label(self.sub_window, text=_pants, font=font2, background="white")
        pants.place(x=513, y=389)
        _shoes = self.read_equipment["-.10"][42] + "성\n" + self.read_equipment["-.10"][43] + "\n" + self.read_equipment["-.10"][44]
        shoes = tk.Label(self.sub_window, text=_shoes, font=font2, background="white")
        shoes.place(x=513, y=460)

        _ear = self.read_equipment["-.16"][42] + "성\n" + self.read_equipment["-.16"][43] + "\n" + self.read_equipment["-.16"][44]
        ear = tk.Label(self.sub_window, text=_ear, font=font2, background="white")
        ear.place(x=602, y=249)
        _shoulder = self.read_equipment["-.17"][42] + "성\n" + self.read_equipment["-.17"][43] + "\n" + self.read_equipment["-.17"][44]
        shoulder = tk.Label(self.sub_window, text=_shoulder, font=font2, background="white")
        shoulder.place(x=602, y=318)
        _gloves = self.read_equipment["-.11"][42] + "성\n" + self.read_equipment["-.11"][43] + "\n" + self.read_equipment["-.11"][44]
        gloves = tk.Label(self.sub_window, text=_gloves, font=font2, background="white")
        gloves.place(x=602, y=389)

        _pendant1 = self.read_equipment["-.19"][42] + "성\n" + self.read_equipment["-.19"][43] + "\n" + self.read_equipment["-.19"][44]
        pendant1 = tk.Label(self.sub_window, text=_pendant1, font=font2, background="white")
        pendant1.place(x=428, y=180)
        _pendant2 = self.read_equipment["-.20"][42] + "성\n" + self.read_equipment["-.20"][43] + "\n" + self.read_equipment["-.20"][44]
        pendant2 = tk.Label(self.sub_window, text=_pendant2, font=font2, background="white")
        pendant2.place(x=428, y=249)
        _weapon = self.read_equipment["-.28"][42] + "성\n" + self.read_equipment["-.28"][43] + "\n" + self.read_equipment["-.28"][44]
        weapon = tk.Label(self.sub_window, text=_weapon, font=font2, background="white")
        weapon.place(x=428, y=318)
        _belt = self.read_equipment["-.13"][42] + "성\n" + self.read_equipment["-.13"][43] + "\n" + self.read_equipment["-.13"][44]
        belt = tk.Label(self.sub_window, text=_belt, font=font2, background="white")
        belt.place(x=428, y=389)

        _ring1 = self.read_equipment["-.21"][42] + "성\n" + self.read_equipment["-.21"][43] + "\n" + self.read_equipment["-.21"][44]
        ring1 = tk.Label(self.sub_window, text=_ring1, font=font2, background="white")
        ring1.place(x=340, y=112)
        _ring2 = self.read_equipment["-.22"][42] + "성\n" + self.read_equipment["-.22"][43] + "\n" + self.read_equipment["-.22"][44]
        ring2 = tk.Label(self.sub_window, text=_ring2, font=font2, background="white")
        ring2.place(x=340, y=180)
        _ring3 = self.read_equipment["-.23"][42] + "성\n" + self.read_equipment["-.23"][43] + "\n" + self.read_equipment["-.23"][44]
        ring3 = tk.Label(self.sub_window, text=_ring3, font=font2, background="white")
        ring3.place(x=340, y=249)
        _ring4 = self.read_equipment["-.24"][42] + "성\n" + self.read_equipment["-.24"][43] + "\n" + self.read_equipment["-.24"][44]
        ring4 = tk.Label(self.sub_window, text=_ring4, font=font2, background="white")
        ring4.place(x=340, y=318)
        _pocket = self.read_equipment["-.25"][42] + "성\n" + self.read_equipment["-.25"][43] + "\n" + self.read_equipment["-.25"][44]
        pocket = tk.Label(self.sub_window, text=_pocket, font=font2, background="white")
        pocket.place(x=340, y=389)

        _emblem = self.read_equipment["-.30"][42] + "성\n" + self.read_equipment["-.30"][43] + "\n" + self.read_equipment["-.30"][44]
        emblem = tk.Label(self.sub_window, text=_emblem, font=font2, background="white")
        emblem.place(x=690, y=112)
        _badge = self.read_equipment["-.27"][42] + "성\n" + self.read_equipment["-.27"][43] + "\n" + self.read_equipment["-.27"][44]
        badge = tk.Label(self.sub_window, text=_badge, font=font2, background="white")
        badge.place(x=690, y=180)
        _medal = self.read_equipment["-.26"][42] + "성\n" + self.read_equipment["-.26"][43] + "\n" + self.read_equipment["-.26"][44]
        medal = tk.Label(self.sub_window, text=_medal, font=font2, background="white")
        medal.place(x=690, y=249)
        _sub_weapon = self.read_equipment["-.29"][42] + "성\n" + self.read_equipment["-.29"][43] + "\n" + self.read_equipment["-.29"][44]
        sub_weapon = tk.Label(self.sub_window, text=_sub_weapon, font=font2, background="white")
        sub_weapon.place(x=690, y=318)
        _cape = self.read_equipment["-.12"][42] + "성\n" + self.read_equipment["-.12"][43] + "\n" + self.read_equipment["-.12"][44]
        cape = tk.Label(self.sub_window, text=_cape, font=font2, background="white")
        cape.place(x=690, y=389)
        _heart = self.read_equipment["-.18"][42] + "성\n" + self.read_equipment["-.18"][43] + "\n" + self.read_equipment["-.18"][44]
        heart = tk.Label(self.sub_window, text=_heart, font=font2, background="white")
        heart.place(x=690, y=460)

        _one_plus_plusA = tk.Label(self.sub_window, text="1++A", width=5, height=1, font=font)
        _one_plus_plusA.place(x=830, y=118)
        _one_plus_plusA_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[0]) + "개", width=5, height=1, font=font, anchor="w")
        _one_plus_plusA_print.place(x=890, y=118)
        _one_plusA = tk.Label(self.sub_window, text="1+A", width=5, height=1, font=font)
        _one_plusA.place(x=830, y=148)
        _one_plusA_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[1]) + "개", width=5, height=1, font=font, anchor="w")
        _one_plusA_print.place(x=890, y=148)
        _oneA = tk.Label(self.sub_window, text="1A", width=5, height=1, font=font)
        _oneA.place(x=830, y=178)
        _oneA_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[2]) + "개", width=5, height=1, font=font, anchor="w")
        _oneA_print.place(x=890, y=178)

        _one_plus_plusB = tk.Label(self.sub_window, text="1++B", width=5, height=1, font=font)
        _one_plus_plusB.place(x=830, y=218)
        _one_plus_plusB_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[3]) + "개", width=5, height=1, font=font, anchor="w")
        _one_plus_plusB_print.place(x=890, y=218)
        _one_plusB = tk.Label(self.sub_window, text="1+B", width=5, height=1, font=font)
        _one_plusB.place(x=830, y=248)
        _one_plusB_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[4]) + "개", width=5, height=1, font=font, anchor="w")
        _one_plusB_print.place(x=890, y=248)
        _oneB = tk.Label(self.sub_window, text="1B", width=5, height=1, font=font)
        _oneB.place(x=830, y=278)
        _oneB_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[5]) + "개", width=5, height=1, font=font, anchor="w")
        _oneB_print.place(x=890, y=278)

        _one_C = tk.Label(self.sub_window, text="1C", width=5, height=1, font=font)
        _one_C.place(x=830, y=318)
        _one_C_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[6]) + "개", width=5, height=1, font=font, anchor="w")
        _one_C_print.place(x=890, y=318)
        _NA = tk.Label(self.sub_window, text="N/A", width=5, height=1, font=font)
        _NA.place(x=830, y=348)
        _NA_print = tk.Label(self.sub_window, text=": " + str(self.cow_rank[7]) + "개", width=5, height=1, font=font, anchor="w")
        _NA_print.place(x=890, y=348)

        self.sub_window.mainloop()

    # 별 개수
    def calc_star_rank(self, x, target):
        print("star", x, "|rank:", self._rank[x], target)

        if target == 0:
            self._rank[x] = self._rank[x] + 0
        elif 1 <= target <= 2:
            self._rank[x] = self._rank[x] + 1
        elif 3 <= target <= 5:
            self._rank[x] = self._rank[x] + 2
        elif 6 <= target <= 8:
            self._rank[x] = self._rank[x] + 3
        elif 9 <= target <= 10:
            self._rank[x] = self._rank[x] + 4
        elif target == 11:
            self._rank[x] = self._rank[x] + 5
        elif target == 12:
            self._rank[x] = self._rank[x] + 6
        elif target == 13:
            self._rank[x] = self._rank[x] + 7
        elif target == 14:
            self._rank[x] = self._rank[x] + 8
        elif target == 15:
            self._rank[x] = self._rank[x] + 10
        elif target == 16:
            self._rank[x] = self._rank[x] + 12
        elif target == 17:
            self._rank[x] = self._rank[x] + 14
        elif target == 18:
            self._rank[x] = self._rank[x] + 17
        elif target == 19:
            self._rank[x] = self._rank[x] + 20
        elif target == 20:
            self._rank[x] = self._rank[x] + 24
        elif target == 21:
            self._rank[x] = self._rank[x] + 29
        elif target == 22:
            self._rank[x] = self._rank[x] + 35
        elif target == 23:
            self._rank[x] = self._rank[x] + 50
        elif target == 24:
            self._rank[x] = self._rank[x] + 70
        elif target == 25:
            self._rank[x] = self._rank[x] + 100

    # 잠재등급
    def calc_potential_rank(self, x, target):
        print("potential", x, "|rank:", self._rank[x], target)
        if target == "레어":
            self._rank[x] = self._rank[x] + 0
        if target == "에픽":
            self._rank[x] = self._rank[x] + 3
        elif target == "유니크":
            self._rank[x] = self._rank[x] + 10
        elif target == "레전":
            self._rank[x] = self._rank[x] + 20

    # 에디잠재등급
    def calc_editional_potential_rank(self, x, target):
        print("editional_potential", x, "|rank:", self._rank[x], target)
        if target == "레어":
            self._rank[x] = self._rank[x] + 3
        if target == "에픽":
            self._rank[x] = self._rank[x] + 7
        elif target == "유니크":
            self._rank[x] = self._rank[x] + 15
        elif target == "레전":
            self._rank[x] = self._rank[x] + 30

    # 잠재수치
    def calc_potential_percent(self, x, target):
        print("potential_percent", x, "| rank :", self._rank[x], "%:", target)

        if target == 0:
            self._rank[x] = self._rank[x] + 0
        elif 1 <= target <= 3:
            self._rank[x] = self._rank[x] + 1
        elif 4 <= target <= 6:
            self._rank[x] = self._rank[x] + 2
        elif 7 <= target <= 9:
            self._rank[x] = self._rank[x] + 3
        elif 10 <= target <= 12:
            self._rank[x] = self._rank[x] + 5
        elif 13 <= target <= 15:
            self._rank[x] = self._rank[x] + 8
        elif 16 <= target <= 18:
            self._rank[x] = self._rank[x] + 12
        elif 19 <= target <= 21:
            self._rank[x] = self._rank[x] + 16
        elif 22 <= target <= 24:
            self._rank[x] = self._rank[x] + 21
        elif 25 <= target <= 27:
            self._rank[x] = self._rank[x] + 27
        elif 28 <= target <= 30:
            self._rank[x] = self._rank[x] + 35
        elif 31 <= target <= 33:
            self._rank[x] = self._rank[x] + 45
        elif 34 <= target <= 36:
            self._rank[x] = self._rank[x] + 60

    # 에디잠재수치
    def calc_editional_potential_percent(self, x, target1, target2):
        target = target1 + (target2 * 0.4)
        print("editional_potential_percent", x, "| rank :", self._rank[x], "sp %:", target1, "am %:", target2, "total %:", target)

        if 0.0 <= target < 1.0:
            self._rank[x] = self._rank[x] + 0
        elif 1.0 <= target < 3.0:
            self._rank[x] = self._rank[x] + 1
        elif 3.0 <= target < 4.0:
            self._rank[x] = self._rank[x] + 2
        elif 4.0 <= target < 5.0:
            self._rank[x] = self._rank[x] + 3
        elif 5.0 <= target < 6.0:
            self._rank[x] = self._rank[x] + 4
        elif 6.0 <= target < 7.0:
            self._rank[x] = self._rank[x] + 6
        elif 7.0 <= target < 8.0:
            self._rank[x] = self._rank[x] + 8
        elif 8.0 <= target < 9.0:
            self._rank[x] = self._rank[x] + 10
        elif 9.0 <= target < 10.0:
            self._rank[x] = self._rank[x] + 13
        elif 10.0 <= target < 11.0:
            self._rank[x] = self._rank[x] + 16
        elif 11.0 <= target < 12.0:
            self._rank[x] = self._rank[x] + 19
        elif 12.0 <= target < 13.0:
            self._rank[x] = self._rank[x] + 23
        elif 13.0 <= target < 14.0:
            self._rank[x] = self._rank[x] + 27
        elif 14.0 <= target < 15.0:
            self._rank[x] = self._rank[x] + 31
        elif 15.0 <= target < 16.0:
            self._rank[x] = self._rank[x] + 35
        elif 16.0 <= target < 17.0:
            self._rank[x] = self._rank[x] + 40
        elif 17.0 <= target < 18.0:
            self._rank[x] = self._rank[x] + 45
        elif 18.0 <= target < 19.0:
            self._rank[x] = self._rank[x] + 50
        elif 19.0 <= target < 20.0:
            self._rank[x] = self._rank[x] + 56
        elif 20.0 <= target < 21.0:
            self._rank[x] = self._rank[x] + 62
        elif 21.0 <= target < 22.0:
            self._rank[x] = self._rank[x] + 68
        elif 22.0 <= target < 23.0:
            self._rank[x] = self._rank[x] + 75
        elif 23.0 <= target < 24.0:
            self._rank[x] = self._rank[x] + 83
        elif 24.0 <= target:
            self._rank[x] = self._rank[x] + 92

    # 무기류 수치 공/마% 보공% 방무%
    def calc_weapon_option(self, job):
        _len = ["-.7", "-.8", "-.9", "-.10", "-.11", "-.12",
                "-.13", "-.14", "-.15", "-.16", "-.17", "-.18",
                "-.19", "-.20", "-.21", "-.22", "-.23", "-.24",
                "-.25", "-.26", "-.27", "-.28", "-.29", "-.30"]

        for x in range(len(_len) - 3, len(_len)):
            star_cal = 0

            # 별 개수
            try:
                star_cal = int(self.read_equipment[_len[x]][42])
            except TypeError:
                pass
            except ValueError:
                pass

            self.calc_star_rank(x, star_cal)
            self.calc_potential_rank(x, self.read_equipment[_len[x]][43])
            self.calc_editional_potential_rank(x, self.read_equipment[_len[x]][44])
            print("\n[rank : ", self._rank[x], "], ", self.read_equipment[_len[x]][1], "\n--------\n")

        # 한줄당 점수매겨야함
        # 21 무기 22 보조 23 엠블
        weapon_index = ["-.1", "-.2", "-.3"]
        target_index = [17, 18, 19, 20, 21, 22, 23, 24]
        print("before", self._rank)

        for x in range(0, len(weapon_index)):
            for y in range(0, len(target_index)):
                try:
                    target = self.read_equipment[weapon_index[x]][target_index[y]]

                    if not job == 3:
                        if target.find("공격력") >= 0:
                            _target = int(target[target.find("+") + 1:target.find("%")])
                            print("공격", _target)
                            if 0 <= _target <= 3:
                                self._rank[21 + x] = self._rank[21 + x] + 1
                            elif 4 <= _target <= 6:
                                self._rank[21 + x] = self._rank[21 + x] + 7
                            elif 7 <= _target <= 9:
                                self._rank[21 + x] = self._rank[21 + x] + 13
                            elif 10 <= _target <= 12:
                                self._rank[21 + x] = self._rank[21 + x] + 20
                    else:
                        if target.find("마력") >= 0:
                            _target = int(target[target.find("+") + 1:target.find("%")])
                            print("마력", _target)
                            if 0 <= _target <= 3:
                                self._rank[21 + x] = self._rank[21 + x] + 1
                            elif 4 <= _target <= 6:
                                self._rank[21 + x] = self._rank[21 + x] + 7
                            elif 7 <= _target <= 9:
                                self._rank[21 + x] = self._rank[21 + x] + 13
                            elif 10 <= _target <= 12:
                                self._rank[21 + x] = self._rank[21 + x] + 20

                    if target.find("보스") >= 0:
                        _target = int(target[target.find("+") + 1:target.find("%")])
                        print("보스", _target)
                        if _target == 20:
                            self._rank[21 + x] = self._rank[21 + x] + 5
                        elif _target == 30:
                            self._rank[21 + x] = self._rank[21 + x] + 10
                        elif _target == 35:
                            self._rank[21 + x] = self._rank[21 + x] + 15
                        elif _target == 40:
                            self._rank[21 + x] = self._rank[21 + x] + 20

                    if target.find("방어율") >= 0:
                        _target = int(target[target.find("+") + 1:target.find("%")])
                        print("방어", _target)
                        if _target == 15:
                            self._rank[21 + x] = self._rank[21 + x] + 3
                        elif _target == 30:
                            self._rank[21 + x] = self._rank[21 + x] + 10
                        elif _target == 35:
                            self._rank[21 + x] = self._rank[21 + x] + 15
                        elif _target == 40:
                            self._rank[21 + x] = self._rank[21 + x] + 20

                except TypeError:
                    pass
                except ValueError:
                    pass

        print("after", self._rank)


class STAT:
    def __init__(self, name):
        self.stat_damage = 0
        self.stat_sum = 0
        self.stat_percent = 0
        self.attack_magic_sum = 0
        self.attack_magic_percent = 0
        self.calculate_stat(name)

        self.sub_window = tkinter.Tk()
        self.sub_window.title("Maple_STAT_TEST")
        self.sub_window.geometry("1096x782")
        self.sub_window.resizable(False, False)

        background = tk.PhotoImage(file="./img/bg2.gif", master=self.sub_window)
        bg_label = tk.Label(self.sub_window, image=background)
        bg_label.place(x=0, y=0)

        font = "Colossus", 13, "bold"
        path = "./" + name + "/img/"

        _character_image = tk.PhotoImage(file=path + name + ".png", master=self.sub_window)
        character_image = tk.Label(self.sub_window, image=_character_image, width=167, height=167)
        character_image.place(x=110, y=110)

        _hat = tk.PhotoImage(file=path + "0.gif", master=self.sub_window)
        hat = tk.Label(self.sub_window, image=_hat)
        hat.place(x=518, y=122)
        _face = tk.PhotoImage(file=path + "7.gif", master=self.sub_window)
        face = tk.Label(self.sub_window, image=_face)
        face.place(x=525, y=190)
        _eyes = tk.PhotoImage(file=path + "8.gif", master=self.sub_window)
        eyes = tk.Label(self.sub_window, image=_eyes)
        eyes.place(x=525, y=258)
        _shirt = tk.PhotoImage(file=path + "1.gif", master=self.sub_window)
        shirt = tk.Label(self.sub_window, image=_shirt)
        shirt.place(x=522, y=328)
        _pants = tk.PhotoImage(file=path + "2.gif", master=self.sub_window)
        pants = tk.Label(self.sub_window, image=_pants)
        pants.place(x=521, y=398)
        _shoes = tk.PhotoImage(file=path + "3.gif", master=self.sub_window)
        shoes = tk.Label(self.sub_window, image=_shoes)
        shoes.place(x=521, y=468)

        _ear = tk.PhotoImage(file=path + "9.gif", master=self.sub_window)
        ear = tk.Label(self.sub_window, image=_ear)
        ear.place(x=608, y=258)
        _shoulder = tk.PhotoImage(file=path + "10.gif", master=self.sub_window)
        shoulder = tk.Label(self.sub_window, image=_shoulder)
        shoulder.place(x=608, y=327)
        _gloves = tk.PhotoImage(file=path + "4.gif", master=self.sub_window)
        gloves = tk.Label(self.sub_window, image=_gloves)
        gloves.place(x=610, y=401)

        _pendant1 = tk.PhotoImage(file=path + "12.gif", master=self.sub_window)
        pendant1 = tk.Label(self.sub_window, image=_pendant1)
        pendant1.place(x=436, y=190)
        _pendant2 = tk.PhotoImage(file=path + "13.gif", master=self.sub_window)
        pendant2 = tk.Label(self.sub_window, image=_pendant2)
        pendant2.place(x=436, y=258)
        _weapon = tk.PhotoImage(file=path + "21.gif", master=self.sub_window)
        weapon = tk.Label(self.sub_window, image=_weapon)
        weapon.place(x=436, y=328)
        _belt = tk.PhotoImage(file=path + "6.gif", master=self.sub_window)
        belt = tk.Label(self.sub_window, image=_belt)
        belt.place(x=436, y=398)

        _ring1 = tk.PhotoImage(file=path + "14.gif", master=self.sub_window)
        ring1 = tk.Label(self.sub_window, image=_ring1)
        ring1.place(x=344, y=118)
        _ring2 = tk.PhotoImage(file=path + "15.gif", master=self.sub_window)
        ring2 = tk.Label(self.sub_window, image=_ring2)
        ring2.place(x=344, y=187)
        _ring3 = tk.PhotoImage(file=path + "16.gif", master=self.sub_window)
        ring3 = tk.Label(self.sub_window, image=_ring3)
        ring3.place(x=344, y=252)
        _ring4 = tk.PhotoImage(file=path + "17.gif", master=self.sub_window)
        ring4 = tk.Label(self.sub_window, image=_ring4)
        ring4.place(x=348, y=325)
        _pocket = tk.PhotoImage(file=path + "18.gif", master=self.sub_window)
        pocket = tk.Label(self.sub_window, image=_pocket)
        pocket.place(x=348, y=396)

        _emblem = tk.PhotoImage(file=path + "23.gif", master=self.sub_window)
        emblem = tk.Label(self.sub_window, image=_emblem)
        emblem.place(x=690, y=118)
        _badge = tk.PhotoImage(file=path + "20.gif", master=self.sub_window)
        badge = tk.Label(self.sub_window, image=_badge)
        badge.place(x=690, y=187)
        _medal = tk.PhotoImage(file=path + "19.gif", master=self.sub_window)
        medal = tk.Label(self.sub_window, image=_medal)
        medal.place(x=690, y=252)
        _sub_weapon = tk.PhotoImage(file=path + "22.gif", master=self.sub_window)
        sub_weapon = tk.Label(self.sub_window, image=_sub_weapon)
        sub_weapon.place(x=690, y=325)
        _cape = tk.PhotoImage(file=path + "5.gif", master=self.sub_window)
        cape = tk.Label(self.sub_window, image=_cape)
        cape.place(x=690, y=396)
        _heart = tk.PhotoImage(file=path + "11.gif", master=self.sub_window)
        heart = tk.Label(self.sub_window, image=_heart)
        heart.place(x=690, y=469)

        _stat_damage_image = tk.PhotoImage(file="./img/_스공.png", master=self.sub_window)
        _stat_damage = tk.Label(self.sub_window, image=_stat_damage_image)
        _stat_damage.place(x=780, y=118)
        stat_damage_print = tk.Label(self.sub_window, text=self.stat_damage, width=18, height=1, font=font)
        stat_damage_print.place(x=840, y=118)
        _stat_sum_image = tk.PhotoImage(file="./img/_스탯합.png", master=self.sub_window)
        _stat_sum = tk.Label(self.sub_window, image=_stat_sum_image)
        _stat_sum.place(x=780, y=168)
        stat_sum_print = tk.Label(self.sub_window, text=self.stat_sum, width=5, height=1, font=font)
        stat_sum_print.place(x=840, y=168)
        _stat_percent_image = tk.PhotoImage(file="./img/_스탯.png", master=self.sub_window)
        _stat_percent = tk.Label(self.sub_window, image=_stat_percent_image)
        _stat_percent.place(x=910, y=168)
        stat_percent_print = tk.Label(self.sub_window, text=self.stat_percent, width=5, height=1, font=font)
        stat_percent_print.place(x=970, y=168)
        _attack_magic_sum_image = tk.PhotoImage(file="./img/_공마합.png", master=self.sub_window)
        _attack_magic_sum = tk.Label(self.sub_window, image=_attack_magic_sum_image)
        _attack_magic_sum.place(x=780, y=198)
        attack_magic_print = tk.Label(self.sub_window, text=self.attack_magic_sum, width=5, height=1, font=font)
        attack_magic_print.place(x=840, y=198)
        _attack_magic_percent_image = tk.PhotoImage(file="./img/_공마%.png", master=self.sub_window)
        _attack_magic_percent = tk.Label(self.sub_window, image=_attack_magic_percent_image)
        _attack_magic_percent.place(x=910, y=198)
        attack_magic_percent_print = tk.Label(self.sub_window, text=self.attack_magic_percent, width=5, height=1, font=font)
        attack_magic_percent_print.place(x=970, y=198)

        self.sub_window.mainloop()

    # 장비옵션계산
    def calculate_stat(self, name):
        # dataset read
        base_dir = "./"
        file_name = name + "/" + name + " profile.xlsx"
        file_dir = os.path.join(base_dir, file_name)
        read_equipment = pd.read_excel(file_dir, sheet_name='Sheet1')
        read_equipment = read_equipment.fillna("-")

        _len = ["-.7", "-.8", "-.9", "-.10", "-.11", "-.12", "-.13",
                "-.14", "-.15", "-.16", "-.17", "-.18", "-.19", "-.20",
                "-.21", "-.22", "-.23", "-.24", "-.25", "-.26", "-.27",
                "-.28", "-.29", "-.30", "-.31"]

        job = [
            "마법사", "플레임위자드", "배틀메이지", "에반", "루미너스", "키네시스", "일리움",
            "전사", "바이퍼", "캐논마스터", "미하일", "소울마스터", "스트라이커", "블래스터", "데몬슬레이어", "아란", "은월", "카이저", "아델", "아크", "제로",
            "궁수", "캡틴", "윈드브레이커", "와일드헌터", "메카닉", "메르세데스", "엔젤릭버스터",
            "도적", "나이트워커", "팬텀", "카데나", "호영",
            "제논", "데몬어벤저",
        ]

        _str = [2, 11, 15, 27, 31]
        _dex = [3, 12, 15, 28, 31]
        _int = [4, 13, 15, 29, 31]
        _luk = [5, 14, 15, 30, 31]

        _str_p = [6, 16, 20, 32, 36]
        _dex_p = [6, 17, 20, 33, 36]
        _int_p = [6, 18, 20, 34, 36]
        _luk_p = [6, 19, 20, 35, 36]

        _attack = [7, 21, 37]
        _magic = [8, 22, 38]
        _attack_p = [23, 39]
        _magic_p = [24, 40]

        job_number = 0  # 1전사 2궁수 3마법사 4도적 5기타
        self.stat_damage = read_equipment["-.1"][5]

        for x in range(0, len(job)):
            if read_equipment["-.1"][2].find(job[x]) >= 0:
                if 7 > x:
                    job_number = 3
                    break

                elif 7 <= x & x < 21:
                    job_number = 1
                    break

                elif 21 <= x & x < 28:
                    job_number = 2
                    break

                elif 28 <= x & x < 33:
                    job_number = 4
                    break

                elif x >= 33:
                    job_number = 5
                    break

        if job_number == 1:  # str
            for x in range(0, len(_len)):  # 아이템 순수 스탯
                for y in range(0, len(_str)):
                    try:
                        self.stat_sum = self.stat_sum + int(read_equipment[_len[x]][_str[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 스탯% 스탯% + 올스탯%
                for y in range(0, len(_str_p)):
                    try:
                        self.stat_percent = self.stat_percent + int(read_equipment[_len[x]][_str_p[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마
                for y in range(0, len(_attack)):
                    try:
                        self.attack_magic_sum = self.attack_magic_sum + int(read_equipment[_len[x]][_attack[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마%
                for y in range(0, len(_attack_p)):
                    try:
                        self.attack_magic_percent = self.attack_magic_percent + int(read_equipment[_len[x]][_attack_p[y]])
                    except ValueError:
                        pass

        if job_number == 2:  # dex
            for x in range(0, len(_len)):  # 아이템 순수 스탯
                for y in range(0, len(_dex)):
                    try:
                        self.stat_sum = self.stat_sum + int(read_equipment[_len[x]][_dex[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 스탯% 스탯% + 올스탯%
                for y in range(0, len(_dex_p)):
                    try:
                        self.stat_percent = self.stat_percent + int(read_equipment[_len[x]][_dex_p[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마
                for y in range(0, len(_attack)):
                    try:
                        self.attack_magic_sum = self.attack_magic_sum + int(read_equipment[_len[x]][_attack[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마%
                for y in range(0, len(_attack_p)):
                    try:
                        self.attack_magic_percent = self.attack_magic_percent + int(read_equipment[_len[x]][_attack_p[y]])
                    except ValueError:
                        pass

        if job_number == 3:  # int
            for x in range(0, len(_len)):  # 아이템 순수 스탯
                for y in range(0, len(_int)):
                    try:
                        self.stat_sum = self.stat_sum + int(read_equipment[_len[x]][_int[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 스탯% 스탯% + 올스탯%
                for y in range(0, len(_int_p)):
                    try:
                        self.stat_percent = self.stat_percent + int(read_equipment[_len[x]][_int_p[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마
                for y in range(0, len(_magic)):
                    try:
                        self.attack_magic_sum = self.attack_magic_sum + int(read_equipment[_len[x]][_magic[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마%
                for y in range(0, len(_magic_p)):
                    try:
                        self.attack_magic_percent = self.attack_magic_percent + int(read_equipment[_len[x]][_magic_p[y]])
                    except ValueError:
                        pass

        if job_number == 4:  # luk
            for x in range(0, len(_len)):  # 아이템 순수 스탯
                for y in range(0, len(_luk)):
                    try:
                        self.stat_sum = self.stat_sum + int(read_equipment[_len[x]][_luk[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 스탯% 스탯% + 올스탯%
                for y in range(0, len(_luk_p)):
                    try:
                        self.stat_percent = self.stat_percent + int(read_equipment[_len[x]][_luk_p[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마
                for y in range(0, len(_attack)):
                    try:
                        self.attack_magic_sum = self.attack_magic_sum + int(read_equipment[_len[x]][_attack[y]])
                    except ValueError:
                        pass

            for x in range(0, len(_len)):  # 아이템 공마%
                for y in range(0, len(_attack_p)):
                    try:
                        self.attack_magic_percent = self.attack_magic_percent + int(read_equipment[_len[x]][_attack_p[y]])
                    except ValueError:
                        pass


MapleStatUi()
