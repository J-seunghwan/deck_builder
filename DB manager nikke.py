# 블라블라 링크
# 니케 캐릭터 정보 db 화 및 리소스 갖고 오는 용도
# 원래는 gemini-cli한테 정보를 가져오는 일을 시키려고 했는데 cli는 링크 접속이 불가능했음
# gemini 웹 버전은 블라블라(니케도감 링크)에 접속해서 니케 등급에 따른 개수를 조사할 수 있었음.
# 그래서 니케 전용 크롤링 프로그램을 만들기로 함.
# 시작 이유
#- 서브컬쳐 게임들의 공통점이 파티를 구성하기 편하도록 필터링 기능을 제공하지 않음.
#- 그래서 파티 구성 프로그램을 만드는 사람들이 있어서 나도 만드려고 함.
import time
import sqlite3
import os
from pathlib import Path
import re

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests

class DBManager:
    def __init__(self, path_db):
        self.driver = None

        # 현재 파일(__file__)의 절대 경로 확인
        current_file = Path(__file__).resolve()
        # 현재 파일이 위치한 폴더 경로
        current_dir = current_file.parent
        self.path_resource = current_dir / "resource"

        # ==== db init
        self.db_connection = sqlite3.connect(path_db)
        self.db_cursor = self.db_connection.cursor()

        # ==== db nikke_info_table element
        # 니케 이름
        self.db_name = "Name"
        # 니케 프로필 이미지 저장 경로
        self.db_path_img_profile = "Path_img_profile"
        # 니케 등급
        self.db_tier = "Tier"
        # 니케 소속 스쿼드
        self.db_squad = "Squad"
        # 니케 코드(속성)
        self.db_code = "Code"
        # 니케 무기
        self.db_weapon_type = "Weapon_type"
        # 니케 무기 장탄수
        self.db_weapon_magazine_capacity = "Weapon_magazine_capacity"
        # 니케 무기 재장전 시간
        self.db_weapon_reload_time = "Weapon_reload_time"
        # 니케 조작타입
        self.db_weapon_operation_type = "Weapon_operation_type"
        # 무기 설명
        self.db_weapon_description = "Weapon_description"
        # 니케 클래스
        self.db_class = "Class"
        # 니케 소속 기업
        self.db_company = "Company"
        # 스킬 1 이름
        self.db_skill_1_name = "Skill_1_name"
        # 스킬 1 설명
        self.db_skill_1_description = "Skill_1_description"
        # 스킬 2 이름
        self.db_skill_2_name = "Skill_2_name"
        # 스킬 2 설명
        self.db_skill_2_description = "Skill_2_description"
        # 버스트 스킬 이름
        self.db_skill_bust_name = "Skill_bust_name"
        # 버스트 스킬 단계
        self.db_skill_bust_step = "Bust_step"
        # 버스트 스킬 쿨타임
        self.db_skill_bust_time = "Skill_bust_time"
        # 버스트 스킬 설명
        self.db_skill_bust_description = "Skill_bust_description"

    def __del__(self):
        # selenium 정리
        if self.driver:
            self.driver.quit()

        # db 연결 끊기
        self.db_connection.close()

    def createDBTable(self):
        '''
        db 파일에 table이 없다면 table 생성
        '''

        # 니케 이름이 중복일 수 있음. ex) 사쿠라 ssr, sr
        self.db_cursor.execute(f'''CREATE TABLE IF NOT EXISTS nikke_info_table(
                               id INTEGER PRIMARY KEY AUTOINCREMENT,
                               {self.db_name} TEXT,
                               {self.db_path_img_profile} TEXT,
                               {self.db_tier} TEXT,
                               {self.db_squad} TEXT,
                               {self.db_code} TEXT,
                               {self.db_weapon_type} TEXT,
                               {self.db_weapon_magazine_capacity} INTEGER,
                               {self.db_weapon_reload_time} REAL,
                               {self.db_weapon_operation_type} TEXT,
                               {self.db_weapon_description} TEXT,
                               {self.db_class} TEXT,
                               {self.db_company} TEXT,
                               {self.db_skill_1_name} TEXT,
                               {self.db_skill_1_description} TEXT,
                               {self.db_skill_2_name} TEXT,
                               {self.db_skill_2_description} TEXT,
                               {self.db_skill_bust_name} TEXT,
                               {self.db_skill_bust_step} TEXT,
                               {self.db_skill_bust_time} REAL,
                               {self.db_skill_bust_description} TEXT
                               )''')
        self.db_connection.commit()
        
        # 스킬 키워드 사전 테이블
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS skill_keywords_table (
                               id INTEGER PRIMARY KEY AUTOINCREMENT,
                               keyword_name TEXT UNIQUE NOT NULL,
                               category TEXT
                               )''')
        self.db_connection.commit()
        
        # 니케-키워드 관계 테이블 (N:M 관계)
        self.db_cursor.execute('''CREATE TABLE IF NOT EXISTS nikke_keywords_table (
                               nikke_id INTEGER,
                               keyword_id INTEGER,
                               PRIMARY KEY (nikke_id, keyword_id),
                               FOREIGN KEY (nikke_id) REFERENCES nikke_info_table (id) ON DELETE CASCADE,
                               FOREIGN KEY (keyword_id) REFERENCES skill_keywords_table (id) ON DELETE CASCADE
                               )''')
        self.db_connection.commit()

    def accessPage(self):
        '''
        크롤링 시작할 페이지로 접속하는 로직
        '''
        # ==== selenium init
        chrome_options = Options()
        # 파이썬이 끝나도 크롬창이 종료되지 않게 하는 옵션
        # vsocde에서도 적용되게 하기 위해서 마지막 코드에 breakpoint 걸면 됨.
        # 그냥 py 프로그램을 실행할 때는 바로 적용됨
        chrome_options.add_experimental_option("detach", True)
        self.driver = webdriver.Chrome(options=chrome_options)

        # 블라블라 접속
        self.driver.get('https://www.blablalink.com/')
        # 동적 로딩 때문에 로딩에 시간이 필요함
        time.sleep(2)

        # 쿠키 창 닫기. 쿠키 창 닫아야 이후 제대로 클릭할 수 있음
        el_cookie_btn = self.driver.find_element(By.XPATH,'//*[@id="onetrust-close-btn-container"]/button')
        el_cookie_btn.click()

        self.wait(By.XPATH,'//*[@id="layout-content"]/div/div[1]/div[2]/div[1]')
        #time.sleep(1)

        # 한국어로 변경 - 링크를 변경하는 방법이 아니어서 필수임
        # 언어 변경 버튼 클릭
        el_lang_change = self.driver.find_element(By.XPATH, '//*[@id="layout-content"]/div/div[1]/div[2]/div[1]')
        el_lang_change.click()

        #self.wait(By.XPATH, '/html/body/div[5]/div[2]/div[3]/div[1]/div[2]/div[3]')
        time.sleep(1)

        # 한국어 선택
        el_lang_korean = self.driver.find_elements(By.XPATH, "//*[text()='한국어']")[0]
        el_lang_korean.click()

        #self.wait(By.XPATH, '/html/body/div[5]/div[2]/div[5]/div[1]')
        time.sleep(1)

        # 확인 누르기
        el_lang_confirm = self.driver.find_elements(By.XPATH, "//*[text()='Confirm']")[0]
        el_lang_confirm.click()

        #self.wait(By.XPATH, '//*[@id="layout-content"]/div/div[3]/div[3]/div/div[2]/div/div[7]')
        time.sleep(3)

        # 니케 도감 페이지 접속
        el_nikkepedia = self.driver.find_element(By.XPATH, '//*[@id="layout-content"]/div/div[3]/div[3]/div/div[2]/div/div[7]')
        el_nikkepedia.click()
    
    def crawlingAllUpdateDB(self, only_new=True):
        '''
        필수 - accessPage로 크롤링을 시작할 페이지에 접속
        크롤링을 해서 db에 넣을 데이터로 만들기
        only_new = True 일 때 기존 db에서 없는 니케만 가져옴
        '''
        # DB에서 기존 (이름, 티어) 정보 가져오기
        self.db_cursor.execute(f"SELECT {self.db_name}, {self.db_tier} FROM nikke_info_table")
        existing_nikkes = set(self.db_cursor.fetchall())

        # 니케 리스트 가져오기
        self.wait(By.CSS_SELECTOR, "[class='cursor-pointer relative nikkes-all-item h-[180px] max-h-[180px] w-[102px] max-w-[22%] my-[4.5px] mx-[5px]']")
        item_list = self.driver.find_elements(By.CSS_SELECTOR, "[class='cursor-pointer relative nikkes-all-item h-[180px] max-h-[180px] w-[102px] max-w-[22%] my-[4.5px] mx-[5px]']")
        print(f"도감에 등록된 니케 총 {len(item_list)}개")

        time.sleep(2)
        # 니케 프로필 이미지가 모두 로딩되도록 스크롤 제어
        for i in range(33, len(item_list), 24):
            temp = self.driver.find_element(By.XPATH, f'//*[@id="layout-content"]/div/div[2]/div[2]/div/div/div[3]/div[{i}]')
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", temp)
            time.sleep(2)

        # 1단계: 전체 리스트에서 이름과 티어 파악 및 새로운 니케 체크
        nikke_list_info = []
        new_nikke_found = False

        for item in item_list:
            el_span = item.find_element(By.XPATH, "./div[1]/span[1]")
            text_name = el_span.text

            text_tier = ""
            try:
                tier_div = item.find_elements(By.TAG_NAME, "div")
                for div in tier_div:
                    class_content = div.get_attribute("class")
                    if class_content:
                        if "yellow" in class_content: text_tier = "SSR"; break
                        elif "purple" in class_content: text_tier = "SR"; break
                        elif "blue" in class_content: text_tier = "R"; break
            except: pass

            nikke_list_info.append((text_name, text_tier))
            if (text_name, text_tier) not in existing_nikkes:
                new_nikke_found = True

        if not new_nikke_found:
            print("새로운 니케가 없습니다.")
            return

        print("새로운 니케 발견! DB 순서 정렬을 위해 재구축을 시작합니다.")

        # 2단계: 기존 테이블 백업 및 새 테이블 생성
        self.db_cursor.execute("ALTER TABLE nikke_info_table RENAME TO nikke_info_table_legacy")
        self.createDBTable() # 새 nikke_info_table 생성
        self.db_connection.commit()

        # 3단계: 순서대로 데이터 채우기
        for i, (name, tier) in enumerate(nikke_list_info):
            # 기존에 있던 니케인지 확인
            self.db_cursor.execute("SELECT * FROM nikke_info_table_legacy WHERE Name=? AND Tier=?", (name, tier))
            legacy_data = self.db_cursor.fetchone()

            if legacy_data:
                # 기존 데이터가 있으면 복사 (id 제외)
                columns = [
                    self.db_name, self.db_path_img_profile, self.db_tier, self.db_squad, self.db_code,
                    self.db_weapon_type, self.db_weapon_magazine_capacity, self.db_weapon_reload_time,
                    self.db_weapon_operation_type, self.db_weapon_description, self.db_class, self.db_company,
                    self.db_skill_1_name, self.db_skill_1_description, self.db_skill_2_name, 
                    self.db_skill_2_description, self.db_skill_bust_name, self.db_skill_bust_step,
                    self.db_skill_bust_time, self.db_skill_bust_description
                ]
                col_str = ", ".join(columns)
                placeholders = ", ".join(["?" for _ in columns])
                
                # legacy_data[0]은 id이므로 제외
                self.db_cursor.execute(f"INSERT INTO nikke_info_table ({col_str}) VALUES ({placeholders})", legacy_data[1:])
                print(f"[{i+1}/{len(item_list)}] {name} ({tier}) 데이터 복사 완료")
            else:
                # 새로운 니케면 상세 페이지 크롤링
                print(f"[{i+1}/{len(item_list)}] {name} ({tier}) 상세 정보 수집 시작")
                
                # 상세 페이지 접속
                current_item_list = self.driver.find_elements(By.CSS_SELECTOR, "[class='cursor-pointer relative nikkes-all-item h-[180px] max-h-[180px] w-[102px] max-w-[22%] my-[4.5px] mx-[5px]']")
                target_item = current_item_list[i]
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_item)
                time.sleep(1)
                target_item.click()
                time.sleep(3)

                # 상세 정보 파싱
                parsed_data = self._parseNikkeDetail(name)
                
                # DB 삽입
                columns = list(parsed_data.keys())
                col_str = ", ".join(columns)
                placeholders = ", ".join([f":{c}" for c in columns])
                self.db_cursor.execute(f"INSERT INTO nikke_info_table ({col_str}) VALUES ({placeholders})", parsed_data)
                
                # 리스트 페이지로 복귀
                el_prev_page_btn = self.driver.find_element(By.XPATH, '//*[@id="layout-content"]/div/div[1]/div/div[1]')
                el_prev_page_btn.click()
                time.sleep(5)
                print(f"[{i+1}/{len(item_list)}] {name} ({tier}) 상세 정보 수집 및 저장 완료")

            self.db_connection.commit()

        # 4단계: 레거시 테이블 삭제
        self.db_cursor.execute("DROP TABLE nikke_info_table_legacy")
        self.db_connection.commit()
        print("DB 순서 정렬 및 업데이트 완료.")

    def _parseNikkeDetail(self, name):
        '''
        니케 상세 페이지에서 정보를 파싱하여 dict로 반환
        '''
        # 레어도
        el_tier = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[2]/p/img')
        url_tier = el_tier.get_attribute("src")
        text_tier = ""
        if "014d90279e0ca7278b0b8f7e9094a8dd" in url_tier: text_tier = "SSR"
        elif "64ec71645699f5fce79d98cfc20a525f" in url_tier: text_tier = "SR"
        elif "1c5d999c42ed640c95da540af7578667" in url_tier: text_tier = "R"

        # 스쿼드 이름
        el_squad = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[2]/div/p/span')
        text_squad = el_squad.text
        
        # 속성
        el_code = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[3]/div[1]/p[1]/img')
        url_code = el_code.get_attribute("src")
        text_code = ""
        if "code-fire" in url_code: text_code = "작열"
        elif "code-water" in url_code: text_code = "수냉"
        elif "code-electronic" in url_code: text_code = "전격"
        elif "code-wind" in url_code: text_code = "풍압"
        elif "code-iron" in url_code: text_code = "철갑"

        # 클래스
        el_class = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[3]/div[3]/p[1]/img')
        url_class = el_class.get_attribute("src")
        text_class = ""
        if "6337300758895c79f9afe8139500adcf" in url_class: text_class = "화력형"
        elif "6c30fdc2d204905edd388fc958359ec3" in url_class: text_class = "방어형"
        elif "e0ff69fe9e3cf29c55232d3590135811" in url_class: text_class = "지원형"

        # 기업
        el_company = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[3]/div[4]/p[1]/img')
        url_company = el_company.get_attribute("src")
        text_company = ""
        if "4aebd6d57c8afd17d334f37130ddc6e1" in url_company: text_company = "엘리시온"
        elif "f3d8f861af997ddc5e1ee81b715ae314" in url_company: text_company = "미실리스"
        elif "8c41a47786bc43a7db71bd7218f4989d" in url_company: text_company = "테트라"
        elif "29d8e5c949e2141795446ca64a867b9b" in url_company: text_company = "필그림"
        elif "2e913befc4fc5e84f9807f8a455232dc" in url_company: text_company = "어브노멀"

        # 버스트 단계
        el_bust_step = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[3]/div[5]/p[1]/img')
        url_bust_step = el_bust_step.get_attribute("src")
        text_bust_step = ""
        if "448005ff9513c9a8afabbece6ad2a05b" in url_bust_step: text_bust_step = "I"
        elif "50f0e38e5306ba4ddae04e494921216a" in url_bust_step: text_bust_step = "II"
        elif "3b98d581bb5776454e084075adf00c4c" in url_bust_step: text_bust_step = "III"
        elif "25bb6a298db42bf89464714a7fd6f159" in url_bust_step: text_bust_step = "A"

        # 무기 정보
        el_weapon_btn = self.driver.find_element(By.XPATH, '//*[@id="nikkes-basics"]/div[2]/div[3]/div[2]')
        el_weapon_btn.click()
        time.sleep(2)

        text_weapon_type = self.driver.find_element(By.XPATH,'//*[@id="nikkes-weapon"]/div/div/div[1]/p[1]/span').text
        text_weapon_magazine_capacity = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[2]/div/div[1]/p[1]/span[2]').text
        text_weapon_reload_time = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[2]/div/div[1]/p[2]/span[2]').text.split("s")[0].strip()
        text_weapon_operation_type = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[2]/div/div[1]/p[3]/span[2]').text
        text_weapon_spec = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[2]/div/div[2]/div').text

        self.driver.find_element(By.XPATH, '/html/body/div[4]/div/a').click() # 닫기
        time.sleep(1)

        # 스킬 정보
        el_skill = self.driver.find_element(By.CSS_SELECTOR, "[id='nikkes-weapon']")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el_skill)
        time.sleep(2)

        # 스킬 1
        text_skill_1_name = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[1]/div[1]/div[1]/p').text
        self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[1]/div[1]/div[1]/div[2]/a[4]').click()
        time.sleep(0.5)
        text_skill_1_spec = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[1]/div[1]/div[2]/div').text
        
        # 스킬 2
        text_skill_2_name = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[1]/div[1]/p').text
        self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[1]/div[1]/div[2]/a[4]').click()
        time.sleep(0.5)
        text_skill_2_spec = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[2]/div[1]/div[2]/div').text

        # 버스트
        text_skill_bust_name = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[3]/div[1]/div[1]/p').text
        text_skill_bust_time = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[3]/div[1]/div[1]/div[3]/span').text.split("s")[0].strip()
        self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[3]/div[1]/div[1]/div[2]/a[4]').click()
        time.sleep(0.5)
        text_skill_bust_spec = self.driver.find_element(By.XPATH, '//*[@id="nikkes-weapon"]/div/div/div[3]/div[1]/div[2]/div').text

        return {
            self.db_name: name,
            self.db_path_img_profile: f"resource/{name.replace(':', '-')}_profile.webp",
            self.db_tier: text_tier,
            self.db_squad: text_squad,
            self.db_code: text_code,
            self.db_weapon_type: text_weapon_type,
            self.db_weapon_magazine_capacity: int(text_weapon_magazine_capacity),
            self.db_weapon_reload_time: text_weapon_reload_time,
            self.db_weapon_operation_type: text_weapon_operation_type,
            self.db_weapon_description: text_weapon_spec,
            self.db_class: text_class,
            self.db_company: text_company,
            self.db_skill_bust_step: text_bust_step,
            self.db_skill_1_name: text_skill_1_name,
            self.db_skill_1_description: text_skill_1_spec,
            self.db_skill_2_name: text_skill_2_name,
            self.db_skill_2_description: text_skill_2_spec,
            self.db_skill_bust_name: text_skill_bust_name,
            self.db_skill_bust_time: text_skill_bust_time,
            self.db_skill_bust_description: text_skill_bust_spec
        }

    def DownloadResource(self):
        '''
        필수 - accessPage 후 진행
        1. 실제 이미지 소스가 없는지 db 쿼리
        2. 없는 이미지 다운로드
        '''
        self.db_cursor.execute("SELECT Path_img_profile FROM nikke_info_table")

        resource_list_db = []
        for item in self.db_cursor.fetchall():
            img_name = item[0].split("/")[-1]
            resource_list_db.append(img_name)

        resource_list_real = os.listdir(self.path_resource)
        resource_list_real.remove('empty_slot_img.png')# 빈 슬릇용 이미지는 제외

        # db에 있지만 실제로 없는 이미지 선별
        download_list = []
        for img_db in resource_list_db: # db에 등록된 이미지 이름과 대조
            if not img_db in resource_list_real: # 실제 이미지가 없다면
                # 저장할 때 이미지 이름에 특수문자가 들어가면 안되는 로직을 적용했기 때문에 복원 후 진행
                img_db_rename = img_db.replace(" - ", " : ")
                img_name = img_db_rename.split("_")[0]
                download_list.append(img_name)

        # 니케 리스트 가져오기.
        # 니케 슬릇에 적용된 css로 전부 가져와서 리스트로 만듦
        self.wait(By.CSS_SELECTOR, "[class='cursor-pointer relative nikkes-all-item h-[180px] max-h-[180px] w-[102px] max-w-[22%] my-[4.5px] mx-[5px]']")
        item_list = self.driver.find_elements(By.CSS_SELECTOR, "[class='cursor-pointer relative nikkes-all-item h-[180px] max-h-[180px] w-[102px] max-w-[22%] my-[4.5px] mx-[5px]']")
        
        time.sleep(0.5)
        # 니케 프로필 이미지가 모두 로딩되도록 스크롤 제어
        for i in range(33, len(item_list), 24):
            temp = self.driver.find_element(By.XPATH, f'//*[@id="layout-content"]/div/div[2]/div[2]/div/div/div[3]/div[{i}]')
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", temp)
            time.sleep(2)

        for item in item_list:
            # 니케 이름 가져오기
            el_span = item.find_element(By.XPATH, "./div[1]/span[1]")
            text_name = el_span.text

            if not text_name in download_list:
                continue

            # 기본 니케 프로필 이미지 가져오기
            file_name = f"resource/{text_name.replace(':', '-')}_profile.webp"

            el_img = item.find_element(By.CSS_SELECTOR, "img[class='nikkes-all-item-img']")
            url_img: str = el_img.get_attribute("src")# 프로필 이미지 링크
        
            response = requests.get(url_img)
            if response.status_code == 200: # http 200 -> 성공
                with open(file_name, 'wb') as file:
                    file.write(response.content)
                time.sleep(0.5)
            else:
                print(f"프로필 이미지 Http error - {response.status_code}")

    def updateSkillDB(self, skill_file):
        '''
        필수 - 크롤링해서 니케 기본 정보가 db로 만들어진 후, 진행해야함
        니케 정보에서 스킬 키워드 생성 및 업데이트
        정보와 스킬 키워드 연결 table도 업데이트
        '''
        # 키워드 파일 경로 (프로젝트 루트)
        project_root = self.path_resource.parent
        keyword_file_path = project_root / skill_file

        if not keyword_file_path.exists():
            print(f"에러: {keyword_file_path} 파일을 찾을 수 없습니다.")
            return

        print("키워드 데이터베이스 구축 시작...")
        
        # 1. 키워드 파일 파싱 및 테이블 삽입
        keywords_data = []
        with open(keyword_file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            current_category = ""
            for line in lines:
                line = line.strip()
                if not line: continue
                # 카테고리 추출 [카테고리명]
                if line.startswith("[") and "]" in line:
                    current_category = line[1:line.find("]")].strip()
                    continue
                # 키워드 추출 - 키워드명
                if line.startswith("-"):
                    kw = line[1:].strip()
                    if kw:
                        keywords_data.append((kw, current_category))

        # 기존 키워드 초기화 후 삽입 (최신화)
        self.db_cursor.execute("DELETE FROM skill_keywords_table")
        for kw, cat in keywords_data:
            # DB 저장 시 '증가' -> '▲', '감소' -> '▼' 로 치환하여 저장 (일관성 및 쿼리 최적화)
            db_kw = kw.replace("증가", "▲").replace("감소", "▼")
            self.db_cursor.execute("""
                INSERT OR IGNORE INTO skill_keywords_table (keyword_name, category)
                VALUES (?, ?)
            """, (db_kw, cat))
        self.db_connection.commit()

        # 2. 니케 스킬 분석 및 관계 매핑
        # 모든 니케의 스킬 설명 가져오기
        self.db_cursor.execute(f"SELECT id, {self.db_skill_1_description}, {self.db_skill_2_description}, {self.db_skill_bust_description} FROM nikke_info_table")
        nikkes = self.db_cursor.fetchall()

        # 모든 키워드 정보 가져오기
        self.db_cursor.execute("SELECT id, keyword_name FROM skill_keywords_table")
        db_keywords = self.db_cursor.fetchall()

        # 관계 테이블 초기화
        self.db_cursor.execute("DELETE FROM nikke_keywords_table")

        import unicodedata
        
        # 기호 정규화용
        SYMBOL_UP = "\u25b2"   # ▲
        SYMBOL_DOWN = "\u25bc" # ▼

        match_total_count = 0
        keyword_match_stats = {} 

        for nikke_id, s1, s2, sb in nikkes:
            # 1. 스킬 설명을 논리적 블록으로 분리
            # 보통 [...] 로 구분되어 있으나, 아닌 경우도 대비하여 줄바꿈/마침표 등으로도 분리
            raw_text = f"{s1 if s1 else ''} {s2 if s2 else ''} {sb if sb else ''}"
            raw_text = unicodedata.normalize('NFC', raw_text).replace("데미지", "대미지")
            
            # 블록 추출 (대괄호 안의 텍스트 또는 문장 단위)
            blocks = re.findall(r'\[([^\]]+)\]', raw_text)
            # 대괄호 밖의 텍스트도 문장 단위로 추가
            outside_text = re.sub(r'\[[^\]]+\]', ' | ', raw_text)
            blocks.extend([b.strip() for b in outside_text.split('|') if b.strip()])

            for kw_id, kw_name in db_keywords:
                # 검색 타겟 생성
                match_orig = re.search(r'\[원문:\s*(.*)\]', kw_name)
                if match_orig:
                    search_targets = [match_orig.group(1).strip()]
                else:
                    raw_parts = kw_name.split('/')
                    search_targets = []
                    first_part = raw_parts[0].strip()
                    search_targets.append(first_part)
                    subject_match = re.match(r'^([^n%▲▼\d]+)', first_part)
                    subject = subject_match.group(1).strip() if subject_match else ""
                    for other_part in raw_parts[1:]:
                        other_part = other_part.strip()
                        if subject and len(other_part) < 15 and subject not in other_part:
                            search_targets.append(f"{subject} {other_part}")
                        else:
                            search_targets.append(other_part)

                is_match = False
                for target in search_targets:
                    # 키워드에서 필수 단어들만 추출
                    target = unicodedata.normalize('NFC', target).replace("데미지", "대미지")
                    target = target.replace("n%", "TOKEN_PCT").replace("n초", "TOKEN_SEC").replace("n발", "TOKEN_VAL")
                    target = target.replace("n회", "TOKEN_CNT").replace("n기", "TOKEN_UNIT").replace("n단계", "TOKEN_STEP")
                    target = target.replace("n개", "TOKEN_EA").replace("n", "TOKEN_NUM")
                    
                    # 필수 단어 목록 (너무 짧은 단어는 제외하여 오탐 방지)
                    # TOKEN_... 패턴을 먼저 매칭하도록 순서 조정
                    essential_words = re.findall(r'TOKEN_[A-Z]+|[가-힣A-Za-z▲▼]{2,}|▲|▼', target)
                    if not essential_words: continue
                    
                    # 각 블록에 대해 필수 단어들이 모두 포함되어 있는지 확인 (순서 무관)
                    for block in blocks:
                        all_found = True
                        for word in essential_words:
                            # 개별 단어 매칭 패턴
                            if word == "TOKEN_PCT": p = r'[\d\.]+\s*%'
                            elif word == "TOKEN_SEC": p = r'[\d\.]+\s*초'
                            elif word == "TOKEN_VAL": p = r'[\d\.]+\s*발'
                            elif word == "TOKEN_CNT": p = r'[\d\.]+\s*회'
                            elif word == "TOKEN_UNIT": p = r'[\d\.]+\s*기'
                            elif word == "TOKEN_STEP": p = r'[\d\.]+\s*단계'
                            elif word == "TOKEN_EA": p = r'[\d\.]+\s*개'
                            elif word == "TOKEN_NUM": p = r'\d+'
                            elif word in ["▲", "증가"]: p = rf'(?:{SYMBOL_UP}|증가)'
                            elif word in ["▼", "감소"]: p = rf'(?:{SYMBOL_DOWN}|감소)'
                            elif word == "대미지": p = r'(?:대미지|데미지)'
                            elif word in ["공격", "최종", "기본", "자신", "아군", "적", "시전자", "배율", "비례"]:
                                continue # 선택적 단어는 체크 패스
                            else: p = re.escape(word)
                            
                            if not re.search(p, block):
                                all_found = False
                                break
                        
                        if all_found:
                            is_match = True
                            break
                    if is_match: break

                if is_match:
                    self.db_cursor.execute("INSERT OR IGNORE INTO nikke_keywords_table (nikke_id, keyword_id) VALUES (?, ?)", (nikke_id, kw_id))
                    match_total_count += 1
                    keyword_match_stats[kw_name] = keyword_match_stats.get(kw_name, 0) + 1

        self.db_connection.commit()
        print(f"키워드 매핑 완료: 총 {match_total_count}개의 관계 생성.")
        unmatched = [kw[1] for kw in db_keywords if kw[1] not in keyword_match_stats]
        if unmatched:
            print(f"매칭되지 않은 키워드 ({len(unmatched)}개): {', '.join(unmatched[:10])}...")

    def wait(self, by_type, element, timeout=10):
        '''
        해당 요소가 위치할 때 까지 대기
        '''
        WebDriverWait(self.driver, timeout=timeout).until(EC.element_to_be_clickable((by_type, element)))

manager = DBManager('./nikke_info.db')

#manager.createDBTable()
#manager.accessPage()
#manager.crawlingAllUpdateDB()
#manager.DownloadResource()

#manager.updateSkillDB("nikke_skill_keyward.txt")