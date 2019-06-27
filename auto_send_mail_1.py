import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import messagebox as msg
from tkinter import *
import requests, json
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from threading import Thread
#=====================================
win = tk.Tk()

win.title("통계 데이터 자동 수집 프로그램")

win.resizable(False, False)
#=====================================
# 전역변수
ID = 'dave.lee@goodoc.co.kr'
PW = ''

#=====================================
# 추출 버튼 클릭 시
def _extracion():
	print(start_date_box.get())
	print(end_date_box.get())

	hospital_name_list = hospital_name_box.get('1.0', END).split()

	driver = webdriver.Chrome('./chromedriver')

	# 접수현황 페이지 입장
	# search_url_sample = 'https://www.goodoc.co.kr/admin/hospital_booking_clients?search_category=hospital_name&keyword=예뻐진의원'

	search_url_sample = 'https://www.goodoc.co.kr/admin/hospital_booking_clients?uid=265751'

	# driver.get('https://www.goodoc.co.kr/admin/hospital_booking_clients?uid=265751')
	driver.get(search_url_sample)

	# id, name, class 속성 값을 찾아 계정을 입력한다.
	elem = driver.find_element_by_name('email')
	elem.send_keys(ID)

	elem = driver.find_element_by_name('password')
	elem.send_keys(PW)

	elem = driver.find_element_by_name('commit')
	elem.click()

	time.sleep(2)

	try:
		for count in range(len(hospital_name_list)):

			search_url_sample_2 = 'https://www.goodoc.co.kr/admin/hospital_booking_clients?search_category=hospital_name&keyword=' + \
			                      hospital_name_list[count]

			# 새 탭 생성1
			#driver.execute_script("window.open('hospital_booking_clients?search_category=hospital_name&keyword=예뻐진의원')")
			driver.execute_script('window.open("about:blank", "_blank");')
			#driver.execute_script('window.open("http://www.naver.com", "_blank");')
			#driver.execute_script('window.open("http://www.naver.com");')
			#driver.execute_script("'window.open('" + search_url_sample_2 + "');'")

			# 새 탭 이동
			handles = driver.window_handles
			print(handles)
			last_tab = driver.window_handles[-1]
			driver.switch_to.window(last_tab)

			# 링크 이동
			driver.get(search_url_sample_2)

			# 시작날짜 박스 선택
			elem = driver.find_element_by_name('start_at')
			elem.click()
			# 시작날짜 지우기
			n = 0
			while n < 10:
				elem.send_keys(Keys.BACKSPACE)
				n += 1
			# 시작날짜 입력
			elem.send_keys(start_date_box.get())

			###############################################

			# 종료날짜 박스 선택
			elem = driver.find_element_by_name('end_at')
			elem.click()
			# 종료날짜 지우기
			n = 0
			while n < 10:
				elem.send_keys(Keys.BACKSPACE)
				n += 1
			# 종료날짜 입력
			elem.send_keys(end_date_box.get())

			###############################################

			time.sleep(1)

			# 접수자 로데이터 뽑아내기 버튼 클릭
			elem = driver.find_element_by_id('btn_export')
			elem.click()

			print('success')

			# 새 탭에서 주소 변경
			#driver.get(search_url_sample)

			# 키 입력
			# ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform();

		input()


	except Exception as e:
		print('앗 이런')
		print(e)
		input()
	finally:
		input()
		#driver.quit()

def click_me():                                       # 추출 버튼 클릭 시 command에 담아야 할 이벤트
	create_thread()
	print("쓰레드 메서드 호출 : ", create_thread)
	print('---')

def create_thread():
	run_thread = Thread(target=_extracion)           # 메서드 대상 지정
	run_thread.setDaemon(True)
	run_thread.start()
	print('검색 쓰레드 시작 : ', run_thread)
	print('---')

#=====================================
# 안내 문구1
ttk.Label(win, text=' * 통계를 출력할 병원 리스트를 불러왔습니다.').grid(column=0, row=0, padx=3, pady=0, sticky='W')

# 병원명 입력창
scroll_w1 = 50
scroll_h1 = 20
hospital_name_box = scrolledtext.ScrolledText(win, width=scroll_w1, height=scroll_h1, wrap=tk.CHAR)      # => wrap option=CHAR/WORD
hospital_name_box.grid(column=0, row=1, columnspan=3, padx=5, pady=5)
#-------------------------------------
# 안내 문구2
ttk.Label(win, text=' * 통계 추출 날짜를 입력해주세요. \n   ex) yyyy-mm-dd').grid(column=0, row=3, padx=3, pady=0, sticky='W')

# 안내 문구2-1
ttk.Label(win, text='   시작일 :').grid(column=0, row=4, padx=3, pady=0, sticky='W')

# 안내 문구2-2
ttk.Label(win, text='   종료일 :').grid(column=0, row=5, padx=3, pady=0, sticky='W')

# 시작 날짜 입력 박스
start_date = tk.StringVar()
start_date_box = ttk.Entry(win, width=12, textvariable=start_date)
start_date_box.grid(column=0, row=4, padx=5, pady=5)

# 종료 날짜 입력 박스
end_date = tk.StringVar()
end_date_box = ttk.Entry(win, width=12, textvariable=end_date)
end_date_box.grid(column=0, row=5, padx=5, pady=5)
#-------------------------------------
# 안내문구 3
ttk.Label(win, text=' * 통계 추출하기!').grid(column=1, row=3, padx=3, pady=0, sticky='W')

# 추출 버튼
action_1 = ttk.Button(win, text="뀨우!", command=click_me).grid(column=1, row=4, ipady=10, pady=3, rowspan=50)

#=====================================
file = open("병원 리스트.txt", 'r')

line = file.read().splitlines()
print(line)

file.close()

for n in range(len(line)):
	hospital_name_box.insert(INSERT, line[n] + '\n')

print(range(len(line)))

#=====================================
win.mainloop()