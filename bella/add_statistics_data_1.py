import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from threading import Thread
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import messagebox as msg
#=====================================
win = tk.Tk()

win.title("벨라 전용 데이터 수집기")

win.resizable(False, False)

frame1 = tk.LabelFrame(win, text=' 1. 통계 추출 날짜 입력하기')
frame1.grid(row=0, column=0, padx=5, pady=5)

frame2 = tk.LabelFrame(win, text=' 2. 통계 파일 수집하기')
frame2.grid(row=0, column=1, padx=5, pady=5)

frame3 = tk.LabelFrame(win, text=' 3. 통계 데이터 읽기')
frame3.grid(row=0, column=2, padx=5, pady=5)

frame4 = tk.LabelFrame(win, text=' 4. 통계 결과 확인')
frame4.grid(row=0, column=3, padx=5, pady=5)

frame5 = tk.LabelFrame(win, text=' 5. 데이터 입력')
frame5.grid(row=0, column=4, padx=5, pady=5)
#=====================================
# 전역변수
ID = 'dave.lee@goodoc.co.kr'
PW = 'dlwlstjr1!'

# 구글스프레드 시트 인증 정보
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cred.json', scope)

# Key 정보 인증
gs = gspread.authorize(credentials)

# CS접수현황 문서 가져오기
doc = gs.open_by_url('https://docs.google.com/spreadsheets/d/1y3kHyN6tnvPUGPP4Ui8E1E2jarp7UvhJpiHbKwDccig/edit?ts=5c296e96#gid=1268013855')
ws = doc.get_worksheet(0)  # 첫번째 시트 선택

#=====================================
# 추출 버튼 클릭 시
def _extracion():
	print(input_date_box.get())

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
		# 시작날짜 박스 선택
		elem = driver.find_element_by_name('start_at')
		elem.click()

		###############################################
		# 시작날짜 지우기
		n = 0
		while n < 10:
			elem.send_keys(Keys.BACKSPACE)
			n += 1
		# 시작날짜 입력
		elem.send_keys(input_date_box.get())

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
		elem.send_keys(input_date_box.get())

		###############################################
		time.sleep(1)

		# 접수자 로데이터 뽑아내기 버튼 클릭
		elem = driver.find_element_by_link_text('일별 접수자 뽑아내기')
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

matrix = list()
sum1 = 0            # 태블릿 접수 건
sum2 = 0            # 태블릿 더미 접수 건
sum3 = 0            # 태블릿 접수 합계 건 = 총 접수건
sum4 = 0            # 태블릿 접수 1건 이상
sum5 = 0            # 태블릿 접수 5건 이상

def click_me_2():
	global sum1, sum2, sum3, sum4, sum5

	# 추출 데이터 초기화
	matrix.clear()
	sum1 = 0
	sum2 = 0
	sum3 = 0
	sum4 = 0
	sum5 = 0

	file = load_workbook('접수자 .xlsx', data_only=True)
	file_2 = file['doro']

	# 변수에 데이터 넣기
	for n in file_2.rows:
		row_value = list()
		for n2 in n:
			row_value.append(n2.value)
		matrix.append(row_value)
	# csv 파일 닫기
	file.close()

	#pprint(matrix)
	print(len(matrix))
	print('---')

	#print(type(matrix[1][0]))
	#print(type(matrix[1][1]))
	#print(type(matrix[1][2]))
	#print(type(matrix[1][3]))
	#print(type(matrix[1][4]))
	#print(type(matrix[1][5]))
	#print(type(matrix[1][6]))

	# 합계 계산
	for v1 in range(1,len(matrix)):
		#print(matrix[v1][4], matrix[v1][5], matrix[v1][6])

		sum1 = sum1 + matrix[v1][4]
		sum2 = sum2 + matrix[v1][5]
		sum3 = sum3 + matrix[v1][6]

		if matrix[v1][4] >= 1:
			sum4 = sum4 + 1
		if matrix[v1][4] >= 5:
			sum5 = sum5 + 1

	print('===')
	print(sum1)
	print(sum2)
	print(sum3)
	print('===')
	print(sum4)
	print(sum5)
	print('===')
	print(type(sum3))

	# 박스 값 초기화
	result_value_box_1.delete(0, tk.END)
	result_value_box_2.delete(0, tk.END)
	result_value_box_3.delete(0, tk.END)
	result_value_box_4.delete(0, tk.END)
	result_value_box_5.delete(0, tk.END)

	# 결과 값 입력
	result_value_box_1.insert(tk.INSERT, sum4)
	result_value_box_2.insert(tk.INSERT, sum5)
	result_value_box_3.insert(tk.INSERT, sum3)
	result_value_box_4.insert(tk.INSERT, sum1)
	result_value_box_5.insert(tk.INSERT, sum2)

def create_thread_2():
	run_thread = Thread(target=click_me_3)           # 메서드 대상 지정
	run_thread.setDaemon(True)
	run_thread.start()
	print('입력 쓰레드 시작 : ', run_thread)
	print('---')

def click_me_3():
	if credentials.access_token_expired:
		gs.login()
		print("restart")
		time.sleep(1)

	# 데이터 넣을 행 찾기
	values_list = ws.col_values(16)
	#print(values_list)
	print(len(values_list))

	row_number = str(len(values_list) + 1)
	row_number_2 = str(len(values_list))

	ws.update_acell('P'+ row_number , result_value_box_1.get())
	ws.update_acell('R' + row_number, result_value_box_2.get())
	ws.update_acell('W' + row_number, result_value_box_3.get())
	ws.update_acell('X' + row_number, result_value_box_4.get())
	ws.update_acell('Y' + row_number, result_value_box_5.get())

	ws.update_acell('N' + '{number}'.format(number=row_number), '=value(left(L{number},4)-P{number})'.format(number=row_number))
	ws.update_acell('O' + '{number}'.format(number=row_number),
	                '=if((N{number}/value(left(L{number},len(L{number})-1)))*100>value(left(O{number2},len(Q{number2})-2)), round(100*sum(N{number}/value(left(L{number},len(L{number})-1))),1)&"%▲", round(100*(N{number}/value(left(L{number},len(L{number})-1))),1)&"%▽")'.format(number=row_number, number2 = row_number_2))
	ws.update_acell('Q' + '{number}'.format(number=row_number),
	                '=if((P{number}/value(left(L{number},len(L{number})-1)))*100>value(left(Q{number2},len(Q{number2})-2)), ROUND(100*(P{number}/value(left(L{number},len(L{number})-1))),1)&"%▲", ROUND(100*(P{number}/value(left(L{number},len(L{number})-1))),1)&"%▽")'.format(number=row_number, number2 = row_number_2))
	ws.update_acell('S' + '{number}'.format(number=row_number),
	                '=if((R{number}/value(left(L{number},len(L{number})-1)))*100>VALUE(left(S{number2},len(S{number2})-2)), ROUND(100*(R{number}/value(left(L{number},len(L{number})-1))),1)&"%▲", ROUND(100*(R{number}/value(left(L{number},len(L{number})-1))),1)&"%▽")'.format(number=row_number, number2 = row_number_2))
	ws.update_acell('Z' + '{number}'.format(number=row_number),
	                '=sum(X{number}/W{number})'.format(number=row_number))
	ws.update_acell('AA' + '{number}'.format(number=row_number),
	                '=if(Z{number}>Z{number}, "▲", "▽")'.format(number=row_number))


	msg.showinfo("알림", "시트에 데이터가 추가되었습니다.")
#=====================================
# 안내 문구 1
ttk.Label(frame1, text=' 시작일 = 종료일 :').grid(column=0, row=1, padx=3, pady=0, sticky='W')

# 시작 날짜 입력 박스
input_date = tk.StringVar()
input_date_box = ttk.Entry(frame1, width=12, textvariable=input_date)
input_date_box.grid(column=1, row=1, padx=5, pady=5, sticky='W')

#-------------------------------------
# 통계 추출 버튼
action_1 = ttk.Button(frame2, text="실행!", command=click_me).grid(column=0, row=0, ipadx=10, ipady=10, padx=20, pady=3)

#-------------------------------------
# 통계 수집 버튼
action_2 = ttk.Button(frame3, text="읽기!", command=click_me_2).grid(column=0, row=0, ipadx=10, ipady=10, padx=20, pady=3)

#-------------------------------------
# 통계 결과 확인
ttk.Label(frame4, text=' 태블릿 접수 1건 이상 :').grid(column=0, row=0, padx=3, pady=0, sticky='W')
ttk.Label(frame4, text=' 태블릿 접수 5건 이상 :').grid(column=0, row=1, padx=3, pady=0, sticky='W')

ttk.Label(frame4, text=' 전체 접수 건 :').grid(column=0, row=2, padx=3, pady=0, sticky='W')

ttk.Label(frame4, text=' 태블릿 접수 건 :').grid(column=0, row=3, padx=3, pady=0, sticky='W')
ttk.Label(frame4, text=' 태블릿 더미 접수건 :').grid(column=0, row=4, padx=3, pady=0, sticky='W')


result_value_1 = tk.IntVar()
result_value_box_1 = ttk.Entry(frame4, width=12, textvariable=result_value_1)
result_value_box_1.grid(column=1, row=0, padx=5, pady=5, sticky='W')

result_value_2 = tk.IntVar()
result_value_box_2 = ttk.Entry(frame4, width=12, textvariable=result_value_2)
result_value_box_2.grid(column=1, row=1, padx=5, pady=5, sticky='W')

result_value_3 = tk.IntVar()
result_value_box_3 = ttk.Entry(frame4, width=12, textvariable=result_value_3)
result_value_box_3.grid(column=1, row=2, padx=5, pady=5, sticky='W')

result_value_4 = tk.IntVar()
result_value_box_4 = ttk.Entry(frame4, width=12, textvariable=result_value_4)
result_value_box_4.grid(column=1, row=3, padx=5, pady=5, sticky='W')

result_value_5 = tk.IntVar()
result_value_box_5 = ttk.Entry(frame4, width=12, textvariable=result_value_5)
result_value_box_5.grid(column=1, row=4, padx=5, pady=5, sticky='W')

#-------------------------------------
# 통계 입력 버튼
action_3 = ttk.Button(frame5, text="입력!", command=create_thread_2).grid(column=0, row=0, ipadx=10, ipady=10, padx=20, pady=3)

#=====================================
win.mainloop()