#import tkinter as tk
#from tkinter import ttk
#from tkinter import scrolledtext
#from tkinter import messagebox as msg
from tkinter import *
import requests, json
import sys
import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import queue
from threading import Thread
import time

#=====================================
class _MainThread():
	def __init__(self):
		self.tkhandler = Tk()
		self.tkhandler.resizable(False, False)
		self.tkhandler.geometry('1000x500')
		self.tkhandler.title('통계 자동 수집/발송 프로그램')


		self.label_title_1 = Label(self.tkhandler, text='병원명을 입력해주세요.')
		self.label_title_1.pack()


	def method_in_a_thread(self):
		print('hello world')

	def beginThread(self):
		self.run_thread = Thread(target=self.method_in_a_thread)
		self.run_thread.start()

	def click_me(self):
		pass


	def run(self):
		self.tkhandler.mainloop()

#=====================================
start = _MainThread()
start.run()