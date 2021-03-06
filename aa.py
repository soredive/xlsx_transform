#-*- coding: utf-8 -*-
import pandas as pd
from pandas import DataFrame
import re
from tkinter import filedialog
from tkinter import *
import time
import sys
import os
from win10toast import ToastNotifier

class ExcelTrans:

	# 번들 경로 저장
	def __init__(self):
		self.appPath = ''
		self.iconPath = ''
		try:
			if getattr(sys, 'frozen', False):
				# 번들되었을때
				bundle_dir = sys._MEIPASS
			else:
				# 일반적인환경
				bundle_dir = os.path.dirname(os.path.abspath(__file__))
			self.appPath = bundle_dir
			if os.path.exists(bundle_dir + '/overwatch.ico') and os.path.isfile(bundle_dir + '/overwatch.ico'):
				self.iconPath = bundle_dir + '/overwatch.ico'
		except Exception as e:
			print("{}".format(str(e)))

	# 읽었던 경로 저장
	def setReadPath(self,path):
		configFile = os.getcwd()+'/market_excel_transform.log'
		if path:
		    f = open(configFile, "wt")
		    f.write(path)
		    f.close()

	# 읽었던 경로 가져오기
	def getReadPath(self):
		path = '/'
		configFile = os.getcwd()+'/market_excel_transform.log'
		if os.path.exists(configFile):
		    f = open(configFile,'r')
		    path = f.read()
		    f.close()
		return path

	# 출력 경로
	def outputPath(self,str):
		basename = os.path.basename(str)
		folder = os.path.dirname(str)
		names = basename.split('.')
		names[0] = names[0] + '-변환'
		if names[-1] == 'xls':
			names[-1] = 'xlsx'
		filename = '.'.join(names)
		path= folder + '/' + filename
		return path, filename, folder

	# 수량이 1개 이상인 경우 초록색
	def qtyOneMore(self,val):
		color = 'green' if val > 2 else 'black'
		return 'color: %s' % color
	
	# 정규식 옵션정보에서 불필요한 내용 제거
	def optionInfo(self,str):
		regex = r"[^:]*:([^/]*)"
		matches = re.finditer(regex, str, re.MULTILINE)
		lists = []
		
		for matchNum, match in enumerate(matches):
			for groupNum in range(0, len(match.groups())):
				groupNum = groupNum + 1
				newstr = match.group(groupNum).strip()
				if newstr:
					lists.append(newstr)
		
		return ' / '.join(lists) if lists else str.strip()	

	def start(self):
		root = Tk()
		root.title("마켓인벤 엑셀변환기")
		try:
			root.iconbitmap(self.iconPath)
		except Exception as e:
			print("Unexpected error: {}".format(str(e)))
		
		root.geometry("500x400")
		menubar = Menu(root)
		menubar.add_command(label="변환하기", command=self.gogogo)
		menubar.add_command(label="종료", command=root.quit)
		root.config(menu=menubar)
		root.mainloop()
		
	def gogogo(self):
		# 파일 읽기
		try:
			openFile = filedialog.askopenfilename(initialdir = self.getReadPath(),title = "파일을 고르세요",filetypes = (("엑셀파일","*.xls"),("모든 파일","*.*")))

			# 취소시 종료
			if not openFile:
				return

			# 읽고 처리하기
			lists = pd.read_excel(openFile, header=1)
			lists = lists.fillna('')

			lists = lists[['배송지','수취인명', '수취인연락처1','배송메세지','상품명', '옵션정보', '수량']]
			lists['옵션정보'] = lists['옵션정보'].apply(self.optionInfo)
			lists = lists[~lists['상품명'].str.contains('정식 라이센스|정식라이센스')]
			# lists.reset_index(drop=True, inplace=True)

			# 주소지가 다른 열마다 빈 공백 추가하기
			oldAddress = ''
			for i, row in lists.iterrows():
				if i != 0 and row.get('배송지') != oldAddress:
					line = DataFrame({'배송지':'','수취인명':'', '수취인연락처1':'','배송메세지':'','상품명':'', '옵션정보':'', '수량':''}, index=[i-0.5])
					lists = lists.append(line,ignore_index=False)
				oldAddress = row.get('배송지')
			lists=lists.sort_index().reset_index(drop=True)
			lists = lists[['배송지','수취인명', '수취인연락처1','배송메세지','상품명', '옵션정보', '수량']]

			# 읽고 나서 저장하기
			path, filename, folder = self.outputPath(openFile)

			# 가져왔던 폴더를 저장하기
			self.setReadPath(folder)

			writer = pd.ExcelWriter(path)
			lists.to_excel(writer,index=False,sheet_name='Sheet1',engine='xlsxwriter')

			workbook = writer.book
			worksheet = writer.sheets['Sheet1']

			worksheet.set_zoom(90)
			worksheet.set_column('A:A', 75)
			worksheet.set_column('B:B', 10)
			worksheet.set_column('C:C', 15)
			worksheet.set_column('D:D', 45)
			worksheet.set_column('E:E', 40)
			worksheet.set_column('F:F', 50)
			worksheet.set_column('G:G', 6)

			# 엑셀 조건부 서식
			number_rows = len(lists.index)
			color_range = "A2:G{}".format(number_rows+1)
			color_range2 = "A2:G{}".format(number_rows+1)
			format1 = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'})
			format2 = workbook.add_format({'bg_color': '#efefef'})
			worksheet.conditional_format(color_range, {'type': 'formula','criteria': '=$G2>1','format': format1})
			worksheet.conditional_format(color_range2, {'type': 'formula','criteria': '=isblank($A2)','format': format2})
			worksheet.freeze_panes(1, 0)
			writer.save()

			# 노티 보내기
			toaster = ToastNotifier()
			toaster.show_toast("마켓인벤 엑셀변환기",
			                   filename + "파일 변환이 완료되었습니다",
							   icon_path=self.iconPath,
			                   duration=5,
			                   threaded=True)
			# while toaster.notification_active(): time.sleep(0.1)

		except Exception as e:
			 print("Unexpected error: {}".format(str(e)))
			 time.sleep(10)

		os.startfile(folder)

if __name__ == "__main__":
	app = ExcelTrans()
	app.start()