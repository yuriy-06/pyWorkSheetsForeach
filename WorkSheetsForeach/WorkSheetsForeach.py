import os
import win32com.client

class WForeach:
	def __init__(self, pathDir):
		self.pathDir = pathDir
		self.files = os.listdir(self.pathDir)
	def eval(self, fn): # параметром метода идет функтор, несущий в себе полезную нагрузку
		self.excel = win32com.client.Dispatch("Excel.Application")
		for file in self.files:
			self.wb = self.excel.Workbooks.Open(self.pathDir + '/' + file)
			self.sheets = self.wb.WorkSheets
			for sh in self.sheets:
				fn(sh) # здесь передадим функтору объект Листа excel, чтобы он над ним произвел действия
			self.wb.Save()
			self.wb.Close()
		self.excel.Quit()