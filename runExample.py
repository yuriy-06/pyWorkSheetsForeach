from WorkSheetsForeach.WorkSheetsForeach import WForeach as wf

s = wf("G:/temp/virtEnvPy/py36-32/xls/Columns 4C,4F")
class Payload():
	def __call__(self, sheet):
		sheet.PageSetup.Orientation = 1
		sheet.PageSetup.Zoom = 92
		val = sheet.Cells(117,1).Value
		if val != None:
			print(val)
			sheet.PageSetup.PrintArea = "$A$1:$T$172"
			sheet.PageSetup.PrintArea = "$A$1:$K$172"
			sheet.PageSetup.PrintArea = "$A$1:$J$172"
		else:
			sheet.PageSetup.PrintArea = "$A$1:$J$122"
			sheet.PageSetup.PrintArea = "$A$1:$J$114"
p = Payload()
s.eval(p)