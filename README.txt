����� ��������� ����������� ��������, ��� ������ ������ ������ Excel, ������������ � ������ �����, ���������� ��� �������� ��� ������������� ������� ������.

��� �������� ��� ������� excel, ������������ � ��������, ������������ ������� ������, ������������ � ������.


������:

import WorkSheetsForeach.WorkSheetsForeach as wf

s = wf("g:/path")
class Payload():
	def __call__(self, sheet):
		sheet.PageSetup.Orientation = 1
		sheet.PageSetup.Zoom = 92
		val = sh.Cells(117,1).Value
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

��� ������ ��������� Python 3.