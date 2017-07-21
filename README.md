*Пакет позволяет проделывать операции, над каждым листом файлов Excel, содержащихся в папке, переданной как параметр при инициализации объекта класса.*

*Эти операции над листами excel, определяются в функторе, передаваемом объекту класса, находящегося в пакете.*


*Пример:*

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
    s.eval(p)`
    
*Для работы требуется Python 3.*