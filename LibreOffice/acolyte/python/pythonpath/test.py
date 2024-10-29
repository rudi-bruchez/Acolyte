def myPAGE():
	doc = XSCRIPTCONTEXT.getDocument()
	text = doc.Text
	cursor = text.createTextCursor()
#===================================================================#
##                           СТРАНИЦА                              ##
#===================================================================#
	style = doc.CurrentController.ViewCursor.PageStyleName
	page = doc.StyleFamilies.getByName("PageStyles").getByName(style)
	page.Width      = 29700  # ширина
	page.Height     = 21000  # высота
	page.LeftMargin = page.RightMargin  = 0  # поля
	page.TopMargin  = page.BottomMargin = 0  # поля
	page.setPropertyValue("BackColor", 0x99FF99)  # цвет
#===================================================================#
##                           ЗАГОЛОВОК                             ##
#===================================================================#
	import locale; from datetime import date
	# русский язык для названия месяца и дней недели Windows
	locale.setlocale(locale.LC_TIME, "RUS")
	# параметры заголовка
	cursor.ParaStyleName = "Heading 4"
	cursor.setPropertyValue("ParaAdjust", 3)  # Center
	# указанный год, указанный месяц, 1-ое число
	y = YEAR; m = MONTH
	dd = date(y,m,1)
	# строка заголовка месяц и год
	text.Start.String = dd.strftime("%B %Y г.")
#===================================================================#
##                            ТАБЛИЦА                              ##
#===================================================================#
	# количество дней в указанном месяце
	if m == 12: x = 31  # для декабря
	else:
		ddd = date(y,m+1,1)	
		x = ddd - dd
		x = x.days  # для всех остальных месяцев
	# количество столбцов
	c = x + 4
	# вставляем таблицу
	TT = doc.createInstance("com.sun.star.text.TextTable")
	TT.initialize(6,c)  # 6 строк и с столбцов
	text.insertTextContent( cursor, TT, 0)
# ширина столбцов таблицы--------------------------------------------
	# создаем коллекцию разделителей
	ZZ = TT.TableColumnSeparators
	# позиция каждого разделителя в процентах
	z = x + 3; y = 10000  # 100%
	# начинаем с последнего разделителя
	while z > 2:
		z -= 1		# или z = z - 1
		y -= 200	# шаг -2%
		ZZ[z].Position = y
	ZZ[2].Position = 3300; ZZ[1].Position = 2800	# 28%
	ZZ[0].Position = 1800	# 18%
	# возвращаем изменения в таблицу
	TT.TableColumnSeparators = ZZ
#===================================================================#
	MsgBox()                              #  вторая часть макроса	
#===================================================================#
##                         ШАПКА ТАБЛИЦЫ                           ##
#===================================================================#
	# цвет строк таблицы в формате RR(TT,index,color)
	RR(TT,0,0x4C1900); RR(TT,1,0x4C1900)
	# заголовки таблицы в формате HH(TT,c,r,text)
	HH(TT,0,1,"ФАМИЛИЯ И.О."); HH(TT,1,1,"ТАБ. №")
	HH(TT,2,1,"+   —"); HH(TT,3,1,"ЧАСЫ")
	# даты в формате HH(TT,c,r1,text); c - column; r - row
	c = 3; r1 = 0; d = 0
	while d < x:
		c += 1	# или c = c + 1
		d += 1
		HH(TT,c,r1,d)
	# дни недели в формате WW(TT,c,r2,text)
	c = 3; r2 = 1; d = 0
	while d < x:
		c += 1
		d += 1
		dd = dd.replace(day=d)
		w = dd.strftime("%a")
		WW(TT,c,r2,w)
#===================================================================#
##                            ГРАФИК                               ##
#===================================================================#
	x += 4	# или x = x + 4
	# для каждой буквы написан отдельный цикл
	# ячейки таблицы в формате CC(TT,c,r3,text)
	c5 = 4; c6 = 5; c7 = 6; c8 = 7; r3  = 2
	while c5 < x: CC(TT,c5,r3,"Д"); c5 += 4
	while c6 < x: CC(TT,c6,r3,"Н"); c6 += 4
	while c7 < x: CC(TT,c7,r3,"—"); c7 += 4
	while c8 < x: CC(TT,c8,r3,"—"); c8 += 4
	# ячейки таблицы в формате CC(TT,c,r4,text)
	c5 = 4; c6 = 5; c7 = 6; c8 = 7; r4  = 3
	while c5 < x: CC(TT,c5,r4,"—"); c5 += 4
	while c6 < x: CC(TT,c6,r4,"Д"); c6 += 4
	while c7 < x: CC(TT,c7,r4,"Н"); c7 += 4
	while c8 < x: CC(TT,c8,r4,"—"); c8 += 4
	# ячейки таблицы в формате CC(TT,c,r5,text)
	c5 = 4; c6 = 5; c7 = 6; c8 = 7; r5  = 4
	while c5 < x: CC(TT,c5,r5,"—"); c5 += 4
	while c6 < x: CC(TT,c6,r5,"—"); c6 += 4
	while c7 < x: CC(TT,c7,r5,"Д"); c7 += 4
	while c8 < x: CC(TT,c8,r5,"Н"); c8 += 4		
	# ячейки таблицы в формате CC(TT,c,r6,text)
	c5 = 4; c6 = 5; c7 = 6; c8 = 7; r6  = 5
	while c5 < x: CC(TT,c5,r6,"Н"); c5 += 4
	while c6 < x: CC(TT,c6,r6,"—"); c6 += 4
	while c7 < x: CC(TT,c7,r6,"—"); c7 += 4
	while c8 < x: CC(TT,c8,r6,"Д"); c8 += 4
	
#===================================================================#
##                       Сокращения для ТАБЛИЦЫ:                   ##
#===================================================================#
#   ТТ - таблица                                                    #
#-------------------------------------------------------------------#
#   RR - строка таблицы                                             #
#-------------------------------------------------------------------#
#   HH - ячейка заголовка                                           #
#   WW - ячейка с днями недели                                      #
#   СС - ячейка с текстом или целыми числами                        #
#-------------------------------------------------------------------#
#   OO - объект (строка или ячейка таблицы)                         #       
#   UU - курсор                                                     #
#   ZZ - разделитель                                                #   
#===================================================================#
##                      ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ                    ##
#===================================================================#
	
def RR(TT,index,color):
	OO = TT.Rows.getByIndex(index)
	OO.setPropertyValue("BackColor", color)
	
def HH(TT,c,r,text):
   	OO = TT.getCellByPosition(c,r)
   	UU = OO.createTextCursor()
   	UU.setPropertyValue("ParaAdjust", 3) # Center
   	UU.setPropertyValue("CharHeight", 11)
   	UU.setPropertyValue("CharColor", 0x99FF99)
   	OO.setString(text)
   
def WW(TT,c,r,text):
   	OO = TT.getCellByPosition(c,r)
   	UU = OO.createTextCursor()
   	UU.setPropertyValue("ParaAdjust", 3)
	UU.setPropertyValue("CharHeight", 9)
	if text == "Сб":
		UU.setPropertyValue("CharColor", 0x00FFFF)
	elif text == "Вс":
		UU.setPropertyValue("CharColor", 0xFF00FF)   
	else:
		UU.setPropertyValue("CharColor", 0x99FF99)
	OO.setString(text)
   
def CC(TT,c,r,text):
	OO = TT.getCellByPosition(c,r)
	UU = OO.createTextCursor()
	UU.setPropertyValue("ParaAdjust", 3)
	OO.setString(text)
   
def MsgBox():
	from ctypes import windll
	msg = windll.user32.MessageBoxA \
	(0, "Конец первой части макроса,\r"
	"выберите дальнейшее действие!",
	"Сообщение на Си!", 0x246)
	# если нажата кнопка
	if   msg ==  2: MsgBox() # Отмена=2
	elif msg == 10: myPAGE() # Повторить=10
	elif msg == 11: 0        # Продолжить=11

g_exportedScripts = myPAGE,	# эту функцию увидим в "Моих макросах"

''' MessageBox.Параметры кнопок и значков:
	MB_OK                =   0x0 # (OK=1)
	MB_OKCANCEL          =   0x1 # (OK=1; Отмена=2)
	MB_ABORTRETRYIGNORE  =   0x2 # (Прервать=3; Повтор=4; Пропустить=5)
	MB_YESNOCANCEL       =   0x3 # (Да=6; Нет=7; Отмена=2)
	MB_YESNO             =   0x4 # (Да=6; Нет=7)
	MB_RETRYCANCEL       =   0x5 # (Повтор=4; Отмена=2)
	MB_CANCELRETRYIGNORE =   0x6 # (Отмена=2; Повторить=10; Продолжить=11)
	
	MB_ICONHAND          =  0x10 # значок ошибки
	MB_ICONQUESTION      =  0x20 # значок вопроса
	MB_ICONEXCLAIMATION  =  0x30 # значок внимания
	MB_ICONASTRISK       =  0x40 # значок информации
	
	DEF_2                = 0x100 # выделить вторую кнопку
	DEF_3                = 0x200 # выделить третью кнопку
'''	
'''	ПРИМЕР:
	0x200 = 512 = DEF_3                - выделить третью кнопку
	 0x40 =  64 = MB_ICONASTRISK       - значок информации
	  0x6 =   6 = MB_CANCELRETRYIGNORE - (Отмена; Повторить; Продолжить)
	-----------------------------------------------------------
	0x246 = 582 = DEF_3 | MB_ICONASTRISK | MB_CANCELRETRYIGNORE
	-----------------------------------------------------------
	Параметры кнопок можно записать любым из этих трех способов,
	но оптимальным является шестнадцатиричное представление - 0x246
'''

