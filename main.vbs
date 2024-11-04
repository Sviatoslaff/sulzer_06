Option Explicit
Public Const firstCol = 39, lastCol = 45

Public Const resNoTemplate = " template not found. Check the template.  "
Public Const resNoBOM = "Nothing is inside this BOM. First make the BOM."

Dim qtn, plant, sorg, template, serno
Dim qtyRows, visibleRows, intRow, grid, bExit, bAbort, txtStatus

'1. Запрашиваем файл QTN и получаем массив значений для последующего заполнения SAP Quotation
Dim excelFile
excelFile = selectExcel()
Dim arrExcel : arrExcel = GetExcelArray()

'2. Вставляем полученный массив значений в SAP Quotation


' Make a structure for a report
Dim qtySerno : qtySerno = UBound(arrSerno)
Dim arrReport, strReport
Dim dicReport : Set dicReport = CreateObject("Scripting.Dictionary")

'WScript.Echo Join(arrSerno)

MsgBox "Script finished! ", vbSystemModal Or vbInformation

'====== Functions ans Subs ========

'returns an unique array from an Excel file chosen by a user
Function GetExcelArray()
	'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
	Dim ArticlesExcel, objWorkbook
	Set ArticlesExcel = CreateObject("Excel.Application")
	Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
	objWorkbook.Sheets("PMU").Activate
	Dim arrExcel()
	
	' Считаем, что в 4 строке - начало таблицы для обработки
	Dim intRow : intRow = 4
	Dim iCol
	' Цикл для каждой строки
	On Error Resume Next
	Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
		ReDim Preserve arrExcel(intRow - 4, 6)
		WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
		For iCol = firstCol to lastCol
			arrExcel(intRow - 4, iCol - firstCol) = ArticlesExcel.Cells(intRow, iCol).Value
		Next 
		intRow = intRow + 1
	Loop
	objWorkbook.Close False
	ArticlesExcel.Quit
	WScript.Echo Join(arrExcel)
	GetExcelArray = arrExcel
End Function


Sub OutputToExcel
	Dim ReportExcel, objWorkbook
	Set ReportExcel = CreateObject("Excel.Application")
	Set objWorkbook = ReportExcel.Workbooks.Add()
	ReportExcel.Visible = True
	
	arrReport = dicReport.Items
	intRow = 0
	For Each serno In arrSerno
		strReport = strReport & serno & " : " & arrReport(intRow) & VbCrLf
		ReportExcel.cells(intRow + 1, 1).value = serno
		ReportExcel.cells(intRow + 1, 2).value = arrReport(intRow)
		intRow = intRow + 1
	Next
End Sub



Sub notused


'StartTransaction("ZIB07")
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0

bAbort = vbFalse
For Each serno In arrSerno
	bExit = vbFalse
	session.findById("wnd[0]/usr/ctxtP_EQUNR").text = serno
	session.findById("wnd[0]/usr/ctxtP_WERKS2").text = plant
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	WScript.Sleep 500 'Delay for SAP processing
	If session.findById("wnd[0]/usr/ctxtP_EQUNR", False) Is Nothing Then
		Do While session.findById("wnd[0]/usr/chkJOB", False) Is Nothing
			If session.findById("wnd[1]/usr/txtLV_MATNR1", False) Is Nothing Then
				dicReport.Add serno, resNoBOM
				bExit = vbTrue
				session.findById("wnd[1]").sendVKey 0
				Exit Do
			Else
				session.findById("wnd[1]/tbar[0]/btn[8]").press 'V
				'session.findById("wnd[1]/tbar[0]/btn[2]").press       'X        
			End If
		Loop
		
		If Not bExit Then
			session.findById("wnd[0]/usr/chkJOB").selected = False
			session.findById("wnd[0]/usr/chkJOB").setFocus
			
			Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")
			
			qtyRows = grid.rowCount - 1
			'MsgBox "Rows amount: " & qtyRows
			visibleRows = grid.VisibleRowCount

			' Цикл для каждой строки
			'On Error Resume Next
			intRow = 0
			Do Until intRow > qtyRows
				'Err.Clear
				'MsgBox "Row: " & intRow
				grid.modifyCell intRow, "TEMPLATE", template
				grid.currentCellRow = intRow
				intRow = intRow + 1
			Loop
			grid.triggerModified
			session.findById("wnd[0]/tbar[1]/btn[8]").press
			'    MsgBox "Next Control - btn[3]", vbSystemModal Or vbInformation

			' It can be error that mat number not found - If for that
			If session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
			Else
				bAbort = vbTrue
				dicReport.Add serno, template & resNoTemplate
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End If
			
			If Not bAbort Then
				session.findById("wnd[0]/tbar[0]/btn[3]").press
				'    MsgBox "Next Control - wnd[1]/tbar[0]/btn[0]", vbSystemModal Or vbInformation
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				dicReport.Add serno, resOK
			End If
		End If
	Else
		' Same selection window - check for status bar
		If session.ActiveWindow.findById("sbar", False) Is Nothing Then
			dicReport.Add serno, resExists
		Else
			txtStatus = session.ActiveWindow.findById("sbar").Text
			dicReport.Add serno, txtStatus
		End If
		
	End If
Next

OutputToExcel

End Sub