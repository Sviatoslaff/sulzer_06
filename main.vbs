Option Explicit
Const xlUp = -4162
Public Const firstCol = 39, lastCol = 45

Public Const resNoTemplate = " template not found. Check the template.  "
Public Const resNoBOM = "Nothing is inside this BOM. First make the BOM."

Dim qtn, plant, sorg, template, serno
Dim qtyRows, visibleRows, intRow, grid, bExit, bAbort, txtStatus

'1. Запрашиваем файл QTN и получаем массив значений для последующего заполнения SAP Quotation
Dim excelFile
excelFile = selectExcel()

qtn = "20330001"
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = qtn
session.findById("wnd[0]").sendVKey 0

'2. Заполняем открытый SAP Quotation
Dim ArticlesExcel, objWorkbook, ws
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
objWorkbook.Sheets("PMU").Activate
Set ws = objWorkbook.Worksheets("PMU")
Dim iLastRow: iLastRow = CInt(0)
iLastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row  
'WScript.Echo iLastRow
On Error Resume Next
Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
	'ReDim Preserve arrExcel(intRow - 4, 6)
	'WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
    Err.Clear
    tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
    Set grid = session.findById(tblArea)
    sapRow = grid.currentRow                'Here is the current visible row of the QTN
MsgBox "sap Row: " & sapRow

	If sapRow > 7 Then
		rowCount = grid.RowCount
		goto_pos = session.findById(tblArea & "/txtVBAP-POSNR[0," & sapRow - 5 & "]").text
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
		session.findById("wnd[1]/usr/txtRV45A-POSNR").text = goto_pos
		session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 3
		session.findById("wnd[1]").sendVKey 0
		WScript.Sleep 300

		tblArea = UserArea.findByName("SAPMV45ATCTRL_U_ERF_KONTRAKT", "GuiTableControl").Id
		Set grid = session.findById(tblArea)
		sapRow = grid.currentRow                'Here is the current visible row of the QTN
		Set cell = grid.GetCell(sapRow + 5, 1)
		cell.setFocus()
		sapRow = grid.currentRow                'Here is the current visible row of the QTN

	MsgBox "new sap Row: " & sapRow

	End If    

	For iCol = firstCol to lastCol
		session.findById(tblArea & "/txtVBAP-POSNR[0," & iCol + 1 - firstCol & "]").text = ArticlesExcel.Cells(intRow, iCol).Value
	Next 
	'WScript.Echo arrExcel(intRow - 4, 0)
	intRow = intRow + 1
	sapRow = sapRow + 1
	Set cell = grid.GetCell(sapRow, 1)
	cell.SetFocus()
	session.findById("wnd[0]").sendVKey 0
Loop


'3. 






MsgBox "Script finished! ", vbSystemModal Or vbInformation

'====== Functions ans Subs ========

'returns an unique array from an Excel file chosen by a user
Function GetExcelArray()
	'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
	Dim ArticlesExcel, objWorkbook, ws
	Set ArticlesExcel = CreateObject("Excel.Application")
	Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
	objWorkbook.Sheets("PMU").Activate
	Set ws = objWorkbook.Worksheets("PMU")
	Dim collTemp : Set collTemp = CreateObject("Scripting.Dictionary")

	Dim iLastRow: iLastRow = CInt(0)
   	iLastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row  
	'WScript.Echo iLastRow

	' Считаем, что в 4 строке - начало таблицы для обработки
	Dim intRow : intRow = 4
	Dim iCol
	' Цикл для каждой строки
	On Error Resume Next
	Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
		'ReDim Preserve arrExcel(intRow - 4, 6)
		'WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
		For iCol = firstCol to lastCol
			arrExcel(intRow - 4, iCol - firstCol) = ArticlesExcel.Cells(intRow, iCol).Value
		Next 
		WScript.Echo arrExcel(intRow - 4, 0)
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