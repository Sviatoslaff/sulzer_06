Option Explicit
Const xlUp = -4162, xlPasteValues = -4163, xlNone = -4142
Public Const firstCol = 39, lastCol = 45
'''Public Const firstCol = 40, lastCol = 40

Public Const resNoTemplate = " template not found. Check the template.  "
Public Const resNoBOM = "Nothing is inside this BOM. First make the BOM."

Dim tblArea
Dim qtn, plant, sorg, template, serno
Dim qtyRows, rowCount, visibleRows, sapRow, goto_pos, grid, cell 
Dim bExit, bAbort, txtStatus
Dim intRow : intRow = 4
Dim iCol

'1. Запрашиваем файл QTN и получаем массив значений для последующего заполнения SAP Quotation
Dim excelFile
excelFile = selectExcel()

'2. Заполняем открытый SAP Quotation
Dim ArticlesExcel, objWorkbook, pmu, TextSheet
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
objWorkbook.Sheets("PMU").Activate
Set pmu = objWorkbook.Worksheets("PMU")
Dim iLastRow: iLastRow = CInt(0)
iLastRow =pmu.Range("A" & pmu.Rows.Count).End(xlUp).Row  
'WScript.Echo iLastRow

transpose


'Подготовка заголовка
objWorkbook.Sheets("Text").Activate
Set TextSheet = objWorkbook.Worksheets("Text")
Dim spaces
Dim arrTexts(5, 1)
For intRow = 1 To 6		'делаем заголовки с пробелами до 18 символов
	spaces = ""
	if Len(TextSheet.Cells(intRow, 1).Value) < 18 Then
		spaces = Space(18 - Len(TextSheet.Cells(intRow, 1).Value))
	End if	
	arrTexts(intRow - 1, 0) = TextSheet.Cells(intRow, 1).Value & spaces
    'MsgBox arrTexts(intRow - 1, 0)
Next


Dim strText
'good session.findById("wnd[0]").sendVKey 2
'good session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08").select	'зашли в тексты позиции
intRow = 1
Do Until TextSheet.Cells(1, intRow + 1).Value = ""
	strText = ""
	For iCol = 1 To 6		'склеиваем заголовки со значениями
		strText = strText & arrTexts(iCol - 1, 0) & TextSheet.Cells(iCol, intRow + 1).Value & vbCrLf 
	Next	
	MsgBox strText
	' goof session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = strText
	if intRow < iLastRow then
		'good session.findById("wnd[0]/tbar[1]/btn[19]").press 'кнопка перехода по позициям
	end if	
    intRow = intRow + 1
Loop

objWorkbook.Close True
ArticlesExcel.Quit
MsgBox "Script finished! ", vbSystemModal Or vbInformation






'====== Functions ans Subs ========

Sub transpose()
  Dim sourceRange
  Dim targetRange

  Set sourceRange = pmu.Range(pmu.Cells(3, 3), pmu.Cells(iLastRow, 8))
  objWorkbook.Sheets("Text").Activate
  Set TextSheet = objWorkbook.Worksheets("Text")
  Set targetRange = TextSheet.Cells(1, 1)

  sourceRange.Copy
  targetRange.PasteSpecial xlPasteValues, xlNone, False, True
  pmu.Cells(3,3).Copy
End Sub

' Диалог выбора файла, создание потоков чтения из файла и записи в файл
Function selectExcel()
    Dim wShell, oExec, result

    Set wShell = CreateObject("WScript.Shell")
    Set oExec  = wShell.Exec("mshta.exe ""about:<input type=file id=FILE accept="".xl*""><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    result = oExec.StdOut.ReadLine
     
    If (result = "") Then  
        WScript.Quit 
    End If
    
    ' excelFile = 
    ' Set objExcel = CreateObject("Excel.Application")
    ' Set objWorkbook = objExcel.Workbooks.Open (excelFile)
    ' Возвращаем нашу книгу
    selectExcel = result ' Полный путь к выбранному файлу
    Set oExec = Nothing
    Set wShell = Nothing
    'MsgBox(result)

End Function