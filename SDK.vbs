Dim WshShell, SapGuiAuto, application, session, Wnd0, Menubar, UserArea, Statusbar, UserName 

' Создаем объект WScript.Shell
Set WshShell = WScript.CreateObject("WScript.Shell")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Подключение к SAP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Создаем объект
Set SapGuiAuto = GetObject("SAPGUI")

' Создаем объект типа GuiApplication (COM-интерфейс)
Set application = SapGuiAuto.GetScriptingEngine()

' Создаем объект типа GuiSession - это сессия, которой соответствует активное окно SAP
' Т.е. при запуске WSF сам скрипт будет выполняться в том же окне SAP, из которого запущен
Set session = application.ActiveSession()

WScript.ConnectObject session,     "on"
WScript.ConnectObject application, "on"

' Создаем объект типа GuiMainWindow
Set Wnd0 = session.findById("wnd[0]")

' Создаем объект типа GuiMenubar
Set Menubar = Wnd0.findById("mbar")

' Создаем объект типа GuiUserArea
Set UserArea = Wnd0.findById("usr")

' Создаем объект типа GuiStatusbar
Set Statusbar = Wnd0.findById("sbar")

' Определяем логин пользователя
UserName = session.Info.User

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Вспомогательные процедуры и функции
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Запуск транзакции
Sub startTransaction(transaction_name)
    session.StartTransaction transaction_name
End Sub

' Нажатие кнопки "Enter"
Sub pressEnter()
    Wnd0.sendVKey 0
End Sub

' Нажатие кнопки F2
Sub pressF2()
    Wnd0.sendVKey 2
End Sub

' Нажатие кнопки F3
Sub pressF3()
    Wnd0.sendVKey 3
End Sub

' Нажатие кнопки F5
Sub pressF5()
    Wnd0.sendVKey 5
End Sub

' Нажатие кнопки F8
Sub pressF8()
    Wnd0.sendVKey 8
End Sub

' Диалог выбора файла, создание потоков чтения из файла и записи в файл
Function selectFile(createOuputFile)
    Set objDialog = CreateObject("UserAccounts.CommonDialog")
    ' Заполняем свойства и открываем диалог
    With objDialog
        .InitialDir = WshShell.SpecialFolders("Desktop") ' Начальная папка - рабочий стол
        .Filter = "Текстовые файлы (*.csv;*.txt)|*.csv;*.txt"
        result = .ShowOpen
    End With
    ' Если файл не выбран - выходим
    If (result = 0) Then WScript.Quit
    inputFile = objDialog.FileName ' Полный путь к выбранному файлу
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputStream = fso.OpenTextFile(inputFile)
    ' Создавать выходной файл?
    If (createOuputFile) Then
        outputFile = Left(inputFile, Len(inputFile) - 3) & "out" & Right(inputFile, 4)
        Set outputStream = fso.CreateTextFile(outputFile, True)
        ' Возвращаем массив из потока чтения из файла и потока записи в файл
        selectFile = Array(inputStream, outputStream)
    Else
        ' Возвращаем поток чтения из файла
        Set selectFile = inputStream
    End If
End Function

' Заполняем одну строку в таблице (для ME51N)
Sub fill_row(row, material, kolvo, zavod, zatreboval)
    Set grid = session.findById(UserArea.findByName("GRIDCONTROL", "GuiCustomControl").Id & "/shellcont/shell")
    grid.modifyCell row, "KNTTP", "K"        ' Тип контировки
    grid.modifyCell row, "MATNR", material   ' Материал
    grid.modifyCell row, "MENGE", kolvo      ' Количество
    grid.modifyCell row, "NAME1", zavod      ' Завод
    grid.modifyCell row, "LGOBE", "0001"     ' Склад
    grid.modifyCell row, "AFNAM", zatreboval ' Затребовал
End Sub

' Set, get and clear ClipBoard text in VBScript
'   CLEAR - QuickClip("")
'   SET - QuickClip("Hello World!")
'   GET - Result = QuickClip(Null)
Function QuickClip(input)
    '@description: A quick way to set and get your clipboard.
    '@author: Jeremy England (SimplyCoded)
    If IsNull(input) Then
        QuickClip = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
        If IsNull(QuickClip) Then QuickClip = ""
    Else
        CreateObject("WScript.Shell").Run "mshta.exe javascript:eval(""document.parentWindow.clipboardData.setData('text','" & Replace(Replace(input, "'", "\\u0027"), """", "\\u0022") & "');window.close()"")", 0, True
    End If
End Function

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