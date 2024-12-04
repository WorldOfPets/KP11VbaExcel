Public PyEXE As String, PyScriptFolder As String
' Установка пути к python и папке с проектом
Function SetPath()
    PyEXE = """C:\Users\Evgen\AppData\Local\Programs\Python\Python311\python.exe"""
    PyScriptFolder = "C:\Users\Evgen\Work\programs\kp11prod\"
End Function

' Поиск накладок по кабинету и преподавателю
Sub ПоискНакладок()
    SetPath
    Dim objShell As Object
    Dim PythonScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    PythonScript = PyScriptFolder & "mainCheckNacladka.py"

    ThisWorkbook.Save
    objShell.Run "cmd /K " & PyEXE & " " & PythonScript & " " & ThisWorkbook.FullName
    Application.OnTime Now + TimeValue("00:00:05"), "ReOpen"
    ThisWorkbook.Close (False)
    
End Sub

' Поиск свободного учителя по строке
Sub РасписаниеУчителя()
    SetPath
    Dim objShell As Object
    Dim PythonScript As String
    Dim row As Long, col As Long
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonScript = PyScriptFolder & "mainScheTeacher.py"
    
    ThisWorkbook.Save
    objShell.Run "cmd /K " & PyEXE & " " & PythonScript & " " & ThisWorkbook.FullName & " " & row & " " & col
    Application.OnTime Now + TimeValue("00:01:00"), "ReOpen"
    ThisWorkbook.Close (False)
End Sub

' Расписание свободного кабинета
Sub РасписаниеКабинета()
    SetPath
    Dim objShell As Object
    Dim PythonScript As String
    Dim row As Long, col As Long
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonScript = PyScriptFolder & "mainScheCabinet.py"
    
    ThisWorkbook.Save
    objShell.Run "cmd /K " & PyEXE & " " & PythonScript & " " & ThisWorkbook.FullName & " " & row & " " & col
    Application.OnTime Now + TimeValue("00:00:30"), "ReOpen"
    ThisWorkbook.Close (False)
End Sub

' Поиск свободного кабиента по строке
Sub ПоискКабинета()
    SetPath
    Dim objShell As Object
    Dim PythonScript As String
    Dim row As Long
    row = ActiveCell.row
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonScript = PyScriptFolder & "mainFindCabinet.py"
    
    objShell.Run "cmd /K " & PyEXE & " " & PythonScript & " " & ThisWorkbook.FullName & " " & row
    
End Sub

' Поиск ошибок в файле (кабинет, преподаватель, предмет)
Sub ПоискОшибок()
    SetPath
    Dim objShell As Object
    Dim PythonScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonScript = PyScriptFolder & "checkError.py"
    
    ThisWorkbook.Save
    objShell.Run "cmd /K " & PyEXE & " " & PythonScript & " " & ThisWorkbook.FullName
    Application.OnTime Now + TimeValue("00:00:05"), "ReOpen"
    ThisWorkbook.Close (False)
    
End Sub

Private Sub ReOpen()
    Application.Workbooks.Open (ThisWorkbook.FullName)
End Sub

