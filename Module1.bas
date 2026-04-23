Attribute VB_Name = "Module1"
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Dim ie As Object

Public Sub load_data(Optional this_month As Boolean = False)
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    
    ' 單一登入
    attempts = 0
    Dim hWnd As Long
    Dim numLockState As Boolean
    numLockState = CBool(GetKeyState(144) And 1)
    Do While attempts < 5
        ie.navigate "http://sso.********.com.tw/"
        Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
        ie.Document.getElementsByName("userid")(0).Focus
        hWnd = FindWindow("IEFrame", vbNullString)
        If hWnd <> 0 Then SetForegroundWindow hWnd
        Application.SendKeys "{DOWN}", True
        Application.Wait (Now + TimeValue("0:00:1"))
        hWnd = FindWindow("IEFrame", vbNullString)
        If hWnd <> 0 Then SetForegroundWindow hWnd
        Application.SendKeys "{DOWN}", True
        Application.Wait (Now + TimeValue("0:00:1"))
        Application.SendKeys "^~", True
        Application.Wait (Now + TimeValue("0:00:1"))
        ie.Document.parentWindow.execScript ("goSubmit();")
        Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
        If InStr(1, ie.LocationURL, "myportal", vbTextCompare) > 0 Then
            Exit Do
        End If
        attempts = attempts + 1
    Loop
    If numLockState Then
        keybd_event 144, 0, 0, 0
        keybd_event 144, 0, 2, 0
    End If

    ' 人事行政管理系統
    Dim elem As Object
    Dim startTime As Double
    ie.navigate "http://sso.********.com.tw/"
    startTime = Timer
    With CreateObject("Shell.Application")
        Do While Timer - startTime < 10
            For Each elem In .Windows
                If InStr(1, elem.FullName, "iexplore.exe", vbTextCompare) > 0 Then
                    Do While elem.Busy Or elem.readyState <> 4: DoEvents: Loop
                    If InStr(1, elem.LocationURL, "EI0100MainClassX", vbTextCompare) > 0 Then
                        Set ie = elem
                        Exit Do
                    End If
                End If
            Next elem
        Loop
    End With
    
    ' 差假管理系統
    Dim allLinks As Object
    Dim link As Object
    Set allLinks = ie.Document.getElementsByName("EItop")(0).contentWindow.Document.all
    For Each link In allLinks
        If link.innerHTML = "差假管理" Then
            link.Click
            Exit For
        End If
    Next link
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
    
    ' 假單查詢
    ie.Document.getElementsByName("top")(0).contentWindow.Document.getElementById("Head7").Click
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
        
    ' 已登錄假單
    With ie.Document.getElementsByName("bottom")(0).contentWindow.Document
        .getElementsByName("frmTools")(0).contentWindow.Document.getElementById("menu2").Click
    End With
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
    With ie.Document.getElementsByName("bottom")(0).contentWindow.Document
        With .getElementsByName("frmContent")(0).contentWindow.Document
            If this_month Then
                .getElementsByName("START_YY")(0).Value = Year(Date) - 1911
                .getElementsByName("END_YY")(0).Value = Year(Date) - 1911
                .getElementsByName("START_MM")(0).Value = Month(Date)
                .getElementsByName("END_MM")(0).Value = Month(Date)
            Else ' last month
                .getElementsByName("START_YY")(0).Value = Year(DateAdd("m", -1, Date)) - 1911
                .getElementsByName("END_YY")(0).Value = Year(DateAdd("m", -1, Date)) - 1911
                .getElementsByName("START_MM")(0).Value = Month(DateAdd("m", -1, Date))
                .getElementsByName("END_MM")(0).Value = Month(DateAdd("m", -1, Date))
            End If
            .forms(0).submit
        End With
    End With
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
    With ie.Document.getElementsByName("bottom")(0).contentWindow.Document
        With .getElementsByName("frmContent")(0).contentWindow.Document
            With .getElementsByName("bottom")(0).contentWindow.Document
                .execCommand "SelectAll"
                .execCommand "Copy"
                Application.Wait (Now + TimeValue("0:00:1"))
                .execCommand "SelectAll"
                .execCommand "Copy"
            End With
        End With
    End With

    ' 貼上資料
    Worksheets("差假資料").Range("A2:G51").Clear
    Worksheets("差假資料").Range("A2").PasteSpecial

    With CreateObject("Shell.Application")
        For Each elem In .Windows
            If InStr(1, elem.FullName, "iexplore.exe", vbTextCompare) > 0 Then
                elem.Quit
            End If
        Next elem
    End With
    
    Set link = Nothing
    Set allLinks = Nothing
    Set ie = Nothing
    Set elem = Nothing
End Sub

Public Sub load_data_this_month()
    load_data True
End Sub

Public Sub load_data_last_month()
    load_data
End Sub
