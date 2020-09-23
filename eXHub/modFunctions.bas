Attribute VB_Name = "modFunctions"
Option Explicit

Public Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uId              As Long
    uFlags           As Long
    uCallBackMessage As Long
    hIcon            As Long
    szTip            As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public nID As NOTIFYICONDATA

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Sub AddSystray(myForm As Form, myTip As String)
    With nID
        .cbSize = Len(nID)
        .hWnd = myForm.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE ' Should call back on click?
        .hIcon = myForm.Icon
        .szTip = myTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nID
End Sub

Public Sub ModifySystray(myForm As Form, myTip As String)
    With nID
        .cbSize = Len(nID)
        .hWnd = myForm.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE ' Should call back on click?
        .hIcon = myForm.Icon
        .szTip = myTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_MODIFY, nID
End Sub

Public Sub RemoveSystray()
    Shell_NotifyIcon NIM_DELETE, nID
End Sub

Function ReadINI(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function IsIDE() As Boolean
    On Error GoTo errHandle
    Debug.Print 1 / 0
    IsIDE = False
    Exit Function
errHandle:
    IsIDE = True
End Function

Function Seconds2Time(Sec As Long) As String
    If Sec <= 0 Or Sec > 2147483647 Then Exit Function
    Seconds2Time = Sec \ 3600 & ":" & Format((Sec Mod 3600) \ 60, "00") & ":" & Format((Sec Mod 3600) Mod 60, "00")
End Function
