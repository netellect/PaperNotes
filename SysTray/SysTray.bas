Attribute VB_Name = "modSysTray"
Option Explicit

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'<12/27/02>
Declare Function DllRegisterMyServer Lib "PNLib2.dll" Alias "DllRegisterServer" () As Long
'<//12/27/02>



Public Const WM_MOUSEMOVE = &H200
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64

Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type

Public nfIconData As NOTIFYICONDATA

