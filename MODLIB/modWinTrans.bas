Attribute VB_Name = "modWinTrans"
Option Explicit

'==================================================================================
' Usage :
'----------------------------------------------------------------------------------
' 1.    Call modWinTrans.twInitialize(Me.hWnd)
'
' 2.    Call modWinTrans.twSetTransparencyLevel(Me.hWnd, nValue) ' nValue = 0..255
'
' 3.    If you want to make the window transparent for mouse clicks also:
'       Call modWinTrans.twAllowThroughClicks(Me.hWnd, True)     ' or False
'----------------------------------------------------------------------------------
'       (c) 2000-2003 by Vlad Kozin
'==================================================================================


' make sure this constant is defined/not defined before compilation.
#Const WINDOWS2000 = 1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

#If WINDOWS2000 Then
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Public lngWindowOrigStyle As Long
Public hWndOrig As Long

Public Function twInitialize(hwnd As Long) As Long
    Err.Clear
    If hwnd = 0 Then GoTo errhand
    On Error GoTo errhand
    #If WINDOWS2000 Then
    hWndOrig = hwnd
    lngWindowOrigStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, lngWindowOrigStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
    #End If
errhand:
    twInitialize = Err.Number
End Function

Public Function twSetTransparencyLevel(hwnd As Long, nValue As Long) As Long
    Err.Clear
    If hwnd = 0 Then GoTo errhand
    If hwnd <> hWndOrig Then GoTo errhand
    On Error GoTo errhand
    #If WINDOWS2000 Then
    If (nValue >= 0) And (nValue <= 255) Then
        SetLayeredWindowAttributes hwnd, 0, nValue, LWA_ALPHA
    Else
        twSetTransparencyLevel = -1
        Exit Function
    End If
    #End If
errhand:
    twSetTransparencyLevel = Err.Number
End Function

Public Function twAllowThroughClicks(hwnd As Long, bAllow As Boolean) As Long
    Err.Clear
    If hwnd = 0 Then GoTo errhand
    If hwnd <> hWndOrig Then GoTo errhand
    On Error GoTo errhand
    #If WINDOWS2000 Then
    If bAllow Then
        SetWindowLong hwnd, GWL_EXSTYLE, lngWindowOrigStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    Else
        SetWindowLong hwnd, GWL_EXSTYLE, 0 Or WS_EX_LAYERED 'lngWindowOrigStyle Or Not WS_EX_TRANSPARENT
    End If
    #End If
errhand:
    twAllowThroughClicks = Err.Number
End Function
