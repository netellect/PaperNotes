VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private tmpParent As NoteLibrary.Notes
Private tmpID As Long
Private fNote As frmNote
Private tmpChanged As Boolean
Private tmpOnTop As Boolean
Private tmpColor As Long
Private tmpVisible As Boolean
Private tmpPassword As String

'----------- CLASS FUNCTIONS ------------------
Private Sub Class_Initialize()
    If fNote Is Nothing Then
        Set fNote = New frmNote
        Set fNote.Parent = Me
    End If
    tmpPassword = ""
    tmpVisible = True
End Sub
Private Sub Class_Terminate()
    '<12/13/02>
    Set fNote.Parent = Nothing
    fNote.Hide
    '<//12/13/02>
    Set fNote = Nothing
End Sub

'----------- PUBLIC FUNCTIONS ------------------
Public Sub Show()
    On Error GoTo errhand
    If fNote Is Nothing Then
        Set fNote = New frmNote
        Set fNote.Parent = Me
    End If
    If tmpVisible Then
        '<new to 1.0.0.11 pnlib2>
        If tmpPassword <> "" Then
            Exit Sub
        End If
        '</new>
        
        fNote.Show
        '<2/14/03>
        If fNote.Left > Screen.Width Then
            fNote.Left = Screen.Width - fNote.Width
        End If
        If fNote.Top > Screen.Height Then
            fNote.Top = Screen.Height - fNote.Height
        End If
        If fNote.Left < 0 Then
            fNote.Left = 0
        End If
        If fNote.Top + fNote.lbTitleBar.Height < 0 Then
            fNote.Top = 0
        End If
        '<//2/14/03>
    End If
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

'<new to 1.0.0.11 pnlib2>
Public Sub Show2(Optional ByVal Password As String = "")
    If tmpVisible Then
        Call Show
        Exit Sub
    End If
    If Password = tmpPassword Then
        tmpVisible = True
        Call Show
    End If
End Sub
Public Sub SetPassword(ByVal Password As String)
    If Not fNote Is Nothing Then
        fNote.Hide
        tmpVisible = False
    End If
End Sub
Public Property Get IsPasswordSet() As Boolean
    If tmpPassword <> "" Then
        IsPasswordSet = True
    Else
        IsPasswordSet = False
    End If
End Property

'</new>

Public Sub Hide()
    If Not fNote Is Nothing Then
        fNote.Hide
        tmpVisible = False
    End If
End Sub

Public Property Get IsChanged() As Boolean
    IsChanged = tmpChanged
End Property

Public Sub Delete()
    tmpParent.DeleteNote tmpID
End Sub

Public Property Get Parent() As Notes
    Set Parent = tmpParent
End Property

Public Property Get NoteID() As Long
    NoteID = tmpID
End Property

Public Property Get OnTop() As Boolean
    OnTop = tmpOnTop
End Property
Public Property Let OnTop(ByVal OnTop As Boolean)
    tmpOnTop = OnTop
End Property

Public Property Get Color() As Long
    Color = fNote.BackColor
End Property
Public Property Let Color(ByVal vNewValue As Long)
    fNote.ChangeColor vNewValue
    'fNote.BackColor = vNewValue
    'fNote.t
End Property

'<NEW to 1.0.11 pnlib2>
Public Property Get TextColor() As Long
    TextColor = fNote.txtNote.ForeColor
End Property
Public Property Let TextColor(ByVal vNewValue As Long)
    fNote.ChangeTextColor vNewValue
End Property
'</NEW>

Public Property Get Left() As Long
    Left = fNote.Left
End Property
Public Property Let Left(ByVal vNewValue As Long)
    fNote.Left = vNewValue
End Property

Public Property Get Top() As Long
    Top = fNote.Top
End Property
Public Property Let Top(ByVal vNewValue As Long)
    fNote.Top = vNewValue
End Property

Public Property Get Height() As Long
    Height = fNote.Height
End Property
Public Property Let Height(ByVal vNewValue As Long)
    fNote.Height = vNewValue
End Property

Public Property Get Width() As Long
    Width = fNote.Width
End Property
Public Property Let Width(ByVal vNewValue As Long)
    fNote.Width = vNewValue
End Property

Public Property Get Text() As String
    Text = fNote.txtNote.Text
End Property
Public Property Let Text(vNewValue As String)
    fNote.txtNote.Text = vNewValue
End Property

Public Property Get Collapsed() As Boolean
    Collapsed = True
End Property

Public Property Let Collapsed(ByVal vNewValue As Boolean)
    fNote.Collapse
End Property

Public Property Get IsVisible() As Boolean
    If tmpVisible Then
        IsVisible = fNote.visible
    Else
        IsVisible = False
    End If
End Property

Public Property Get Title() As String
    Title = fNote.lbTitleBar.caption
End Property

Public Property Let Title(ByVal vNewValue As String)
    fNote.lbTitleBar.caption = vNewValue
End Property


'---------- FRIEND FUNCTIONS -------------------------
Friend Sub SetParent(vNewValue As Object)
    If Not vNewValue Is Nothing Then
        Set tmpParent = vNewValue
    End If
End Sub

Friend Sub SetChanged(Optional ch As Boolean = True)
    tmpChanged = ch
End Sub

Friend Sub SetID(ByVal ID As Long)
    tmpID = ID
End Sub

Friend Sub SetVisible(ByVal visible As Boolean)
    On Error GoTo errhand
    fNote.visible = visible
    tmpVisible = visible
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

Friend Sub SetTransparencyLevel(ByVal level As Long)
    fNote.SetTransparencyLevelAndThroughClicks level, tmpParent.ThroughClicks
End Sub

Friend Sub AllowThroughClicks(ByVal allow As Boolean)
    fNote.SetTransparencyLevelAndThroughClicks tmpParent.TransparencyLevel, allow
End Sub


Friend Sub SetFontProperties(ByVal fp As String)
    If fp = "" Then
        Exit Sub
    End If
    If fNote Is Nothing Then
        Exit Sub
    End If
    fNote.SetFontProps fp
End Sub

Friend Function GetFontProperties() As String
    Dim tmpstr As String
    tmpstr = fNote.GetFontProps
    GetFontProperties = tmpstr
End Function

'<new to 1.0.0.11 pnlib2>
Friend Property Get Password() As String
    Password = tmpPassword
End Property
Friend Property Let Password(ByVal vNewValue As String)
    tmpPassword = vNewValue
End Property
'</new>
