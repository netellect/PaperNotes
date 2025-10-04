VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colNotes As Collection
Private Type udfNote
    caption As String
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    Color As Long
    Text As String
    visible As Boolean
    TextColor As Long
End Type
Private tmpAllowThroughClicks As Boolean
Private tmpTransparencyLevel As Long

Private tmpDefaultNoteStyle As String

Private tmpIsLoading As Boolean


'---------- PUBLIC -----------

Public Sub CollapseAll()
    Dim aNote As Note
    For Each aNote In colNotes
        aNote.Collapsed = True
    Next
End Sub

Public Sub ShowAll()
    Dim aNote As Note
    For Each aNote In colNotes
        aNote.SetVisible True
        aNote.Show
        aNote.AllowThroughClicks tmpAllowThroughClicks
        aNote.SetTransparencyLevel tmpTransparencyLevel
    Next
End Sub

Public Sub HideAll()
    Dim aNote As Note
    For Each aNote In colNotes
        aNote.Hide
    Next
End Sub

Public Function AddNote() As Note
    Dim aNote As Note
    '<2/18/03>
    Dim tmpID As Long
    '<//2/18/03>
    Set aNote = New Note
    aNote.SetParent Me
    ' DEBUG, Unrem On Error GoTo errhand
    On Error GoTo errhand
    '<2/18/03>
    tmpID = CLng(GetTempID) + colNotes.Count
    '<//2/18/03>
    On Error GoTo errhand
    colNotes.Add aNote, CStr(tmpID) 'CStr(colNotes.Count + 1)
    aNote.SetID tmpID 'colNotes.Count
    If Not tmpIsLoading Then
        aNote.SetFontProperties tmpDefaultNoteStyle
    End If
    Set AddNote = aNote
    
errhand:
    If Err Then
        MsgBox Err.Description, vbExclamation, App.ProductName
        Err.Clear
    End If
End Function

Public Function DeleteNote(ID As Long) As Long
    On Error GoTo errhand
    colNotes.Remove CStr(ID)
errhand:
    If Err Then
        App.LogEvent App.Title & " error: " & Err.Description, vbLogEventTypeWarning
        Err.Clear
    End If
End Function

Public Property Get Notes() As Collection
    Set Notes = colNotes
End Property

Public Function SaveAll() As Long
    Dim k As Long
    Dim aNote As Note
    On Error GoTo errhand
    SaveSetting App.Title, "General", "defaultstyle", tmpDefaultNoteStyle
    SaveSetting App.Title, "General", "notes", colNotes.Count
    SaveSetting App.Title, "General", "level", tmpTransparencyLevel
    SaveSetting App.Title, "General", "clickthrough", tmpAllowThroughClicks
    For k = 1 To colNotes.Count
        SaveSetting App.Title, "note" & CStr(k), "title", colNotes(k).Title
        SaveSetting App.Title, "note" & CStr(k), "left", colNotes(k).Left
        SaveSetting App.Title, "note" & CStr(k), "top", colNotes(k).Top
        SaveSetting App.Title, "note" & CStr(k), "width", colNotes(k).Width
        SaveSetting App.Title, "note" & CStr(k), "height", colNotes(k).Height
        SaveSetting App.Title, "note" & CStr(k), "color", colNotes(k).Color
        '<new to 1.0.11 pnlib2>
        SaveSetting App.Title, "note" & CStr(k), "textcolor", colNotes(k).TextColor
        '</new>
        
        SaveSetting App.Title, "note" & CStr(k), "text", colNotes(k).Text
        SaveSetting App.Title, "note" & CStr(k), "visible", CInt(colNotes(k).IsVisible)
        Set aNote = colNotes(k)
        
        SaveSetting App.Title, "note" & CStr(k), "fontproperties", aNote.GetFontProperties
    Next
    SaveAll = 0
errhand:
    If Err Then
        SaveAll = Err.Number
        Err.Clear
    End If
End Function

Public Sub SetOnTop(Optional OnTop As Boolean = True)
    Dim aNote As Note
    For Each aNote In colNotes
        aNote.OnTop = OnTop
    Next
End Sub

Public Sub LoadAll()
    Dim aNote As Note
    Dim k As Long
    Dim totalNotes As Long
    Dim tmpNote As udfNote
    Dim tmpstr As String
    tmpIsLoading = True
    tmpDefaultNoteStyle = GetSetting(App.Title, "General", "defaultstyle", "")
    tmpTransparencyLevel = GetSetting(App.Title, "General", "level", 255)
    tmpAllowThroughClicks = GetSetting(App.Title, "General", "clickthrough", False)
    totalNotes = CLng(GetSetting(App.Title, "General", "notes", 0))
    For k = 1 To totalNotes
        With tmpNote
            .Left = CLng(GetSetting(App.Title, "note" & CStr(k), "left", 0))
            .Top = CLng(GetSetting(App.Title, "note" & CStr(k), "top", 0))
            .Width = CLng(GetSetting(App.Title, "note" & CStr(k), "width", 2500))
            .Height = CLng(GetSetting(App.Title, "note" & CStr(k), "height", 1300))
            .Color = CLng(GetSetting(App.Title, "note" & CStr(k), "color", vbButtonFace))
            '<new to 1.0.11 pnlib2>
            .TextColor = CLng(GetSetting(App.Title, "note" & CStr(k), "textcolor", vbBlack))
            '</new>
            .Text = GetSetting(App.Title, "note" & CStr(k), "text")
            tmpstr = GetSetting(App.Title, "note" & CStr(k), "visible")
            .visible = CBool(IIf((tmpstr = "") Or (tmpstr = "0"), "0", "1"))
            Set aNote = AddNote
            aNote.Title = GetSetting(App.Title, "note" & CStr(k), "title", "")
            aNote.Left = .Left
            aNote.Top = .Top
            aNote.Width = .Width
            aNote.Height = .Height
            aNote.Color = .Color
            aNote.TextColor = .TextColor
            aNote.Text = .Text
            aNote.SetVisible .visible
            aNote.AllowThroughClicks tmpAllowThroughClicks
            aNote.SetTransparencyLevel tmpTransparencyLevel
            tmpstr = GetSetting(App.Title, "note" & CStr(k), "fontproperties", "")
            
            aNote.SetFontProperties tmpstr
        End With
    Next
    tmpIsLoading = False
End Sub

Public Property Get IsChanged() As Boolean
    Dim k As Long
    IsChanged = False
    For k = 1 To colNotes.Count
        If colNotes(k).IsChanged Then
            IsChanged = True
            Exit Property
        End If
    Next
End Property


'---------- PRIVATE -------

Private Sub Class_Initialize()
    Dim k As Long
    
    Set colNotes = New Collection
    tmpAllowThroughClicks = False
    tmpTransparencyLevel = 255
    For k = 0 To UBound(aCustomColors)
        aCustomColors(k) = 0
    Next
    
    tmpDefaultNoteStyle = -1
    
    
End Sub

Private Sub Class_Terminate()
    Set colNotes = Nothing
End Sub

Public Property Get ThroughClicks() As Boolean
    ThroughClicks = tmpAllowThroughClicks
End Property

Public Property Let ThroughClicks(ByVal vNewValue As Boolean)
    Dim aNote As Note
    tmpAllowThroughClicks = vNewValue
    For Each aNote In colNotes
        aNote.AllowThroughClicks tmpAllowThroughClicks
    Next
End Property

Public Property Get TransparencyLevel() As Long
    TransparencyLevel = tmpTransparencyLevel
End Property

Public Property Let TransparencyLevel(ByVal vNewValue As Long)
    Dim aNote As Note
    tmpTransparencyLevel = vNewValue
    For Each aNote In colNotes
        aNote.SetTransparencyLevel tmpTransparencyLevel
    Next
End Property

Private Function GetTempID() As String
    'GetTempID = VBA.Right(CStr(Year(VBA.Date)), 2) &
    GetTempID = CStr(Month(VBA.Date)) & CStr(Day(VBA.Date)) & CStr(Hour(VBA.Time)) & CStr(Minute(VBA.Time)) & CStr(Second(VBA.Time))
End Function

Friend Function DefaultNoteStyle(Optional ByVal DefaultStyle As String = "") As String
    If DefaultStyle <> "" Then
        tmpDefaultNoteStyle = DefaultStyle
    End If
    DefaultNoteStyle = tmpDefaultNoteStyle
End Function
