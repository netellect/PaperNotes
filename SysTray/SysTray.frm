VERSION 5.00
Begin VB.Form frmIcon 
   ClientHeight    =   930
   ClientLeft      =   1440
   ClientTop       =   2025
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SysTray.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   930
   ScaleWidth      =   2835
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuPopup 
      Caption         =   "SysTray Popup Menu"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
         Begin VB.Menu mnuShowAll 
            Caption         =   "Show All"
         End
         Begin VB.Menu mnuHideAll 
            Caption         =   "Hide All"
         End
         Begin VB.Menu mnuDefaultSize 
            Caption         =   "Default Size All"
         End
         Begin VB.Menu mnuShowBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuShowNote 
            Caption         =   "<untitled>"
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSaveAll 
         Caption         =   "Save All"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NOTEDELIM = "-----------" & vbCrLf

Private Const DEFSAVEASEXTENSION = "txt"


Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long


Private aNote As NoteLibrary.Note
Private allNotes As New NoteLibrary.Notes

Private Sub Form_Initialize()
    Dim anObj As Object
    Dim res As Long
    '<2/14/03>
    If App.PrevInstance Then
        MsgBox "Another copy of " & App.Title & " is running", vbCritical, App.Title
        End
    End If
    '<2/14/03>

    '<12/27/02>
    On Error Resume Next
    Set anObj = CreateObject("NoteLibrary.Notes")
    If Not anObj Is Nothing Then
        GoTo ExitProc
    End If
    Err.Clear
    res = DllRegisterMyServer
    Set anObj = CreateObject("NoteLibrary.Notes")
    If Err.Number <> 0 Then
        If Not anObj Is Nothing Then
            MsgBox "Installation complete. Click OK to proceed", vbInformation, App.Title
        Else
            MsgBox "Error registering PNLib2.DLL. Please throw program away, or bother to download the complete package ;)", vbCritical, App.Title
            End
        End If
    End If
ExitProc:
    Set anObj = Nothing
    Err.Clear
    '<//12/27/02>
End Sub

Private Sub Form_Load()
    
    With nfIconData
        .hWnd = Me.hWnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.handle
        .szTip = App.Title & Chr$(0)
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
    
    allNotes.LoadAll
    'allNotes.ShowAll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
        Case 7680               'MouseMove
        Case 7695               'LeftMouseDown
            '<new to 1.0.8>
            Call BringOnTop
            ''Call mnuShowAll_Click
            '<//new to 1.0.8>
        Case 7710               'LeftMouseUp
        Case 7725, 6180         'LeftDblClick
            Call mnuNew_Click
        Case 7740, 6192         'RightMouseDown
            Call SetForegroundWindow(Me.hWnd)
            Call BuildNoteList
            PopupMenu mnuPopup, 0, , , mnuNew
        Case 7755               'RightMouseUp
        Case 7770               'RightDblClick
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allNotes.IsChanged Then
        allNotes.SaveAll
    End If
    Set allNotes = Nothing
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
    End
End Sub

Private Sub mnuAbout_Click()
    MsgBox App.ProductName & " V 1.0." & App.Revision & vbCrLf & vbCrLf & _
    App.LegalCopyright, vbInformation, App.ProductName
End Sub

Private Sub mnuDefaultSize_Click()
    Dim a
    For Each a In allNotes.Notes
        a.Height = 1300
        a.Width = 2500
    Next
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHideAll_Click()
    allNotes.HideAll
End Sub

Private Sub mnuNew_Click()
    Set aNote = allNotes.AddNote
    '<12/13/02>
    Dim k As Long
    Dim x As Long, y As Long
    If Not aNote Is Nothing Then
        x = aNote.Left
        y = aNote.Top
        For k = 1 To allNotes.Notes.Count - 1
            If allNotes.Notes(k).Left = x Then
                aNote.Left = x + 80
                x = aNote.Left
            End If
            If allNotes.Notes(k).Top = y Then
                aNote.Top = y + 80
                y = aNote.Top
            End If
        Next
        '<//12/13/02>
        aNote.Show
    End If
End Sub

Private Sub mnuOptions_Click()
    Dim fOptions As frmOptions
    Set fOptions = New frmOptions
    With fOptions
        .cbClickThrough.Value = IIf(allNotes.ThroughClicks, 1, 0)
        .Slider1.Value = allNotes.TransparencyLevel
        .Left = Screen.Width - Me.Width - 1000
        .Top = Screen.Height - Me.Height - 500
        .Show vbModal
        If .OkPressed Then
            allNotes.ThroughClicks = (.cbClickThrough.Value <> 0)
            allNotes.TransparencyLevel = .Slider1.Value
        End If
    End With
    Set fOptions = Nothing
End Sub

Private Sub mnuSaveAll_Click()
    allNotes.SaveAll
End Sub

Private Sub mnuSaveAs_Click()
    Dim fn As String
    Dim fvar As Long
    Dim aNote As NoteLibrary.Note
    Dim tmpStr As String
    On Error GoTo ErrHand
    fn = ""
    
    '<new>
    Dim res As Long
    Dim sFileOpen As OPENFILENAME
    Dim myFile As String * 1024
    With sFileOpen
        .hwndOwner = Me.hWnd
        .nMaxFile = 1024
        .lStructSize = Len(sFileOpen)
        .lpstrFile = myFile
        .lpstrDefExt = DEFSAVEASEXTENSION
        .lpstrFilter = "Text Files (*.txt)"
        res = GetSaveFileName(sFileOpen)
        If res = 1 Then
            fn = .lpstrFile
        Else
            fn = ""
        End If
    End With
    '</new>
        
    If fn <> "" Then
        fn = StripNulls(fn)
        If LCase(Right(fn, 4)) <> ".txt" Then
            fn = fn & ".txt"
        End If
        fvar = FreeFile
        Open fn For Output As #fvar
        For Each aNote In allNotes.Notes
            Print #fvar, vbCrLf & "Note #" & aNote.NoteID
            Print #fvar, vbCrLf & "Title" & aNote.Title
            Print #fvar, NOTEDELIM
            Print #fvar, aNote.Text
        Next
        Close fvar
        MsgBox "Successfully saved to " & fn, vbInformation, App.Title
    End If
ErrHand:
    Err.Clear
End Sub

Private Sub mnuShowAll_Click()
    allNotes.ShowAll
End Sub

Private Function StripNulls(strItem As String) As String
    Dim nPos As Integer
    
    nPos = InStr(strItem, Chr$(0))
    If nPos Then
        strItem = Left$(strItem, nPos - 1)
    End If
    StripNulls = strItem
End Function

'<new to 1.0.8>
Private Function BringOnTop() As Long
    Dim k As Long
    Dim aNote As Note
    For k = allNotes.Notes.Count To 1 Step -1
        Set aNote = allNotes.Notes(k)
        If aNote.IsVisible Then
            aNote.Show
        End If
    Next
End Function

Private Function BuildNoteList() As Long
    Dim k As Long
    Dim aNote As Note
    Dim nNotes As Long
    Dim nNoteIndex As Long
    
    nNotes = allNotes.Notes.Count
    On Error Resume Next
    For k = mnuShowNote.UBound To 1 Step -1
        Unload mnuShowNote(k)
    Next
    Err.Clear

    For k = nNotes To 1 Step -1
        Set aNote = allNotes.Notes(k)
        If Not aNote.IsVisible Then
            nNoteIndex = mnuShowNote.Count
            Load mnuShowNote(nNoteIndex)
            mnuShowNote(nNoteIndex).Caption = IIf(aNote.Title = "", "<untitled>", aNote.Title)
            mnuShowNote(nNoteIndex).Tag = aNote.NoteID
            mnuShowNote(nNoteIndex).Visible = True
        End If
    Next
    HideBar
End Function

Private Sub mnuShowNote_Click(Index As Integer)
    Dim aNote As Note
    Dim tmpPassword As String
    Set aNote = allNotes.Notes(CStr(mnuShowNote(Index).Tag))
    If Not aNote Is Nothing Then
        If aNote.IsPasswordSet Then
            tmpPassword = InputBox("Note Password:", "Please enter password", "")
            If tmpPassword = "" Then
                Exit Sub
            End If
        End If
        aNote.Show2 tmpPassword
        On Error Resume Next
        Unload mnuShowNote(Index)
        HideBar
        Err.Clear
    End If
End Sub
Private Sub HideBar()
    mnuShowNote(0).Visible = False
    If mnuShowNote.UBound < 1 Then
        mnuShowBar0.Visible = False
    Else
        mnuShowBar0.Visible = True
    End If
End Sub
'</new>

