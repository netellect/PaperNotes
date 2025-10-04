VERSION 5.00
Begin VB.Form frmNote 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   3870
   ClientTop       =   1830
   ClientWidth     =   2280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmNote.frx":0000
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   50
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   855
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   855
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   50
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   855
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   50
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   2175
      MouseIcon       =   "frmNote.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Hide"
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lbTitleBar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   1
      ToolTipText     =   "Right Click for Menu"
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuColor 
         Caption         =   "&Color"
         Begin VB.Menu mnuYellow 
            Caption         =   "&Yellow"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuButtonface 
            Caption         =   "Button&face"
         End
         Begin VB.Menu mnuBackground 
            Caption         =   "B&ackground"
         End
         Begin VB.Menu mnuColorBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuColorCustom 
            Caption         =   "&Custom..."
         End
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnufont 
         Caption         =   "&Font"
         Begin VB.Menu mnuFontCustom 
            Caption         =   "&Custom..."
         End
         Begin VB.Menu mnuFontBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSetAsDefault 
            Caption         =   "&Set as Default Style"
         End
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTitle 
         Caption         =   "&Title..."
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TITLEHEIGHT = 150
Private Const DEFNOTEHEIGHT = 1300
Private Const DEFNOTEWIDTH = 2500
Private Const STHRESHOLD = 150

Private lf As LOGFONT

Private mbMoving As Boolean
Private mbSmall As Boolean
Private oldHeight As Long
Private startX As Long
Private startY As Long

'<sides>
Private Enum ENUM_SIDE
    E_NOTHING
    E_LEFT
    E_RIGHT
    E_BOTTOM
End Enum
Private eSide As ENUM_SIDE
'</sides>

Public Parent As NoteLibrary.Note

Public Sub SetTransparencyLevelAndThroughClicks(level As Long, allow As Boolean)
    '<transp>
    On Error GoTo errhand
    If Me.visible Then
        SetForegroundWindow Me.hwnd
        Call modWinTrans.twInitialize(Me.hwnd)
        'Call modWinTrans.twAllowThroughClicks(Me.hwnd, allow)
        Call modWinTrans.twSetTransparencyLevel(Me.hwnd, level)
        Call modWinTrans.twAllowThroughClicks(Me.hwnd, allow)
    End If
    '</transp>
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

Private Sub Form_Activate()
    Me.BackColor = lbTitleBar.BackColor '= Me.BackColor
    lblClose.BackColor = Me.BackColor
    Picture2.BackColor = Me.BackColor
    Picture3.BackColor = Me.BackColor
    Picture4.BackColor = Me.BackColor
    '<transp>
    'Call modWinTrans.twSetTransparencyLevel(Me.hWnd, 50)
    'Debug.Print Me.hWnd
    '</transp>
End Sub

Private Sub Form_Paint()
    mbMoving = False
End Sub

Private Sub lblClose_Click()
    Parent.Hide
End Sub

Private Sub Form_Load()
    On Error GoTo errhand
    Me.Width = DEFNOTEWIDTH
    Me.Height = DEFNOTEHEIGHT
    Parent.SetChanged False
    mbMoving = False
    Call modWinTrans.twInitialize(Me.hwnd)
    txtNote.Left = Picture2.Width
    ResizeHeight
    ResizeWidth
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

Private Sub Form_Resize()
    With lblClose
        .Width = TITLEHEIGHT
        .Height = TITLEHEIGHT - 8
        .Top = 0
        .Left = Me.ScaleWidth - .Width
    End With
    With lbTitleBar
        .Left = 0
        .Top = 0
        .Height = TITLEHEIGHT
        .Width = Me.ScaleWidth - lblClose.Width
    End With
    ResizeHeight
    ResizeWidth
End Sub

Private Sub lbTitleBar_DblClick()
    Collapse
End Sub

Private Sub lbTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    startX = X
    startY = Y
    mbMoving = True
End Sub

Private Sub lbTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbMoving Then
        Me.Top = Me.Top + Y - startY
        Me.Left = Me.Left + X - startX
    End If
End Sub

Private Sub lbTitleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMoving = False
    If Button = 1 Then
        Call Dock
    Else
        PopupMenu mnuOptions
    End If
End Sub

Private Function FindNearest(Left As Long, Top As Long, Width As Long, Height As Long, Col As Collection, current As Long, ByRef nSide As Long) As Note
    Dim aNote As Note
    Dim ml As Long, mt As Long, mr As Long, mb As Long
    Dim il As Long, it As Long, ir As Long, ib As Long
    Dim d As Long
    Dim side As Long, nnote As Long, min As Long
    Dim k As Long
    If Col.Count = 0 Then Exit Function
    On Error GoTo errhand
    Set aNote = Parent 'Col(1)
    side = 0
    nnote = 0
    min = Screen.Width
    il = 0: ir = 0: it = 0: ib = 0
    With aNote
        ml = Abs(.Left - Left)
        mr = Abs(.Left + .Width - Left - Width)
        mt = Abs(.Top - Top)
        mb = Abs(.Top + .Height - Top - Height)
    End With
    For k = 1 To Col.Count
        Set aNote = Col(k)
        If aNote.IsVisible Then
            If Parent.NoteID <> Col(k).NoteID Then
            'If k <> current Then
                d = Abs(aNote.Left - Left - Width)
                If d > 0 And d < STHRESHOLD Then
                    If d < ml Then
                        ml = d
                        il = k
                    End If
                End If
                d = Abs(aNote.Left + aNote.Width - Left)
                If d > 0 And d < STHRESHOLD Then
                    If d < mr Then
                        mr = d
                        ir = k
                    End If
                End If
                d = Abs(aNote.Top - Top - Height)
                If d > 0 And d < STHRESHOLD Then
                    If d < mt Then
                        mt = d
                        it = k
                    End If
                End If
                d = Abs(aNote.Top + aNote.Height - Top)
                If d > 0 And d < STHRESHOLD Then
                    If d < mb Then
                        mb = d
                        ib = k
                    End If
                End If
            End If
        End If
    Next
    Set aNote = Nothing
    If ml < min Then
        min = ml
        nnote = il
        'side = 1    '1
        side = side + 1
    End If
    If mr < min Then
        min = mr
        nnote = ir
        'side = 2    '2
        side = side + 2
    End If
    If mt < min Then
        min = mt
        nnote = it
        'side = 3    '4
        side = side + 4
    End If
    If mb < min Then
        min = mb
        nnote = ib
        'side = 4    '8
        side = side + 8
    End If
    If nnote > 0 Then
    Set aNote = Col(nnote)
    End If
    nSide = side
    Set FindNearest = aNote
errhand:
    If Err Then
        Err.Clear
    End If
End Function

Function RetMin(ByRef nA As Long, ByRef nB As Long) As Long 'returns Minimum of Absolute Two values
    Dim absA As Long
    Dim absB As Long
    absA = Abs(nA)
    absB = Abs(nB)
    If absA < absB Then
        RetMin = nA
    Else
        RetMin = nB
    End If
End Function

Private Sub FindMoveTo(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, aCol As Collection, MyID As Long, ByRef xMove As Long, ByRef yMove As Long)
    ' returns DeltaX and DeltaY  - values of how far the box should be moved. Sign shows the moving direction.
    Dim k As Long
    Dim aNote As Note
    Dim xMin As Long, yMin As Long
    Dim tmpD As Long
    xMin = Screen.Width
    yMin = Screen.Height
    On Error GoTo errhand
    For k = 1 To aCol.Count
        Set aNote = aCol(k)
        If aNote.IsVisible Then
            If aNote.NoteID <> MyID Then
                tmpD = aNote.Left - nLeft               ' |A1-B1|
                If Abs(tmpD) < STHRESHOLD Then
                    xMin = tmpD
                End If
                tmpD = aNote.Left - (nLeft + nWidth)    ' |A1-B2|
                If Abs(tmpD) < STHRESHOLD Then
                    xMin = RetMin(xMin, tmpD)
                End If
                tmpD = (aNote.Left + aNote.Width) - nLeft ' |A2-B1|
                If Abs(tmpD) < STHRESHOLD Then
                    xMin = RetMin(xMin, tmpD)
                End If
                tmpD = (aNote.Left + aNote.Width) - (nLeft + nWidth) ' |A2-B2|
                If Abs(tmpD) < STHRESHOLD Then
                    xMin = RetMin(xMin, tmpD)
                End If
                
                
                tmpD = aNote.Top - nTop                     ' |A3-B3|
                If Abs(tmpD) < STHRESHOLD Then
                    yMin = tmpD
                End If
                tmpD = aNote.Top - (nTop + nHeight)         ' |A3-B4|
                If Abs(tmpD) < STHRESHOLD Then
                    yMin = RetMin(yMin, tmpD)
                End If
                tmpD = (aNote.Top + aNote.Height) - nTop    ' |A4-B3|
                If Abs(tmpD) < STHRESHOLD Then
                    yMin = RetMin(yMin, tmpD)
                End If
                tmpD = (aNote.Top + aNote.Height) - (nTop + nHeight) ' |A4-B4|
                If Abs(tmpD) < STHRESHOLD Then
                    yMin = RetMin(yMin, tmpD)
                End If
            End If
        End If
    Next
errhand:
    If Err Then
        Err.Clear
    End If
    If Abs(xMin) >= Screen.Width Then xMin = 0
    If Abs(yMin) >= Screen.Height Then yMin = 0
    xMove = xMin
    yMove = yMin
End Sub

Private Sub Dock()
    Dim aCol As Collection
    Dim aNote As Note
    Dim theNearestNote As Note
    Dim nNearestSide As Long    '1,2,3, or 4
    Dim xMove As Long, yMove As Long
    Set aCol = Parent.Parent.Notes
    
    Call FindMoveTo(Me.Left, Me.Top, Me.Width, Me.Height, aCol, Parent.NoteID, xMove, yMove)
    Me.Left = Me.Left + xMove
    Me.Top = Me.Top + yMove

    With Me
        If .Left < STHRESHOLD Then
            .Left = 0
        End If
        If .Top < STHRESHOLD Then
            .Top = 0
        End If
        
        If .Top + .Height > Screen.Height - STHRESHOLD Then
            .Top = Screen.Height - .Height
        End If
        If .Left + .Width > Screen.Width - STHRESHOLD Then
            .Left = Screen.Width - .Width
        End If
    End With
End Sub

Private Function WithinThreshold(a As Long, b As Long, t As Long) As Boolean
    WithinThreshold = True
    If a < b Then
        If (a + t > b) And (b - a < t) Then Exit Function
    Else
        If (b + t > a) And (a - b < t) Then Exit Function
    End If
    WithinThreshold = False
End Function

Friend Sub ChangeColor(Color As Long)
    Me.BackColor = Color
    txtNote.BackColor = Color
    lbTitleBar.BackColor = Me.BackColor
    lblClose.BackColor = Color
    Picture2.BackColor = Color 'Me.BackColor
    Picture3.BackColor = Color 'Me.BackColor
    Picture4.BackColor = Color 'Me.BackColor
End Sub

Private Sub mnuBackground_Click()
    ChangeColor vbDesktop
End Sub

Private Sub mnuBlue_Click()
    ChangeColor &HFFFF80    'vbBlue
End Sub

Private Sub mnuButtonface_Click()
    ChangeColor vbButtonFace
End Sub

Private Sub mnuColorCustom_Click()
    Dim ptrCustomColors As Variant
    Dim Colors As modDialogs.CHOOSECOLORSTRUCT
    Dim k As Long
    Dim res As Long
    Dim rgbCurrent
    rgbCurrent = Me.BackColor
    With Colors
        .flags = modDialogs.CC_RGBINIT
        .hwndOwner = Me.hwnd
        ptrCustomColors = aCustomColors
        .lpCustColors = VarPtr(ptrCustomColors)
        .rgbResult = rgbCurrent
        .lStructSize = Len(Colors)
    End With
    res = modDialogs.ChooseColor(Colors)
    If res <> 0 Then
        ChangeColor Colors.rgbResult
    End If
End Sub

Public Sub ShowFontProps()
    Dim cf As modDialogs.CHOOSEFONTSTRUCT
    Dim sStyle As String
    Dim sFontName As String
    Dim k As Long
    Dim res As Long
    Dim vLf As Variant
    Dim vFaceName() As Byte
    
    ReDim vFaceName(32)
    
    ''sStyle = "Regular" & vbNullString
    sFontName = Me.txtNote.Font.Name & vbNullString
    
    With cf
        .flags = CF_SCREENFONTS Or CF_EFFECTS Or _
            CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .hwndOwner = Me.hwnd
        
        For k = 0 To Len(sFontName) - 1
            vFaceName(k) = Asc(Mid(sFontName, k + 1, 1))
        Next k
        vFaceName(k + 1) = 0
        For k = 0 To 31
            lf.lfFaceName(k) = vFaceName(k)
        Next k
        lf.lfCharSet = Me.txtNote.Font.Charset
        lf.lfItalic = Me.txtNote.Font.Italic
        lf.lfStrikeOut = Me.txtNote.Font.Strikethrough
        lf.lfUnderline = Me.txtNote.Font.Underline
        lf.lfHeight = Round(Me.txtNote.Font.Size * 1.32) '* 10
        Dim lhwndMem As Long
        Dim lptrMem As Long
        lhwndMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lf))
        lptrMem = GlobalLock(lhwndMem)
        CopyMemory ByVal lptrMem, lf, Len(lf)
        
        .lpLogFont = lptrMem
        
        ''.lpszStyle = sStyle
        .iPointSize = Me.txtNote.Font.Size * 10
        
        '.nSizeMin = 4                 'minimum point size
        .nSizeMax = 36                 'maximum point size
        
        .lStructSize = Len(cf)
    End With
    
    res = modDialogs.ChooseFont(cf)
    
    
    
    If res <> 0 Then
        CopyMemory lf, ByVal lptrMem, Len(lf)
        sFontName = ""
        For k = 0 To UBound(lf.lfFaceName) - 1
            If lf.lfFaceName(k) <> 0 Then
                sFontName = sFontName & Chr(lf.lfFaceName(k))
            Else
                Exit For
            End If
        Next k
        'MsgBox sFontName
        Me.txtNote.Font.Name = sFontName
        Me.txtNote.Font.Size = cf.iPointSize \ 10
        Me.txtNote.FontBold = (lf.lfWeight > 500)
        Me.txtNote.FontItalic = lf.lfItalic
        Me.txtNote.FontStrikethru = lf.lfStrikeOut
        Me.txtNote.FontUnderline = lf.lfUnderline
        Me.txtNote.Font.Charset = lf.lfCharSet
        Me.txtNote.ForeColor = cf.rgbColors
    End If
    res = GlobalUnlock(lhwndMem)
    res = GlobalFree(lhwndMem)
End Sub

Private Sub mnuFontCustom_Click()
    Call ShowFontProps
End Sub

Private Sub mnuGreen_Click()
    ChangeColor &H80FF80    'vbGreen
End Sub

Private Sub mnuSetAsDefault_Click()
    Dim tmpStyle As String
    tmpStyle = GetFontProps
    Call Parent.Parent.DefaultNoteStyle(tmpStyle)
End Sub

Private Sub mnuYellow_Click()
    ChangeColor &HC0FFFF    'vbYellow
End Sub

Private Sub mnuDelete_Click()
    Parent.Delete
    Unload Me
End Sub

Private Sub mnuTitle_Click()
    Dim tmpstr As String
    Dim fInput As dlgInputbox
    On Error GoTo errhand
    Set fInput = New dlgInputbox
    With fInput
        .caption = "Note Title"
        .lblPrompt.caption = "Please enter new Title:"
        .txtText.Text = Parent.Title
        .Show vbModal, Me
        If .OkPressed Then
            tmpstr = .txtText.Text
            Parent.Title = tmpstr
        End If
    End With
errhand:
    If Err Then
        Err.Clear
    End If
    Set fInput = Nothing
End Sub

Private Sub txtNote_Change()
    Parent.SetChanged
End Sub

Public Sub Collapse()
    If mbSmall Then
        If oldHeight <= Me.Height Then
            mbSmall = False
            Me.Height = DEFNOTEHEIGHT
        Else
            mbSmall = False
            Me.Height = oldHeight
        End If
    Else
        oldHeight = Me.Height
        Me.Height = lbTitleBar.Height + 80
        mbSmall = True
    End If
End Sub

'<sides>
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_LEFT
End Sub
Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_BOTTOM
End Sub
Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_RIGHT
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If eSide = E_LEFT Then
        If X >= 0 And Me.Width < 470 Then Exit Sub
        On Error GoTo errhand
        Me.Left = Me.Left + X
        Me.Width = Me.Width - X
        Picture4.Left = Me.ScaleWidth - Picture4.Width
        txtNote.Width = Me.ScaleWidth - Picture2.Width - Picture4.Width
        Picture3.Width = Me.ScaleWidth
    End If
errhand:
    If Err Then
        Err.Clear
    End If
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If eSide = E_BOTTOM Then
        If Y <= 0 And Me.Height < 270 Then Exit Sub
        On Error GoTo errhand
        Me.Height = Picture3.Top + Picture3.Height + Y
        ResizeHeight
    End If
errhand:
    If Err Then
        Err.Clear
    End If
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If eSide = E_RIGHT Then
        If X <= 0 And Me.Width < 470 Then Exit Sub
        On Error GoTo errhand
        Me.Width = Picture4.Left + Picture4.Width + X
        Picture4.Left = Me.ScaleWidth - Picture4.Width
        txtNote.Width = Me.ScaleWidth - Picture2.Width - Picture4.Width
        Picture3.Width = Me.ScaleWidth
    End If
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_NOTHING
    txtNote.SetFocus
End Sub
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_NOTHING
    txtNote.SetFocus
End Sub
Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eSide = E_NOTHING
    txtNote.SetFocus
End Sub

Private Sub ResizeWidth()
    On Error GoTo errhand
    Picture3.Width = Me.ScaleWidth
    Picture2.Left = 0
    Picture4.Left = Me.ScaleWidth - Picture4.Width
    txtNote.Width = Me.ScaleWidth - Picture2.Width - Picture4.Width
    txtNote.Left = Picture2.Width
errhand:
    If Err Then
        Err.Clear
    End If
End Sub
Private Sub ResizeHeight()
    On Error GoTo errhand
    Picture3.Top = Me.ScaleHeight - Picture3.Height
    txtNote.Top = lbTitleBar.Height
    txtNote.Height = Me.ScaleHeight - Picture3.Height - lbTitleBar.Height
    Picture2.Top = Me.lbTitleBar.Height
    Picture4.Top = Picture2.Top
    Picture2.Height = Me.ScaleHeight - Me.lbTitleBar.Height
    Picture4.Height = Picture2.Height
    'Original
'    txtNote.height = Me.ScaleHeight - Picture3.height - txtNote.top
'    Picture2.height = Me.ScaleHeight
'    Picture4.height = Me.ScaleHeight
errhand:
    If Err Then
        Err.Clear
    End If
End Sub

'</sides>
Private Sub txtNote_KeyPress(KeyAscii As Integer)
    Dim res As Long
    On Error GoTo errhand
    Select Case KeyAscii
    Case 6      ' Ctrl+F
        '<2/14/03>
        FindString
        '<//2/14/03>
    Case 1      ' Ctrl+A
        txtNote.SelStart = 0
        txtNote.SelLength = Len(txtNote.Text)
    Case 19
        res = MsgBox("Save all notes ?", vbQuestion Or vbOKCancel, App.ProductName)
        If res = vbOK Then
            res = Parent.Parent.SaveAll
            If res <> 0 Then
                MsgBox "Error " & CStr(res) & " saving data" & vbCrLf & vbCrLf & _
                "Please use SaveAs and export all notes to" & vbCrLf & _
                "text file, or some data may be lost", vbExclamation, App.ProductName
            End If
        End If
    Case Else
        Debug.Print KeyAscii
    End Select
errhand:
    If Err Then
        MsgBox Err.Description, vbExclamation, App.ProductName
        Err.Clear
    End If
End Sub

'<2/14/03>
Private Sub FindString(Optional SkipShowDialog As Boolean = False)
    Dim tmpstr As String
    Dim fFind As dlgInputbox
    Dim AllText As String
    Dim pos As Long
    
    pos = 1
    On Error GoTo errhand
    Set fFind = New dlgInputbox
    With fFind
        .caption = "Find"
        .lblPrompt.caption = "Find what:"
        .txtText.Text = Me.txtNote.SelText
        pos = Me.txtNote.SelStart + 1
        
        If Not SkipShowDialog Then
            .Show vbModal, Me
            If .OkPressed Then
                tmpstr = .txtText.Text
                If tmpstr = "" Then
                    GoTo errhand
                End If
            End If
        Else
            tmpstr = Me.txtNote.SelText
        End If
        AllText = Me.txtNote.Text
        pos = InStr(pos + 1, AllText, tmpstr, vbTextCompare)
        If pos = 0 Then
            MsgBox "Cannot find """ & tmpstr & """", vbInformation, App.ProductName
        Else
            Me.txtNote.SelStart = pos - 1
            Me.txtNote.SelLength = Len(tmpstr)
        End If
    End With
errhand:
    If Err Then
        Err.Clear
    End If
    Set fFind = Nothing
End Sub

'<//2/14/03>

Private Sub txtNote_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then 'f3 - continue search
        FindString True
    End If
End Sub


Public Sub SetFontProps(ByVal fp As String)
    Dim k As Long
    Dim sFontName As String
    If fp = "" Then
        Exit Sub
    End If
    On Error GoTo errhand
    With lf
        .lfWeight = IIf(Mid(fp, 1, 1) = "0", 0, 700)
        .lfItalic = Asc(Mid(fp, 2, 1))
        .lfUnderline = Asc(Mid(fp, 3, 1))
        .lfStrikeOut = Asc(Mid(fp, 4, 1))
        .lfCharSet = Asc(Mid(fp, 5, 1))
        .lfOutPrecision = Asc(Mid(fp, 6, 1))
        .lfClipPrecision = Asc(Mid(fp, 7, 1))
        .lfQuality = Asc(Mid(fp, 8, 1))
        .lfPitchAndFamily = Asc(Mid(fp, 9, 1))
        .lfHeight = Asc(Mid(fp, 10, 1))
        For k = 0 To 31
            On Error Resume Next
            .lfFaceName(k) = Asc(Mid(fp, 11 + k, 1))
        Next k

        sFontName = ""
        For k = 0 To UBound(.lfFaceName) - 1
            If .lfFaceName(k) <> 0 Then
                sFontName = sFontName & Chr(.lfFaceName(k))
            Else
                Exit For
            End If
        Next k
        Me.txtNote.Font.Name = IIf(sFontName = "", "Arial", sFontName)
        Me.txtNote.Font.Size = .lfHeight ' 8.25 ' cf.PointSize \ 10
        Me.txtNote.FontBold = (.lfWeight > 500)
        Me.txtNote.FontItalic = .lfItalic
        Me.txtNote.FontStrikethru = .lfStrikeOut
        Me.txtNote.FontUnderline = .lfUnderline
        Me.txtNote.Font.Charset = .lfCharSet
    End With
errhand:
    Err.Clear
End Sub

Public Function GetFontProps() As String
    Dim k As Long
    Dim tmpstr As String
    tmpstr = ""
    With lf
        tmpstr = tmpstr & IIf(.lfWeight > 500, "1", "0")
        tmpstr = tmpstr & Chr(.lfItalic)
        tmpstr = tmpstr & Chr(.lfUnderline)
        tmpstr = tmpstr & Chr(.lfStrikeOut)
        tmpstr = tmpstr & Chr(.lfCharSet)
        tmpstr = tmpstr & Chr(.lfOutPrecision)
        tmpstr = tmpstr & Chr(.lfClipPrecision)
        tmpstr = tmpstr & Chr(.lfQuality)
        tmpstr = tmpstr & Chr(.lfPitchAndFamily)
        tmpstr = tmpstr & Chr(Me.txtNote.Font.Size)
        'Debug.Print Me.txtNote.FontName
        For k = 0 To 31
            tmpstr = tmpstr & Chr(.lfFaceName(k))
        Next k
    End With
    GetFontProps = tmpstr
End Function


Friend Sub ChangeTextColor(Color As Long)
    Me.txtNote.ForeColor = Color
End Sub
