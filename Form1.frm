VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   5760
   ClientTop       =   8385
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "def size All"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "collapse all"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "show all"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hide all"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "create object"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private allNotes As New Notes
Private aNote As Note

Private Sub Command1_Click()
    Set aNote = allNotes.AddNote
    aNote.Text = Me.Text1.Text
    aNote.Title = Me.txtTitle.Text
    aNote.Show
End Sub

Private Sub Command2_Click()
    allNotes.HideAll
End Sub

Private Sub Command3_Click()
    allNotes.ShowAll
End Sub

Private Sub Command4_Click()
    allNotes.CollapseAll
End Sub

Private Sub Command5_Click()
    Dim a
    For Each a In allNotes.Notes
        a.Height = 1300
        a.Width = 2500
    Next
End Sub
