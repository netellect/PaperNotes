VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1590
   ClientLeft      =   4890
   ClientTop       =   7590
   ClientWidth     =   3615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CheckBox cbClickThrough 
         Caption         =   "Allow Through Clicks"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   57
         SmallChange     =   19
         Min             =   10
         Max             =   255
         SelStart        =   255
         TickFrequency   =   19
         Value           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Transparency Level"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OkPressed As Boolean

Private Sub cmdCancel_Click()
    OkPressed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    OkPressed = True
    Me.Hide
End Sub
