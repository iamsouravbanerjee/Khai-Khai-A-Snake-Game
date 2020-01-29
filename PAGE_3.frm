VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form LOADING_PAGE 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14370
   LinkTopic       =   "PAGE_3"
   ScaleHeight     =   8055
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   13080
      Top             =   1920
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   4920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "American Captain"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   8280
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "American Captain"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   8100
      Left            =   0
      Picture         =   "PAGE_3.frx":0000
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "LOADING_PAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label1.Caption = "LOADING.........."
Label2.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
LEVEL_PAGE.Show
LOADING_PAGE.Hide
End If
End Sub
