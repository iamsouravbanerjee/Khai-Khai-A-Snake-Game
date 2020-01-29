VERSION 5.00
Begin VB.Form LEVEL_PAGE 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   14115
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optHard 
      BackColor       =   &H00C000C0&
      Caption         =   "HARD"
      BeginProperty Font 
         Name            =   "American Captain"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
   End
   Begin VB.OptionButton optMedium 
      BackColor       =   &H000040C0&
      Caption         =   "MEDIUM"
      BeginProperty Font 
         Name            =   "American Captain"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.OptionButton optEasy 
      BackColor       =   &H008080FF&
      Caption         =   "EASY"
      BeginProperty Font 
         Name            =   "American Captain"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6000
      TabIndex        =   1
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10680
      TabIndex        =   0
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   8100
      Left            =   0
      Picture         =   "PAGE_2.frx":0000
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "LEVEL_PAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MAIN_PAGE.Show
LEVEL_PAGE.Hide
End Sub

Private Sub Command2_Click()
MAIN_PAGE.Show
LEVEL_PAGE.Hide
End Sub

Private Sub Command3_Click()
MAIN_PAGE.Show
LEVEL_PAGE.Hide
End Sub

Private Sub cmdStart_Click()
    If optEasy.Value = True Then
        Level = 5
    End If
    
    If optMedium.Value = True Then
        Level = 10
    End If
    
    If optHard.Value = True Then
        Level = 20
    End If
    
    Unload Me
    MAIN_PAGE.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
