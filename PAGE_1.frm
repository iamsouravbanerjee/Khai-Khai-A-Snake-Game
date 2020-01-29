VERSION 5.00
Begin VB.Form START_PAGE 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   Picture         =   "PAGE_1.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   14355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "BatmanForeverAlternate"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "BatmanForeverAlternate"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
End
Attribute VB_Name = "START_PAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LOADING_PAGE.Show
START_PAGE.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
