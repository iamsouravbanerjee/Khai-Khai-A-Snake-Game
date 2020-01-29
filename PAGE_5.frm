VERSION 5.00
Begin VB.Form ABOUT_PAGE 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL RIGHT RESERVED."
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © 2016, SSKV. "
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   3840
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   4800
      Picture         =   "PAGE_5.frx":0000
      Top             =   1200
      Width           =   3300
   End
End
Attribute VB_Name = "ABOUT_PAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
