VERSION 5.00
Begin VB.Form MAIN_PAGE 
   BackColor       =   &H00C0FFC0&
   Caption         =   "PAGE_4"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8160
      Top             =   6720
   End
   Begin VB.Image imgCellPicture 
      Height          =   375
      Index           =   6
      Left            =   5880
      Picture         =   "PAGE_4.frx":0000
      Stretch         =   -1  'True
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SNAKE RELOADED 4.0"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11655
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Level: "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   8280
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblscore 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   8280
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   0
      Left            =   4680
      Picture         =   "PAGE_4.frx":0312
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   1
      Left            =   3360
      Picture         =   "PAGE_4.frx":0A14
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   2
      Left            =   2160
      Picture         =   "PAGE_4.frx":1116
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   3
      Left            =   2760
      Picture         =   "PAGE_4.frx":1818
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   4
      Left            =   3960
      Picture         =   "PAGE_4.frx":1F1A
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCellPicture 
      Height          =   360
      Index           =   5
      Left            =   5280
      Picture         =   "PAGE_4.frx":261C
      Top             =   7800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   299
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   298
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   297
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   296
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   295
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   294
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   293
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   292
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   291
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   290
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   289
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   288
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   287
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   286
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   285
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   284
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   283
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   282
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   281
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   280
      Left            =   720
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   279
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   278
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   277
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   276
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   275
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   274
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   273
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   272
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   271
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   270
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   269
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   268
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   267
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   266
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   265
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   264
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   263
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   262
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   261
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   260
      Left            =   720
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   259
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   258
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   257
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   256
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   255
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   254
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   253
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   252
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   251
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   250
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   249
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   248
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   247
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   246
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   245
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   244
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   243
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   242
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   241
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   240
      Left            =   720
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   239
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   238
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   237
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   236
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   235
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   234
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   233
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   232
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   231
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   230
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   229
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   228
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   227
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   226
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   225
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   224
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   223
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   222
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   221
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   220
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   219
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   218
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   217
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   216
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   215
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   214
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   213
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   212
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   211
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   210
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   209
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   208
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   207
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   206
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   205
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   204
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   203
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   202
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   201
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   200
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   199
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   198
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   197
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   196
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   195
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   194
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   193
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   192
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   191
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   190
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   189
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   188
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   187
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   186
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   185
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   184
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   183
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   182
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   181
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   180
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   179
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   178
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   177
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   176
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   175
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   174
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   173
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   172
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   171
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   170
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   169
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   168
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   167
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   166
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   165
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   164
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   163
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   162
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   161
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   160
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   159
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   158
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   157
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   156
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   155
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   154
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   153
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   152
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   151
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   150
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   149
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   148
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   147
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   146
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   145
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   144
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   143
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   142
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   141
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   140
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   139
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   138
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   137
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   136
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   135
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   134
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   133
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   132
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   131
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   130
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   129
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   128
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   127
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   126
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   125
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   124
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   123
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   122
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   121
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   120
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   119
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   118
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   117
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   116
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   115
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   114
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   113
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   112
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   111
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   110
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   109
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   108
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   107
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   106
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   105
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   104
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   103
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   102
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   101
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   100
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   99
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   98
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   97
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   96
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   95
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   94
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   93
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   92
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   91
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   90
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   89
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   88
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   87
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   86
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   85
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   84
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   83
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   82
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   81
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   80
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   79
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   78
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   77
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   76
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   75
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   74
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   73
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   72
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   71
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   70
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   69
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   68
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   67
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   66
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   65
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   64
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   63
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   62
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   61
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   60
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   59
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   58
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   57
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   56
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   55
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   54
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   53
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   52
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   51
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   50
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   49
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   48
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   47
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   46
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   45
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   44
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   43
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   42
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   41
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   40
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   39
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   38
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   37
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   36
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   35
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   34
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   33
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   32
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   31
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   30
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   29
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   28
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   27
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   26
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   25
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   24
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   23
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   22
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   21
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   20
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   19
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   18
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   17
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   16
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   15
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   14
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   13
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   12
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   11
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   10
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   9
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   8
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   7
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   6
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   4
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   2
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   1
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgCells 
      Height          =   375
      Index           =   0
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   7
      Height          =   5415
      Left            =   720
      Top             =   1920
      Width           =   7215
   End
   Begin VB.Menu FILE 
      Caption         =   "FILE"
      Begin VB.Menu CT 
         Caption         =   "CONTROLS"
      End
      Begin VB.Menu DASH 
         Caption         =   "-"
      End
      Begin VB.Menu ET 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu EDIT 
      Caption         =   "EDIT"
      Begin VB.Menu PG 
         Caption         =   "PAUSE GAME         ESC"
      End
   End
   Begin VB.Menu HS 
      Caption         =   "HIGH SCORE"
   End
   Begin VB.Menu ABOUT 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu HELP 
      Caption         =   "HELP"
   End
End
Attribute VB_Name = "MAIN_PAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCell As Integer
Dim Direction As Integer
Dim curFoodPosition As Integer
Dim SnakeMain(299) As Integer
Dim curLength As Integer
Dim Score As Integer
Dim Paused As Boolean

Private Sub ABOUT_Click()
ABOUT_PAGE.Show
MAIN_PAGE.Hide
End Sub

Private Sub CT_Click()
CONTROL_PAGE.Show
MAIN_PAGE.Hide
End Sub

Private Sub ET_Click()
Unload Me
End Sub

Private Sub HELP_Click()
HELP_PAGE.Show
MAIN_PAGE.Hide
End Sub

Private Sub Form_Activate()
    Init
End Sub
Private Sub Init()
    'Initializes the background cells
    Dim a As Integer
    For a = 0 To 299
        imgCells(a).Picture = imgCellPicture(BlankCell).Picture
    Next
    
    'the snake has no body
    curLength = 0
    'the score is 0
    Score = 0
    Paused = False
    
    Select Case Level
        Case 5
            lblLevel.Caption = "Level: Easy"
            Timer1.Interval = 10000
        Case 10
            lblLevel.Caption = "Level: Medium"
            Timer1.Interval = 4000
        Case 20
            lblLevel.Caption = "Level: Hard"
            Timer1.Interval = 1500
    End Select
    lblscore.Caption = "Score: 0"
    
    'sets the timer interval
    Timer1.Interval = 800 / Level
    'Sets the direction of the snake head
    curCell = 0
    imgCells(0).Picture = imgCellPicture(HeadRight).Picture
    Direction = vbKeyRight
    Timer1.Enabled = True
    'Set the food's location
    curFoodPosition = setFoodPosition()
    imgCells(curFoodPosition).Picture = imgCellPicture(TheFood).Picture
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Do nothing if what user entered is not an arrow key
    If KeyCode = vbKeyEscape Then
        Paused = True
        Timer1.Enabled = False
    End If
    
    If isArrowKey(KeyCode) = False Then
        Exit Sub
    End If
    
    If KeyCode = Direction Then
        'Do nothing if the user keeps on pressing the same direction
        If Paused Then
            Direction = KeyCode
            Paused = False
        Else
            Exit Sub
        End If
    End If
    
    If isOppositeDirection(KeyCode) = True Then
    'calls a function that checks if the user tried to move in the opposite direction
        Exit Sub
    Else
        Timer1.Enabled = False
        MoveCell KeyCode
    End If
End Sub

Private Function isOppositeDirection(key As Integer) As Boolean
    Select Case key
        Case vbKeyUp
            If Direction = vbKeyDown Then
                isOppositeDirection = True
            Else
                isOppositeDirection = False
            End If
        Case vbKeyDown
            If Direction = vbKeyUp Then
                isOppositeDirection = True
            Else
                isOppositeDirection = False
            End If
        Case vbKeyRight
            If Direction = vbKeyLeft Then
                isOppositeDirection = True
            Else
                isOppositeDirection = False
            End If
        Case vbKeyLeft
            If Direction = vbKeyRight Then
                isOppositeDirection = True
            Else
                isOppositeDirection = False
            End If
    End Select
End Function

Private Function isArrowKey(a As Integer) As Boolean
    If a <> vbKeyUp And a <> vbKeyDown And a <> vbKeyLeft And a <> vbKeyRight Then
        isArrowKey = False
    Else
        isArrowKey = True
    End If
End Function
Private Sub MoveCell(d As Integer)
    Timer1.Enabled = True
    Dim newCell As Integer

    Direction = d
    Select Case d
    Case vbKeyUp
        If curCell < 20 Then
            HitEdge
        Else
            newCell = curCell - 20
            If imgCells(newCell).Picture = imgCellPicture(TheBody).Picture Then
            'tangled with itself
                Tangle
                Exit Sub
            End If
            imgCells(newCell).Picture = imgCellPicture(HeadUp).Picture
        End If
    Case vbKeyDown
        If curCell >= 280 And curCell <= 299 Then
            HitEdge
        Else
            newCell = curCell + 20
            If imgCells(newCell).Picture = imgCellPicture(TheBody).Picture Then
            'tangled with itself
                Tangle
                Exit Sub
            End If
            
            imgCells(newCell).Picture = imgCellPicture(HeadDown).Picture
        End If
    Case vbKeyLeft
        If curCell Mod 20 = 0 Then
            HitEdge
        Else
            newCell = curCell - 1
            If imgCells(newCell).Picture = imgCellPicture(TheBody).Picture Then
            'tangled with itself
                Tangle
                Exit Sub
            End If
            
            imgCells(newCell).Picture = imgCellPicture(HeadLeft).Picture
        End If
    Case vbKeyRight
        If curCell Mod 20 = 19 Then
            HitEdge
        Else
            newCell = curCell + 1
            If imgCells(newCell).Picture = imgCellPicture(TheBody).Picture Then
            'tangled with itself
                Tangle
                Exit Sub
            End If
            imgCells(newCell).Picture = imgCellPicture(HeadRight).Picture
        End If
    End Select
    
    'the new cell is loaded with the snake's head
    'what we need to do is to the body one level higher
    Dim OldCell As Integer
    OldCell = curCell
    curCell = newCell
        
    'Checks if the snake has eaten the food
    If curFoodPosition = newCell Then
        curFoodPosition = setFoodPosition()
        imgCells(curFoodPosition) = imgCellPicture(TheFood).Picture
        curLength = curLength + 1
        Score = Score + Level
        lblscore.Caption = "Score: " & Score
    End If
    
    If curLength = 0 Then
        imgCells(OldCell).Picture = imgCellPicture(BlankCell).Picture
    Else
        If curLength = 1 Then
            imgCells(OldCell).Picture = imgCellPicture(BlankCell).Picture
        End If
        MoveSnake
    End If
    
End Sub

Private Sub MoveSnake()
    Dim temp(299) As Integer
    Dim rear As Integer
    'Copies the contents of snakemain to temp
    Dim i As Integer
    temp(0) = curCell
    For i = 0 To curLength
        temp(i + 1) = SnakeMain(i)
    Next
    rear = SnakeMain(curLength)
    'Copies the contents back to the SnakeMain
    For i = 0 To curLength
        SnakeMain(i) = temp(i)
    Next
    
    'Fills the image with snake body
    For i = 1 To curLength
        imgCells(SnakeMain(i)).Picture = imgCellPicture(TheBody).Picture
    Next
    imgCells(SnakeMain(curLength)).Picture = imgCellPicture(BlankCell).Picture
    
End Sub
Private Sub HitEdge()
    Dim num As Integer
    Timer1.Enabled = False
    num = MsgBox("YOU  HAVE  HIT  THE  EDGE  !!!     YOUR  SCORE  IS   " & Score & ",      WANT  TO  START  A  NEW  GAME ?", vbYesNo + vbExclamation, "GAME OVER")
    If num = vbYes Then
        Init
    Else
        Unload Me
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.Width = 12800
    Me.Height = 8900
End Sub

Private Sub lblExit_Click()
    Dim num As Integer
    num = MsgBox("ARE  YOU  SURE, YOU  WANT  TO  EXIT?", vbYesNo + vbQuestion, "EXIT")
    If num = vbYes Then
        Unload Me
    End If
End Sub

Private Sub NG_Click()
MAIN_PAGE.Show
End Sub

Private Sub Timer1_Timer()
    MoveCell Direction
    
End Sub

Private Function setFoodPosition()
    Randomize
    Dim tempPos As Integer
    Dim lValid As Boolean
    lValid = False
    While Not lValid
        tempPos = Int((299 + 0 - 1) * Rnd + 0)
        If imgCells(tempPos).Picture = imgCellPicture(BlankCell).Picture Then
            setFoodPosition = tempPos
            lValid = True
        Else
            lValid = False
        End If
    Wend
End Function

Private Sub Tangle()
    Dim num As Integer
    Timer1.Enabled = False
    num = MsgBox("YOU  HAVE  KNOCKED  INTO  YOURSELF     !!!    YOUR  SCORE  IS  " & Score & ",   START  A  NEW  GAME?", vbYesNo + vbExclamation, "Game Over")
    If num = vbYes Then
        Init
    Else
        Unload Me
    End If
End Sub
