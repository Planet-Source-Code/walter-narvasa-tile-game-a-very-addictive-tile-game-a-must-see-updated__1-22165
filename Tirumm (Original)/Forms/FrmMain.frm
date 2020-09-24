VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tirumm: The Tile Rummy Game by Walter A. Narvasa"
   ClientHeight    =   10020
   ClientLeft      =   2040
   ClientTop       =   1860
   ClientWidth     =   12840
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":08CA
   ScaleHeight     =   10020
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox MoveStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   840
      Left            =   120
      TabIndex        =   408
      Top             =   8160
      Width           =   9975
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":133F2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   399
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":136FC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   398
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":13A06
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   397
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":13D10
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   396
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1401A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   395
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":14324
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   394
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1462E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   393
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":14938
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   392
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":14C42
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   391
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":14F4C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   390
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":15256
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   389
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":15560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   388
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1586A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   387
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":15B74
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   386
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":15E7E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   385
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":16188
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   384
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":16492
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   383
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":1679C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   382
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":16AA6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   381
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":16DB0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   380
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":170BA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   379
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":173C4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   378
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":176CE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   377
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":179D8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   376
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":17CE2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   375
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":17FEC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   374
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":182F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   373
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":18600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   372
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1890A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   371
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   240
      MouseIcon       =   "FrmMain.frx":18C14
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   370
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   240
      MouseIcon       =   "FrmMain.frx":18F1E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   369
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   240
      MouseIcon       =   "FrmMain.frx":19228
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   368
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   240
      MouseIcon       =   "FrmMain.frx":19532
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   367
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1983C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   366
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   240
      MouseIcon       =   "FrmMain.frx":19B46
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   365
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   240
      MouseIcon       =   "FrmMain.frx":19E50
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   364
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1A15A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   363
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1A464
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   362
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1A76E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   361
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1AA78
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   360
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1AD82
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   359
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   240
      MouseIcon       =   "FrmMain.frx":1B08C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   358
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1B396
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   357
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1B6A0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   356
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1B9AA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   355
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1BCB4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   354
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1BFBE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   353
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1C2C8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   352
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1C5D2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   351
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1C8DC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   350
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1CBE6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   349
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1CEF0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   348
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1D1FA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   347
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1D504
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   346
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1D80E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   345
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   9600
      MouseIcon       =   "FrmMain.frx":1DB18
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   344
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":1DE22
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   343
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":1E12C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   342
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":1E436
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   341
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":1E740
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   340
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":1EA4A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   339
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":1ED54
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   338
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":1F05E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   337
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":1F368
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   336
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":1F672
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   335
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":1F97C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   334
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":1FC86
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   333
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":1FF90
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   332
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":2029A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   331
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":205A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   330
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":208AE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   329
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":20BB8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   328
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":20EC2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   327
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":211CC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   326
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":214D6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   325
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":217E0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   324
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":21AEA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   323
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":21DF4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   322
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":220FE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   321
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":22408
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   320
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":22712
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   319
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":22A1C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   318
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":22D26
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   317
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":23030
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   316
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   240
      MouseIcon       =   "FrmMain.frx":2333A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   315
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   240
      MouseIcon       =   "FrmMain.frx":23644
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   314
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   240
      MouseIcon       =   "FrmMain.frx":2394E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   313
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   240
      MouseIcon       =   "FrmMain.frx":23C58
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   312
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   240
      MouseIcon       =   "FrmMain.frx":23F62
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   311
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   240
      MouseIcon       =   "FrmMain.frx":2426C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   310
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   240
      MouseIcon       =   "FrmMain.frx":24576
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   309
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   240
      MouseIcon       =   "FrmMain.frx":24880
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   308
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   240
      MouseIcon       =   "FrmMain.frx":24B8A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   307
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   240
      MouseIcon       =   "FrmMain.frx":24E94
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   306
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   240
      MouseIcon       =   "FrmMain.frx":2519E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   305
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   240
      MouseIcon       =   "FrmMain.frx":254A8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   304
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   240
      MouseIcon       =   "FrmMain.frx":257B2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   303
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   240
      MouseIcon       =   "FrmMain.frx":25ABC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   298
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":25DC6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   295
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":260D0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   294
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":263DA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   293
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":266E4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   292
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":269EE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   291
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":26CF8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   290
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":27002
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   289
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":2730C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   288
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":27616
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   287
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":27920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   286
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":27C2A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   285
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":27F34
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   284
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":2823E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   283
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":28548
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   282
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":28852
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   281
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":28B5C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   280
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":28E66
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   279
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":29170
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   278
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":2947A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   277
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":29784
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   276
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":29A8E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   275
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":29D98
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   274
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":2A0A2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   273
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":2A3AC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   272
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":2A6B6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   271
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":2A9C0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   270
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":2ACCA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   269
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":2AFD4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   268
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2B2DE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   267
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2B5E8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   266
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2B8F2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   265
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2BBFC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   264
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2BF06
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   263
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2C210
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   262
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2C51A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   261
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2C824
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   260
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2CB2E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   259
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2CE38
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2D142
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   257
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2D44C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   256
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2D756
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   255
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Cover 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      MouseIcon       =   "FrmMain.frx":2DA60
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   254
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer GameClocker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   9240
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2DD6A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2E074
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   251
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2E37E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   250
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2E688
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2E992
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   248
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2EC9C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   247
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2EFA6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   12240
      MouseIcon       =   "FrmMain.frx":2F2B0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   245
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":2F5BA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":2F8C4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":2FBCE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":2FED8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":301E2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   240
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":304EC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":307F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":30B00
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":30E0A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":31114
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":3141E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":31728
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":31A32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":31D3C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":32046
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":32350
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":3265A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   11760
      MouseIcon       =   "FrmMain.frx":32964
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":32C6E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":32F78
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":33282
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":3358C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":33896
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":33BA0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":33EAA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":341B4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":344BE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":347C8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":34AD2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":34DDC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":350E6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":353F0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":356FA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":35A04
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":35D0E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   11280
      MouseIcon       =   "FrmMain.frx":36018
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":36322
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":3662C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":36936
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":36C40
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":36F4A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":37254
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":3755E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":37868
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":37B72
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":37E7C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":38186
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":38490
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":3879A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":38AA4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":38DAE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":390B8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":393C2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   10800
      MouseIcon       =   "FrmMain.frx":396CC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":399D6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":39CE0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":39FEA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3A2F4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3A5FE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3A908
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3AC12
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3AF1C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3B226
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3B530
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3B83A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3BB44
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3BE4E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3C158
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3C462
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3C76C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3CA76
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DropTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   10320
      MouseIcon       =   "FrmMain.frx":3CD80
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":3D08A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":3D394
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":3D69E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":3D9A8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":3DCB2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":3DFBC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":3E2C6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":3E5D0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":3E8DA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":3EBE4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":3EEEE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":3F1F8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":3F502
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8160
      MouseIcon       =   "FrmMain.frx":3F80C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":3FB16
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7200
      MouseIcon       =   "FrmMain.frx":3FE20
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   6720
      MouseIcon       =   "FrmMain.frx":4012A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6240
      MouseIcon       =   "FrmMain.frx":40434
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      MouseIcon       =   "FrmMain.frx":4073E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5280
      MouseIcon       =   "FrmMain.frx":40A48
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      MouseIcon       =   "FrmMain.frx":40D52
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      MouseIcon       =   "FrmMain.frx":4105C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3840
      MouseIcon       =   "FrmMain.frx":41366
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      MouseIcon       =   "FrmMain.frx":41670
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      MouseIcon       =   "FrmMain.frx":4197A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2400
      MouseIcon       =   "FrmMain.frx":41C84
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":41F8E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":42298
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":425A2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":428AC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":42BB6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":42EC0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":431CA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":434D4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":437DE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":43AE8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":43DF2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":440FC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":44406
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer3Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8880
      MouseIcon       =   "FrmMain.frx":44710
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   960
      MouseIcon       =   "FrmMain.frx":44A1A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   960
      MouseIcon       =   "FrmMain.frx":44D24
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   960
      MouseIcon       =   "FrmMain.frx":4502E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   960
      MouseIcon       =   "FrmMain.frx":45338
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   960
      MouseIcon       =   "FrmMain.frx":45642
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   960
      MouseIcon       =   "FrmMain.frx":4594C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   960
      MouseIcon       =   "FrmMain.frx":45C56
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      MouseIcon       =   "FrmMain.frx":45F60
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      MouseIcon       =   "FrmMain.frx":4626A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      MouseIcon       =   "FrmMain.frx":46574
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      MouseIcon       =   "FrmMain.frx":4687E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      MouseIcon       =   "FrmMain.frx":46B88
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      MouseIcon       =   "FrmMain.frx":46E92
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton BufferTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      MouseIcon       =   "FrmMain.frx":4719C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton BufferTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5400
      MouseIcon       =   "FrmMain.frx":474A6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton BufferTile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5880
      MouseIcon       =   "FrmMain.frx":477B0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "&Drop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      MouseIcon       =   "FrmMain.frx":47ABA
      MousePointer    =   99  'Custom
      TabIndex        =   115
      Top             =   9240
      Width           =   1365
   End
   Begin VB.CommandButton AIPlayer4Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":47DC4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AIPlayer1Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      MouseIcon       =   "FrmMain.frx":480CE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   113
      Top             =   9765
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15064
            MinWidth        =   15064
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton AIPlayer2Tile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      MouseIcon       =   "FrmMain.frx":483D8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000080FF&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   105
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000080FF&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   104
      Left            =   2640
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   103
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   102
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   101
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   100
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   99
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   98
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   97
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   96
      Left            =   3360
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   95
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   94
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   93
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   92
      Left            =   2400
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   91
      Left            =   5880
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   90
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   89
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   88
      Left            =   4920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   87
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   86
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   85
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   84
      Left            =   4800
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   83
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   82
      Left            =   3480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   81
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   2880
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   2640
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   3000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   7920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   3840
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   3120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   3120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   3480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   2160
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   2520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   7920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   3960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   4320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   2040
      MouseIcon       =   "FrmMain.frx":486E2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   3000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   7920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   4560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   3120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   3120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   7920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   3480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   2520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   4440
      MouseIcon       =   "FrmMain.frx":489EC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   2640
      MouseIcon       =   "FrmMain.frx":48CF6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   2160
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   6480
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   7560
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   2640
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4560
      MouseIcon       =   "FrmMain.frx":49000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   2040
      MouseIcon       =   "FrmMain.frx":4930A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      MouseIcon       =   "FrmMain.frx":49614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   3120
      MouseIcon       =   "FrmMain.frx":4991E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   2640
      MouseIcon       =   "FrmMain.frx":49C28
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5400
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7440
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3600
      MouseIcon       =   "FrmMain.frx":49F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Tiles 
      BackColor       =   &H000000FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7920
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Current Status:"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   120
      TabIndex        =   409
      Top             =   7920
      Width           =   1065
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   10200
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P l  a y e  r  3  R a c k "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3960
      Left            =   9000
      TabIndex        =   407
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P l  a y e  r  1  R a c k "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3960
      Left            =   1080
      TabIndex        =   406
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P l  a y e  r  3  P e  n a  l  t  y  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5160
      Left            =   9720
      TabIndex        =   405
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P l  a y e  r  1  P e  n a  l  t  y  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5160
      Left            =   360
      TabIndex        =   404
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Penalty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 4 Penalty"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   403
      Top             =   7440
      Width           =   7095
   End
   Begin VB.Label Rack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 4 Rack"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   402
      Top             =   6720
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2 Penalty"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   401
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2 Rack"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   400
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   7
      Left            =   1680
      TabIndex        =   302
      Top             =   7440
      Width           =   165
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   6
      Left            =   9720
      TabIndex        =   301
      Top             =   600
      Width           =   165
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   5
      Left            =   1680
      TabIndex        =   300
      Top             =   240
      Width           =   165
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   4
      Left            =   360
      TabIndex        =   299
      Top             =   600
      Width           =   165
   End
   Begin VB.Label GameClockerCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timer->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   6960
      TabIndex        =   297
      Top             =   9240
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label GameClockerTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   8160
      TabIndex        =   296
      Top             =   9240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   6840
      X2              =   6840
      Y1              =   9120
      Y2              =   9720
   End
   Begin VB.Label TurnStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Player's Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   253
      Top             =   9240
      Width           =   2715
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   2400
      X2              =   2400
      Y1              =   9120
      Y2              =   9720
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   2
      Left            =   9000
      TabIndex        =   108
      Top             =   600
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   7215
      Index           =   10
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Players Dropped Tiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   705
      Left            =   10440
      TabIndex        =   117
      Top             =   120
      Width           =   2085
      WordWrap        =   -1  'True
   End
   Begin VB.Label TileStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Tiles Picked ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   2520
      TabIndex        =   116
      Top             =   9240
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      X1              =   10250
      X2              =   12720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      Height          =   9675
      Left            =   10230
      Top             =   60
      Width           =   2505
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   9240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   9735
      Left            =   10170
      Top             =   0
      Width           =   2625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   8640
      X2              =   8640
      Y1              =   9120
      Y2              =   9720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   615
      Left            =   15
      Top             =   9120
      Width           =   10170
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   107
      Top             =   960
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   3
      Left            =   1680
      TabIndex        =   109
      Top             =   6720
      Width           =   165
   End
   Begin VB.Label TurnsCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   106
      Top             =   600
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   7215
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   7215
      Index           =   7
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   615
      Index           =   5
      Left            =   1590
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   7215
      Index           =   2
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   7215
      Index           =   8
      Left            =   9510
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   7215
      Index           =   4
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   7215
      Index           =   9
      Left            =   870
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   7215
      Index           =   11
      Left            =   8790
      Shape           =   4  'Rounded Rectangle
      Top             =   525
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   615
      Index           =   13
      Left            =   1590
      Shape           =   4  'Rounded Rectangle
      Top             =   870
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   615
      Index           =   15
      Left            =   1590
      Shape           =   4  'Rounded Rectangle
      Top             =   6630
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   615
      Index           =   6
      Left            =   1590
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   7095
   End
   Begin VB.Menu mnuDeal 
      Caption         =   "&Deal"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Help Topics"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuScore 
      Caption         =   "&Score"
      Begin VB.Menu mnuTopScore 
         Caption         =   "&Top Score"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuSoundsOn 
         Caption         =   "Sounds &On"
         Checked         =   -1  'True
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSoundsOff 
         Caption         =   "Sounds O&ff"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutAuthor 
         Caption         =   "&About Author"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Tirumm: The Tile Rummy Game fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Tirumm: The Tile Rummy Game. Contact me for
' additional help/suggestions
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

Dim xCount

' STARTUP
Private Sub Form_Load()
    SoundOn = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

' START/RESTART NEW GAME
Private Sub mnuNewGame_Click()
    Dim xTiles As Integer, xVal As Integer
    Dim xTop, xLeft, Message
    
StartAgain:
    xCurrent_WinnerName = InputBox("Please Enter Your Name:", "Current Player - Input Name")
    If xCurrent_WinnerName = "" Then
        Message = MsgBox("Do you wish to cancel?", vbCritical + vbYesNo + vbQuestion, "Current Player - Cancel")
        If Message = vbYes Then
            Exit Sub
        Else
            GoTo StartAgain
        End If
    Else
        Me.Rack.Caption = "Player 4 (" & Trim(xCurrent_WinnerName) & ") Rack"
        Me.Penalty.Caption = "Player 4 (" & Trim(xCurrent_WinnerName) & ") Penalty"
    End If
    
    mnuNewGame.Enabled = False
    
    Call Initialize_Variables
    
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Start.wav", 1)
    End If
    
    For xTiles = 0 To 105
        Tiles(xTiles).Visible = True
    Next xTiles
    
    For x = 1 To 80
        DropTile(x).Visible = True
    Next x
        
    ' PLAYER 1 RANDOM PICK A TILE FROM CENTER
    'MsgBox "Please wait while Player 1 is picking a tile", vbOKOnly + vbInformation, "Player 1's Turn"
    MoveStatus.AddItem "Player 1 just pick a tile"
    MoveStatus.ListIndex = MoveStatus.NewIndex
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
    End If
    Call Random_Pick(1)
    For x = 0 To 13
        AIPlayer1Tile(x).Visible = True
        AIPlayer1Cover(x).Visible = True
    Next x
    
    ' PLAYER 2 RANDOM PICK A TILE FROM CENTER
    'MsgBox "Please wait while Player 2 is picking a tile", vbOKOnly + vbInformation, "Player 2's Turn"
    MoveStatus.AddItem "Player 2 just pick a tile"
    MoveStatus.ListIndex = MoveStatus.NewIndex
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
    End If
    Call Random_Pick(2)
    For x = 0 To 13
        AIPlayer2Tile(x).Visible = True
        AIPlayer2Cover(x).Visible = True
    Next x
    
    ' PLAYER 3 RANDOM PICK A TILE FROM CENTER
    'MsgBox "Please wait while Player 3 is picking a tile", vbOKOnly + vbInformation, "Player 3's Turn"
    MoveStatus.AddItem "Player 3 just pick a tile"
    MoveStatus.ListIndex = MoveStatus.NewIndex
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
    End If
    Call Random_Pick(3)
    For x = 0 To 13
        AIPlayer3Tile(x).Visible = True
        AIPlayer3Cover(x).Visible = True
    Next x
    
    ' PLAYER 4 RANDOM PICK A TILE FROM CENTER
    'MsgBox "Please pick a tile by clicking a tile", vbOKOnly + vbInformation, "Player 4's Turn"
    MoveStatus.AddItem "Player 4 is currently picking a tile"
    MoveStatus.ListIndex = MoveStatus.NewIndex
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
    End If
    Call Random_Pick(4)
    For x = 0 To 13
        AIPlayer4Tile(x).Visible = True
    Next x
    
    GameReset = False
    
    xTemp(0) = Val(AIPlayer1Tile(0).Caption)
    xTemp(1) = Val(AIPlayer2Tile(0).Caption)
    xTemp(2) = Val(AIPlayer3Tile(0).Caption)
    xTemp(3) = Val(AIPlayer4Tile(0).Caption)
    
    Call BubbleSortVariantArray(xTemp(), True)
    
    Call Refresh_PlayerTempArrays
                
    If xTemp(0) = Val(AIPlayer1Tile(0).Caption) Then
        x1stMove = "Player 1"
        TurnStatus.Caption = "Player 1's Turn"
        'MsgBox x1stMove & " has draw the first move.", vbOKOnly + vbInformation, "Readme"
        MoveStatus.AddItem x1stMove & " has draw the first move."
        MoveStatus.ListIndex = MoveStatus.NewIndex
        If SoundOn = True Then
            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
        End If
        CurrentTurn = "P1"
        TurnStatus.Caption = "Player 1's Turn"
        Call ValidateTile(1)
    ElseIf xTemp(0) = Val(AIPlayer2Tile(0).Caption) Then
        x1stMove = "Player 2"
        TurnStatus.Caption = "Player 2's Turn"
        'MsgBox x1stMove & " has draw the first move.", vbOKOnly + vbInformation, "Readme"
        MoveStatus.AddItem x1stMove & " has draw the first move."
        MoveStatus.ListIndex = MoveStatus.NewIndex
        If SoundOn = True Then
            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
        End If
        CurrentTurn = "P2"
        TurnStatus.Caption = "Player 2's Turn"
        Call ValidateTile(2)
    ElseIf xTemp(0) = Val(AIPlayer3Tile(0).Caption) Then
        x1stMove = "Player 3"
        TurnStatus.Caption = "Player 3's Turn"
        'MsgBox x1stMove & " has draw the first move.", vbOKOnly + vbInformation, "Readme"
        MoveStatus.AddItem x1stMove & " has draw the first move."
        MoveStatus.ListIndex = MoveStatus.NewIndex
        If SoundOn = True Then
            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
        End If
        CurrentTurn = "P3"
        Call ValidateTile(3)
    ElseIf xTemp(0) = Val(AIPlayer4Tile(0).Caption) Then
        x1stMove = "Player 4"
        TurnStatus.Caption = "Player 4's Turn"
        'MsgBox x1stMove & " has draw the first move." & vbCrLf & _
        "Please select a Straight or Trio Combinations!", vbOKOnly + vbInformation, "Readme"
        MoveStatus.AddItem x1stMove & " has draw the first move."
        MoveStatus.AddItem "Select a Straight or Trio Combinations!"
        MoveStatus.ListIndex = MoveStatus.NewIndex
        If SoundOn = True Then
            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
        End If
        For x = 0 To 2
            BufferTile(x).Visible = True
        Next x
        TileStatus.Visible = True
        GameClockerTime.Visible = True
        GameClockerCaption.Visible = True
        GameClocker.Enabled = True
        cmdDrop.Enabled = True
        CurrentTurn = "P4"
        TurnStatus.Caption = "Player 4's Turn"
    End If
End Sub

' EVENTS WHEN CURRENT CENTERED TILES ARE BEING CLICKED
Private Sub Tiles_Click(Index As Integer)
    If xPickPlayer1Now = True Then
        Tiles(Index).Visible = False
    End If
    If xPickPlayer2Now = True Then
        Tiles(Index).Visible = False
    End If
    If xPickPlayer3Now = True Then
        Tiles(Index).Visible = False
    End If
    If xPickPlayer4Now = True Then
    End If
End Sub

' PLAYER 4 TILES CLICK EVENT (HUMAN PLAYER)
Private Sub AIPlayer4Tile_Click(Index As Integer)
    If CurrentTurn = "P4" Then
        If xCount = 0 Then
            BufferTile(0).Caption = AIPlayer4Tile(Index).Caption
            BufferTile(0).BackColor = AIPlayer4Tile(Index).BackColor
            BufferTile(0).ToolTipText = AIPlayer4Tile(Index).Index
            xCount = 1
        ElseIf xCount = 1 Then
            BufferTile(1).Caption = AIPlayer4Tile(Index).Caption
            BufferTile(1).BackColor = AIPlayer4Tile(Index).BackColor
            BufferTile(1).ToolTipText = AIPlayer4Tile(Index).Index
            xCount = 2
        ElseIf xCount = 2 Then
            BufferTile(2).Caption = AIPlayer4Tile(Index).Caption
            BufferTile(2).BackColor = AIPlayer4Tile(Index).BackColor
            BufferTile(2).ToolTipText = AIPlayer4Tile(Index).Index
            xCount = 3
        End If
    End If
    If Index = 14 Or Index = 15 Or Index = 16 Or Index = 17 Or Index = 18 Or Index = 19 Or _
        Index = 20 Or Index = 21 Or Index = 22 Or Index = 23 Or Index = 24 Or Index = 25 Or _
        Index = 26 Or Index = 27 Then
        MsgBox "You can select from this set of tiles!" & vbCrLf & _
                "This is a set of penalty tiles.", vbOKOnly + vbInformation, "Player 1's Warning"
    End If
End Sub

' PLAYER 4 DROP EVENT VALIDATION
Private Sub cmdDrop_Click()
    If Not GameReset Then
        If SoundOn = True Then
            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
        End If
        If xCount = 3 Then
                If Val(BufferTile(0).Caption) = Val(BufferTile(1).Caption) And _
                    Val(BufferTile(1).Caption) = Val(BufferTile(2).Caption) Or _
                    (Val(BufferTile(0).Caption) - Val(BufferTile(1).Caption)) = _
                    (Val(BufferTile(1).Caption) - Val(BufferTile(2).Caption)) And _
                    (BufferTile(0).BackColor = BufferTile(1).BackColor And _
                    BufferTile(1).BackColor = BufferTile(2).BackColor) Then
                    AIPlayer4Tile(Val(BufferTile(0).ToolTipText)).Visible = False
                    AIPlayer4Tile(Val(BufferTile(1).ToolTipText)).Visible = False
                    AIPlayer4Tile(Val(BufferTile(2).ToolTipText)).Visible = False
                    'MsgBox ("Player Four xDpVal=>" & xDpVal)
                    DropTile(xDpVal).Caption = Val(BufferTile(0).Caption)
                    DropTile(xDpVal).BackColor = BufferTile(0).BackColor
                    DropTile(xDpVal).ToolTipText = Val(BufferTile(0).ToolTipText)
                    DropTile(xDpVal + 1).Caption = Val(BufferTile(1).Caption)
                    DropTile(xDpVal + 1).BackColor = BufferTile(1).BackColor
                    DropTile(xDpVal + 1).ToolTipText = Val(BufferTile(1).ToolTipText)
                    DropTile(xDpVal + 2).Caption = Val(BufferTile(2).Caption)
                    DropTile(xDpVal + 2).BackColor = BufferTile(2).BackColor
                    DropTile(xDpVal + 2).ToolTipText = Val(BufferTile(2).ToolTipText)
                    xDpVal = xDpVal + 3
                    AIPlayer4Tile(Val(BufferTile(0).ToolTipText)).Caption = ""
                    AIPlayer4Tile(Val(BufferTile(1).ToolTipText)).Caption = ""
                    AIPlayer4Tile(Val(BufferTile(2).ToolTipText)).Caption = ""
                    FillInTiles = True
                    Call Random_Pick(4)
                    cmdDrop.Enabled = False
                    fMain.GameClocker.Enabled = False
                    GameClockerTime.Caption = 0
                    If Not GameReset Then
                        'MsgBox "Congratulations! Player 4 has completed a combination of Trio." & vbCrLf & _
                        '        "Player 1 will make the next move.", vbOKOnly + vbInformation, "Player 4's Alert"
                        MoveStatus.AddItem "Congratulations! Player 4 has completed a tile combination."
                        MoveStatus.AddItem "Player 1 will make the next move."
                        MoveStatus.ListIndex = MoveStatus.NewIndex
                        If SoundOn = True Then
                            Call sndPlaySound(App.Path & "\Sounds\Click.wav", 1)
                        End If
                        AIPlayer4Tile(Val(BufferTile(0).ToolTipText)).Visible = True
                        AIPlayer4Tile(Val(BufferTile(1).ToolTipText)).Visible = True
                        AIPlayer4Tile(Val(BufferTile(2).ToolTipText)).Visible = True
                        GameLoopCount = GameLoopCount + 1
                        fMain.TileStatus.Visible = False
                        fMain.GameClockerTime.Visible = False
                        fMain.GameClockerCaption.Visible = False
                        For i = 0 To 2
                            BufferTile(i).Visible = True
                            BufferTile(i).Caption = ""
                            BufferTile(i).ToolTipText = ""
                            BufferTile(i).BackColor = &H0&
                            BufferTile(i).Visible = False
                        Next i
                        CurrentTurn = "P1"
                        TurnStatus.Caption = "Player 1's Turn"
                        Call ValidateTile(1)
                    Else
                        Call ClearAllTiles
                    End If
                Else
                    If SoundOn = True Then
                        Call sndPlaySound(App.Path & "\Sounds\Alert.wav", 1)
                    End If
                    For x = 0 To 2
                        BufferTile(x).Visible = True
                        BufferTile(x).Caption = ""
                        BufferTile(x).ToolTipText = ""
                        BufferTile(x).BackColor = &H0&
                    Next x
                    CurrentTurn = "P4"
                End If
        xCount = 0
        End If
    End If
End Sub

' PLAYER 4 GAME CLOCK TIMER LIMIT
Private Sub GameClocker_Timer()
    If Not GameReset Then
        GameClockerTime.Caption = Val(GameClockerTime.Caption) + 1
        If Val(GameClockerTime.Caption) >= 10 Then
            fMain.cmdDrop.Enabled = False
            fMain.TileStatus.Visible = False
            fMain.GameClockerTime.Visible = False
            fMain.GameClockerCaption.Visible = False
            fMain.GameClocker.Enabled = False
            GameClockerTime.Caption = 0
            For i = 0 To 2
                BufferTile(i).Visible = False
                BufferTile(i).Caption = ""
                BufferTile(i).ToolTipText = ""
                BufferTile(i).BackColor = &H0&
            Next i
            'MsgBox "Player 4 has no move." & vbCrLf & _
            '        "Player 4 have exceeded the time limit!", vbOKOnly + vbInformation, "Player 4's Alert"
            MoveStatus.AddItem "Player 4 has no move."
            MoveStatus.AddItem "Player 4 have exceeded the time limit!"
            MoveStatus.ListIndex = MoveStatus.NewIndex
            If SoundOn = True Then
                Call sndPlaySound(App.Path & "\Sounds\Alert.wav", 1)
            End If
            Call Force_PickTiles(4)
            xCount = 0
            CurrentTurn = "P1"
            TurnStatus.Caption = "Player 1's Turn"
            Call ValidateTile(1)
        End If
    End If
End Sub

' HELP
Private Sub mnuHelpTopics_Click()
    Dim msg
    msg = MsgBox("TIRUMM can be played by one player only." & vbCrLf & _
                 "It is played with 106 tiles including two jokers." & vbCrLf & _
                 "It is divided in 4 different colors,26 tiles each" & vbCrLf & _
                 "of the 4 different colors. The tiles were numbered" & vbCrLf & _
                 "1 to 13 in 4 diff. colors. So for each color the" & vbCrLf & _
                 "value will be twice the number. The object of the" & vbCrLf & _
                 "game is to be the first player to dispose all the" & vbCrLf & _
                 "tiles from their rack, and gain a high score by" & vbCrLf & _
                 "catching opponents with full racks.", vbOKOnly + vbInformation, "Help Information - Part 1")
    msg = MsgBox("Each player draws one tile to decide the order of the" & vbCrLf & _
                 "play. The player who draws the highest will be the" & vbCrLf & _
                 "first to play. To commence laying tiles on the board," & vbCrLf & _
                 "each player must lay their arrange manipulation, either" & vbCrLf & _
                 "GROUP which consist of 3 tiles of the same number value" & vbCrLf & _
                 "but with each tile a different color or RUN consists of" & vbCrLf & _
                 "3 tiles in numerical sequence, all in the same color at" & vbCrLf & _
                 "every turn and if they failed to they must take 3 tiles" & vbCrLf & _
                 "from the pool as a penalty.", vbOKOnly + vbInformation, "Help Information - Part 2")
    msg = MsgBox("The losing players will add up the value of all the" & vbCrLf & _
                 "tiles they still holds on the racks and score this as" & vbCrLf & _
                 "a minus amount, the winner of the round will received" & vbCrLf & _
                 "a positive score equal to the total value of all the" & vbCrLf & _
                 "losers point. The player with the highest score is the" & vbCrLf & _
                 "overall winner. And if in case the player didn't" & vbCrLf & _
                 "dispose all the racks they have, and there is no tiles" & vbCrLf & _
                 "left on the board, the values of all the tiles will be" & vbCrLf & _
                 "counted and the one who got the lowest value may" & vbCrLf & _
                 "considered a winner.", vbOKOnly + vbInformation, "Help Information - Part 3")
    msg = MsgBox("Game Instructions: The first move will depend on the" & vbCrLf & _
                 "highest tile value picked. All player tiles are all" & vbCrLf & _
                 "automatically picked up to 13 tiles each. The object of" & vbCrLf & _
                 "the game is to complete a Trio or Straight combinations" & vbCrLf & _
                 "with a limit of 3 tiles each combination & drop it as" & vbCrLf & _
                 "quickly as possible to win the game with less tile value." & vbCrLf & _
                 "If a player does not have a move 3 tiles will be added to" & vbCrLf & _
                 "his Penalty Rack. Player 4 has a time limit of 10 seconds" & vbCrLf & _
                 "to move or else he be on penalty. The chosen winner will" & vbCrLf & _
                 "be added to the Top Score.", vbOKOnly + vbInformation, "Help Information - Part 4")
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Effects.wav", 1)
    End If
End Sub

' ABOUT
Private Sub mnuAboutAuthor_Click()
    If SoundOn = True Then
        Call sndPlaySound(App.Path & "\Sounds\Effects.wav", 1)
    End If
    fAbout.Show
End Sub

Public Function Initialize_Variables()
    xVT = 0: xP1 = 0: xP2 = 1: xP3 = 0: GameClocker.Enabled = False: GameClockerTime.Caption = 0
    xP1Ctr = 0: xP2Ctr = 0: xP3Ctr = 0: xP4Ctr = 0: CurrentTurn = "": TurnStatus.Caption = "Player's Turn"
    xCount = 0: xPickPlayer1Now = False: xPickPlayer2Now = False: xPickPlayer3Now = False: xPickPlayer4Now = False
    xPlayer1TilePickCount = 0: xPlayer2TilePickCount = 0: xPlayer3TilePickCount = 0: xPlayer4TilePickCount = 0
    GameLoopCount = 0: xDpVal = 1: xP1Val = 0: xP2Val = 0: xP3Val = 0: xP4Val = 0: FillInTiles = False: MoveStatus.Clear
    Call ClearAllTiles
End Function
    
Public Function ClearAllTiles()
    For x = 0 To 27
        AIPlayer1Tile(x).Visible = False
        AIPlayer1Tile(x).Caption = ""
        AIPlayer1Tile(x).ToolTipText = ""
        AIPlayer1Tile(x).BackColor = &H0&
        AIPlayer2Tile(x).Visible = False
        AIPlayer2Tile(x).Caption = ""
        AIPlayer2Tile(x).ToolTipText = ""
        AIPlayer2Tile(x).BackColor = &H0&
        AIPlayer3Tile(x).Visible = False
        AIPlayer3Tile(x).Caption = ""
        AIPlayer3Tile(x).ToolTipText = ""
        AIPlayer3Tile(x).BackColor = &H0&
        AIPlayer4Tile(x).Visible = False
        AIPlayer4Tile(x).Caption = ""
        AIPlayer4Tile(x).ToolTipText = ""
        AIPlayer4Tile(x).BackColor = &H0&
        AIPlayer1Cover(x).Visible = False
        AIPlayer2Cover(x).Visible = False
        AIPlayer3Cover(x).Visible = False
    Next x
    For x = 0 To 105
        Tiles(x).Visible = False
    Next x
    For x = 1 To 80
        DropTile(x).Visible = False
        DropTile(x).Caption = ""
        DropTile(x).ToolTipText = ""
        DropTile(x).BackColor = &H0&
    Next x
    cmdDrop.Enabled = False
    TileStatus.Visible = False
    GameClockerTime.Visible = False
    GameClockerCaption.Visible = False
    GameClocker.Enabled = False
    For i = 0 To 2
        BufferTile(i).Visible = False
    Next i
    'MsgBox ("Wait Current Turn-" & CurrentTurn & "Game Reset-" & GameReset)
End Function
 
' EXIT
Private Sub mnuExit_Click()
    End
End Sub

' SET SOUND OFF
Private Sub mnuSoundsOff_Click()
    SoundOn = False: mnuSoundsOn.Checked = False: mnuSoundsOff.Checked = True
End Sub

' SET SOUND ON
Private Sub mnuSoundsOn_Click()
    SoundOn = True: mnuSoundsOn.Checked = True: mnuSoundsOff.Checked = False
End Sub

' DISPLAY HIGHEST SCORE
Private Sub mnuTopScore_Click()
    fHighestScore.Show: fHighestScore.Visible = True
End Sub

