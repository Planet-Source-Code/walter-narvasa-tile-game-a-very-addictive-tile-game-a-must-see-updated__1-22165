VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About - TIRUMM: The Tile Rummy Game"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2753
      MouseIcon       =   "FrmAbout.frx":08CA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer ReDrawTimer 
      Interval        =   1
      Left            =   7620
      Top             =   5760
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   113
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "fAbout"
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

Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCCOPY = &HCC0020

Dim Tempstring(1 To 3000) As Variant
Dim ipicHeight As Integer
Dim ipicWidth As Integer
Dim lYOffset As Integer
Dim iColorCur As Single
Dim iColorStep As Single
Dim NumLines As Integer
Dim lX As Long
Dim lY As Long
Dim strRead As String

Private Sub Form_Load()
    On Error Resume Next
    Dim iLine As Integer
    NumLines = 1
    fAbout.ScaleMode = vbPixels
    picBuffer.ScaleMode = vbPixels
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    picBuffer.Visible = False
    Open (App.Path & "\Database\CREDITS.TXT") For Input As #1
    Do Until EOF(1)
        Line Input #1, Tempstring(NumLines)
        NumLines = NumLines + 1
    Loop
    Close #1
    NumLines = NumLines - 1
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    GradiantBackground picBackBuffer
    ReDrawTimer.Interval = 5
    ReDrawTimer.Enabled = True
End Sub


Private Function GradiantBackground(picBox As PictureBox)
    ipicWidth = picBox.ScaleWidth
    ipicHeight = picBox.ScaleHeight
    iColorCur = 255
    iColorStep = 5 * (0 - 255) / ipicHeight
    For lYOffset = 0 To ipicHeight Step 5
        picBox.Line (-1, lYOffset - 1)-(ipicWidth, lYOffset + 5), RGB(0, 0, iColorCur), BF
        iColorCur = iColorCur + iColorStep
    Next lYOffset
End Function

Private Sub RedrawTimer_Timer()
    Dim l As Long
    Dim j As Long
    On Error Resume Next
    l = BitBlt(picBuffer.hDC, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hDC, 0, 0, SRCCOPY)
    For j = 1 To NumLines Step 1
        picBuffer.CurrentY = lY + (j * picBuffer.FontSize + (6 * j))
        picBuffer.CurrentX = (picBuffer.ScaleWidth / 2) - (picBuffer.TextWidth(Tempstring(j)) / 2)
        picBuffer.ForeColor = vbWhite
        If picBuffer.CurrentY < 245 Then
            If picBuffer.CurrentY > 15 Then
                picBuffer.ForeColor = RGB((((255 / 235) * picBuffer.CurrentY)), (((255 / 235) * picBuffer.CurrentY)), (((255 / 25) * picBuffer.CurrentY)))
            Else
                picBuffer.ForeColor = vbBlack
                If j = NumLines And picBuffer.CurrentY < -25 Then
                    ReDrawTimer.Enabled = False
                    Unload Me
                End If
            End If
        End If
        picBuffer.Print Tempstring(j)
    Next
    l = BitBlt(picOut.hDC, 0, picOut.ScaleTop, picOut.ScaleWidth, picOut.ScaleHeight, picBuffer.hDC, 0, 0, SRCCOPY)
    picOut.Refresh
    lY = lY - 1
End Sub

Private Sub cmdOk_Click()
    ReDrawTimer.Enabled = False
    Unload Me
End Sub


