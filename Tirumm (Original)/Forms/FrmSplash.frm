VERSION 5.00
Begin VB.Form fSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   1245
   ClientTop       =   1440
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":000C
   ScaleHeight     =   4290
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   4275
      Left            =   15
      Top             =   15
      Width           =   7560
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright c (Your Company)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is protected by national and international copyright laws as described in Help About."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   7275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPlatformAndVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "for Win X Version x.x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   5445
   End
End
Attribute VB_Name = "fSplash"
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

'Stores the number of seconds elapsed since midnight to determine the display time of the Splash window.
Private msngSplashDisplayStartTime As Single

'Platform the application runs on (e.g. "Win 95").
Public Platform As String

Private Sub Form_Activate()
    On Error GoTo HandleErrors
    Refresh
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Private Sub Form_Load()
    On Error GoTo HandleErrors
    Call StopMuzik
    Call NonstopMuzik(App.Path + "\Sounds\Intro.wav")
    'Change the Screens mouse pointer for this application as we dont want the user thinking that
    'they can start working on another form while the Splash form is still displayed.
    Screen.MousePointer = vbHourglass
    '------------------------------------Assign propertys for the Splash form---------------------------------------------
    lblPlatformAndVersion = "For " & App.FileDescription
    
    lblCopyright = "Copyright " & Chr(169) & " " & Year(Now()) & " " & "Your Company"
    'Include the Applications Title in the splash forms caption so that when displayed at run-time the
    'applications title will appear in the Windows Task Bar. This is especially important with applications
    'with lengthy startup code because the user needs to be informed that the application has indeed started.
    'Note that the Splash forms 'ShowInTaskbar' property must be set to TRUE at design-time to achieve this.
    Caption = App.Title
    '----------------------------------------------------------------------------------------------------------------------
    'Assign start time of the display of the Splash window.
    msngSplashDisplayStartTime = Timer
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo HandleErrors
    'Num. sec. the Splash Window is displayed.
    Const cintDisplayTimeSeconds As Integer = 3
    'Loop until the Display Time has elpased - if the applications loading time took longer than
    'the display time it will not enter this loop.
    Do Until (Timer - msngSplashDisplayStartTime) > cintDisplayTimeSeconds
    Loop
    Screen.MousePointer = vbNormal
ExitHandleErrors:
    Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

