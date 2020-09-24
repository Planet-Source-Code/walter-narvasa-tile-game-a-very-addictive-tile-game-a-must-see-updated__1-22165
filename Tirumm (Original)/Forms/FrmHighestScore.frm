VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fHighestScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Highest Score - TIRUMM: The Tile Rummy Game"
   ClientHeight    =   7035
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "FrmHighestScore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7035
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      MouseIcon       =   "FrmHighestScore.frx":08CA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   6280
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   6890
         _ExtentX        =   12144
         _ExtentY        =   11086
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   65535
         HeadLines       =   1
         RowHeight       =   24
         RowDividerStyle =   0
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "fHighestScore"
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

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1

Private Sub Form_Load()
    On Error GoTo ErrorLoad
    Dim db As Connection
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path + "\Database\TIRUMM.MDB;Jet OLEDB:Database Password=3773;"
    Set adoPrimaryRS = New Recordset
    adoPrimaryRS.Open "select Winner_Name,Winner_Score from Overall_Standings " & _
                    "Order by Winner_Score", db, adOpenStatic, adLockOptimistic
    adoPrimaryRS.Requery
    If adoPrimaryRS.RecordCount <> 0 Then
        Set grdDataGrid.DataSource = adoPrimaryRS
        grdDataGrid.Columns(0).Caption = "Player Name": grdDataGrid.Columns(1).Caption = "Score"
        grdDataGrid.Columns(0).Width = 5000: grdDataGrid.Columns(1).Width = 5000
    End If
    Exit Sub
ErrorLoad:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Warning"
End Sub

Public Sub UpdateDatabase()
    adoPrimaryRS.AddNew
    adoPrimaryRS("Winner_Name") = xCurrent_WinnerName
    adoPrimaryRS("Winner_Score") = xCurrent_WinnerScore
    adoPrimaryRS.UpdateBatch adAffectAll
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub
