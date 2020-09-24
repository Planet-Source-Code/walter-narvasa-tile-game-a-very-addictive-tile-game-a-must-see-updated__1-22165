Attribute VB_Name = "mStartup"
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

Global xVT As Integer
Global xP1 As Integer
Global xP2 As Integer
Global xP3 As Integer
Global xP4 As Integer
Global xAIPlayer1Tile As Integer
Global xAIPlayer2Tile As Integer
Global xAIPlayer3Tile As Integer
Global xAIPlayer4Tile As Integer
Global xPickPlayer1Now As Boolean
Global xPickPlayer2Now As Boolean
Global xPickPlayer3Now As Boolean
Global xPickPlayer4Now As Boolean
Global xPlayer1TilePickCount As Integer
Global xPlayer2TilePickCount As Integer
Global xPlayer3TilePickCount As Integer
Global xPlayer4TilePickCount As Integer
Global x1stMove As String
Global xTemp(3)
Global xPlayer1Temp(13) As String
Global xPlayer2Temp(13) As String
Global xPlayer3Temp(13) As String
Global xPlayer1Out(106)
Global xPlayer2Out(106)
Global xPlayer3Out(106)
Global xPlayer4Out(106)
Global xDpVal As Integer
Global xP1Val As Integer
Global xP2Val As Integer
Global xP3Val As Integer
Global xP4Val As Integer
Global HaltAITimer As Boolean
Global GameReset As Boolean
Global GameLoopCount As Integer
Global CurrentTurn As String
Global FillInTiles As Boolean
Global xP1Ctr As Integer
Global xP2Ctr As Integer
Global xP3Ctr As Integer
Global xP4Ctr As Integer
Global SoundOn As Boolean
Global xCurrent_WinnerName As String
Global xCurrent_WinnerScore As Integer

Sub Main()
    On Error GoTo HandleErrors
    fSplash.Platform = pcstrAppPlatform
    fSplash.Show
    'Ensure the Splash form is refreshed prior to displaying the Main form.
    DoEvents
    '---------------------------------------------------------------------------------------------------------------------
    'Perform other start up tasks here...
    'For demo purposes we add a delay to simulate a typical applications initialisation.
    Call SplashDelay
  '---------------------------------------------------------------------------------------------------------------------
    fMain.Show
    fHighestScore.Show: fHighestScore.Visible = False
    DoEvents
    Unload fSplash
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Public Sub SplashDelay()
    On Error Resume Next
    Dim sngStartTime As Single
    sngStartTime = Timer
    Do Until (Timer - sngStartTime) > 4
          DoEvents
    Loop
End Sub




