Attribute VB_Name = "mFunctions"
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

' FOR PLAYING SOUNDS
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim sound As String
Const SND_SYNC = &H0&
Const SND_ASYNC = &H1&
Const SND_NODEFAULT = &H2&
Const SND_LOOP = &H8&
Const SND_NOSTOP = &H10&

' NON-STOP CURRENT PLAYING SOUNDS
Function NonstopMuzik(Soundfile As String)
    wFlags% = SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
    x% = sndPlaySound(Soundfile$, wFlags%)
End Function

' STOP CURRENT SOUNDS
Function StopMuzik()
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    x% = sndPlaySound(Soundfile$, wFlags%)
End Function

Public Function iGetTileValue(iTileNum As Integer) As String
    Select Case iTileNum
            ' BATCH 1 (RED TILE)
            Case 0: iGetTileValue = "1"
            Case 1: iGetTileValue = "2"
            Case 2: iGetTileValue = "3"
            Case 3: iGetTileValue = "4"
            Case 4: iGetTileValue = "5"
            Case 5: iGetTileValue = "6"
            Case 6: iGetTileValue = "7"
            Case 7: iGetTileValue = "8"
            Case 8: iGetTileValue = "9"
            Case 9: iGetTileValue = "10"
            Case 10: iGetTileValue = "11"
            Case 11: iGetTileValue = "12"
            Case 12: iGetTileValue = "13"
            Case 13: iGetTileValue = "1"
            Case 14: iGetTileValue = "2"
            Case 15: iGetTileValue = "3"
            Case 16: iGetTileValue = "4"
            Case 17: iGetTileValue = "5"
            Case 18: iGetTileValue = "6"
            Case 19: iGetTileValue = "7"
            Case 20: iGetTileValue = "8"
            Case 21: iGetTileValue = "9"
            Case 22: iGetTileValue = "10"
            Case 23: iGetTileValue = "11"
            Case 24: iGetTileValue = "12"
            Case 25: iGetTileValue = "13"
            ' BATCH 2 (GREEN TILE)
            Case 26: iGetTileValue = "1"
            Case 27: iGetTileValue = "2"
            Case 28: iGetTileValue = "3"
            Case 29: iGetTileValue = "4"
            Case 30: iGetTileValue = "5"
            Case 31: iGetTileValue = "6"
            Case 32: iGetTileValue = "7"
            Case 33: iGetTileValue = "8"
            Case 34: iGetTileValue = "9"
            Case 35: iGetTileValue = "10"
            Case 36: iGetTileValue = "11"
            Case 37: iGetTileValue = "12"
            Case 38: iGetTileValue = "13"
            Case 39: iGetTileValue = "1"
            Case 40: iGetTileValue = "2"
            Case 41: iGetTileValue = "3"
            Case 42: iGetTileValue = "4"
            Case 43: iGetTileValue = "5"
            Case 44: iGetTileValue = "6"
            Case 45: iGetTileValue = "7"
            Case 46: iGetTileValue = "8"
            Case 47: iGetTileValue = "9"
            Case 48: iGetTileValue = "10"
            Case 49: iGetTileValue = "11"
            Case 50: iGetTileValue = "12"
            Case 51: iGetTileValue = "13"
            ' BATCH 3 (BLUE TILE)
            Case 52: iGetTileValue = "1"
            Case 53: iGetTileValue = "2"
            Case 54: iGetTileValue = "3"
            Case 55: iGetTileValue = "4"
            Case 56: iGetTileValue = "5"
            Case 57: iGetTileValue = "6"
            Case 58: iGetTileValue = "7"
            Case 59: iGetTileValue = "8"
            Case 60: iGetTileValue = "9"
            Case 61: iGetTileValue = "10"
            Case 62: iGetTileValue = "11"
            Case 63: iGetTileValue = "12"
            Case 64: iGetTileValue = "13"
            Case 65: iGetTileValue = "1"
            Case 66: iGetTileValue = "2"
            Case 67: iGetTileValue = "3"
            Case 68: iGetTileValue = "4"
            Case 69: iGetTileValue = "5"
            Case 70: iGetTileValue = "6"
            Case 71: iGetTileValue = "7"
            Case 72: iGetTileValue = "8"
            Case 73: iGetTileValue = "9"
            Case 74: iGetTileValue = "10"
            Case 75: iGetTileValue = "11"
            Case 76: iGetTileValue = "12"
            Case 77: iGetTileValue = "13"
            ' BATCH 4 (YELLOW TILE)
            Case 78: iGetTileValue = "1"
            Case 79: iGetTileValue = "2"
            Case 80: iGetTileValue = "3"
            Case 81: iGetTileValue = "4"
            Case 82: iGetTileValue = "5"
            Case 83: iGetTileValue = "6"
            Case 84: iGetTileValue = "7"
            Case 85: iGetTileValue = "8"
            Case 86: iGetTileValue = "9"
            Case 87: iGetTileValue = "10"
            Case 88: iGetTileValue = "11"
            Case 89: iGetTileValue = "12"
            Case 90: iGetTileValue = "13"
            Case 91: iGetTileValue = "1"
            Case 92: iGetTileValue = "2"
            Case 93: iGetTileValue = "3"
            Case 94: iGetTileValue = "4"
            Case 95: iGetTileValue = "5"
            Case 96: iGetTileValue = "6"
            Case 97: iGetTileValue = "7"
            Case 98: iGetTileValue = "8"
            Case 99: iGetTileValue = "9"
            Case 100: iGetTileValue = "10"
            Case 101: iGetTileValue = "11"
            Case 102: iGetTileValue = "12"
            Case 103: iGetTileValue = "13"
            ' BATCH 5 (ORANGE TILE)
            Case 104: iGetTileValue = "J"
            Case 105: iGetTileValue = "J"
    End Select
End Function

Public Function BubbleSortVariantArray(ByRef varArray As Variant, Optional ByVal boolDesc As Boolean)
    ' Bubble-sorts the passed variant array
    Dim intLBound As Integer, intUBound As Integer
    Dim intX As Integer, intY As Integer
    On Error Resume Next
    intLBound = LBound(varArray)
    intUBound = UBound(varArray)
    If intUBound >= 0 Then
        For intX = intLBound To intUBound - 1
            DoEvents
            For intY = intX + 1 To intUBound
                DoEvents
                AEvalSwap varArray, intX, intY, boolDesc
            Next intY
        Next intX
    End If
End Function

Public Sub QuickSortVariantArray(ByRef varArray As Variant, ByVal intLBound As Integer, ByVal intUBound As Integer, Optional ByVal boolDesc As Boolean)
    ' Quicksorts the passed array of Variants
    Dim intX As Integer, intY As Integer, varMidBound As Variant
    Dim varTemp As Variant
    On Error Resume Next
    If intUBound >= 0 And intUBound > intLBound Then
        ' Calculate the value of the middle array element
        varMidBound = varArray((intLBound + intUBound) \ 2)
        intX = intLBound
        intY = intUBound
        ' Split the array into halves
        Do While intX <= intY
            DoEvents
            If boolDesc Then
                If varArray(intX) <= varMidBound And varArray(intY) >= varMidBound Then
                    ASwap varArray, intX, intY
                    intX = intX + 1
                    intY = intY - 1
                Else
                    If varArray(intX) > varMidBound Then intX = intX + 1
                    If varArray(intY) < varMidBound Then intY = intY - 1
                End If
            Else
                If varArray(intX) >= varMidBound And varArray(intY) <= varMidBound Then
                    ASwap varArray, intX, intY
                    intX = intX + 1
                    intY = intY - 1
                Else
                    If varArray(intX) < varMidBound Then intX = intX + 1
                    If varArray(intY) > varMidBound Then intY = intY - 1
                End If
            End If
        Loop
        ' Sort the lower half of the array
        QuickSortVariantArray varArray, intLBound, intY, boolDesc
        ' Sort the upper half of the array
        QuickSortVariantArray varArray, intX, intUBound, boolDesc
    End If
End Sub

Public Sub ShellSortVariantArray(ByRef varArray As Variant, Optional ByVal boolDesc As Boolean)
    ' Sorts the passed variant array using the shell-sort algorithm
    Dim intLBound As Integer, intUBound As Integer
    Dim intX As Integer, intY As Integer
    On Error Resume Next
    intLBound = LBound(varArray)
    intUBound = UBound(varArray)
    If intUBound >= 0 Then
        ' Get the middle of the array
        intY = (intUBound - intLBound + 1) \ 2
        Do While intY > 0
            DoEvents
            ' Sort the lower portion of the array
            For intX = intLBound To intUBound - intY
                DoEvents
                AEvalSwap varArray, intX, intX + intY, boolDesc
            Next intX
            ' Sort the upper portion of the array
            For intX = intUBound - intY To intLBound Step -1
                DoEvents
                AEvalSwap varArray, intX, intX + intY, boolDesc
            Next intX
            ' Divide the array into smaller portions for the next loop
            intY = intY \ 2
        Loop
    End If
End Sub

Private Sub AEvalSwap(ByRef varArray As Variant, ByVal intX As Integer, ByVal intY As Integer, Optional ByVal boolDesc As Boolean)
    ' Evaluate the swap to support ascending or descending sort
    On Error Resume Next
    If boolDesc Then
        ' Sort descending
        If varArray(intX) < varArray(intY) Then ASwap varArray, intX, intY
    Else
        ' Sort ascending
        If varArray(intX) > varArray(intY) Then ASwap varArray, intX, intY
    End If
End Sub

Private Sub ASwap(ByRef varArray As Variant, ByVal intX As Integer, ByVal intY As Integer)
    ' Swap values between two items in an array
    Dim varTemp As Variant
    On Error Resume Next
    varTemp = varArray(intX)
    varArray(intX) = varArray(intY)
    varArray(intY) = varTemp
End Sub

Public Sub ASort(ByRef varArray As Variant, Optional ByVal boolDesc As Boolean, Optional ByVal boolMin As Boolean)
    ' Sort an array ascending by default
    ' Although the sort procedures called here are declared
    '   Public, use ASort() to more effectively implement them
    ' The default sorting procedure uses the quick sort algorithm
    On Error Resume Next
    If UBound(varArray) <= 10 And boolMin Then
        ' Ideal for small arrays
        BubbleSortVariantArray varArray, boolDesc
    ElseIf UBound(varArray) <= 100 And boolMin Then
        ' Faster than bubble sort but still ideal for small arrays
        ShellSortVariantArray varArray, boolDesc
    Else
        ' Ideal for large arrays but may risk memory and stack overflow
        QuickSortVariantArray varArray, LBound(varArray), UBound(varArray), boolDesc
    End If
End Sub

Public Sub SortStringArray(ByRef Arr() As String, ByVal ascending As Boolean)
    Dim l As Long
    Dim r As Long
    l = 0
    r = UBound(Arr)
    If (ascending) Then
        Call QuickSort(Arr, l, r, 1)
    Else
        Call QuickSort(Arr, l, r, -1)
    End If
End Sub

Private Sub QuickSort(ByRef Arr() As String, ByVal l As Long, ByVal r As Long, ByVal flag As Integer)
    If (r <= l) Then Exit Sub
    Dim i As Long
    Dim j As Long
    Dim temp As String
    Dim ret As Integer
    i = l - 1
    j = r
    Do While (True)
        Do
            i = i + 1
            ret = StrComp(Arr(i), Arr(r))
            ret = ret * flag
        Loop While (ret < 0)
        Do While (j > 0)
            j = j - 1
            ret = StrComp(Arr(j), Arr(r))
            ret = ret * flag
            If (ret <= 0) Then Exit Do
        Loop
        If (i > j) Then Exit Do
        temp = Arr(i)
        Arr(i) = Arr(j)
        Arr(j) = temp
    Loop
    temp = Arr(i)
    Arr(i) = Arr(r)
    Arr(r) = temp
    Call QuickSort(Arr, l, i - 1, flag)
    Call QuickSort(Arr, i + 1, r, flag)
End Sub

Function ExArg(ArgNum As Integer, srchstr As String, Delim As String) As String
    'Extract an argument or token from a string based on its position and a delimiter.
    On Error GoTo Err_ExArg
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr
        Do While InStr(srchstr, Delim) > 0
            Pos = InStr(LastPos, srchstr, Delim)
        If Pos = 0 Then
            'No More Args found
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1
            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    '---------
    ExArg = Arg
    Exit Function
Err_ExArg:
    MsgBox "Error " & Err & ": " & Error
    Resume Next
End Function

