Attribute VB_Name = "Module1"
Option Explicit
'Code courtsey of JPEG Strip
Public sFiles() As String
Public bCancelFlag As Byte

Dim bIn() As Byte
Dim bOut() As Byte
Dim lPos As Long
Dim sLogMsg As String
Dim lFileSize As Long
Dim lFileOutSize As Long

Const OKAY As Long = 0
Const ERROR As Long = -1
Const DONE As Long = -2
Const PROBLEM As Long = -3

Private Function DoAFile(sFileIn As String) As Long
    Dim lret As Long
    Dim boolProbFlag As Boolean
    Dim sFileOut As String
    sFileOut = sFileIn & ".tmp"
    lFileOutSize = 0
    lPos = 1
    lret = ReadFileSize(sFileIn)                ' returns filesize or ERROR
    If lret = ERROR Then
        sLogMsg = "Couldn't open input file."
        GoTo UhOh
    End If
    lFileSize = lret
    ReDim bIn(1 To lFileSize + 10)              ' dim variables, 1 based, with some extra space
    ReDim bOut(1 To lFileSize + 10)
    lret = ReadFile(sFileIn)                        ' read the file into bIn
    If lret = ERROR Then
        sLogMsg = "Couldn't open input file."
        GoTo UhOh
    End If
    lret = FindJpgHeader()                         ' find the jpg header
    If lret = ERROR Then
        sLogMsg = "Not a valid JPEG file."
        GoTo UhOh
    End If
    lret = 0
    Do Until lret = DONE Or lret = ERROR Or lret = PROBLEM
        lret = GetMarkers()                         ' copy needed data
    Loop
    If lret = ERROR Then
        sLogMsg = "Problem parsing file."
        GoTo UhOh
    End If
    If lret = PROBLEM Then boolProbFlag = True
    lret = WriteOutFile(sFileOut)                   ' write output file
    If lret = ERROR Then
        sLogMsg = "Could not write output file"
        GoTo UhOh
    End If
    lret = KillInFile(sFileIn)                          ' delete input file
    If lret = ERROR Then
        sLogMsg = "Could not delete original file."
        GoTo UhOh
    End If
    lret = ReNameOutFile(sFileOut, sFileIn)    ' rename output file to original name
    If lret = ERROR Then
        sLogMsg = "Could not rename output file to original filename."
        GoTo UhOh
    End If
SkipRename:
    If boolProbFlag = True Then
        sLogMsg = "processed successfully, but problems encountered. Check file."
        DoAFile = PROBLEM
    Else
        sLogMsg = "Processed successfully."
        DoAFile = OKAY
    End If
Exit Function
UhOh:                                                   ' Houston, we have a problem
    DoAFile = ERROR
End Function

Private Function ReadFileSize(sFileName As String) As Long
    On Error GoTo HandleIt
    ReadFileSize = FileLen(sFileName)
Exit Function
HandleIt:
    ReadFileSize = ERROR
End Function

Private Function ReadFile(sFileName As String) As Long
    Dim iFN As Integer
    On Error GoTo HandleIt
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bIn()
    Close iFN
    ReadFile = OKAY
    Exit Function
HandleIt:
    ReadFile = ERROR
End Function

Private Function FindJpgHeader() As Long
    Do
        If bIn(lPos) = &HFF And bIn(lPos + 1) = &HD8 And bIn(lPos + 2) = &HFF Then
            FindJpgHeader = OKAY
            Exit Do
        End If
        If lPos >= lFileSize Then
            FindJpgHeader = ERROR
            Exit Do
        End If
        lPos = lPos + 1
    Loop
    lPos = lPos + 1
End Function

Private Function GetMarkers() As Long
    Dim lSkip As Long
    Dim lTemp As Long
    Dim bFlag As Byte
    Select Case bIn(lPos)
        Case &HD8
            WriteArray &HFF
            WriteArray &HD8
            WriteArray &HFF
            lSkip = 2
        Case &HE0, &HDB, &HC0 To &HCB, &HDD
            lSkip = Mult(bIn(lPos + 2), bIn(lPos + 1)) + 1
            If lSkip + lPos >= lFileSize Then GoTo Oops
            For lTemp = lPos To lPos + lSkip
                WriteArray bIn(lTemp)
            Next lTemp
        Case &HDA
            bFlag = 1
            Do
                WriteArray bIn(lPos)
                If bIn(lPos + 1) = &HFF And bIn(lPos + 2) = &HD9 Then Exit Do
                lPos = lPos + 1
                If lPos > lFileSize Then
                    bFlag = 2
                    Exit Do
                End If
            Loop
            WriteArray &HFF
            WriteArray &HD9
        Case Else
            lSkip = Mult(bIn(lPos + 2), bIn(lPos + 1)) + 1
            If lSkip + lPos > lFileSize Then GoTo Oops
    End Select
    lPos = lPos + lSkip
    Do
        If bIn(lPos) <> &HFF Then Exit Do
        lPos = lPos + 1
        If lPos > lFileSize Then GoTo Oops
    Loop
    If bFlag = 0 Then GetMarkers = OKAY
    If bFlag = 1 Then GetMarkers = DONE
    If bFlag = 2 Then GetMarkers = PROBLEM
Exit Function
Oops:
GetMarkers = ERROR
End Function

Private Function WriteOutFile(sFileName As String) As Long
    Dim iFN As Integer
    On Error GoTo NoOpen
    iFN = FreeFile
    ReDim Preserve bOut(1 To lFileOutSize)
    Open sFileName For Binary As iFN
    On Error GoTo Opened
    Put #iFN, , bOut()
    Close iFN
    WriteOutFile = OKAY
Exit Function
NoOpen:
    WriteOutFile = ERROR
Exit Function
Opened:
    Close iFN
    WriteOutFile = ERROR
End Function

Private Function KillInFile(sFileName As String) As Long
    On Error GoTo HandleIt
    Kill sFileName
    KillInFile = OKAY
Exit Function
HandleIt:
    KillInFile = ERROR
End Function

Private Function ReNameOutFile(sFileOld As String, sFileNew As String) As Long
    On Error GoTo HandleIt
    Name sFileOld As sFileNew
    ReNameOutFile = OKAY
Exit Function
HandleIt:
    ReNameOutFile = ERROR
End Function

Private Sub WriteArray(bData As Byte)
    lFileOutSize = lFileOutSize + 1
    bOut(lFileOutSize) = bData
End Sub

Private Function Mult(lsb As Byte, msb As Byte) As Long
    Mult = CLng(lsb) + (CLng(msb) * 256&)
End Function
Public Sub DoIt(lNumber As Long)
    On Error Resume Next
    Dim lCount As Long
    Dim iFN As Integer
    Dim lret As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim lDiff As Long
    Dim lTotal As Long
    Dim sDone As String
    Dim lFilesDone As Long
    iFN = FreeFile
    Open "js.log" For Output As iFN
    For lCount = 0 To lNumber
        lBefore = FileLen(sFiles(lCount))
        lret = DoAFile(sFiles(lCount))
        lAfter = FileLen(sFiles(lCount))
        frmMain.lblMessage.Caption = sFiles(lCount)
        frmMain.pUpdate lNumber, lCount
        lDiff = lBefore - lAfter
        lTotal = lTotal + lDiff
        Print #iFN, "**************"
        Print #iFN, sFiles(lCount)
        Print #iFN, sLogMsg
        Print #iFN, lBefore
        Print #iFN, lAfter
        Print #iFN, lDiff
        DoEvents
        lFilesDone = lFilesDone + 1
        If bCancelFlag = 1 Then Exit For
    Next lCount
    bCancelFlag = 0
    sDone = "--------------------------------------" & vbCrLf & "Files processed:" & lFilesDone _
        & vbCrLf & "Total bytes saved:" & lTotal & vbCrLf & "--------------------------------------"
    Print #iFN, sDone
    MsgBox sDone, vbOKOnly, "Finished JPEG Saver"
    Close iFN
End Sub


