VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPEG Image Saver - By VB Dude"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About JPEG Saver"
      Height          =   330
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   330
      Left            =   6240
      TabIndex        =   6
      Top             =   3840
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3360
      TabIndex        =   5
      Top             =   3840
      Width           =   945
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Begin Saver"
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Width           =   1425
   End
   Begin VB.PictureBox pbProgress 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H8000000D&
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7230
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4695
      Width           =   7290
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3600
      MultiSelect     =   2  'Extended
      Pattern         =   "*.jp*"
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0442
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   5640
      Left            =   0
      Picture         =   "frmMain.frx":1F3D
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iResizeFlag As Integer
Dim lButt As Long
Dim lDir As Long
Dim lFile As Long
Dim lFileLeft As Long
Dim lFileWidth As Long
Dim lLabelWidth As Long
Dim lLabelTop As Long
Dim lButtonTop As Long
Dim iNumSelected As Integer
Dim bWorking As Boolean

Private Sub cmdAbout_Click()
MsgBox "-------------------------------------------------------------------------------------------------------" & _
vbCrLf & "JPEG Image Saver - By VB Dude" & _
vbCrLf & "-------------------------------------------------------------------------------------------------------" & _
vbCrLf & "(c) VB Dude 2000. By Reynard Chan. Age: 12" & _
vbCrLf & "Made in: Thursday, 4th January, 2001. 8:28pm AUS Time" & _
vbCrLf & "This code can be accessed at: www.planet-source-code.com/vb/" & _
vbCrLf & "and in search type in: JPEG Image Saver" & _
vbCrLf & "-------------------------------------------------------------------------------------------------------" & _
vbCrLf & "If you are going to use this code for commercial use or for fun, please " & _
vbCrLf & "E-mail me at: reychan@hotmail.com and I might help you out. You need to ask!!!", vbInformation, "JPEG Image Saver"
End Sub

Private Sub cmdCancel_Click()
    bCancelFlag = 1
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    bWorking = True
    Dim iCount As Integer
    Dim iFiles As Integer
    ChDrive Drive1.Drive
    ChDir Dir1.Path
    Dir1.Enabled = False
    Drive1.Enabled = False
    File1.Enabled = False
    MousePointer = vbHourglass
    cmdQuit.Enabled = False
    cmdStart.Enabled = False
    cmdCancel.Enabled = True
    ReDim sFiles(iNumSelected)
    For iCount = 0 To File1.ListCount - 1
        If File1.Selected(iCount) = True Then
            sFiles(iFiles) = File1.List(iCount)
            iFiles = iFiles + 1
        End If
    Next iCount
    If iFiles = 0 Then GoTo skip:
    DoIt (iFiles - 1)
skip:
    cmdQuit.Enabled = True
    cmdCancel.Enabled = False
    Dir1.Enabled = True
    Drive1.Enabled = True
    File1.Enabled = True
    MousePointer = vbDefault
    bWorking = False
    pbProgress.Cls
    RefreshFiles
End Sub

Private Sub Dir1_Change()
    On Error GoTo er
    ChDir Dir1.Path
    RefreshFiles
Exit Sub
er:
    NewPath
End Sub

Private Sub Drive1_Change()
    On Error GoTo er
    ChDrive Drive1.Drive
    RefreshFiles
Exit Sub
er:
    NewPath
End Sub

Sub RefreshFiles()
    Drive1.Drive = CurDir
    Dir1.Path = CurDir
    File1.Path = CurDir
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
    CountSelectedFiles
End Sub

Private Sub File1_Click()
    CountSelectedFiles
End Sub

Private Sub Form_Load()
    iResizeFlag = 0
    AdjustResize
    iResizeFlag = 1
    CountSelectedFiles
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bWorking = True Then
        Cancel = vbCancel
    End If
End Sub

Sub Form_Resize()
    Dim lFormHeight As Long
    Dim lFormWidth As Long
    lFormHeight = frmMain.Height
    lFormWidth = frmMain.Width
    If iResizeFlag = 1 Then
        If lFormHeight < 4000 Then lFormHeight = 4000
        If lFormWidth < 4000 Then lFormWidth = 4000
        lblMessage.Top = lFormHeight - lLabelTop
        Dir1.Height = lFormHeight - lDir
        File1.Height = lFormHeight - lFile
        File1.Width = lFormWidth - lFileWidth
        lblMessage.Width = lFormWidth - lLabelWidth
        cmdStart.Top = lFormHeight - lButtonTop
        cmdCancel.Top = lFormHeight - lButtonTop
        cmdQuit.Top = lFormHeight - lButtonTop
    End If
End Sub

Private Sub AdjustResize()
    lButt = frmMain.Height - cmdStart.Top
    lDir = frmMain.Height - Dir1.Height
    lFile = frmMain.Height - File1.Height
    lFileLeft = frmMain.Width - File1.Left
    lFileWidth = frmMain.Width - File1.Width
    lLabelWidth = frmMain.Width - lblMessage.Width
    lLabelTop = frmMain.Height - lblMessage.Top
    lButtonTop = frmMain.Height - cmdStart.Top
End Sub
Public Sub pUpdate(total As Long, progress As Long)
    If progress > pbProgress.ScaleWidth Then
        progress = pbProgress.ScaleWidth
    End If
    If total < 1 Then total = 1
    pbProgress.ScaleWidth = total
    pbProgress.Line (0, 0)-(progress, pbProgress.ScaleHeight), pbProgress.ForeColor, BF
    DoEvents
End Sub
Private Sub CountSelectedFiles()
    Dim iCount As Integer
    Dim iSelected As Integer
    For iCount = 0 To File1.ListCount - 1
        If File1.Selected(iCount) = True Then
            iSelected = iSelected + 1
        End If
    Next iCount
    iNumSelected = iSelected
    lblMessage.Caption = iSelected & "  selected,     " & File1.ListCount & "  total"
    If iNumSelected < 1 Then
        cmdStart.Enabled = False
    Else
        cmdStart.Enabled = True
    End If
End Sub
Private Sub NewPath()
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
    RefreshFiles
End Sub

