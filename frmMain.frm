VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "AEI Data Scrubber"
   ClientHeight    =   1080
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4680
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdScrub 
      Caption         =   "Scrub"
      Height          =   372
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   2532
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   372
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtAEIpath 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AEIFilePath As String

Private Sub cmdBrowse_Click()
    cmdlg.Filter = "AEI Data (*.txt) | *.txt"
    cmdlg.ShowOpen
    
    If cmdlg.FileName <> "" Then
        AEIFilePath = cmdlg.FileName
        txtAEIpath.Text = AEIFilePath
    End If
End Sub

Private Sub cmdScrub_Click()
    ScrubDATA (AEIFilePath)
End Sub


Sub ScrubDATA(FILEPATH As String)
    Dim FSO As New Scripting.FileSystemObject
    Dim FWrite, FRead As TextStream
    Dim SourceDirectory, SourceFileName, PathStr() As String
    Dim INDEX As Integer
    Dim TempData, CarInitial, TextOut, PrevCarInitial, AEIdata() As String
    Dim OrdinalNum, CarNum, AxleTimeStamp, TimeAdjustmentFactor, PrevCarNum As Double
    
    PathStr() = Split(FILEPATH, "\")
    SourceFileName = PathStr(UBound(PathStr))
    SourceDirectory = PathStr(0)
    If UBound(PathStr) > 2 Then
        For INDEX = 1 To (UBound(PathStr) - 1)
            SourceDirectory = SourceDirectory & "\" & PathStr(INDEX)
        Next INDEX
    End If
    
    Set FWrite = FSO.CreateTextFile(SourceDirectory & "\TEMP.dat", True)
    Set FRead = FSO.OpenTextFile(FILEPATH, ForReading, False)
    FWrite.WriteLine FRead.ReadLine 'Export the first line without any processing
    Do
        TempData = FRead.ReadLine
        AEIdata = Split(TempData, vbTab)
        OrdinalNum = Val(AEIdata(0))
        CarInitial = AEIdata(1)
        CarNum = Val(AEIdata(2))
        AxleTimeStamp = Val(AEIdata(3))
    Loop Until (CarNum >= 10000)
    TimeAdjustmentFactor = AxleTimeStamp
    PrevCarNum = CarNum
    PrevCarInitial = CarInitial
    OrdinalNum = 1
    'Export the first axle of the first car
    TextOut = OrdinalNum & vbTab & CarInitial & vbTab & CarNum & vbTab & Round((AxleTimeStamp - TimeAdjustmentFactor) * 1#, 4)
    FWrite.Write TextOut
    Do
        TempData = FRead.ReadLine
        If (Trim(TempData) = "") Then Exit Do
        AEIdata = Split(TempData, vbTab)
        CarInitial = AEIdata(1)
        CarNum = Val(AEIdata(2))
        AxleTimeStamp = Val(AEIdata(3))
        If CarInitial <> "RRRR" Then
            If CarNum <> PrevCarNum Then
                OrdinalNum = OrdinalNum + 1
                PrevCarNum = CarNum
                PrevCarInitial = CarInitial
            End If
            TextOut = vbCrLf & OrdinalNum & vbTab & CarInitial & vbTab & CarNum & vbTab & Round((AxleTimeStamp - TimeAdjustmentFactor) * 2.12, 4)
        Else
            TextOut = vbCrLf & OrdinalNum & vbTab & PrevCarInitial & vbTab & PrevCarNum & vbTab & Round((AxleTimeStamp - TimeAdjustmentFactor) * 2.12, 4)
        End If
        FWrite.Write TextOut
    Loop Until (FRead.AtEndOfStream)
    FRead.Close
    FWrite.Close
    
    If (MsgBox("Do you want to delete Original AEI file?", vbYesNo, "Delete") = vbYes) Then
        Kill FILEPATH
        FSO.MoveFile SourceDirectory & "\TEMP.dat", FILEPATH
    End If
    Set FSO = Nothing
End Sub
