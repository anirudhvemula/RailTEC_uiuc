VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmFirst 
   Caption         =   "Train Scoring System"
   ClientHeight    =   4536
   ClientLeft      =   6660
   ClientTop       =   3840
   ClientWidth     =   7800
   Icon            =   "frmFirst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4536
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdAbort 
      BackColor       =   &H000000FF&
      Caption         =   "Abort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.CommandButton cmdSetOutPath 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6360
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtOutFName 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox chkAutoClose 
      Caption         =   "Close command window automatically"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   4245
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   508
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13251
            MinWidth        =   5292
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUmler 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtUmler 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox txtOutPath 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   5295
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H0080FF80&
      Caption         =   "Start generating Scoring Sytem result"
      Height          =   495
      Left            =   360
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   6375
   End
   Begin MSComDlg.CommonDialog cdlgPath 
      Left            =   7200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGap 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAero 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtGap 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtAero 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4. Specify output path & filename"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "3. Specify MiniUmler database"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2145
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2. Specify TMS Output/Load Timestamp"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. Specify AEI Axle Timestamp"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2160
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE TO SELF FOR FUTURE:
'1) Use errMsg system for error output and reporting, don't insert them among other operations.
'2) These need to be seperated from each other: Core Operation, Generating Output, Validating Input, Error Handling and Reporting. Use integer ID or bit-flags to link them.

Option Explicit
Option Base 1

'Global Variables Declaration
'============================

'Data columns in input file for AeroDynamic executable
Dim InputColumn(16) As String

'Dynamic data array to remember the number of axles
Dim AxlesNumCount() As Integer

'Multi-purpose temporary variables
Dim tempStr, gapStr, inputStr As String
Dim tempVal As Double

'Variables for train-scoring
Dim CarInit, CarNum, Attr1, Attr2, Attr3, CarType, AeroInput As String
Dim CargoLength, carStart, carStop, CarLength, TempDouble, BestUpper, BestLower, BaseTotalUpper, BaseTotalLower, AeroScore, AccAeroScore, timeStart, timeStop, axleStart, axleEnd, bStart, bEnd, PrevTimeStamp As Double
Dim UnitCount, SlotCount, AxlesCount, EmptyCount, Base, rCount, cCount, MaxrCount, AxlesNum, Ordinal As Integer

'File paths
Dim fnGap, fnAxles, fnUmler, fnAeroIn, fnAeroInF, fnFinal, fnFinalF, fnProcess, fnReport As String

'File handlers
Dim fso, fAxles, fGap, fOut, fOutF, fFinal, fFinalF, fReport

'Databases
Dim dbUmler, dbProcess As Database

'Recordsets
Dim rsUmler, rsBestLoad, rsAxles, rsGap As Recordset

'Flags
Dim IgnorePart2, LoadFit, MuteAERO, GapEnd, TimeStampError As Boolean
'******************************************************************************
'************************ Start modification/addition by JR *******************
'******************************************************************************
Dim sQueryString As String
Dim sTempName As String
Dim BlindMode As Boolean
Dim AbortRun As Boolean

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" _
    (ByVal hObject As Long) As Long

Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Integer) As _
    Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal _
    lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar _
    As Byte) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" _
    (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
    (iccex As tagInitCommonControlsEx) As Boolean


Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const KEYEVENTF_KEYDOWN      As Long = &H0
Private Const KEYEVENTF_KEYUP        As Long = &H2

Private Const ICC_USEREX_CLASSES = &H200

Private Type VKType
    VKCode As Integer
    scanCode As Integer
    Control As Boolean
    Shift As Boolean
    Alt As Boolean
End Type

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
'******************************************************************************
'************************** End modification/addition by JR *******************
'******************************************************************************

'******************************************************************************
'************************ Start modification/addition by ANI ******************
'******************************************************************************

Dim TimeAdjustmentConstant As Double
Dim TestCar As String

'******************************************************************************
'************************** End modification/addition by ANI ******************
'******************************************************************************


Private Sub cmdAbort_Click()
  AbortRun = True
  Sbar.Panels(1).Text = "Aborting..."
End Sub

Private Sub cmdGenerate_Click()
    AbortRun = False    ' just in case...
    DoGenerate
End Sub

Private Sub DoGenerate()
On Error GoTo LocalErrorHandler ' Ensure clean exit

EnableButtons (False)   ' Disable buttons to avoid problems
chkAutoClose.SetFocus

'=================================================================
'Initialize data columns in input file for AeroDynamic executable.
'=================================================================
InputColumn(1) = "  0.0000"
InputColumn(2) = "  0.0000"
InputColumn(3) = "  0.0000"
InputColumn(4) = "  0.0000"
InputColumn(5) = "  0.0000"
InputColumn(6) = "  0.0000"
InputColumn(7) = "  0.0000"
InputColumn(8) = "  0.0000"
InputColumn(9) = "  0.0000"
InputColumn(10) = "  0.0000"
InputColumn(11) = "  0.0000"
InputColumn(12) = "  0.0000"
InputColumn(13) = "  0.0000"
InputColumn(14) = "  0.0000"
InputColumn(15) = "  0.0000"
InputColumn(16) = "  0.0000"
'=================================================================


'==============================
'Initialize necessary variables
'==============================
rCount = 0
cCount = 0
UnitCount = 0
SlotCount = 0
AccAeroScore = 0
AxlesCount = 0
EmptyCount = 0
Base = 0
BaseTotalUpper = 0
BaseTotalLower = 0
'==============================

'==============================
' JR - Pseudoconstants
sTempName = "TempFile"
'==============================


'===================
'Preparing I/O Files
'===================
SetStatus "Status: Preparing I/O Files"
fnGap = txtGap.Text 'input file containing gap-length data
fnAxles = txtAero.Text 'input file containing axles data
fnUmler = txtUmler.Text 'input file containing car umler data

'Database to store process information
fnProcess = App.Path & "\" & "dbProcess.mdb"

'Output files as input for AeroDynamic executable
fnAeroIn = App.Path & "\" & Trim(sTempName) & ".1" 'temporary file
fnAeroInF = App.Path & "\AERO\" & Trim(sTempName) & ".acn" 'final output

'Final output files for user
fnFinal = App.Path & "\" & Trim(sTempName) & ".2" 'temporary file
fnFinalF = App.Path & "\" & Trim(sTempName) & ".txt" 'final output

'Report file
fnReport = txtOutPath.Text & "\" & Mid(txtOutFName.Text, 1, Len(txtOutFName.Text) - 4) & "_Report.txt"

'Preparing file handlers
Set fso = CreateObject("Scripting.FileSystemObject")

'Delete existing temporary file
If fso.FileExists(fnAeroIn) Then
    fso.DeleteFile fnAeroIn, True
End If
If fso.FileExists(fnFinal) Then
    fso.DeleteFile fnFinal, True
End If
If fso.FileExists(fnReport) Then
    fso.DeleteFile fnReport, True
End If


'Opening files
Set fAxles = fso.OpenTextFile(fnAxles, 1)
Set fGap = fso.OpenTextFile(fnGap, 1)
Set fOut = fso.OpenTextFile(fnAeroIn, 8, True)
Set fFinal = fso.OpenTextFile(fnFinal, 8, True)
Set fReport = fso.OpenTextFile(fnReport, 8, True)
'===================


'======================
'Connecting to database
'======================
SetStatus "Status: Connecting to Database"
Set dbUmler = DBEngine.OpenDatabase(fnUmler)
Set dbProcess = DBEngine.OpenDatabase(fnProcess)
'======================


'==========================
'Preparing Process Database
'==========================
Set rsAxles = dbProcess.OpenRecordset("select * from Axles")
Set rsGap = dbProcess.OpenRecordset("select * from Gap")

'Clearing databases
If rsAxles.RecordCount <> 0 Then
    rsAxles.MoveFirst
    Do While rsAxles.EOF = False
        rsAxles.Delete
        rsAxles.MoveNext
    Loop
End If

If rsGap.RecordCount <> 0 Then
    rsGap.MoveFirst
    Do While rsGap.EOF = False
        rsGap.Delete
        rsGap.MoveNext
    Loop
End If
'==========================


'================================
'Processing gap-length input file
'================================
'Read the first line which contains the path of video
'Write it to temporary user's view file
gapStr = fGap.ReadLine
If (InStr(gapStr, "Unit") = 0) Then 'Path provided, write to report
    fFinal.WriteLine gapStr
    gapStr = fGap.ReadLine
End If

'Collect all gap-length information
SetStatus "Status: Processing gap-length"

Dim CarGapInfo() As String
Dim ContainerInfo() As String

Do While fGap.AtEndOfStream = False And Trim(gapStr) <> "Level1 Gaps:"
    rsGap.AddNew
    UnitCount = UnitCount + 1
    SetStatus "Status: Processing gap-length (" + Str$(UnitCount) + ")"
    rsGap.Fields("UnitNum").Value = UnitCount
    gapStr = fGap.ReadLine
    
    If InStr(6, gapStr, vbTab) <> 0 Then
        Dim ContainerType() As String
        ContainerInfo = Split(gapStr, vbTab)
        ContainerType = Split(ContainerInfo(5))
        
        rsGap.Fields("Type").Value = ContainerType(2)
        rsGap.Fields("TimeStart").Value = ContainerInfo(6)
        rsGap.Fields("TimeStop").Value = ContainerInfo(7)
        
        If ContainerInfo(2) = "DoubleStackedContainers" Then
            rsGap.Fields("TimeStartUpper").Value = ContainerInfo(8)
            rsGap.Fields("TimeStopUpper").Value = ContainerInfo(9)
        End If
    Else
        ContainerInfo = Split(gapStr)
        rsGap.Fields("Type").Value = ContainerInfo(2)
        rsGap.Fields("TimeStart").Value = ContainerInfo(3)
        rsGap.Fields("TimeStop").Value = ContainerInfo(4)
        
        'MsgBox ContainerInfo(3) & vbCrLf & ContainerInfo(4)
        
        If ContainerInfo(2) = "DoubleStackedContainers" Then
            rsGap.Fields("TimeStartUpper").Value = ContainerInfo(5)
            rsGap.Fields("TimeStopUpper").Value = ContainerInfo(6)
        End If
    End If
    
    'MsgBox "TimeStart: " & vbTab & rsGap.Fields("TimeStart").Value & vbCrLf & "TimeStop: " & vbTab & rsGap.Fields("TimeStop").Value & vbCrLf & "TimeStartUpper: " & vbTab & rsGap.Fields("TimeStartUpper").Value & vbCrLf & "TimeStopUpper: " & vbTab & rsGap.Fields("TimeStopUpper").Value
    gapStr = fGap.ReadLine
    CarGapInfo = Split(Right(gapStr, Len(gapStr) - InStr(gapStr, ":")), "ft")
    'remove blank spaces, and read off the value ignoring parenthesis
    
    rsGap.Fields("CarStart").Value = Val(Right(Trim(CarGapInfo(0)), Len(Trim(CarGapInfo(0))) - 1))
    rsGap.Fields("CarStop").Value = Val(Right(Trim(CarGapInfo(1)), Len(Trim(CarGapInfo(1))) - 1))
    rsGap.Fields("CarBottom").Value = Val(Right(Trim(CarGapInfo(2)), Len(Trim(CarGapInfo(2))) - 2))
    rsGap.Fields("CarTop").Value = Val(Right(Trim(CarGapInfo(3)), Len(Trim(CarGapInfo(3))) - 1))
    
    'MsgBox "CarStart: " & vbTab & rsGap.Fields("CarStart").Value & vbCrLf & "CarStop" & vbTab & rsGap.Fields("CarStop").Value
    
    If ContainerInfo(2) = "DoubleStackedContainers" Then
        gapStr = fGap.ReadLine
        CarGapInfo = Split(Right(gapStr, Len(gapStr) - InStr(gapStr, ":")), "ft")
        rsGap.Fields("CarStartUpper").Value = Val(Right(Trim(CarGapInfo(0)), Len(Trim(CarGapInfo(0))) - 1))
        rsGap.Fields("CarStopUpper").Value = Val(Right(Trim(CarGapInfo(1)), Len(Trim(CarGapInfo(1))) - 1))
        rsGap.Fields("CarBottomUpper").Value = Val(Right(Trim(CarGapInfo(2)), Len(Trim(CarGapInfo(2))) - 2))
        rsGap.Fields("CarTopUpper").Value = Val(Right(Trim(CarGapInfo(3)), Len(Trim(CarGapInfo(3))) - 1))
    End If
    
    rsGap.Update
    gapStr = fGap.ReadLine
    If (InStr(gapStr, "Level") <> 0) Then
        Exit Do
    End If
Loop

'Close file
fGap.Close

'===========================
'Processing axles input file
'===========================
AxlesCount = 0

'Skip first line
tempStr = fAxles.ReadLine
SetStatus "Status: Processing axles input file"


' ANI -- Modification:
' First line contains the column indices
Dim OldCarInit, AxleData() As String
Dim TimeStampCorrFactor, deltaT As Double

Dim AEIScrubbed As Boolean
Dim OldCarNum, OldOrdinalNum As Long
Dim CarAxlesCount() As Integer

Do
    tempStr = fAxles.ReadLine
    AxleData = Split(tempStr, vbTab)
Loop While Val(AxleData(2)) < 10000

OldOrdinalNum = 1
OldCarInit = AxleData(1)
OldCarNum = AxleData(2)
deltaT = Val(AxleData(3))

'Multiplier for correcting difference between AEI reader and TMS timestamps
TimeStampCorrFactor = 2.12  'empirical constant

ReDim CarAxlesCount(OldOrdinalNum)
CarAxlesCount(OldOrdinalNum) = 0

If deltaT = 0 And Val(AxleData(0)) = 1 Then AEIScrubbed = True

Do
    'Saving axles info
    rsAxles.AddNew
    
    If AxleData(1) <> "RRRR" And AxleData(2) <> OldCarNum Then
        OldOrdinalNum = OldOrdinalNum + 1
        ReDim Preserve CarAxlesCount(OldOrdinalNum)
        CarAxlesCount(OldOrdinalNum) = 1
        OldCarInit = AxleData(1)
        OldCarNum = AxleData(2)
    Else
        CarAxlesCount(OldOrdinalNum) = CarAxlesCount(OldOrdinalNum) + 1
    End If
    
    SetStatus "Status: Processing axles input file (" + Str$(OldOrdinalNum) + ")"
    
    rsAxles.Fields("CarInit").Value = OldCarInit
    rsAxles.Fields("CarNum").Value = OldCarNum
    rsAxles.Fields("Ordinal").Value = OldOrdinalNum
    
    If Not AEIScrubbed Then
        rsAxles.Fields("AxleTimeStamp").Value = Round((Val(AxleData(3)) - deltaT) * TimeStampCorrFactor, 4)
    Else
        rsAxles.Fields("AxleTimeStamp").Value = Round(Val(AxleData(3)), 4)
    End If
    rsAxles.Update
    tempStr = fAxles.ReadLine
    AxleData = Split(tempStr, vbTab)
Loop While fAxles.AtEndOfStream = False

CarAxlesCount(OldOrdinalNum) = CarAxlesCount(OldOrdinalNum) + 1

rsAxles.AddNew
rsAxles.Fields("Ordinal").Value = OldOrdinalNum
rsAxles.Fields("CarInit").Value = AxleData(1)
rsAxles.Fields("CarNum").Value = AxleData(2)

If Not AEIScrubbed Then
    rsAxles.Fields("AxleTimeStamp").Value = Round((Val(AxleData(3)) - deltaT) * TimeStampCorrFactor, 4)
Else
    rsAxles.Fields("AxleTimeStamp").Value = Round(Val(AxleData(3)), 4)
End If

rsAxles.Update
fAxles.Close
AxlesNumCount = CarAxlesCount

'===========================

'*****************MAIN LOOP******************

'Initialize variables, flags and database
AxlesCount = 0
rCount = 0
UnitCount = 0
LoadFit = True
GapEnd = False
TimeStampError = False
PrevTimeStamp = 0

Set rsAxles = dbProcess.OpenRecordset("select * from Axles order by Ordinal asc, AxleTimeStamp asc")
Set rsGap = dbProcess.OpenRecordset("select * from Gap order by UnitNum asc")
rsAxles.MoveFirst
rsGap.MoveFirst

MainLoop:
Do While (rsAxles.EOF = False) And (Not AbortRun)

    'Reinitialize List
    InputColumn(1) = "  0.0000"
    InputColumn(2) = "  0.0000"
    InputColumn(3) = "  0.0000"
    InputColumn(4) = "  0.0000"
    InputColumn(5) = "  0.0000"
    InputColumn(6) = "  0.0000"
    InputColumn(7) = "  0.0000"
    InputColumn(8) = "  0.0000"
    InputColumn(9) = "  0.0000"
    InputColumn(10) = "  0.0000"
    InputColumn(11) = "  0.0000"
    InputColumn(12) = "  0.0000"
    InputColumn(13) = "  0.0000"
    InputColumn(14) = "  0.0000"
    InputColumn(15) = "  0.0000"
    InputColumn(16) = "  0.0000"

    'Determine the index to read from AxlesNumCount
    AxlesCount = AxlesCount + 1
    
    SetStatus "Status: Processing cars (" + Str$(AxlesCount) + ")"
    'Get CarInit, CarNum, Axles
    CarInit = rsAxles.Fields("CarInit").Value
    CarNum = rsAxles.Fields("CarNum").Value
    
    'Number of Axles this car has
    AxlesNum = AxlesNumCount(AxlesCount)
    
    'Get Car Info
    sQueryString = "select * from MiniUmler where CarInitial = '" & CarInit & "' and CarNumber = '" & CarNum & "'"
    
    Set rsUmler = dbUmler.OpenRecordset(sQueryString)
    
    'Reporting
    fReport.WriteLine "---------------------------------------------------------------------------------"
    fReport.WriteLine "Processing Car: [" & CarInit & " " & CarNum & "] Number of Axles: " & AxlesNum
    
    If AxlesNum = "6" And rsUmler.RecordCount = 0 Then
    'This is a locomotive
        
        'Reporting
        fReport.WriteLine "[Unit #" & rsGap.Fields("Unitnum").Value & " Type: " & rsGap.Fields("Type").Value & " treated as LOCOMOTIVE"
        
        'Writing locomotive output
        inputStr = "  6.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000  0.0000"
        fOut.WriteLine (inputStr)
        
        'Skip 1 unit in rsGap
        rsGap.MoveNext
        Call TimeStampCheck
        
        'Skip 5 axles in rsAxles
        For cCount = 1 To 5
            rsAxles.MoveNext
        Next
        
        'UnitCount is number of load counted so far
        UnitCount = UnitCount + 1
    
    Else
    'NOT a locomotive
    
        If rsUmler.RecordCount <> 0 Then
            rsUmler.MoveFirst
            
            'get Attr, Cartype
            Attr1 = Trim(rsUmler.Fields("FirstAttr").Value)
            Attr2 = Trim(rsUmler.Fields("SecondAttr").Value)
            Attr3 = Trim(rsUmler.Fields("ThirdAttr").Value)
            CarType = Trim(rsUmler.Fields("CarType").Value)
            CarLength = Val(Trim(rsUmler.Fields("Length").Value))
            
'            MsgBox Attr1 & vbTab & Attr2 & vbTab & Attr3 & vbTab & CarType & vbTab & CarLength
            
            'Reporting
            fReport.WriteLine "Car found in UMLER. Car Info:"
            fReport.WriteLine "--> [CarType = " & CarType & "] [CarLength = " & CarLength & "] [Attr1 = " & Attr1 & "] [Attr2 = " & Attr2 & "] [Attr3 = " & Attr3 & "]"
            fReport.WriteLine "Now at Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                    
            Select Case CarType

            Case "S"
                If (HandleType_S = 1) Then GoTo MainLoop

            Case "Q"
                'MsgBox HandleType_Q
                If (HandleType_Q = 1) Then GoTo MainLoop
                
            Case "F", "P"
                'MsgBox HandleType_FP
                If (HandleType_FP = 1) Then GoTo MainLoop
                        
            Case Else
                If Not BlindMode Then MsgBox "Cartype " & CarType & " of Car " & CarInit & " " & CarNum & " has not yet coded to be handled", vbCritical, "Unexpected Data"
                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Car not yet coded to be handled."
                GoTo SkipCar
            End Select
            
        Else
            If Not BlindMode Then MsgBox "Car " & CarInit & " " & CarNum & " doesn't exist in MiniUmler", vbCritical, "Missing Information"
            AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & "Car not found in MiniUmler."
            
            'Reporting
            fReport.WriteLine "Car " & CarInit & " " & CarNum & " doesn't exist in MiniUmler, and not a locomotive. Car skipped."
            
SkipCar:
            'AxlesCount = AxlesCount + 1
            'Ordinal = rsAxles.Fields("Ordinal").Value
            Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
                'Reporting
                fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                                
                rsAxles.MoveNext
                                
            Loop
            fFinal.WriteLine (AeroInput)
            
            'Reporting
            fReport.WriteLine "Output >> " & AeroInput
            
            GoTo MainLoop
        End If
    End If
    
    If rsAxles.EOF = False Then
        rsAxles.MoveNext
    End If
Loop
'*****************END MAIN LOOP******************

'Close files
fOut.Close
fFinal.Close
fReport.Close

If AbortRun And Not BlindMode Then
    MsgBox "Process aborted", vbCritical, "Finished"
    End
End If
    
'======================
'Preparing output files
'======================
SetStatus "Status: Running Aero.exe ***PLEASE WAIT ***"
'Finish up fFinal & fOut
Set fOutF = fso.OpenTextFile(fnAeroInF, 2, True)
Set fOut = fso.OpenTextFile(fnAeroIn, 1)
fOutF.WriteLine ("VERSION 3.0")
fOutF.WriteLine UnitCount
Do While fOut.AtEndOfStream = False
    tempStr = fOut.ReadLine
    fOutF.WriteLine tempStr
Loop
fOut.Close
fOutF.Close

' Run Aero.exe
Dim TestID As Integer
Dim i As Integer
Dim S As String
Dim n As Integer

'***************************************************************************************
' Avoid breakpoints in this area... Focus MUST be kept on cmd window
S = "cmd /k chdir /d " + App.Path + "\AERO"
TestID = Shell(S, 2)   ' Run Calculator.
If TestID = 0 Then
    MsgBox "Problem running Aero.exe", vbCritical, "File Missing"
    End
  Else
    WaitSec (2)         ' Allow some time to start program
    AppActivate TestID      ' Activate command window
    SendString "title TEM~", 0.5 ' Change cmd window title
    WaitSec (0.1)           '
    SendString "~"          ' Need for double-enter?
    SendString "Aero.exe~", 0.5 ' Start aero.exe
    WaitSec (1)
    SendString "~~"         ' Skip introduction string
    SendString "4~"         ' Option #4
    SendString "TempFile~", 0.7 ' Name of the file (constant)
    SendString "~~"         ' Start processing, default temp and pressure
    SendString "N~"         ' No need to print
    SendString "5~"         ' Quit
    WaitSec (3)             ' Let user see something
    If chkAutoClose.Value = 1 Then
        SendString "exit~"      ' Close cmd window
    End If
    CloseHandle TestID      ' Release handle
    TestID = 0
End If
'***************************************************************************************

SetStatus "Status: Finalizing"
'Getting AeroScores
Set fFinalF = fso.OpenTextFile(fnFinalF, 2, True)
Set fFinal = fso.OpenTextFile(fnFinal, 1)
fnAxles = App.Path & "\AERO\" & Trim(sTempName) & ".drg"

Set fAxles = fso.OpenTextFile(fnAxles, 1)
tempStr = fAxles.ReadLine
Do While Left(tempStr, 9) <> "      0.0"
    tempStr = fAxles.ReadLine
Loop
fFinalF.WriteLine "C (lbs/mph/mph) = " & Trim(Right(tempStr, 11))
fFinalF.WriteLine "Average AeroScore = " & FormatNumber(AccAeroScore / SlotCount, 2, vbFalse) & "%"
Do While fFinal.AtEndOfStream = False
    tempStr = fFinal.ReadLine
    fFinalF.WriteLine tempStr
Loop

'Close files
fFinal.Close
fFinalF.Close
fAxles.Close

'**************************************************************************************
'Create output file
fso.CopyFile fnFinalF, fso.BuildPath(txtOutPath, txtOutFName)
'**************************************************************************************
EnableButtons (True)
SetStatus "Status: Done"

If Not BlindMode Then SafeMsgBox "Done. " & Val(EmptyCount) & " empty slot detected.", vbOKOnly, "Progress"

Kill (fnFinal)
Kill (fnFinalF)
Kill (fnAeroIn)

If BlindMode Then End

LocalErrorHandler:
    EnableButtons (True)
End Sub

Private Sub cmdSetOutPath_Click()
Dim tmp As String

'http://techrepublic.com.com/5208-10878-0.html?forumID=87&threadID=186276&messageID=1905937
    Dim shlShell As Variant, shlFolder As Variant
    Set shlShell = New Shell32.Shell
    Set shlFolder = shlShell.BrowseForFolder(Me.hWnd, "Select a Folder", 1)
    If shlFolder Is Nothing Then
        Exit Sub
    End If
    tmp = shlFolder.Items.Item.Path ' (full drive and path)
    If tmp <> "" Then
        txtOutPath.Text = tmp + "\"
    End If
End Sub

Private Sub cmdUmler_Click()
cdlgPath.DialogTitle = "Specify MiniUmler database"
cdlgPath.Filter = "UMLER Database (*.mdb) | *.mdb"
cdlgPath.ShowOpen
txtUmler.Text = cdlgPath.FileName
End Sub

Function getData(tempStr As String, indData As Integer) As String

'tempStr is the string passed, an array of data seperated by spacebar or tab.
'        All spaces and tab at the end of this string will be truncated.
'indData is the index of data to read for output
'indData = 1 means the first data

If indData < 1 Then
    MsgBox "getData received indData < 1", vbExclamation, "Error"
End If

Dim DataCount, startInd, endInd As Integer

DataCount = 1
startInd = 1
endInd = 1

'remove space AND tab at the end of string
Do While Left(tempStr, 1) = " " Or Left(tempStr, 1) = Chr(9)
    tempStr = Right(tempStr, Len(tempStr) - 1)
Loop

Do While Right(tempStr, 1) = " " Or Right(tempStr, 1) = Chr(9)
    tempStr = Left(tempStr, Len(tempStr) - 1)
Loop


Do While DataCount < indData
    Do While Mid(tempStr, startInd, 1) <> " " And Mid(tempStr, startInd, 1) <> Chr(9)
        startInd = startInd + 1
        'If startInd > Len(tempStr) Then
        '    MsgBox "Index requested out of bound. TempStr = {" & tempStr & "} indData = {" & indData & "}"
        '    End
        'End If
    Loop
    
    DataCount = DataCount + 1
    
    Do While Mid(tempStr, startInd, 1) = " " Or Mid(tempStr, startInd, 1) = Chr(9)
        startInd = startInd + 1
        'If startInd > Len(tempStr) Then
        '    MsgBox "Index requested out of bound. TempStr = {" & tempStr & "} indData = {" & indData & "}"
        '    End
        'End If
    Loop
Loop

endInd = startInd

Do While (Mid(tempStr, endInd, 1) <> " ") And (Mid(tempStr, endInd, 1) <> "") And Mid(tempStr, endInd, 1) <> Chr(9)
    endInd = endInd + 1
Loop

getData = Mid(tempStr, startInd, endInd - startInd)

End Function

Function HandleType_S() As Integer
'=============================
'Getting best load information
'=============================
Set rsBestLoad = dbProcess.OpenRecordset("select * from S where SecondNum like '*" & Attr2 & "*' and ThirdNum like '*" & Attr3 & "*'")

If rsBestLoad.RecordCount = 1 Then
            
    'Integrity Check
    rsBestLoad.MoveLast
    rsBestLoad.MoveFirst
    If rsBestLoad.RecordCount <> 1 Then
        If Not BlindMode Then MsgBox "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount, vbCritical, "Car Umler Data Error"
        
        'Reporting
        fReport.WriteLine "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount
    End If
    'End Integrity Check
    
    BestUpper = rsBestLoad.Fields("BestUp").Value
    BestLower = rsBestLoad.Fields("BestLow").Value
    
    'Get special case Base value
    If IsNull(rsBestLoad.Fields("Base").Value) Then
        Base = 0
    Else
        Base = rsBestLoad.Fields("Base").Value
    End If
    
'///////////////////////////
'MsgBox "<" & rsBestLoad.Fields("Base").Value & ">" & vbCrLf & "SBase Empty? " & IsNull(rsBestLoad.Fields("Base").Value)
    
    BaseTotalUpper = 0
    BaseTotalLower = 0
    
    IgnorePart2 = False
    MuteAERO = False
    
    'Reporting
    fReport.WriteLine "Best Load information:"
    fReport.WriteLine "--> [BestUpper = " & BestUpper & "] [BestLower = " & BestLower & "] [Base = " & Base & "]"
    
Else
    
    'This type of car should be ignored.
    IgnorePart2 = True
    MuteAERO = False

    'Reporting
    fReport.WriteLine "Best Load information not found. Car ignored."


End If
'=============================


'===========================
'Getting the number of wells
'===========================
Select Case Val(Attr2)

Case 1
    rCount = 1
Case 2
    rCount = 2
Case 3
    rCount = 3
Case 4
    rCount = 4
Case 5 To 8
    rCount = 5
    
Case Else
    'Car not specified or not-used (0 or 9), ignore this car
    MuteAERO = False
    IgnorePart2 = True
    
    'Skip Axles to next car, do not proceed in gap. Generate error message in output file.
    'AxlesCount = AxlesCount + 1
    'Ordinal = rsAxles.Fields("Ordinal").Value
    Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
        'Reporting
        fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
        
        rsAxles.MoveNext
        
    Loop
    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to 0/9 Second Numeric."
    fFinal.WriteLine (AeroInput)
    
    'Reporting
    fReport.WriteLine "Output >> " & AeroInput
    
    HandleType_S = 1
    Exit Function
End Select

'Reporting
fReport.WriteLine "Number of well = " & rCount

'Remember the number of wells
MaxrCount = rCount

'================
'Processing wells
'================

'skip first axle, because it is always at the end
rsAxles.MoveNext

'Start to match loads
Do While rCount <> 0

    'Reinitialize descriptors
    InputColumn(1) = "  0.0000"
    InputColumn(2) = "  0.0000"
    InputColumn(3) = "  0.0000"
    InputColumn(4) = "  0.0000"
    InputColumn(5) = "  0.0000"
    InputColumn(6) = "  0.0000"
    InputColumn(7) = "  0.0000"
    InputColumn(8) = "  0.0000"
    InputColumn(9) = "  0.0000"
    InputColumn(10) = "  0.0000"
    InputColumn(11) = "  0.0000"
    InputColumn(12) = "  0.0000"
    InputColumn(13) = "  0.0000"
    InputColumn(14) = "  0.0000"
    InputColumn(15) = "  0.0000"
    InputColumn(16) = "  0.0000"

    
    If GapEnd = False Then
        'Unit Start
        timeStart = rsGap.Fields("TimeStart").Value
            
        'Unit Stop
        timeStop = rsGap.Fields("TimeStop").Value
    
        'Axle Start
        axleStart = rsAxles.Fields("AxleTimeStamp").Value
        
'ANI
'-------------------------------
        'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
        Do While (timeStop < axleStart) And GapEnd = False
            'Reporting
            fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
            
            rsGap.MoveNext
            Call TimeStampCheck
            
            If rsGap.EOF = True Then
                GapEnd = True
            Else
                timeStart = rsGap.Fields("TimeStart").Value
                timeStop = rsGap.Fields("TimeStop").Value
            End If
        Loop
    End If

    If GapEnd = True Then
        'Force empty scenario
        timeStart = 0
        timeStop = 0
    End If
    
    'Read next Axles
    rsAxles.MoveNext
    axleEnd = rsAxles.Fields("AxleTimeStamp").Value
    
    'Matching load
    If (axleStart < ((timeStart + timeStop) / 2)) And (axleEnd > ((timeStart + timeStop) / 2)) Then
        LoadFit = True 'Load Matched
        
        'Reporting
        fReport.WriteLine "Load Matched:"
        fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & axleEnd & "]"
    Else
        LoadFit = False
    End If
    
    'Determine Layout and match load
    '===============================
    If rCount > 1 Then
        If AxlesNum = Val(MaxrCount * 4) Then
            'read 3 axle ahead
            rsAxles.MoveNext
            rsAxles.MoveNext
            rsAxles.MoveNext
            
            'Reporting
            fReport.WriteLine "Layout: AxlesNum = (# of Well) * 4"
            
        ElseIf AxlesNum = Val(MaxrCount * 2 + 2) Then
            'read 1 axle ahead
            rsAxles.MoveNext
            
            'Reporting
            fReport.WriteLine "Layout: AxlesNum = (# of Well) * 2 + 2"
            
        Else
            'Ignore for both output
            If Not BlindMode Then MsgBox "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum, vbCritical, "This car will be ignored."
            IgnorePart2 = True
            MuteAERO = True
            
            'Reporting
            fReport.WriteLine "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum & ". Car ignored."
            
            'Skip this car to next car, don't proceed in gap file. Generate error message in output file.
            'AxlesCount = AxlesCount + 1
            'Ordinal = rsAxles.Fields("Ordinal").Value
            Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
                
                'Reporting
                fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                
                rsAxles.MoveNext
                
            Loop
            AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to incorrect Axles Number."
            fFinal.WriteLine (AeroInput)
            
            'Reporting
            fReport.WriteLine "Output >> " & AeroInput
            
            LoadFit = False
            HandleType_S = 1
            Exit Function
        End If
    Else
        'last slot, skip only one axle
        rsAxles.MoveNext
    End If
    
    '===============
    'ONLY IF MATCHED
    '===============
    If LoadFit = True Then
        
        If rsGap.Fields("Type").Value = "SingleContainer" Then
            cCount = 1
            InputColumn(2) = "  2.0000"
            
            'Reporting
            fReport.WriteLine "Load fit as SingleContainer."
            
        ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
            cCount = 2
            InputColumn(2) = "  3.0000"
        
            'Reporting
            fReport.WriteLine "Load fit as DoubleStackedContainers."
        
        ElseIf rsGap.Fields("Type").Value = "Trailer" Then
            cCount = 1
            InputColumn(2) = "  1.0000"
        
            'Reporting
            fReport.WriteLine "Load fit as Trailer."
        
        Else
            If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
            cCount = 0
            IgnorePart2 = True
            MuteAERO = True
            
            'Reporting
            fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
            
        End If

        'If SingleStack, means upper slot is empty
        If cCount = 1 Then
            InputColumn(2) = "  0.0000"
            InputColumn(3) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
            InputColumn(7) = "  0.0000"
            EmptyCount = EmptyCount + 1
    
            If IgnorePart2 = False And TimeStampError = False Then
            
                CargoLength = 0
                AeroScore = 0
                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
                fFinal.WriteLine (AeroInput)
                
                'Reporting
                fReport.WriteLine "Matching load not found. Well is empty."
                fReport.WriteLine "Output >> " & AeroInput
                
                SlotCount = SlotCount + 1
                
                'Reporting
                fReport.WriteLine "Number of slot is increased to " & SlotCount
            
            End If
        End If
        
        

        Do While cCount <> 0
                        
            'Get CargoLength
            If cCount = 2 Then
                carStart = rsGap.Fields("CarStartUpper").Value
                carStop = rsGap.Fields("CarStopUpper").Value
            Else
                carStart = rsGap.Fields("CarStart").Value
                carStop = rsGap.Fields("CarStop").Value
            End If
            
            
            CargoLength = carStop - carStart
            
'            MsgBox "CarStart: " & carStart & vbTab & "CarStop: " & carStop & vbTab & "Cargolength: " & CargoLength
            
            fReport.WriteLine "Raw CargoLength = " & CargoLength
    
            'Convert into discrete values
            If CargoLength <= 24 Then
                CargoLength = 20
            ElseIf CargoLength <= 34 Then
                CargoLength = 28
            ElseIf CargoLength <= 42.5 Then
                CargoLength = 40
            ElseIf CargoLength <= 46.5 Then
                CargoLength = 45
            ElseIf CargoLength <= 50.5 Then
                CargoLength = 48
            ElseIf CargoLength <= 55 Then
                CargoLength = 53
            Else
                CargoLength = 57
            End If
            
            'Reporting
            fReport.WriteLine "CargoLength = " & CargoLength
        
            '====================
            'Adding to TrainScore
            '====================
            If IgnorePart2 = False And TimeStampError = False Then
            
                If Base = 0 Then
            
                    If cCount = 2 Then
                        AeroScore = FormatNumber((CargoLength / BestUpper) * 100, 2)
                    Else
                        AeroScore = FormatNumber((CargoLength / BestLower) * 100, 2)
                    End If
                    
                    If AeroScore > 100 Then
                        AeroScore = FormatNumber(100, 2)
                    End If

                    If cCount = 2 Then
                        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestUpper & Chr(9) & CargoLength & Chr(9) & AeroScore & "%"
                    Else
                        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%"
                    End If
                    fFinal.WriteLine (AeroInput)
                    
                    'Reporting
                    fReport.WriteLine "Output >> " & AeroInput
                    
                    AccAeroScore = AccAeroScore + AeroScore
                    SlotCount = SlotCount + 1
                    
                    'Reporting
                    fReport.WriteLine "Number of slot is increased to " & SlotCount
                
                Else 'The special Base 5 case
                
                    'Accumulate cargolength
                    If cCount = 2 Then
                        BaseTotalUpper = BaseTotalUpper + CargoLength
                    Else
                        BaseTotalLower = BaseTotalLower + CargoLength
                    End If
                    
                    SlotCount = SlotCount + 1
                
                    'Reporting
                    fReport.WriteLine "Number of slot is increased to " & SlotCount
                    fReport.WriteLine "Handling Base 5 Case, accumulating CargoLength:"
                    fReport.WriteLine "Current BaseTotalUpper = " & BaseTotalUpper
                    fReport.WriteLine "Current BaseTotallower = " & BaseTotalLower
                End If
                
            Else
            
                If cCount = 1 Then
                    
                    If TimeStampError = False Then
                        AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to missing best load information."
                    Else
                        AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to illogical timestamp in gap-length information."                    '
                    End If
                    
                    fFinal.WriteLine (AeroInput)

                    'Reporting
                    fReport.WriteLine "Output >> " & AeroInput
                End If
                
            End If
            '====================
                        
            cCount = cCount - 1
                        
        Loop
                        
        'CargoLength here is always the lower one
        InputColumn(7) = FormatNumber(CargoLength, 4, vbFalse)

        'Done Matching, move to next unit
        rsGap.MoveNext
        Call TimeStampCheck
    
    '==============
    'IF NOT MATCHED
    '==============
    Else
    
        For cCount = 1 To 2
            InputColumn(2) = "  0.0000"
            InputColumn(3) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
            InputColumn(7) = "  0.0000"
            EmptyCount = EmptyCount + 1
    
            If IgnorePart2 = False And TimeStampError = False Then
            
                CargoLength = 0
                AeroScore = 0
                If cCount = 2 Then
                    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestUpper & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
                Else
                    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
                End If
                fFinal.WriteLine (AeroInput)
                
                'Reporting
                fReport.WriteLine "Matching load not found. Well is empty."
                fReport.WriteLine "Output >> " & AeroInput
                
                SlotCount = SlotCount + 1
                
                'Reporting
                fReport.WriteLine "Number of slot is increased to " & SlotCount
            Else
            
                If cCount = 1 Then
                    If TimeStampError = False Then
                        AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to missing best load information."
                    Else
                        AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to illogical timestamp in gap-length information."                    '
                    End If
                    fFinal.WriteLine (AeroInput)
        
                    'Reporting
                    fReport.WriteLine "Output >> " & AeroInput
                End If
            End If
        Next
    End If
        
    'Need to do these no matter the load is matched or not
    '=====================================================
    If Base = 5 And rCount = 1 Then 'Last slot, output to AeroScore
            
        'If DoubleStacked
        If InputColumn(2) = "  3.0000" Then
            AeroScore = FormatNumber((BaseTotalUpper / BestUpper) * 100, 2)
                        
            If AeroScore > 100 Then
                AeroScore = FormatNumber(100, 2)
            End If

            If IgnorePart2 = False Then
                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestUpper & Chr(9) & BaseTotalUpper & Chr(9) & AeroScore & "%"
            Else
                AeroScore = 0
            End If
            
            fFinal.WriteLine (AeroInput)
            
            'Reporting
            fReport.WriteLine "Output >> " & AeroInput
            
            AccAeroScore = AccAeroScore + AeroScore
            
            'Reporting
            fReport.WriteLine "Current Accumulated AeroScore = " & AccAeroScore
            
        End If
        
        'Handle the lower load
        AeroScore = FormatNumber((BaseTotalLower / BestLower) * 100, 2)
                                                
        If AeroScore > 100 Then
            AeroScore = FormatNumber(100, 2)
        End If

        If IgnorePart2 = False And TimeStampError = False Then
            AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & BaseTotalLower & Chr(9) & AeroScore & "%"
        Else
            AeroScore = 0
            If TimeStampError = False Then
                AeroInput = "Base 5 case: Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to missing best load information."
            Else
                AeroInput = "Base 5 case: Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to illogical timestamp in gap-length information."                    '
            End If
        End If
        
        fFinal.WriteLine (AeroInput)
            
        'Reporting
        fReport.WriteLine "Output >> " & AeroInput
            
        AccAeroScore = AccAeroScore + AeroScore
        
        'Reporting
        fReport.WriteLine "Current Accumulated AeroScore = " & AccAeroScore
        
    End If
    
    'Formatting Output
    InputColumn(3) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
    If InputColumn(3) = ".0000" Then
        InputColumn(3) = "0.0000"
    End If
                        
    If Len(InputColumn(3)) = 6 Then
        InputColumn(3) = "  " & InputColumn(3)
    ElseIf Len(InputColumn(3)) = 7 Then
        InputColumn(3) = " " & InputColumn(3)
    End If
    
    If rCount = MaxrCount Or rCount = 1 Then
        InputColumn(4) = "  1.5000"
    Else
        InputColumn(4) = "  1.0000"
    End If
                        
    If Len(InputColumn(7)) = 6 Then
        InputColumn(7) = "  " & InputColumn(7)
    ElseIf Len(InputColumn(7)) = 7 Then
        InputColumn(7) = " " & InputColumn(7)
    End If
                                                        
    InputColumn(8) = "  2.0000"

    'Need to do this only if load is matched
    If LoadFit = True Then
        InputColumn(12) = FormatNumber((Val(InputColumn(3)) - Val(InputColumn(7))) / 2, 4, vbFalse)
    
        If Val(InputColumn(12)) < 0 Then
            InputColumn(12) = "0.0000"
        End If
    
        If Len(InputColumn(12)) = 6 Then
            InputColumn(12) = "  " & InputColumn(12)
        ElseIf Len(InputColumn(12)) = 7 Then
            InputColumn(12) = " " & InputColumn(12)
        End If
    Else
        InputColumn(12) = "  0.0000"
    End If

    InputColumn(5) = "  0.0000"
    InputColumn(6) = "  0.0000"
    InputColumn(9) = "  0.0000"
    InputColumn(10) = "  0.0000"
    InputColumn(11) = "  0.0000"
    InputColumn(13) = "  0.0000"
    
    inputStr = InputColumn(1) & InputColumn(2) & InputColumn(3) & InputColumn(4) & InputColumn(5) & InputColumn(6) & InputColumn(7) & InputColumn(8) & InputColumn(9) & InputColumn(10) & InputColumn(11) & InputColumn(12) & InputColumn(13) & InputColumn(14) & InputColumn(15) & InputColumn(16)
    If MuteAERO = False Then
        fOut.WriteLine (inputStr)
    End If
        
    If LoadFit = True Then
        UnitCount = UnitCount + 1
    End If

    rCount = rCount - 1

Loop

HandleType_S = 0

End Function

Function HandleType_Q()



TestCar = "553828"

'Getting best load information
Set rsBestLoad = dbUmler.OpenRecordset("select * from Q where FirstNum like '*" & Attr1 & "*' and ThirdNum like '*" & Attr3 & "*'")

If rsBestLoad.RecordCount = 1 Then
    
    'Integrity Check
    rsBestLoad.MoveLast
    rsBestLoad.MoveFirst
    If rsBestLoad.RecordCount <> 1 Then
        If Not BlindMode Then MsgBox "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount, vbCritical, "Car Umler Data Error"
        
        'Reporting
        fReport.WriteLine "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount
    End If
    'End Integrity Check

    BestLower = rsBestLoad.Fields("Best").Value
    
    'Get special case Base value
    If IsNull(rsBestLoad.Fields("Base").Value) Then
        Base = 0
    Else
        Base = rsBestLoad.Fields("Base").Value
    End If
    
'    If CarNum = TestCar Then MsgBox "BestLoad: " & rsBestLoad.Fields("Best").Value & vbCrLf & "Base: " & Base
'/////////////////////////////


'MsgBox "<" & rsBestLoad.Fields("Base").Value & ">" & vbCrLf & "QBase Empty? " & IsNull(rsBestLoad.Fields("Base").Value)
    
    BaseTotalLower = 0
    
    IgnorePart2 = False
    MuteAERO = False
    
    'Reporting
    fReport.WriteLine "Best Load information:"
    fReport.WriteLine "--> [Best = " & BestLower & "] [Base = " & Base & "]"
    
Else
    
    'This type of car should be ignored.
    IgnorePart2 = True
    MuteAERO = False
    
    'Reporting
    fReport.WriteLine "Best Load information not found. Car ignored."

End If

'======== Branch out to handle two different Q Case=============
If Base = 0 Then

    'Reporting
'    If CarNum = TestCar Then MsgBox "Handling Type Q Base 0 case."
    
    fReport.WriteLine "Handling Type Q Base 0 case."
    
'=======Handling Car Type: Q Normal Case ===================
    InputColumn(1) = "  1.0000"
    
    
    
    'Determine number of well
    Select Case Val(Attr2)
        
    Case 1 To 9
        rCount = Val(Attr2)
    Case 0
        rCount = 10 'Can be more than 10 cars, but was told to ignore
    Case Else
        'Car not specified or not-used (0 or 9), ignore this car
        MuteAERO = False
        IgnorePart2 = True
        'Skip Axles to next car, do not proceed in gap. Generate error message in output file.
        'AxlesCount = AxlesCount + 1
        'Ordinal = rsAxles.Fields("Ordinal").Value
        Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
            'Reporting
            fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
            
            rsAxles.MoveNext
                        
        Loop
        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to unrecognizable Second Numeric."
        fFinal.WriteLine (AeroInput)
        
        'Reporting
        fReport.WriteLine "Output >> " & AeroInput
        
        HandleType_Q = 1
        Exit Function
    End Select
    
'    If CarNum = TestCar Then MsgBox "# of wells: " & rCount

    'Reporting
    fReport.WriteLine "Number of well = " & rCount

    'Remember the number of wells
    MaxrCount = rCount
    
    'skip first axle, because it is always at the end
    rsAxles.MoveNext
    
    'Start to match loads
    Do While rCount <> 0
    
        'Reinitialize descriptors
        InputColumn(1) = "  1.0000"
        InputColumn(2) = "  0.0000"
        InputColumn(3) = "  0.0000"
        InputColumn(4) = "  0.0000"
        InputColumn(5) = "  0.0000"
        InputColumn(6) = "  0.0000"
        InputColumn(7) = "  0.0000"
        InputColumn(8) = "  0.0000"
        InputColumn(9) = "  0.0000"
        InputColumn(10) = "  0.0000"
        InputColumn(11) = "  0.0000"
        InputColumn(12) = "  0.0000"
        InputColumn(13) = "  0.0000"
        InputColumn(14) = "  0.0000"
        InputColumn(15) = "  0.0000"
        InputColumn(16) = "  0.0000"
                
        If GapEnd = False Then
            'Unit Start
            timeStart = rsGap.Fields("TimeStart").Value
        
            'Unit Stop
            timeStop = rsGap.Fields("TimeStop").Value
        
            'Axle Start
            axleStart = rsAxles.Fields("AxleTimeStamp").Value
            
            'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
            Do While (timeStop < axleStart) And GapEnd = False
                'Reporting
                fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                
                rsGap.MoveNext
                Call TimeStampCheck
                If rsGap.EOF = True Then
                    GapEnd = True
                Else
                    timeStart = rsGap.Fields("TimeStart").Value
                    timeStop = rsGap.Fields("TimeStop").Value
                End If
            Loop
            
        End If
        
        If GapEnd = True Then
            'Force empty scenario
            timeStart = 0
            timeStop = 0
        End If
        
        'Read next Axle
        rsAxles.MoveNext
        axleEnd = rsAxles.Fields("AxleTimeStamp").Value
        
        If (axleStart < ((timeStart + timeStop) / 2)) And (axleEnd > ((timeStart + timeStop) / 2)) Then
            LoadFit = True 'Load Matched
            
            'Reporting
            fReport.WriteLine "Load Matched:"
            fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & axleEnd & "]"
            
        Else
            LoadFit = False
        End If
        
        'Determine Layout and match load
        '===============================
        If rCount > 1 Then
            If AxlesNum = Val(MaxrCount * 4) Then
                'read 3 axles ahead
                rsAxles.MoveNext
                rsAxles.MoveNext
                rsAxles.MoveNext
                
                'Reporting
                fReport.WriteLine "Layout: AxlesNum = (# of Well) * 4"
                
            ElseIf AxlesNum = Val(MaxrCount * 2 + 2) Then
                'read 1 axle ahead
                rsAxles.MoveNext
                
                'Reporting
                fReport.WriteLine "Layout: AxlesNum = (# of Well) * 2 + 2"
                
            Else
                'Ignore for both output
                If Not BlindMode Then MsgBox "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum, vbCritical, "This car will be ignored."
                IgnorePart2 = True
                MuteAERO = True
                
                'Reporting
                fReport.WriteLine "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum & ". Car Ignored."
                
                'Skip this car to next car, don't proceed in gap file. Generate error message in output file.
                'AxlesCount = AxlesCount + 1
                'Ordinal = rsAxles.Fields("Ordinal").Value

                
                Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
                    'Reporting
                    fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                    rsAxles.MoveNext
                Loop
                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to incorrect Axles Number."
                
                'Reporting
                fReport.WriteLine "Output >> " & AeroInput
                
                fFinal.WriteLine (AeroInput)
                LoadFit = False
                HandleType_Q = 1
                
                Exit Function
            End If
        Else
            'last slot, skip only one axle
            rsAxles.MoveNext
        End If
        
        
        'ONLY IF MATCHED
        '===============
        If LoadFit = True Then
            
            If rsGap.Fields("Type").Value = "SingleContainer" Then
                cCount = 1
                InputColumn(2) = "  2.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as SingleContainer."
                
            ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
                cCount = 2
                InputColumn(2) = "  2.0000"
                If Not BlindMode Then MsgBox "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer.", vbCritical, "Internal Data Error"
                
                'Reporting
                fReport.WriteLine "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer."
            ElseIf rsGap.Fields("Type").Value = "Trailer" Then
                cCount = 1
                InputColumn(2) = "  1.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as Trailer."
            Else
                If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
                cCount = 0
                IgnorePart2 = True
                MuteAERO = True
                
                'Reporting
                fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
            End If
    
            'Get cargolength
            If cCount = 2 Then
                carStart = rsGap.Fields("CarStartUpper").Value
                carStop = rsGap.Fields("CarStopUpper").Value
            Else
                carStart = rsGap.Fields("CarStart").Value
                carStop = rsGap.Fields("CarStop").Value
            End If
            CargoLength = carStop - carStart
            
            fReport.WriteLine "Raw CargoLength = " & CargoLength
    
            'Convert into discrete values
            
        'MsgBox CargoLength

        If CargoLength <= 24 Then CargoLength = 20
        
        If (CargoLength <= 34 And CargoLength > 24) Then CargoLength = 28
        
        If (CargoLength <= 42.5 And CargoLength > 34) Then CargoLength = 40
        
        If (CargoLength <= 46.5 And CargoLength > 42.5) Then CargoLength = 45
        
        If (CargoLength <= 50.5 And CargoLength > 50.5) Then CargoLength = 48
        
        If CargoLength <= 55 Then
            CargoLength = 53
        Else
            CargoLength = 57
        End If
        
'        MsgBox CargoLength

'
'            ANI--
'
'            If CargoLength <= 24 Then
'                CargoLength = 20
'            ElseIf CargoLength <= 34 Then
'                CargoLength = 28
'            ElseIf CargoLength <= 42.5 Then
'                CargoLength = 40
'            ElseIf CargoLength <= 46.5 Then
'                CargoLength = 45
'            ElseIf CargoLength <= 50.5 Then
'                CargoLength = 48
'            ElseIf CargoLength <= 55 Then
'                CargoLength = 53
'            Else
'                CargoLength = 57
'            End If
        
            'Reporting
            fReport.WriteLine "CargoLength = " & CargoLength
            
'            If CarNum = TestCar Then MsgBox IgnorePart2 & vbCrLf & TimeStampError
        
            '=====Adding to TrainScore=====
            If IgnorePart2 = False And TimeStampError = False Then
            
                AeroScore = FormatNumber((CargoLength / BestLower) * 100, 2)
                
'                If CarNum = TestCar Then MsgBox AeroScore
                    
                If AeroScore > 100 Then
                    AeroScore = FormatNumber(100, 2)
                End If

                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%"
                fFinal.WriteLine (AeroInput)
                
                'Reporting
                fReport.WriteLine "Output >> " & AeroInput
            
                AccAeroScore = AccAeroScore + AeroScore
                SlotCount = SlotCount + 1
                
                'Reporting
                fReport.WriteLine "Number of slot is increased to " & SlotCount
                fReport.WriteLine "Current Accumulated AeroScore = " & AccAeroScore
                                                
            Else
                If TimeStampError = False Then
                    AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to missing best load information."
                Else
                    AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to illogical timestamp in gap-length information."                    '
                End If
                fFinal.WriteLine (AeroInput)
 
                'Reporting
                fReport.WriteLine "Output >> " & AeroInput
            End If
            '===============================
                                                    
            'CargoLength here is always the lower one
            InputColumn(7) = FormatNumber(CargoLength, 4, vbFalse)
            
            'Done Matching, move to next unit
            rsGap.MoveNext
            Call TimeStampCheck
    
        'IF NOT MATCHED
        '===============
        Else
            InputColumn(2) = "  0.0000"
            InputColumn(7) = "  0.0000"
            EmptyCount = EmptyCount + 1
        
            If IgnorePart2 = False And TimeStampError = False Then
            
                CargoLength = 0
                AeroScore = 0
                AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & MaxrCount - rCount + 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
                fFinal.WriteLine (AeroInput)
                
                'Reporting
                fReport.WriteLine "Output >> " & AeroInput
                
                SlotCount = SlotCount + 1
                
                'Reporting
                fReport.WriteLine "Number of slot is increased to " & SlotCount
            Else
                If TimeStampError = False Then
                    AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to missing best load information."
                Else
                    AeroInput = "Slot #" & (MaxrCount - rCount + 1) & " on " & CarInit & CarNum & " is ignored due to illogical timestamp in gap-length information."                    '
                End If
                fFinal.WriteLine (AeroInput)

                'Reporting
                fReport.WriteLine "Output >> " & AeroInput
            End If
        End If
        
        
        'Need to do these no matter the load is matched or not
        '=====================================================
        InputColumn(3) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
        If InputColumn(3) = ".0000" Then
            InputColumn(3) = "0.0000"
        End If
                            
        If Len(InputColumn(3)) = 6 Then
            InputColumn(3) = "  " & InputColumn(3)
        ElseIf Len(InputColumn(3)) = 7 Then
            InputColumn(3) = " " & InputColumn(3)
        End If
        
        If rCount = MaxrCount Or rCount = 1 Then
            InputColumn(4) = "  1.5000"
        Else
            InputColumn(4) = "  1.0000"
        End If
                            
        If Len(InputColumn(7)) = 6 Then
            InputColumn(7) = "  " & InputColumn(7)
        ElseIf Len(InputColumn(7)) = 7 Then
            InputColumn(7) = " " & InputColumn(7)
        End If
                                                            
        InputColumn(8) = "  2.0000"
    
        'Need to do this only if load is matched
        If LoadFit = True Then
            InputColumn(11) = FormatNumber((Val(InputColumn(3)) - Val(InputColumn(7))) / 2, 4, vbFalse)
        
            If Val(InputColumn(11)) < 0 Then
                InputColumn(11) = "0.0000"
            End If
        
            If Len(InputColumn(11)) = 6 Then
                InputColumn(11) = "  " & InputColumn(11)
            ElseIf Len(InputColumn(11)) = 7 Then
                InputColumn(11) = " " & InputColumn(11)
            End If
        Else
            InputColumn(7) = "  0.0000"
            InputColumn(11) = InputColumn(7)
        End If
    
        InputColumn(12) = "  0.0000"
        InputColumn(5) = "  0.0000"
        InputColumn(6) = "  0.0000"
        InputColumn(9) = "  0.0000"
        InputColumn(10) = "  0.0000"
        
        inputStr = InputColumn(1) & InputColumn(2) & InputColumn(3) & InputColumn(4) & InputColumn(5) & InputColumn(6) & InputColumn(7) & InputColumn(8) & InputColumn(9) & InputColumn(10) & InputColumn(11) & InputColumn(12) & InputColumn(13) & InputColumn(14) & InputColumn(15) & InputColumn(16)
        If MuteAERO = False Then
            fOut.WriteLine (inputStr)
        End If
            
        If LoadFit = True Then
            UnitCount = UnitCount + 1
        End If
    
        rCount = rCount - 1

    Loop

'======== Branch out to handle two different Q Case=============
Else

'=======Handling Car Type: Q Base 2 Case ===================
'    If CarNum = TestCar Then MsgBox "Handling Type Q Base 2 case."
    'Reporting
    fReport.WriteLine "Handling Type Q Base 2 Case."
    
    
    'Reinitializing descriptors
    InputColumn(1) = "  4.0000"
    InputColumn(2) = "  0.0000"
    InputColumn(3) = "  0.0000"
    InputColumn(4) = "  0.0000"
    InputColumn(5) = "  0.0000"
    InputColumn(6) = "  0.0000"
    InputColumn(7) = "  0.0000"
    InputColumn(8) = "  0.0000"
    InputColumn(9) = "  0.0000"
    InputColumn(10) = "  0.0000"
    InputColumn(11) = "  0.0000"
    InputColumn(12) = "  0.0000"
    InputColumn(13) = "  0.0000"
    InputColumn(14) = "  0.0000"
    InputColumn(15) = "  0.0000"
    InputColumn(16) = "  0.0000"
    
    'Always have 2 unit
    rCount = 2
    
    'Reporting
    fReport.WriteLine "Automatically assume Number of Well = 2."

    'Remember the number of unit
    MaxrCount = rCount
    
    'skip first axle, because it is always at the end
    rsAxles.MoveNext
    
    'Axle Start
    axleStart = rsAxles.Fields("AxleTimeStamp").Value
    
    'Accumulate Cargolength
    BaseTotalLower = 0
    
    'Determine Layout
    '================
    If AxlesNum = Val(MaxrCount * 4) Then
        'read 5 axle ahead
        rsAxles.MoveNext
        bEnd = rsAxles.Fields("AxleTimeStamp").Value
        rsAxles.MoveNext
        rsAxles.MoveNext
        rsAxles.MoveNext
        bStart = rsAxles.Fields("AxleTimeStamp").Value
        rsAxles.MoveNext
        
        'Reporting
        fReport.WriteLine "Layout: AxlesNum = (# of Well) * 4"
        
    ElseIf AxlesNum = Val(MaxrCount * 2 + 2) Then
        'read 3 axle ahead
        rsAxles.MoveNext
        bEnd = rsAxles.Fields("AxleTimeStamp").Value
        rsAxles.MoveNext
        bStart = rsAxles.Fields("AxleTimeStamp").Value
        rsAxles.MoveNext
        
        'Reporting
        fReport.WriteLine "Layout: AxlesNum = (# of Well) * 2 + 2"
    Else
        'Ignore for both output
        If Not BlindMode Then MsgBox "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum, vbCritical, "This car will be ignored."
        IgnorePart2 = True
        MuteAERO = True
        
        'Reporting
        fReport.WriteLine "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum & ". Car ignored."
        
        'Skip this car to next car, don't proceed in gap file. Generate error message in output file.
        'AxlesCount = AxlesCount + 1
        'Ordinal = rsAxles.Fields("Ordinal").Value
        Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
            'Reporting
            fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                        
            rsAxles.MoveNext
                
        Loop
        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to incorrect Axles Number."
        fFinal.WriteLine (AeroInput)
        
        'Reporting
        fReport.WriteLine "Output >> " & AeroInput
        
        LoadFit = False
        HandleType_Q = 1
        Exit Function
    End If

    'Read end axles
    axleEnd = rsAxles.Fields("AxleTimeStamp").Value
    
'If CarNum = TestCar Then MsgBox axleEnd
    
    'Skip to last axle since we know 2 unit.
    rsAxles.MoveNext
    
    'Start to match loads for first unit
    '===================================
    CargoLength = 0
    Do While rCount <> 1
                
        If GapEnd = False Then
            'Unit Start
            timeStart = rsGap.Fields("TimeStart").Value
        
            'Unit Stop
            timeStop = rsGap.Fields("TimeStop").Value
                                
            'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
            Do While (timeStop < axleStart) And GapEnd = False
                'Reporting
                fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                
                rsGap.MoveNext
                Call TimeStampCheck
                If rsGap.EOF = True Then
                    GapEnd = True
                Else
                    timeStart = rsGap.Fields("TimeStart").Value
                    timeStop = rsGap.Fields("TimeStop").Value
                End If
            Loop
            
        End If
        
        If GapEnd = True Then
            'Force empty scenario
            rCount = 1
        End If
        
        'Try fit on first unit
        If (axleStart < ((timeStart + timeStop) / 2)) And (bEnd > ((timeStart + timeStop) / 2)) Then
            LoadFit = True 'Load Matched
            
            'Reporting
            fReport.WriteLine "Load Matched:"
            fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & bEnd & "]"
            
        Else
            LoadFit = False
            If ((timeStart + timeStop) / 2) > bEnd Then 'this load belongs to 2nd unit
                rCount = 1
            End If
        End If
                            
        'ONLY IF MATCHED
        '===============
'If CarNum = TestCar Then MsgBox "LoadFit? " & LoadFit
        
        If LoadFit = True Then
        
            UnitCount = UnitCount + 1
            
            If rsGap.Fields("Type").Value = "SingleContainer" Then
                cCount = 1
                InputColumn(4) = "  2.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as SingleContainer."
                
            ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
                cCount = 2
                InputColumn(4) = "  2.0000"
                If Not BlindMode Then MsgBox "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer.", vbCritical, "Internal Data Error"
                
                'Reporting
                fReport.WriteLine "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer."
            ElseIf rsGap.Fields("Type").Value = "Trailer" Then
                cCount = 1
                InputColumn(4) = "  1.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as Trailer."
            Else
            
                'Reporting
                fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
                
                If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
                cCount = 0
                IgnorePart2 = True
                MuteAERO = True
            End If
    
            'Get cargolength
            If cCount = 2 Then
                carStart = rsGap.Fields("CarStartUpper").Value
                carStop = rsGap.Fields("CarStopUpper").Value
            Else
                carStart = rsGap.Fields("CarStart").Value
                carStop = rsGap.Fields("CarStop").Value
            End If
            CargoLength = carStop - carStart
            
            fReport.WriteLine "Raw CargoLength = " & CargoLength
            
            'Get discrete values
            '===================
            If CargoLength <= 24 Then
                CargoLength = 20
            ElseIf CargoLength <= 34 Then
                CargoLength = 28
            ElseIf CargoLength <= 42.5 Then
                CargoLength = 40
            ElseIf CargoLength <= 46.5 Then
                CargoLength = 45
            ElseIf CargoLength <= 50.5 Then
                CargoLength = 48
            ElseIf CargoLength <= 55 Then
                CargoLength = 53
            Else
                CargoLength = 57
            End If
            
            'Reporting
            fReport.WriteLine "CargoLength = " & CargoLength
            
            'Accumulate cargolength
            tempVal = CargoLength
            
            'Done Matching, move to next unit
            rsGap.MoveNext
            Call TimeStampCheck
        End If
    Loop
        
    'accumulate for AeroScore
    BaseTotalLower = BaseTotalLower + CargoLength
    
    'Reporting
    fReport.WriteLine "Current Accumulated CargoLength = " & BaseTotalLower
    
    'Unit is not empty
    If CargoLength <> 0 Then
        'CargoLength here is always the lower one
        InputColumn(6) = FormatNumber(CargoLength, 4, vbFalse)
        InputColumn(5) = FormatNumber((CarLength - CargoLength) / 2, 4, vbFalse)
    Else
        InputColumn(6) = "  0.0000"
        InputColumn(5) = "  0.0000"
        EmptyCount = EmptyCount + 1
        AeroScore = 0
        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & rCount & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
        fFinal.WriteLine (AeroInput)
        
        'Reporting
        fReport.WriteLine "Output >> " & AeroInput
    End If
    
    'Need to do these no matter unit is empty or not
    '===============================================
                                                
    InputColumn(2) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
    
    If InputColumn(3) = ".0000" Then
        InputColumn(3) = "0.0000"
    End If
                        
    If Len(InputColumn(3)) = 6 Then
        InputColumn(3) = "  " & InputColumn(3)
    ElseIf Len(InputColumn(3)) = 7 Then
        InputColumn(3) = " " & InputColumn(3)
    End If
                                            
    If Len(InputColumn(7)) = 6 Then
        InputColumn(7) = "  " & InputColumn(7)
    ElseIf Len(InputColumn(7)) = 7 Then
        InputColumn(7) = " " & InputColumn(7)
    End If
                                                                            
    inputStr = InputColumn(1) & InputColumn(2) & InputColumn(3) & InputColumn(4) & InputColumn(5) & InputColumn(6) & InputColumn(7) & InputColumn(8) & InputColumn(9) & InputColumn(10) & InputColumn(11) & InputColumn(12) & InputColumn(13) & InputColumn(14) & InputColumn(15) & InputColumn(16)
    If MuteAERO = False Then
        fOut.WriteLine (inputStr)
    End If
            
    'Handle the load in-between first and 2nd unit
    '=============================================
    Do While rCount <> 2
        
        If GapEnd = False Then
            'Unit Start
            timeStart = rsGap.Fields("TimeStart").Value
        
            'Unit Stop
            timeStop = rsGap.Fields("TimeStop").Value
                                
            'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
            Do While (timeStop < axleStart) And GapEnd = False
                'Reporting
                fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                
                rsGap.MoveNext
                Call TimeStampCheck
                If rsGap.EOF = True Then
                    GapEnd = True
                Else
                    timeStart = rsGap.Fields("TimeStart").Value
                    timeStop = rsGap.Fields("TimeStop").Value
                End If
            Loop
        End If
        
        If GapEnd = True Then
            'Force empty scenario
            rCount = 2
        End If
        
        'Try fit between 1st and 2nd unit
        If (bEnd < ((timeStart + timeStop) / 2)) And (bStart > ((timeStart + timeStop) / 2)) Then
            LoadFit = True 'Load Matched
            
            'Reporting
            fReport.WriteLine "Load Matched:"
            fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & bStart & "]"
            
        Else
            LoadFit = False
            If ((timeStart + timeStop) / 2) > bStart Then 'this load belongs to 2nd unit
                rCount = 2
            End If
        End If
                            
        'ONLY IF MATCHED
        '===============
        If LoadFit = True Then
        
            UnitCount = UnitCount + 1
            
            If rsGap.Fields("Type").Value = "SingleContainer" Then
                cCount = 1
                InputColumn(4) = "  2.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as SingleContainer."
                
            ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
                cCount = 2
                InputColumn(4) = "  2.0000"
                If Not BlindMode Then MsgBox "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer.", vbCritical, "Internal Data Error"
                
                'Reporting
                fReport.WriteLine "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer."
            
            ElseIf rsGap.Fields("Type").Value = "Trailer" Then
                cCount = 1
                InputColumn(4) = "  1.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as Trailer."
            
            Else
                If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
                cCount = 0
                IgnorePart2 = True
                MuteAERO = True
                
                'Reporting
                fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
            End If
    
            'Get cargolength
            If cCount = 2 Then
                carStart = rsGap.Fields("CarStartUpper").Value
                carStop = rsGap.Fields("CarStopUpper").Value
            Else
                carStart = rsGap.Fields("CarStart").Value
                carStop = rsGap.Fields("CarStop").Value
            End If
            CargoLength = carStop - carStart
            
            fReport.WriteLine "Raw CargoLength = " & CargoLength
    
            'Get discrete values
            '===================
            If CargoLength <= 24 Then
                CargoLength = 20
            ElseIf CargoLength <= 34 Then
                CargoLength = 28
            ElseIf CargoLength <= 42.5 Then
                CargoLength = 40
            ElseIf CargoLength <= 46.5 Then
                CargoLength = 45
            ElseIf CargoLength <= 50.5 Then
                CargoLength = 48
            ElseIf CargoLength <= 55 Then
                CargoLength = 53
            Else
                CargoLength = 57
            End If
            
        
            If cCount = 2 Then
                If Not BlindMode Then MsgBox cCount
                'fGap.ReadLine
            End If
    
            'Reporting
            fReport.WriteLine "CargoLength = " & CargoLength
            
            'Accumulate cargolength
            CargoLength = CargoLength + TempDouble
            
            'Done Matching, move to next unit
            rsGap.MoveNext
            Call TimeStampCheck
        End If
    Loop
        
    'accumulate for AeroScore
    BaseTotalLower = BaseTotalLower + CargoLength
    
    'Reporting
    fReport.WriteLine "Current Accumulated CargoLength = " & BaseTotalLower
    
    'Start matching load for 2nd unit
    '=================================
    CargoLength = 0
    Do While rCount <> 0
        If GapEnd = False Then
            'Unit Start
            timeStart = rsGap.Fields("TimeStart").Value
        
            'Unit Stop
            timeStop = rsGap.Fields("TimeStop").Value
                                
            'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
            Do While (timeStop < axleStart) And GapEnd = False
                'Reporting
                fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
                
                rsGap.MoveNext
                Call TimeStampCheck
                If rsGap.EOF = True Then
                    GapEnd = True
                Else
                    timeStart = rsGap.Fields("TimeStart").Value
                    timeStop = rsGap.Fields("TimeStop").Value
                End If
            Loop
        End If
        
        If GapEnd = True Then
            'Force empty scenario
            rCount = 0
        End If
        
        'Try fit on first unit
        If (bStart < ((timeStart + timeStop) / 2)) And (axleEnd > ((timeStart + timeStop) / 2)) Then
            LoadFit = True 'Load Matched
            
            'Reporting
            fReport.WriteLine "Load Matched:"
            fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & axleEnd & "]"

        Else
            LoadFit = False
            If ((timeStart + timeStop) / 2) > axleEnd Then 'Load doesn't belong to this car
                rCount = 0
            End If
        End If
                            
        'ONLY IF MATCHED
        '===============
        If LoadFit = True Then
        
            UnitCount = UnitCount + 1
            
            If rsGap.Fields("Type").Value = "SingleContainer" Then
                cCount = 1
                InputColumn(4) = "  2.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as SingleContainer."
            
            ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
                cCount = 2
                InputColumn(4) = "  2.0000"
                If Not BlindMode Then MsgBox "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer.", vbCritical, "Internal Data Error"
                
                'Reporting
                fReport.WriteLine "DoubleStackedContainers detected on Q type car. Will be treated as SingleContainer."
            
            ElseIf rsGap.Fields("Type").Value = "Trailer" Then
                cCount = 1
                InputColumn(4) = "  1.0000"
                
                'Reporting
                fReport.WriteLine "Load fit as Trailer."
            Else
                If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
                cCount = 0
                IgnorePart2 = True
                MuteAERO = True
                
                'Reporting
                fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
            End If
    
            'Get cargolength
            If cCount = 2 Then
                carStart = rsGap.Fields("CarStartUpper").Value
                carStop = rsGap.Fields("CarStopUpper").Value
            Else
                carStart = rsGap.Fields("CarStart").Value
                carStop = rsGap.Fields("CarStop").Value
            End If
            CargoLength = carStop - carStart
            
            fReport.WriteLine "Raw CargoLength = " & CargoLength
    
            'Get discrete values
            '===================
            If CargoLength <= 24 Then
                CargoLength = 20
            ElseIf CargoLength <= 34 Then
                CargoLength = 28
            ElseIf CargoLength <= 42.5 Then
                CargoLength = 40
            ElseIf CargoLength <= 46.5 Then
                CargoLength = 45
            ElseIf CargoLength <= 50.5 Then
                CargoLength = 48
            ElseIf CargoLength <= 55 Then
                CargoLength = 53
            Else
                CargoLength = 57
            End If
            
            'Reporting
            fReport.WriteLine "CargoLength = " & CargoLength
                        
            'Accumulate cargolength
            TempDouble = CargoLength
                        
            'Done Matching, move to next unit
            rsGap.MoveNext
            Call TimeStampCheck
        End If
    Loop
        
    'accumulate for AeroScore
    BaseTotalLower = BaseTotalLower + CargoLength
    
    'Reporting
    fReport.WriteLine "Current Accumulated CargoLength = " & BaseTotalLower
    
    'Unit is not empty
    If CargoLength <> 0 Then
        'CargoLength here is always the lower one
        InputColumn(6) = FormatNumber(CargoLength, 4, vbFalse)
        InputColumn(5) = FormatNumber((CarLength - CargoLength) / 2, 4, vbFalse)
    Else
        InputColumn(6) = "  0.0000"
        InputColumn(5) = "  0.0000"
        EmptyCount = EmptyCount + 1
        AeroScore = 0
        AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & rCount & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
        fFinal.WriteLine (AeroInput)
        
        'Reporting
        fReport.WriteLine "Output >> " & AeroInput
    End If
    
    'Output to AeroScore
    AeroScore = FormatNumber((BaseTotalLower / BestLower) * 100, 2)
                        
    If AeroScore > 100 Then
        AeroScore = FormatNumber(100, 2)
    End If

    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & 2 & Chr(9) & BestLower & Chr(9) & BaseTotalLower & Chr(9) & AeroScore & "%"
    fFinal.WriteLine (AeroInput)
    
    'Reporting
    fReport.WriteLine "Output >> " & AeroInput
    
    AccAeroScore = AccAeroScore + AeroScore
    SlotCount = SlotCount + 1

    'Reporting
    fReport.WriteLine "Number of slot is increased to " & SlotCount
    fReport.WriteLine "Current Accumulated AeroScore = " & AccAeroScore

    'Need to do these no matter unit is empty or not
    '===============================================
                                                
    InputColumn(2) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)
    
    If InputColumn(3) = ".0000" Then
        InputColumn(3) = "0.0000"
    End If
                        
    If Len(InputColumn(3)) = 6 Then
        InputColumn(3) = "  " & InputColumn(3)
    ElseIf Len(InputColumn(3)) = 7 Then
        InputColumn(3) = " " & InputColumn(3)
    End If
                                            
    If Len(InputColumn(7)) = 6 Then
        InputColumn(7) = "  " & InputColumn(7)
    ElseIf Len(InputColumn(7)) = 7 Then
        InputColumn(7) = " " & InputColumn(7)
    End If
                                                                            
    inputStr = InputColumn(1) & InputColumn(2) & InputColumn(3) & InputColumn(4) & InputColumn(5) & InputColumn(6) & InputColumn(7) & InputColumn(8) & InputColumn(9) & InputColumn(10) & InputColumn(11) & InputColumn(12) & InputColumn(13) & InputColumn(14) & InputColumn(15) & InputColumn(16)
    If MuteAERO = False Then
        fOut.WriteLine (inputStr)
    End If

End If

HandleType_Q = 0

End Function

Function HandleType_FP()

If CarType = "F" Then
    Set rsBestLoad = dbProcess.OpenRecordset("select * from F where ThirdNum like '*" & Attr3 & "*'")
Else 'Cartype = P
    Set rsBestLoad = dbProcess.OpenRecordset("select * from P where FirstNum like '*" & Attr1 & "*' and SecondNum like '*" & Attr2 & "*' and ThirdNum like '*" & Attr3 & "*'")
End If
    
'MsgBox Attr1 & vbCrLf & Attr2 & vbCrLf & Attr3
'MsgBox "BestLoad RecordCount:" & vbTab & rsBestLoad.RecordCount

If rsBestLoad.RecordCount = 1 Then
    
    'Integrity Check
    rsBestLoad.MoveLast
    rsBestLoad.MoveFirst
    If rsBestLoad.RecordCount <> 1 Then
        If Not BlindMode Then MsgBox "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount, vbCritical, "Car Umler Data Error"
        
        'Reporting
        fReport.WriteLine "Multiple best-load records found for CarInfo: " & CarInit & " " & CarNum & ". Amount: " & rsBestLoad.RecordCount
    End If
    'End Integrity Check
    
    BestLower = rsBestLoad.Fields("Best").Value
    
'///////////////////////////////////////////////////
    
'Get special case Base value
'    If IsNull(rsBestLoad.Fields("Base").Value) Then
'        Base = 0
'    Else
'        Base = rsBestLoad.Fields("Base").Value
'    End If
    
    Base = 0
    
    BaseTotalLower = 0
    
    'Reporting
    fReport.WriteLine "Best Load information:"
    fReport.WriteLine "--> [Best = " & BestLower & "] [Base = " & Base & "]"

    IgnorePart2 = False
    MuteAERO = False
    
Else
    
    'This type of car should be ignored.
    IgnorePart2 = True
    MuteAERO = False

    'Reporting
    fReport.WriteLine "Best Load information not found. Car ignored."

End If

'Reinitializing descriptors
InputColumn(1) = "  4.0000"
InputColumn(2) = "  0.0000"
InputColumn(3) = "  0.0000"
InputColumn(4) = "  0.0000"
InputColumn(5) = "  0.0000"
InputColumn(6) = "  0.0000"
InputColumn(7) = "  0.0000"
InputColumn(8) = "  0.0000"
InputColumn(9) = "  0.0000"
InputColumn(10) = "  0.0000"
InputColumn(11) = "  0.0000"
InputColumn(12) = "  0.0000"
InputColumn(13) = "  0.0000"
InputColumn(14) = "  0.0000"
InputColumn(15) = "  0.0000"
InputColumn(16) = "  0.0000"

'Always have 1 unit
rCount = 1

'Reporting
fReport.WriteLine "Automatically assume Number of Well = 1"

'Remember the number of unit
MaxrCount = rCount

'skip first axle, because it is always at the end
rsAxles.MoveNext

'Axle Start
axleStart = rsAxles.Fields("AxleTimeStamp").Value
                                
'Read Axles End
rsAxles.MoveNext
axleEnd = rsAxles.Fields("AxleTimeStamp").Value



'Determine Layout
'================
If AxlesNum = Val(MaxrCount * 4) Then
    'read 1 axle ahead, only 1 unit
    rsAxles.MoveNext
    
    'Reporting
    fReport.WriteLine "Layout: AxlesNum = (# of Well) * 4"

Else
    'Ignore for both output
    If Not BlindMode Then MsgBox "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum, vbCritical, "This car will be ignored."
    
    'Reporting
    fReport.WriteLine "Number of axles is not correct: # of Axles: " & AxlesNum & ", # of Car Unit: " & MaxrCount & ", CarInfo:  " & CarInit & " " & CarNum
    
    IgnorePart2 = True
    MuteAERO = True
    'Skip this car to next car, don't proceed in gap file. Generate error message in output file.
    'AxlesCount = AxlesCount + 1
    'Ordinal = rsAxles.Fields("Ordinal").Value
    Do While rsAxles.Fields("Ordinal").Value <> AxlesCount + 1
        'Reporting
        fReport.WriteLine "Skipping Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
        
        rsAxles.MoveNext
                        
    Loop
    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & Chr(9) & "Ignored due to incorrect Axles Number."
    fFinal.WriteLine (AeroInput)
    
    'Reporting
    fReport.WriteLine "Output >> " & AeroInput
    
    LoadFit = False
    HandleType_FP = 1
    Exit Function
End If
                                
'Start to match loads for first unit
'===================================
CargoLength = 0

'Debugging

Do While rCount <> 0
    If GapEnd = False Then
        'Unit Start
        timeStart = rsGap.Fields("TimeStart").Value
    
        'Unit Stop
        timeStop = rsGap.Fields("TimeStop").Value
        
        'MsgBox timeStart & vbCrLf & timeStop, vbOKOnly, BlindMode
        
        'MsgBox axleStart
        
        'If previously ignored a car, possibly load is not read ahead due to insufficient information of car.
        Do While (timeStop < axleStart) And GapEnd = False
            If Not BlindMode Then MsgBox "looping"
            'Reporting
            
            fReport.WriteLine "Skipping Load Unit #" & rsGap.Fields("UnitNum").Value & " for Axle Ordinal #" & rsAxles.Fields("Ordinal").Value & " AxleTimeStamp = " & rsAxles.Fields("AxleTimeStamp").Value
            
            rsGap.MoveNext
            Call TimeStampCheck
            If rsGap.EOF = True Then
                GapEnd = True
            Else
                timeStart = rsGap.Fields("TimeStart").Value
                timeStop = rsGap.Fields("TimeStop").Value
            End If
        Loop
        
        'MsgBox "Out of GAP END"
        
    End If
    

'    MsgBox rsGap.AbsolutePosition & "/" & rsGap.RecordCount
    
    If GapEnd = True Then
        'Force empty scenario
        rCount = 0
    End If
    
    'MsgBox "GapEnd " & GapEnd & vbCrLf & "rCount " & rCount
    
    
    'Try fit on first unit
    If (axleStart < ((timeStart + timeStop) / 2)) And (axleEnd > ((timeStart + timeStop) / 2)) Then
        LoadFit = True 'Load Matched
        
        'Reporting
        fReport.WriteLine "Load Matched:"
        fReport.WriteLine "--> Unit #" & rsGap.Fields("UnitNum").Value & " between [AxleTimeStamp: " & axleStart & ", " & axleEnd & "]"

    Else
        LoadFit = False

        If ((timeStart + timeStop) / 2) > axleEnd Then
            rCount = 0
        End If
    End If
                        
    'MsgBox "LoadFit?" & vbTab & LoadFit & vbCrLf & "rCount?" & rCount

                        
    'ONLY IF MATCHED
    '===============
    If LoadFit = True Then
    
        UnitCount = UnitCount + 1
        
        If rsGap.Fields("Type").Value = "SingleContainer" Then
            cCount = 1
            InputColumn(4) = "  2.0000"
            
            'Reporting
            fReport.WriteLine "Load fit as SingleContainer."
            If Not BlindMode Then MsgBox "Load fit as SingleContainer."
            
        ElseIf rsGap.Fields("Type").Value = "DoubleStackedContainers" Then
            cCount = 2
            InputColumn(4) = "  2.0000"
            If Not BlindMode Then MsgBox "DoubleStackedContainers detected on F type car. Will be treated as SingleContainer.", vbCritical, "Internal Data Error"
            
            'Reporting
            fReport.WriteLine "DoubleStackedContainers detected on F type car. Will be treated as SingleContainer."
            
        ElseIf rsGap.Fields("Type").Value = "Trailer" Then
            cCount = 1
            InputColumn(4) = "  1.0000"
            
            'Reporting
            fReport.WriteLine "Load fit as Trailer."
            If Not BlindMode Then MsgBox "Load fit as Trailer."
        Else
            If Not BlindMode Then MsgBox "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value, vbCritical, "Error"
            
            'Reporting
            fReport.WriteLine "This load is ignored. Unknown load-type: " & rsGap.Fields("Type").Value
            
            cCount = 0
            IgnorePart2 = True
            MuteAERO = True
        End If

        'Get cargolength
        If cCount = 2 Then
            carStart = rsGap.Fields("CarStartUpper").Value
            carStop = rsGap.Fields("CarStopUpper").Value
        Else
            carStart = rsGap.Fields("CarStart").Value
            carStop = rsGap.Fields("CarStop").Value
        End If
        CargoLength = carStop - carStart
        
        fReport.WriteLine "Raw CargoLength = " & CargoLength
        

        'Get discrete values
        '===================
        If CargoLength <= 24 Then CargoLength = 20
        
        If (CargoLength <= 34 And CargoLength > 24) Then CargoLength = 28
        
        If (CargoLength <= 42.5 And CargoLength > 34) Then CargoLength = 40
        
        If (CargoLength <= 46.5 And CargoLength > 42.5) Then CargoLength = 45
        
        If (CargoLength <= 50.5 And CargoLength > 50.5) Then CargoLength = 48
        
        If CargoLength <= 55 Then
            CargoLength = 53
        Else
            CargoLength = 57
        End If
        
    
        'Reporting
        fReport.WriteLine "CargoLength = " & CargoLength
    
        'Accumulate cargolength
        TempDouble = CargoLength
        
        'Done Matching, move to next unit
        rsGap.MoveNext
        Call TimeStampCheck
        
    End If
    rsGap.MoveNext
Loop

'/////////////////////////////////////////
                    
'Unit is not empty
If CargoLength <> 0 Then
    'CargoLength here is always the lower one
    InputColumn(6) = FormatNumber(CargoLength, 4, vbFalse)
    InputColumn(5) = FormatNumber((CarLength - CargoLength) / 2, 4, vbFalse)
    
    'Output to AeroScore
    AeroScore = FormatNumber((CargoLength / BestLower) * 100, 2)
                        
    If AeroScore > 100 Then
        AeroScore = FormatNumber(100, 2)
    End If

    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & 1 & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%"
    fFinal.WriteLine (AeroInput)
                
    'Reporting
    fReport.WriteLine "Output >> " & AeroInput
                
    AccAeroScore = AccAeroScore + AeroScore
    SlotCount = SlotCount + 1
    
    'Reporting
    fReport.WriteLine "Number of slot is increased to " & SlotCount
    fReport.WriteLine "Current Accumulated AeroScore = " & AccAeroScore
    
Else
    InputColumn(6) = "  0.0000"
    InputColumn(5) = "  0.0000"
    EmptyCount = EmptyCount + 1
    AeroScore = 0
    AeroInput = CarInit & Chr(9) & CarNum & Chr(9) & CarType & rCount & Chr(9) & BestLower & Chr(9) & CargoLength & Chr(9) & AeroScore & "%" & Chr(9) & "EMPTY"
    fFinal.WriteLine (AeroInput)
    
    fReport.WriteLine "Output >> " & AeroInput
End If

'Need to do these no matter unit is empty or not
'===============================================
                                            
InputColumn(2) = FormatNumber(CarLength / MaxrCount, 4, vbFalse)

If InputColumn(3) = ".0000" Then
    InputColumn(3) = "0.0000"
End If
                    
If Len(InputColumn(3)) = 6 Then
    InputColumn(3) = "  " & InputColumn(3)
ElseIf Len(InputColumn(3)) = 7 Then
    InputColumn(3) = " " & InputColumn(3)
End If
                                        
If Len(InputColumn(7)) = 6 Then
    InputColumn(7) = "  " & InputColumn(7)
ElseIf Len(InputColumn(7)) = 7 Then
    InputColumn(7) = " " & InputColumn(7)
End If
                                                                        
inputStr = InputColumn(1) & InputColumn(2) & InputColumn(3) & InputColumn(4) & InputColumn(5) & InputColumn(6) & InputColumn(7) & InputColumn(8) & InputColumn(9) & InputColumn(10) & InputColumn(11) & InputColumn(12) & InputColumn(13) & InputColumn(14) & InputColumn(15) & InputColumn(16)
If MuteAERO = False Then
    fOut.WriteLine (inputStr)
End If

HandleType_FP = 0

End Function

Private Sub cmdAero_Click()
cdlgPath.DialogTitle = "Specify AeroInput file"
cdlgPath.Filter = "AEI Data (*.txt) | *.txt"
cdlgPath.ShowOpen

txtAero.Text = cdlgPath.FileName

End Sub

Private Sub cmdGap_Click()
cdlgPath.DialogTitle = "Specify Gap Measurement file"
cdlgPath.Filter = "TMS Data (*.txt) | *.txt"
cdlgPath.ShowOpen

txtGap.Text = cdlgPath.FileName

End Sub

Function TimeStampCheck()

If rsGap.EOF = False Then
    'TimeStampCheck
    If rsGap.Fields("TimeStart").Value < PrevTimeStamp Then
            
        'Ignore due to timestamp error
        TimeStampError = True
        
        'Reporting
        fReport.WriteLine "WARNING >> Unit #" & rsGap.Fields("UnitNum").Value & " has invalid timestamp information compared with the previous unit."
        
    Else
        TimeStampError = False
    End If

    PrevTimeStamp = rsGap.Fields("TimeStop").Value
End If

End Function

'******************************************************************************
'************************ Start modification/addition by JR *******************
'******************************************************************************
'--------------------------------------------------------------------
' Routine: ShellAndWait
' Author:  Intelligent Solutions Inc.
' Modified JR

Public Function ShellandWait(ExeFullPath As String, _
                             Optional WindowStyle As Integer = vbMinimizedNoFocus, _
                             Optional TimeOutValue As Long = 0)
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, WindowStyle)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
        End If
    Loop While lExitCode = STATUS_PENDING
    
    'ShellandWait = True
    ShellandWait = lInst
    Exit Function
   
ErrorHandler:
    'ShellandWait = False
    ShellandWait = 0
    Exit Function

End Function

'--------------------------------------------------------------------
' Routine: WaitSec
' Author: JR
' Goal: General purpose pause, time in seconds

Private Sub WaitSec(secs)
Dim wait
    wait = Timer + secs
    Do
      DoEvents
    Loop While Timer <= wait
End Sub

'--------------------------------------------------------------------
' Routine: SendString
' Author: JR
' Goal: Slow down SendKeys

Private Sub SendString(sString As String, Optional PauseSec As Long = 0.3)
Dim i, j As Integer

For i = 1 To Len(sString)
  VbSendKeys (Mid$(sString, i, 1))
  WaitSec (PauseSec)
Next i

End Sub

'---------------------------------------------------------------------
' Routine: VbSendKeys()
'
' Author:  Bryan Wolf, 1999
'
' Purpose: Imitate VB's internal SendKeys statement, but add the
'          ability to send keypresses to a DOS application.  You
'          can use SendKeys, to paste ASCII characters to a DOS
'          window from the clipboard, but you can't send function
'          keys.  This module solves that problem and makes sending
'          any keys to any Application, DOS or Windows, easy.
'
' Arguments: Keystrokes.  Note that this does not implement the
'            SendKeys's 'wait' argument.  If you need to wait,
'            try using a timing loop.
'
'            The syntax for specifying keystrokes is the
'            same as that of SendKeys - Please refer to VB's
'            documentation for an in-depth description.  Support
'            for the following codes has been added, in addition
'            to the standard set of codes suppored by SendKeys:
'
'            KEY                  CODE
'            break                {CANCEL}
'            escape               {ESCAPE}
'            left mouse button    {LBUTTON}
'            right mouse button   {RBUTTON}
'            middle mouse button  {MBUTTON}
'            clear                {CLEAR}
'            shift                {SHIFT}
'            control              {CONTROL}
'            alt                  {MENU} or {ALT}
'            pause                {PAUSE}
'            space                {SPACE}
'            select               {SELECT}
'            execute              {EXECUTE}
'            snapshot             {SNAPSHOT}
'            number pad 0         {NUMPAD0}
'            number pad 1         {NUMPAD1}
'            number pad 2         {NUMPAD2}
'            number pad 3         {NUMPAD3}
'            number pad 4         {NUMPAD4}
'            number pad 5         {NUMPAD5}
'            number pad 6         {NUMPAD6}
'            number pad 7         {NUMPAD7}
'            number pad 8         {NUMPAD8}
'            number pad 9         {NUMPAD9}
'            number pad multiply  {MULTIPLY}
'            number pad add       {ADD}
'            number pad separator {SEPARATOR}
'            number pad subtract  {SUBTRACT}
'            number pad decimal   {DECIMAL}
'            number pad divide    {DIVIDE}
'
' Sample Calls:
'   VbSendKeys "Dir~"   ' View a directory of in DOS
'
' NOTE: there is a minor difference with SendKeys syntax. You can
'       group multiple characters under the same shift key using
'       curly brackets, while VB's SendKeys uses regular brackets.
'       For example, to keep the SHIFT key pressed while you type
'       A, B, and C keys, you must run the following statement
'           VBSendKeys "+{abc}"
'       while the syntax of the built-in VB function is
'           SendKeys "+(abc)"
'---------------------------------------------------------------------

Sub VbSendKeys(ByVal sKeystrokes As String)
    Dim iKeyStrokesLen As Integer
    Dim lRepetitions As Long
    Dim bShiftKey As Boolean
    Dim bControlKey As Boolean
    Dim bAltKey As Boolean
    Dim lResult As Long
    Dim sKey As String
    Dim iAsciiKey As Integer
    Dim iVirtualKey As Integer
    Dim i As Long
    Dim j As Long
  
    Static bInitialized As Boolean
    Static AsciiKeys(0 To 255) As VKType
    Static VirtualKeys(0 To 255) As VKType
  
    On Error GoTo 0

    If Not bInitialized Then
        Dim iKey As Integer
        Dim OEMChar As String
        Dim keyScan As Integer
        
        ' Initialize AsciiKeys()
        For iKey = LBound(AsciiKeys) To UBound(AsciiKeys)
            keyScan = VkKeyScan(iKey)
            AsciiKeys(iKey).VKCode = keyScan And &HFF   ' low-byte of key scan
                                                        ' code
            AsciiKeys(iKey).Shift = (keyScan And &H100)
            AsciiKeys(iKey).Control = (keyScan And &H200)
            AsciiKeys(iKey).Alt = (keyScan And &H400)
            ' Get the ScanCode
            OEMChar = "  " ' 2 Char
            CharToOem Chr(iKey), OEMChar
            AsciiKeys(iKey).scanCode = OemKeyScan(Asc(OEMChar)) And &HFF
        Next iKey
        
        ' Initialize VirtualKeys()
        For iKey = LBound(VirtualKeys) To UBound(VirtualKeys)
            VirtualKeys(iKey).VKCode = iKey
            VirtualKeys(iKey).scanCode = MapVirtualKey(iKey, 0)
            ' no use in initializing remaining elements
        Next iKey
        bInitialized = True     ' don't run this code twice
    End If    ' End of initialization routine
  
    ' Parse the string in the same way that SendKeys() would
    Do While Len(sKeystrokes)
        lRepetitions = 1 ' Default number of repetitions for each character
        bShiftKey = False
        bControlKey = False
        bAltKey = False
        
        ' Pull off Control, Alt or Shift specifiers
        sKey = Left$(sKeystrokes, 1)
        sKeystrokes = Mid$(sKeystrokes, 2)
        
        Do While InStr(" ^%+", sKey) > 1 ' The space in " ^%+" is necessary
            If sKey = "+" Then
                bShiftKey = True
            ElseIf sKey = "^" Then
                bControlKey = True
            ElseIf sKey = "%" Then
                bAltKey = True
            End If
            sKey = Left$(sKeystrokes, 1)
            sKeystrokes = Mid$(sKeystrokes, 2)
        Loop
        
        ' Look for "{}"
        If sKey = "{" Then
            ' Look for the  "}"
            i = InStr(sKeystrokes, "}")
            If i > 0 Then
                sKey = Left$(sKeystrokes, i - 1) ' extract the content between
                                                 ' the {}
                sKeystrokes = Mid$(sKeystrokes, i + 1) ' Remove the }
            End If
        
            ' Look for repetitions
            i = Len(sKey)
            Do While Mid$(sKey, i, 1) >= "0" And Mid$(sKey, i, _
                1) <= "9" And i >= 3
                i = i - 1
            Loop
        
            If i < Len(sKey) Then ' If any digits were found...
                If i >= 2 Then ' If there is something preceding it...
                    If Mid$(sKey, i, 1) = " " Then  ' If a space precedes the
                                                    ' digits...
                        On Error Resume Next ' On overflow, ignore the value
                        lRepetitions = CLng(Mid$(sKey, i + 1))
                        On Error GoTo 0
                        sKey = Left$(sKey, i - 1)
                    End If
                End If
            End If
        End If
        
        ' Look for special words
        Select Case UCase$(sKey)
            Case "LBUTTON" ' New
                iVirtualKey = vbKeyLButton
            Case "RBUTTON" ' New
                iVirtualKey = vbKeyRButton
            Case "BREAK", "CANCEL"
                iVirtualKey = vbKeyCancel
            Case "MBUTTON" ' New
                iVirtualKey = vbKeyMButton
            Case "BACKSPACE", "BS", "BKSP"
                iVirtualKey = vbKeyBack
            Case "TAB"
                iVirtualKey = vbKeyTab
            Case "CLEAR" ' New
                iVirtualKey = vbKeyClear
            Case "ENTER", "~"
                iVirtualKey = vbKeyReturn
            Case "SHIFT" ' New
                iVirtualKey = vbKeyShift
            Case "CONTROL" ' New
                iVirtualKey = vbKeyControl
            Case "MENU", "ALT" ' New
                iVirtualKey = vbKeyMenu
            Case "PAUSE" ' New
                iVirtualKey = vbKeyPause
            Case "CAPSLOCK"
                iVirtualKey = vbKeyCapital
            Case "ESCAPE", "ESC"
                iVirtualKey = vbKeyEscape
            Case "SPACE" ' New
                iVirtualKey = vbKeySpace
            Case "PGUP"
                iVirtualKey = vbKeyPageUp
            Case "PGDN"
                iVirtualKey = vbKeyPageDown
            Case "END"
                iVirtualKey = vbKeyEnd
            Case "HOME" ' New
                iVirtualKey = vbKeyHome
            Case "LEFT"
                iVirtualKey = vbKeyLeft
            Case "UP"
                iVirtualKey = vbKeyUp
            Case "RIGHT"
                iVirtualKey = vbKeyRight
            Case "DOWN"
                iVirtualKey = vbKeyDown
            Case "SELECT" ' New
                iVirtualKey = vbKeySelect
            Case "PRTSC"
                iVirtualKey = vbKeyPrint
            Case "EXECUTE" ' New
                iVirtualKey = vbKeyExecute
            Case "SNAPSHOT" ' New
                iVirtualKey = vbKeySnapshot
            Case "INSERT", "INS"
                iVirtualKey = vbKeyInsert
            Case "DELETE", "DEL"
                iVirtualKey = vbKeyDelete
            Case "HELP"
                iVirtualKey = vbKeyHelp
            Case "NUMLOCK"
                iVirtualKey = vbKeyNumlock
            Case "SCROLLLOCK"
                iVirtualKey = vbKeyScrollLock
            Case "NUMPAD0" ' New
                iVirtualKey = vbKeyNumpad0
            Case "NUMPAD1" ' New
                iVirtualKey = vbKeyNumpad1
            Case "NUMPAD2" ' New
                iVirtualKey = vbKeyNumpad2
            Case "NUMPAD3" ' New
                iVirtualKey = vbKeyNumpad3
            Case "NUMPAD4" ' New
                iVirtualKey = vbKeyNumpad4
            Case "NUMPAD5" ' New
                iVirtualKey = vbKeyNumpad5
            Case "NUMPAD6" ' New
                iVirtualKey = vbKeyNumpad6
            Case "NUMPAD7" ' New
                iVirtualKey = vbKeyNumpad7
            Case "NUMPAD8" ' New
                iVirtualKey = vbKeyNumpad8
            Case "NUMPAD9" ' New
                iVirtualKey = vbKeyNumpad9
            Case "MULTIPLY" ' New
                iVirtualKey = vbKeyMultiply
            Case "ADD" ' New
                iVirtualKey = vbKeyAdd
            Case "SEPARATOR" ' New
                iVirtualKey = vbKeySeparator
            Case "SUBTRACT" ' New
                iVirtualKey = vbKeySubtract
            Case "DECIMAL" ' New
                iVirtualKey = vbKeyDecimal
            Case "DIVIDE" ' New
                iVirtualKey = vbKeyDivide
            Case "F1"
                iVirtualKey = vbKeyF1
            Case "F2"
                iVirtualKey = vbKeyF2
            Case "F3"
                iVirtualKey = vbKeyF3
            Case "F4"
                iVirtualKey = vbKeyF4
            Case "F5"
                iVirtualKey = vbKeyF5
            Case "F6"
                iVirtualKey = vbKeyF6
            Case "F7"
                iVirtualKey = vbKeyF7
            Case "F8"
                iVirtualKey = vbKeyF8
            Case "F9"
                iVirtualKey = vbKeyF9
            Case "F10"
                iVirtualKey = vbKeyF10
            Case "F11"
                iVirtualKey = vbKeyF11
            Case "F12"
                iVirtualKey = vbKeyF12
            Case "F13"
                iVirtualKey = vbKeyF13
            Case "F14"
                iVirtualKey = vbKeyF14
            Case "F15"
                iVirtualKey = vbKeyF15
            Case "F16"
                iVirtualKey = vbKeyF16
            Case Else
                ' Not a virtual key
                iVirtualKey = -1
        End Select
        
        ' Turn on CONTROL, ALT and SHIFT keys as needed
        If bShiftKey Then
            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                VirtualKeys(vbKeyShift).scanCode, KEYEVENTF_KEYDOWN, 0
        End If
        
        If bControlKey Then
            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                VirtualKeys(vbKeyControl).scanCode, KEYEVENTF_KEYDOWN, 0
        End If
        
        If bAltKey Then
            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                VirtualKeys(vbKeyMenu).scanCode, KEYEVENTF_KEYDOWN, 0
        End If
        
        ' Send the keystrokes
        For i = 1 To lRepetitions
            If iVirtualKey > -1 Then
                ' Virtual key
                keybd_event VirtualKeys(iVirtualKey).VKCode, _
                    VirtualKeys(iVirtualKey).scanCode, KEYEVENTF_KEYDOWN, 0
                keybd_event VirtualKeys(iVirtualKey).VKCode, _
                    VirtualKeys(iVirtualKey).scanCode, KEYEVENTF_KEYUP, 0
            Else
                ' ASCII Keys
                For j = 1 To Len(sKey)
                    iAsciiKey = Asc(Mid$(sKey, j, 1))
                    ' Turn on CONTROL, ALT and SHIFT keys as needed
                    If Not bShiftKey Then
                        If AsciiKeys(iAsciiKey).Shift Then
                            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                                VirtualKeys(vbKeyShift).scanCode, _
                                KEYEVENTF_KEYDOWN, 0
                        End If
                    End If
        
                    If Not bControlKey Then
                        If AsciiKeys(iAsciiKey).Control Then
                            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                                VirtualKeys(vbKeyControl).scanCode, _
                                KEYEVENTF_KEYDOWN, 0
                        End If
                    End If
        
                    If Not bAltKey Then
                        If AsciiKeys(iAsciiKey).Alt Then
                            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                                VirtualKeys(vbKeyMenu).scanCode, _
                                KEYEVENTF_KEYDOWN, 0
                        End If
                    End If
        
                    ' Press the key
                    keybd_event AsciiKeys(iAsciiKey).VKCode, _
                        AsciiKeys(iAsciiKey).scanCode, KEYEVENTF_KEYDOWN, 0
                    keybd_event AsciiKeys(iAsciiKey).VKCode, _
                        AsciiKeys(iAsciiKey).scanCode, KEYEVENTF_KEYUP, 0
        
                    ' Turn on CONTROL, ALT and SHIFT keys as needed
                    If Not bShiftKey Then
                        If AsciiKeys(iAsciiKey).Shift Then
                            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                                VirtualKeys(vbKeyShift).scanCode, _
                                KEYEVENTF_KEYUP, 0
                        End If
                    End If
        
                    If Not bControlKey Then
                        If AsciiKeys(iAsciiKey).Control Then
                            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                                VirtualKeys(vbKeyControl).scanCode, _
                                KEYEVENTF_KEYUP, 0
                        End If
                    End If
        
                    If Not bAltKey Then
                        If AsciiKeys(iAsciiKey).Alt Then
                            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                                VirtualKeys(vbKeyMenu).scanCode, _
                                KEYEVENTF_KEYUP, 0
                        End If
                    End If
                Next j ' Each ASCII key
            End If  ' ASCII keys
        Next i ' Repetitions
        
        ' Turn off CONTROL, ALT and SHIFT keys as needed
        If bShiftKey Then
            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                VirtualKeys(vbKeyShift).scanCode, KEYEVENTF_KEYUP, 0
        End If
        
        If bControlKey Then
            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                VirtualKeys(vbKeyControl).scanCode, KEYEVENTF_KEYUP, 0
        End If
        
        If bAltKey Then
            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                VirtualKeys(vbKeyMenu).scanCode, KEYEVENTF_KEYUP, 0
        End If
        
    Loop ' sKeyStrokes
End Sub

'------------------------------------------------------------------------------
' To allow using XP manifests

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function
'******************************************************************************
'************************* End modification/addition by JR ********************
'******************************************************************************

Private Sub Form_Initialize()

'ANI -- MODIFICATION
'Define the time adjustment factor: 0.3 sec

    TimeAdjustmentConstant = 0.3

'END OF MODIFICATION

    Dim fso As FileSystemObject
    
    InitCommonControlsVB
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists("LastRun.txt") Then 'Clear log file
        fso.DeleteFile ("LastRun.txt")
    End If
    
    txtUmler.Text = App.Path & "\" & "miniUMLER.mdb"
    txtOutFName.Text = "Results.txt"
End Sub

Private Sub EnableButtons(NewState As Boolean)
    cmdAero.Enabled = NewState
    cmdGap.Enabled = NewState
    cmdGenerate.Enabled = NewState
    cmdUmler.Enabled = NewState
    cmdSetOutPath.Enabled = NewState
End Sub

Private Sub SetStatus(S As String)
    Sbar.Panels(1).Text = S
    If BlindMode Then AppendToLog (S)
    DoEvents
End Sub

Private Function ValidateFile(S As String) As String
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(S) Then
      ValidateFile = S
     Else
      ValidateFile = ""
    End If
End Function

Private Function GetPath(ByVal S As String) As String
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    S = fso.GetParentFolderName(S)
    If fso.FolderExists(S) Then
        GetPath = fso.BuildPath(S, "") 'Includes last back slash
      Else
        GetPath = ""
    End If
End Function

Private Function GetFileName(ByVal S As String) As String
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    GetFileName = fso.GetFileName(S)
End Function

Private Sub Form_Load()
    Dim strArgs() As String
    Dim OutPath, OutFName, tmp As String
    Dim i, delta As Integer
    
    strArgs = Split(Command$, " ")
    If UBound(strArgs) < 0 Then
        txtOutPath.Text = App.Path
        Exit Sub
    End If
    
    BlindMode = (strArgs(0) = "-b") Or (strArgs(0) = "/b")
    
    If BlindMode Then
        delta = 1
    Else
        delta = 0
    End If
    
    If UBound(strArgs) >= 0 + delta Then
        'At least axles filename
        txtAero.Text = ValidateFile(strArgs(0 + delta))
    End If
    
    If UBound(strArgs) >= 1 + delta Then
        'At least axles and gaps
        txtGap.Text = ValidateFile(strArgs(1 + delta))
    End If
    
    If UBound(strArgs) >= 2 + delta Then
        'Axles, gaps and database
        txtUmler.Text = ValidateFile(strArgs(2 + delta))
    End If
    
    If UBound(strArgs) >= 3 + delta Then
        'Output directory and filename (ignore extensions)
        tmp = GetPath(strArgs(3 + delta))
        If tmp <> "" Then
            txtOutPath.Text = tmp
            txtOutFName.Text = GetFileName(strArgs(3 + delta))
        Else
            txtOutPath.Text = App.Path
            If Not BlindMode Then MsgBox "Path is invalid or not specified. Using application path instead", _
                vbExclamation, "Invalid paramter found"
       End If
    End If
    
    If BlindMode Then
        If (txtAero.Text <> "") And (txtGap.Text <> "") And (txtUmler.Text <> "") _
              And (txtOutPath.Text <> "") Then
            Timer1.Enabled = True   ' Allow for the form to be painted before processing
        Else
            MsgBox "Can not enter blind mode with missing parameters", _
                vbCritical, "Quitting"
            End
        End If
    End If
    
    AbortRun = False
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    DoGenerate
End Sub

Private Function SafeMsgBox(prompt, buttons, title)
    If Not (BlindMode) Then
        SafeMsgBox = MsgBox(prompt, buttons, title)
      Else
        AppendToLog ("[" & title & "] " & prompt)
        SafeMsgBox = vbIgnore
    End If
End Function

Private Sub AppendToLog(S As String)
    Dim fso As FileSystemObject, F
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.OpenTextFile("LastRun.txt", ForAppending, True)
    F.WriteLine (S)
    F.Close
End Sub




