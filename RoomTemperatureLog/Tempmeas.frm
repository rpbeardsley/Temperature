VERSION 5.00
Begin VB.Form CentTempMeas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Centigrade temperature measurement"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Turn on sensor 2"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Turn on sensor 1"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combodelay 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Time interval units"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox combotime 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Run time units"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox inputinterval 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox inputTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "End"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMeasure 
      Caption         =   "Measure Temperature"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Sensor 2"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Sensor 1"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Temperature (C):"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Time to next measurement (Seconds):"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblinterval 
      Caption         =   "Time interval between measurements"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblHowlong 
      Caption         =   "Run time"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "CentTempMeas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim dmm As Long
    Dim dmm2 As Long
    
Private Sub Form_Load()

    Dim Termin As TermT
    Dim ieee As Long
    Dim response As String
    Dim B As Long
    
    combotime.AddItem ("Days")
    combotime.AddItem ("Hours")
    combotime.AddItem ("Minutes")
    Combodelay.AddItem ("Hours")
    Combodelay.AddItem ("Minutes")
    Combodelay.AddItem ("Seconds")
    
    Check1.Value = 0
    Check2.Value = 0
    
    '-1 is not a byte as makenewdevice expects, yet this is advised?
    'dmm = IOTIEEE.MakeNewDevice("IEEE0", "DMM", 12, -1, Termin, Termin, 1000)
    'If dmm = -1 Then
       'MsgBox ("Can't create DMM")
       'End
    'End If
    
    
    dmm = OpenName("DMM")
    dmm2 = OpenName("DMM2")
    
    If dmm = -1 Then
        MsgBox ("Can't open DMM")
        End
    End If
        
    If dmm2 = -1 Then
        MsgBox ("Can't open DMM2")
        End
    End If
    
    response = ""
    B = GetError(ieee, response)
    ioClear (dmm)
    ioClear (dmm2)
    
End Sub

Private Sub cmdAbort_Click()
    
    ioClear (dmm)
    ioClear (dmm2)
    
    If dmm > -1 Then
        ioClose (dmm)
    End If
    
    If dmm2 > -1 Then
        ioClose (dmm2)
    End If
    
    End

End Sub

Private Sub cmdMeasure_Click()
    
    Dim Timehours As Integer
    Dim Time As Long
    Dim Measdelayminutes As Integer
    Dim Measdelay As Integer
    Dim Npoints As Long
    Dim Mpoints As Single
    Dim count As Integer
    Dim Timedays As Integer
    Dim Timeminutes As Integer
    Dim Measdelayhours As Integer
    Dim Message As String
    Dim Saveas As String
    Dim DataDir As String
    Dim a As Single
    Dim Complete As Integer
    Dim toofast As Integer
    Dim Info As Integer
    Dim dmm As Long
    Dim Termin As TermT
    Dim Compr As Long
    Dim Comps As Long
    Dim Send As Long
    Dim Voltage As Long
    Dim data As String
    Dim Temp As Single
    Dim strtosend As String
    Dim oldTimer As Single
    Dim currentTimer As Single
    Dim infocheck As Integer
    Dim Correction As Single
    Dim Send2 As Long
    Dim Voltage2 As Long
    Dim data2 As String
    Dim Temp2 As Single
    Dim Saveas2 As String
    Dim Xpos()
    Dim Ypos()
    Dim TempPlot As Single
    Dim DayCnt As Integer
    
    If combotime = "Hours" Then
        Timehours = Val(inputTime.Text)
        Time = Timehours * 3600
    ElseIf combotime = "Days" Then
        Timedays = Val(inputTime.Text)
        Time = Timedays * 86400
    ElseIf combotime = "Minutes" Then
        Timeminutes = Val(inputTime.Text)
        Time = Timeminutes * 60
    End If
        
    If Combodelay = "Hours" Then
        Measdelayhours = Val(inputinterval.Text)
        Measdelay = Measdelayhours * 3600
    ElseIf Combodelay = "Minutes" Then
        Measdelayminutes = Val(inputinterval.Text)
        Measdelay = Measdelayminutes * 60
    ElseIf Combodelay = "Seconds" Then
        Measdelay = Val(inputinterval.Text)
    End If
    
    If inputTime.Text = "" Or inputinterval.Text = "" Or Combodelay = "" Or combotime = "" Then
        Info = MsgBox("You have not entered all of the information. Try again", 48, "Temperature Measurment")
        GoTo ext
    End If
    
    If Check1.Value = 0 And Check2.Value = 0 Then
        infocheck = MsgBox("Turn on one or both of the senors", 48, "Temperature Measurment")
        GoTo ext
    End If
    
    If Measdelay < 180 Then
        toofast = MsgBox("The response time of the LM35 temperature sensor is 3" _
        & vbCrLf & "minutes in still air your measurement interval is lower than this")
    End If
    
    Mpoints = Time / Measdelay
    Npoints = Mpoints
      
    If Mpoints < 1 Then
        Message = MsgBox("The interval you have selected is larger than the full measurement time", vbRetryCancel, "Error")
        GoTo ext
    Else
        Message = 0
    End If
    
    DataDir = InputBox("Save data in:", "Temperature Data", "C:\WINDOWS\Desktop\Temperature Measurement\Temperature")
    
    If DataDir = "" Then
       GoTo ext
    End If
            
    If (Right(DataDir, 1) <> "\") Then DataDir = DataDir + "\"
    Saveas = DataDir + Format(Day(Now), "00") + "-" + Format(Month(Now), "00") + "-" + Format(Year(Now)) + ".txt"
        
    Open Saveas For Output As #1
    Print #1, ":Time(S)", "Sensor 1 (C)", "Sensor 2 (C)"
    
    inputTime.Enabled = False
    inputinterval.Enabled = False
    combotime.Enabled = False
    Combodelay.Enabled = False
    cmdMeasure.Enabled = False
    
    'Terminator structure
    Termin.EOI = 1
    Termin.EightBits = 1
    Termin.nChar = 2
    Termin.Term1 = 10
    Termin.Term2 = 13
    
    count = 0
    a = 100
    data = String(255, " ")
    Comps = 1
    Compr = 1
    DayCnt = 0
    
    If Check1.Value = 0 Then Check1.Enabled = False
    If Check2.Value = 0 Then Check2.Enabled = False
    
    strtosend = "U0M0R2I3N0T0"
    Send = IOTIEEE.OutputXdll(dmm, strtosend, Len(strtosend), 1, 1, Termin, 0, Comps)
    Send2 = IOTIEEE.OutputXdll(dmm2, strtosend, Len(strtosend), 1, 1, Termin, 0, Comps)
    
    ReDim Xpos(0 To 0)
    ReDim Ypos(0 To 0)
    
    oldTimer = 0
    For count = 1 To Npoints
        
        If Check1.Value = 1 Then
            Send = IOTIEEE.OutputXdll(dmm, "G", 1, 1, 1, ByVal 0&, 0, Comps)
            Label1.Caption = "Reading..."
            DoEvents
            data = String(9, " ")
            Voltage = IOTIEEE.EnterXdll(dmm, data, Len(data), 1, Termin, 0, Compr)
            currentTimer = Timer
        
            Temp = Val(data) * a
            Label3.Caption = Temp
        Else
            Temp = 0
        End If
        
        If Check2.Value = 1 Then
            Send2 = IOTIEEE.OutputXdll(dmm2, "G", 1, 1, 1, ByVal 0&, 0, Comps)
            DoEvents
            data2 = String(9, " ")
            Voltage2 = IOTIEEE.EnterXdll(dmm2, data2, Len(data), 1, Termin, 0, Compr)
            currentTimer = Timer
        
            Temp2 = Val(data2) * a
            Label5.Caption = Temp2
        Else
            Temp2 = 0
        End If
        
        If Check1.Value = False Then
            TempPlot = Temp2
        Else
            TempPlot = Temp
        End If
        
        ReDim Preserve Xpos(0 To count - 1)
        ReDim Preserve Ypos(0 To count - 1)
        Xpos(count - 1) = currentTimer + (DayCnt * 86400)
        Ypos(count - 1) = TempPlot
        TraceForm.PlotTrace Xpos, Ypos
        
        If currentTimer < oldTimer Then
            Close #1
            If (Saveas = DataDir + Format(Day(Now), "00") + "-" + Format(Month(Now), "00") + "-" + Format(Year(Now)) + ".txt") Then
                MsgBox "Error! - I'm trying to overwrite my last file! - Quiting", vbCritical, "Error!!"
                GoTo ext
            End If
    
            Saveas = DataDir + Format(Day(Now), "00") + "-" + Format(Month(Now), "00") + "-" + Format(Year(Now)) + ".txt"
            Open Saveas For Output As #1
            Print #1, ":Time(S)", "Sensor 1 (C)", "Sensor 2 (C)"
            
            If DayCnt = 0 Then
                DayCnt = 1
            ElseIf DayCnt <> 0 Then
                DayCnt = DayCnt + 1
            End If
            
        End If
            
        oldTimer = currentTimer
        Print #1, Timer, Temp, Temp2
        
        If Check1.Value = 1 And Check2.Value = 1 Then
            Correction = 0.8
        Else
            Correction = 0.4
        End If
        
        EventPause (Measdelay - Correction) 'This must be -0.8 for both sensors (replace 0.4 with Correction)
             
    Next count
                                                                                                                                                                                                                                 
    Close #1
                                                                                                                                                                                                                                    
    Complete = MsgBox("Measurement complete your data is saved at:" _
    & vbCrLf & _
    vbCrLf & Saveas, vbOKOnly, "Temperature Measurement")
                                                                                                                                                                                                                             
    ioClear (dmm)
    
    inputTime.Enabled = True
    inputinterval.Enabled = True
    combotime.Enabled = True
    Combodelay.Enabled = True
    cmdMeasure.Enabled = True
    
ext:
                                                                                                                                                                                                                             
    'IOTIEEE.RemoveDevice (dmm)
                                                                                                                                                                                                                             
End Sub

Private Function DoesFileExist(FilePath As String, Optional FileAttr As VbFileAttribute) As Boolean

    If Len(Dir$(FilePath, FileAttr)) > 0 Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If
    
End Function

Private Function EventPause(Seconds As Single)

    Dim starttimer As Single
    Dim offset As Single
    
    starttimer = Timer
    offset = 0
    While (Timer + offset < (starttimer + Seconds))
    
        If Timer < starttimer Then offset = 86400 'If we go past midnight add a day's worth of seconds onto Timer
        Label1.Caption = (starttimer + Seconds - (Timer + offset))
        DoEvents
    
    Wend
      
End Function

Private Sub Form_Terminate()

    If dmm > -1 Then
        ioClose (dmm)
    End If
    
    If dmm2 > -1 Then
        ioClose (dmm2)
    End If
    
    'IOTIEEE.RemoveDevice (dmm)
    
End Sub

