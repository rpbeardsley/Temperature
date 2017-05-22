VERSION 5.00
Begin VB.Form frmPID 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PID Temperature Controller"
   ClientHeight    =   3720
   ClientLeft      =   1830
   ClientTop       =   2835
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3720
   ScaleWidth      =   7650
   Begin VB.ComboBox cmbYdevice 
      Height          =   315
      Left            =   6510
      TabIndex        =   27
      Text            =   "Combo1"
      Top             =   2760
      Width           =   945
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   4650
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1620
      Width           =   2865
   End
   Begin VB.TextBox txtSampleTime 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5220
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   2910
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   6990
      Top             =   870
   End
   Begin VB.TextBox txtSweepRate 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1710
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1980
      Width           =   1005
   End
   Begin VB.TextBox txtOutputV 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1710
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   6990
      Top             =   450
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6990
      Top             =   30
   End
   Begin VB.TextBox txtSetT 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1710
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   1005
   End
   Begin VB.CommandButton cmdTakeData 
      Caption         =   "Take data"
      Height          =   405
      Left            =   3930
      TabIndex        =   15
      Top             =   2400
      Width           =   2205
   End
   Begin VB.PictureBox Spin1 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   5820
      ScaleHeight     =   315
      ScaleWidth      =   285
      TabIndex        =   19
      Top             =   2880
      Width           =   345
   End
   Begin VB.CommandButton cmdSetOutputV 
      Caption         =   "Set Output Voltage"
      Height          =   405
      Left            =   300
      TabIndex        =   9
      Top             =   2580
      Width           =   2205
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   405
      Left            =   3900
      TabIndex        =   5
      Top             =   840
      Width           =   2205
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Parameters"
      Height          =   405
      Left            =   3900
      TabIndex        =   4
      Top             =   150
      Width           =   2205
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Regulate Temperature"
      Height          =   405
      Left            =   300
      TabIndex        =   3
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label15 
      Caption         =   "Y device"
      Height          =   255
      Left            =   6540
      TabIndex        =   28
      Top             =   2490
      Width           =   885
   End
   Begin VB.Line Line2 
      X1              =   3570
      X2              =   7530
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label13 
      Caption         =   "File name"
      Height          =   255
      Left            =   3750
      TabIndex        =   25
      Top             =   1650
      Width           =   885
   End
   Begin VB.Line Line3 
      X1              =   3570
      X2              =   3570
      Y1              =   120
      Y2              =   3570
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3570
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label Label4 
      Caption         =   "V"
      Height          =   255
      Left            =   2790
      TabIndex        =   24
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "K"
      Height          =   285
      Left            =   2790
      TabIndex        =   23
      Top             =   1530
      Width           =   285
   End
   Begin VB.Label Label11 
      Caption         =   "K"
      Height          =   255
      Left            =   2790
      TabIndex        =   22
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label Label10 
      Caption         =   "K / s"
      Height          =   255
      Left            =   2790
      TabIndex        =   21
      Top             =   2010
      Width           =   525
   End
   Begin VB.Label Label9 
      Caption         =   "K"
      Height          =   255
      Left            =   2790
      TabIndex        =   20
      Top             =   660
      Width           =   225
   End
   Begin VB.Label lblYdata 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5220
      TabIndex        =   18
      Top             =   3300
      Width           =   1365
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data value"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3930
      TabIndex        =   17
      Top             =   3330
      Width           =   1275
   End
   Begin VB.Label lblCurrTargetT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1710
      TabIndex        =   16
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Target Temp"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   300
      TabIndex        =   14
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time Interval"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3930
      TabIndex        =   12
      Top             =   2910
      Width           =   1275
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Max sweep rate"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   300
      TabIndex        =   10
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output Voltage"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   330
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblCurrT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1710
      TabIndex        =   6
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current temp."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   690
      Width           =   1245
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sweep to"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   1110
      Width           =   1395
   End
End
Attribute VB_Name = "frmPID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdParams_Click()
Timer1.Enabled = False
Timer2.Enabled = True
cmdStartStop.Caption = "Regulate Temperature"
frmParams.Show
End Sub

Private Sub cmdQuit_Click()
Dim bool As Integer
bool = WritePrivateProfileString(ByVal "PID Parameters", ByVal "PROP", ByVal Str$(gProp), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "PID Parameters", ByVal "DIFF", ByVal Str$(gDiff), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "PID Parameters", ByVal "INTG", ByVal Str$(gIntg), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "Others", ByVal "SensorCurrent", ByVal Str$(gSensorCurrent), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "Others", ByVal "MaxOutputVoltage", ByVal Str$(gMaxV), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "Others", ByVal "MaxSweepRate", ByVal Str$(gMaxSweepRate), ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
bool = WritePrivateProfileString(ByVal "Others", ByVal "FileName", ByVal gFileName, ByVal "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")

SetOutputV (0)
End
End Sub

Private Sub cmdSetOutputV_Click()
SetOutputV (Val(txtOutputV.Text))
End Sub

Private Sub cmdSetup_Click()
Timer1.Enabled = False
Timer2.Enabled = False
frmPID!cmdStartStop.Caption = "Regulate Temperature"
frmParams.Show
End Sub

Private Sub cmdStartStop_Click()
If cmdStartStop.Caption = "Regulate Temperature" Then
    cmdStartStop.Caption = "Stop"
    gIntgValue = 0
    gSweepT = Val(lblCurrT.Caption)
    Timer2.Enabled = False
    Timer1.Enabled = True
    Else
    SetOutputV (0)
    cmdStartStop.Caption = "Regulate Temperature"
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub cmdSweep_Click()
Timer2.Enabled = False
Timer1.Enabled = True

End Sub

Private Sub cmdTakeData_Click()
gFileName = txtFileName.Text
gYdevice = cmbYdevice.Text


If cmdTakeData.Caption = "Take Data" Then
    gFileNum = FreeFile
    Open gFileName For Output As gFileNum
    cmdTakeData.Caption = "Stop Data"
    Timer2.Enabled = False
    Timer3.Enabled = True
Else
    Close gFileNum
    cmdTakeData.Caption = "Take Data"
    Timer3.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Command3D1_Click()
sendIEEEcmd (Str$(psu))
sendIEEEdata ("X1V1A")
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

Timer1.Interval = 100
Timer2.Interval = 2000
Timer3.Interval = 1000
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = False

gFileName = GetINIstring("PID Parameters", "FileName")
gYdevice = "Dvm1"
gScaleP = True


gProp = Val(GetINIstring("PID Parameters", "PROP"))
gDiff = Val(GetINIstring("PID Parameters", "DIFF"))
gIntg = Val(GetINIstring("PID Parameters", "INTG"))
gSensorCurrent = Val(GetINIstring("Others", "SensorCurrent"))
gMaxV = Val(GetINIstring("Others", "MaxOutputVoltage"))
gMaxSweepRate = Val(GetINIstring("Others", "MaxSweepRate"))
gFileName = GetINIstring("Others", "FileName")

lblCurrT.Caption = Format$(GetTemperature(), "##.##")
txtSetT.Text = lblCurrT.Caption
lblCurrTargetT.Caption = lblCurrT.Caption
txtSweepRate.Text = Str$(gMaxSweepRate)
txtSampleTime.Text = 1
lblYdata.Caption = Str$(GetYData())
cmbYdevice.AddItem "Dvm1"
cmbYdevice.AddItem "Scope cursors"
cmbYdevice.ListIndex = 0
txtFileName.Text = gFileName

sendIEEEcmd ("TO 8 END=ON EOS=10")
SetOutputV (0)
End Sub

Private Function GetDvm1()
'Do
    sendIEEEcmd (Str$(dvm1))
    sendIEEEdata ("S")
    L$ = RecvIEEEdata$(255)
'Loop Until Val(L$) <> 0
GetDvm1 = Val(L$)
End Function

Private Function GetLockin() As Single
'Do
sendIEEEcmd (Str$(lck))
sendIEEEdata ("OUTP?1")
L$ = RecvIEEEdata$(255)
'Loop Until Val(L$) <> 0
GetLockin = Val(L$)
End Function

Private Function GetTemperature() As Single
Dim R, T As Single
'Do
R = GetLockin() / (gSensorCurrent * 0.000001)
Rem T = AllenBradley(R)
If R = 0 Or R < 0 Then
    MsgBox "Invalid lockin reading!"
    GetTemperature = 0
    Exit Function
    End If
T = Cernox(R)
'Loop Until T > 0
GetTemperature = T
End Function

Private Function GetYData()

Select Case gYdevice
Case "Scope cursors"
    sendIEEEcmd (Str$(scope))
    sendIEEEdata (":MEASURE:TSTART?")
    t1 = Val(RecvIEEEdata$(255))
    sendIEEEcmd (Str$(scope))
    sendIEEEdata (":MEASURE:TSTOP?")
    t2 = Val(RecvIEEEdata$(255))
    sendIEEEcmd (Str$(scope))
    sendIEEEdata (":MEASURE:VTIME? " + Str$(t1))
    v1 = Val(RecvIEEEdata$(255))
    sendIEEEcmd (Str$(scope))
    sendIEEEdata (":MEASURE:VTIME? " + Str$(t2))
    v2 = Val(RecvIEEEdata$(255))
    YData = v2 - v1
Case "Dvm1"
    YData = GetDvm1()
End Select
GetYData = YData
End Function

Private Sub SetOutputV(V As Single)
sendIEEEcmd (Str$(psu))
'set psu output voltage without waiting for verification
sendIEEEdata ("V " + Format$(V))
txtOutputV.Text = Format$(V, "##.#")
End Sub

Private Sub Spin1_SpinDown()
txtSampleTime.Text = Str$(0.5 * Val(txtSampleTime.Text))
End Sub

Private Sub Spin1_SpinUp()
txtSampleTime.Text = Str$(2 * Val(txtSampleTime.Text))
End Sub

Private Sub Timer1_Timer()
Static OldT, LastTime As Single
DoEvents

ThisTime = Timer
If LastTime = 0 Then LastTime = ThisTime
DeltaTime = ThisTime - LastTime

SetT = Val(txtSetT.Text)
NewT = GetTemperature()
If gScaleP = True Then
    Prop = gProp * NewT / 300   'scale proportional term with current temperature
    Else
    Prop = gProp
    End If
lblCurrT.Caption = Format$(NewT, "##.##")
If OldT = 0 Then OldT = NewT

If gSweepT < SetT Then
    gSweepT = gSweepT + gMaxSweepRate * DeltaTime
ElseIf gSweepT > SetT Then
    gSweepT = gSweepT - gMaxSweepRate * DeltaTime
End If
lblCurrTargetT.Caption = Format$(gSweepT, "###.#")

If Abs(SetT - gSweepT) < (gMaxSweepRate * DeltaTime) Then gSweepT = SetT

gIntgValue = gIntgValue + (gSweepT - NewT)
OutputVoltage = Prop * (gSweepT - NewT) + gDiff * (NewT - OldT) + gIntg * gIntgValue

If OutputVoltage > gMaxV Then OutputVoltage = gMaxV
If OutputVoltage < 0 Then OutputVoltage = 0

SetOutputV (OutputVoltage)

OldT = NewT
LastTime = ThisTime
End Sub

Private Sub Timer2_Timer()
DoEvents
lblCurrT.Caption = Format$(GetTemperature(), "##.##")
lblYdata.Caption = GetYData()
End Sub

Private Sub Timer3_Timer()
DoEvents
Temperature = GetTemperature()
YData = GetYData()
lblYdata.Caption = Str$(YData)
Print #gFileNum, Temperature, YData
Timer3.Interval = 1000 * Val(txtSampleTime.Text)
End Sub

Private Sub txtSweepRate_Change()
gMaxSweepRate = Val(txtSweepRate.Text)
End Sub

