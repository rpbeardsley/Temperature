VERSION 4.00
Begin VB.Form frmScope 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Phillips Scope"
   ClientHeight    =   2490
   ClientLeft      =   1665
   ClientTop       =   3675
   ClientWidth     =   7365
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   2895
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   7365
   Top             =   3330
   Width           =   7485
   Begin VB.TextBox txtRecv 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   150
      TabIndex        =   2
      Top             =   870
      Width           =   4275
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   4275
   End
   Begin Threed.SSCommand cmdQuit 
      Height          =   525
      Left            =   5070
      TabIndex        =   6
      Top             =   1770
      Width           =   1935
      _version        =   65536
      _extentx        =   3413
      _extenty        =   926
      _stockprops     =   78
      caption         =   "Quit"
   End
   Begin Threed.SSCommand cmdRecv 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   1020
      Width           =   1905
      _version        =   65536
      _extentx        =   3360
      _extenty        =   873
      _stockprops     =   78
      caption         =   "Receive"
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   525
      Left            =   5010
      TabIndex        =   1
      Top             =   270
      Width           =   1905
      _version        =   65536
      _extentx        =   3360
      _extenty        =   926
      _stockprops     =   78
      caption         =   "Send"
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Recieved data"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Send Data"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "frmScope"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRecv_Click()
txtRecv.Text = RecvIEEEdata(255)
End Sub

Private Sub cmdSend_Click()
sendIEEEcmd (Str$(7))
sendIEEEdata (txtSend.Text)
End Sub

Private Sub cmdStart_Click()
If txtDtime.Text = "" Then txtDtime.Text = "1"
If cmdStart.Caption = "Start Experiment" Then
    filenum = FreeFile
    Open "c:\tmp\tmp.dat" For Output As filenum
    cmdStart.Caption = "Stop"
    Timer1.Interval = 1000 * Val(txtDtime.Text)
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
    cmdStart.Caption = "Start Experiment"
    Close #filenum
End If
End Sub

Private Sub Spin1_SpinDown()
txtDtime.Text = Str$(0.5 * Val(txtDtime.Text))
End Sub

Private Sub Spin1_SpinUp()
txtDtime.Text = Str$(2 * Val(txtDtime.Text))
End Sub

Private Sub Timer1_Timer()
Do
    sendIEEEcmd (Str$(dvm1))
    AB = Val(RecvIEEEdata$(255))
Loop While AB = 0

sendIEEEcmd (Str$(scope))
sendIEEEdata ("SPL CURSOR,DVOLT ?")
dvolt = Val(Mid$(RecvIEEEdata$(255), 7))

Print #filenum, AB, dvolt
labCursorData.Caption = Str$(dvolt)
labDVM1Data.Caption = Str$(AB)
Timer1.Interval = 1000 * Val(txtDtime.Text)

gphPlot.NumPoints = tmp + 1
tmp = tmp + 1
gphPlot.ThisSet = 1
gphPlot.ThisPoint = tmp
gphPlot.XPosData = AB

gphPlot.ThisSet = 1
gphPlot.ThisPoint = tmp
gphPlot.GraphData = dvolt

gphPlot.DrawMode = 2
End Sub

