VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmParams 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Parameters"
   ClientHeight    =   2955
   ClientLeft      =   7290
   ClientTop       =   1995
   ClientWidth     =   5265
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   5265
   Begin VB.CheckBox chkScaleP 
      Caption         =   "Scale with T (this value at 300K)"
      Height          =   465
      Left            =   2910
      TabIndex        =   11
      Top             =   150
      Width           =   1875
   End
   Begin VB.TextBox txtMaxV 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1500
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2370
      Width           =   1335
   End
   Begin VB.TextBox txtSensorCurrent 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1500
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1800
      Width           =   1305
   End
   Begin VB.TextBox txtInt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1020
      Width           =   1305
   End
   Begin VB.TextBox txtDiff 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   1305
   End
   Begin VB.TextBox txtProp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdDone 
      Height          =   435
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
      Caption         =   "Done"
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3000
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Maximum output voltage"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   180
      TabIndex        =   9
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cernox test current / uA"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Integral"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Differential"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proportional"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1065
   End
End
Attribute VB_Name = "frmParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDone_Click()
If chkScaleP.Value = 1 Then
    gScaleP = True
    Else
    gScaleP = False
    End If
Unload frmParams
frmPID!Timer2.Enabled = True
End Sub

Private Sub Form_Load()
txtProp.Text = Str$(gProp)
txtDiff.Text = Str$(gDiff)
txtInt.Text = Str$(gIntg)
txtSensorCurrent.Text = Str$(gSensorCurrent)
txtMaxV = Str$(gMaxV)
If gScaleP = True Then
    chkScaleP.Value = 1
    Else
    chkScaleP.Value = 0
    End If
End Sub

Private Sub txtDiff_Change()
gDiff = Val(txtDiff.Text)
End Sub

Private Sub txtInt_Change()
gIntg = Val(txtInt.Text)
End Sub

Private Sub txtMaxV_Change()
gMaxV = Val(txtMaxV.Text)
End Sub

Private Sub txtProp_Change()
gProp = Val(txtProp.Text)
End Sub

Private Sub txtSensorCurrent_Change()
gSensorCurrent = Val(txtSensorCurrent.Text)
End Sub

