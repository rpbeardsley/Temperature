Attribute VB_Name = "MODPID1"
' Required declarations for the Pro488 IEEE board
Declare Function Pro488Present Lib "pro488.dll" () As Integer
Declare Function Pro488ReadData Lib "pro488.dll" (ByVal Buffer$, ByVal NumOfBytes%) As Integer
Declare Function Pro488WriteData Lib "pro488.dll" (ByVal Buffer$, ByVal NumOfBytes%) As Integer
Declare Function Pro488ReadCtrl Lib "pro488.dll" (ByVal Buffer$) As Integer
Declare Function Pro488WriteCtrl Lib "pro488.dll" (ByVal Buffer$, ByVal NumOfBytes%) As Integer

' Win API function declarations
Declare Function Getprivateprofilestring Lib "Kernel32" (ByVal lpApplicationName As String, lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFilename As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer

' Define a few global variables and constants
Global Const dvm1 = 16
Global Const psu = 11
Global Const lck = 9
Global Const scope = 7

Global Const mirrorscale = 384

Global gStop As Integer
Global gProp
Global gDiff
Global gIntg
Global gSensorCurrent
Global gMaxV
Global gMaxSweepRate
Global gYdevice
Global gSweepT
Global gFileNum As Integer
Global gFileName As String
Global gScaleP As Integer
Function AllenBradley(R)
Rem calibration for 300ohm Allen Bradley taken on 12/11/97
log10R = Log(R) / Log(10)
T = 1 / (4.33 / log10R - 3.18 + 0.589 * log10R)
AllenBradley = T
End Function

Function GetINIInt(header$, subheader$)
GetINIInt = GetPrivateProfileInt(header$, subheader$, 0, "c:\users\phil\software - homebrew\pid\pid.ini")
End Function

Function GetINIstring(header$, subheader$)
buf$ = String(255, 0)
ret% = Getprivateprofilestring(header$, ByVal subheader$, "", buf$, 255, "C:\Documents and Settings\Ryan\My Documents\1st year\Analysis\PID Temperature Control\pid.ini")
GetINIstring = Left$(buf$, ret%)
End Function

Sub pause(tmp)
start = Timer
Do
DoEvents
If gStop = 1 Then Exit Sub
Loop Until Timer > start + tmp
End Sub

Function RecvIEEEdata$(length)
inbuf$ = String(length, 0)
ret% = Pro488ReadData(inbuf$, length)
RecvIEEEdata$ = Left$(inbuf$, ret%)
End Function

Function Cernox(R)
Dim coefs(10) As Single

R = Abs(R)  'make sure we don't get any -ve values

If R > 632.65 Then
    a = 8.026026
    b = -16.961066
    n = 7
    coefs(0) = -0.00019
    coefs(1) = 0.00051
    coefs(2) = 0.0003
    coefs(3) = -0.00847
    coefs(4) = 0.01748
    coefs(5) = 0.11329
    coefs(6) = -1.51295
    coefs(7) = 3.60533

Else
        a = 9.137652
        b = -15.278483
        n = 8
        coefs(0) = 0.000053
        coefs(1) = 0.000089
        coefs(2) = -0.000145
        coefs(3) = -0.000665
        coefs(4) = 0.002312
        coefs(5) = -0.002648
        coefs(6) = -0.069569
        coefs(7) = -1.239404
        coefs(8) = 9.057825
End If
R = Log(Log(R))
x = a * R + b
For i = 0 To n
g = x * f - d + coefs(i)
e = d
d = f
f = g
Next

Cernox = Exp(0.5 * (g - e))


End Function

Sub sendIEEEcmd(dat$)
ret% = Pro488WriteCtrl(dat$ + Chr$(13), Len(dat$) + 1)
End Sub

Sub sendIEEEdata(dat$)
ret% = Pro488WriteData(dat$ + Chr$(13) + Chr$(10), Len(dat$) + 2)
End Sub

