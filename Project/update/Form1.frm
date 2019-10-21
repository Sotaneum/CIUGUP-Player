VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Updayt체크"
   ClientHeight    =   990
   ClientLeft      =   8745
   ClientTop       =   7125
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4695
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "업데이트"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "최신 버전 : "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "현제 버전 : "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Q As String

Private Sub Command1_Click()
Shell "explorer.exe http://gnyontu.netne.net/ver/CIUGUPPlayer/setup.zip"
End

End Sub


Private Sub Timer1_Timer()
If PB.Value <= 50 Then
Dim i As String
i = PB.Value
i = i + 2
PB.Value = i
Else
If PB.Value >= 50 Then
Dim O, c
Q = "http://gnyontu.netne.net/ver/CIUGUPPlayer/ver.day"
O = App.Major & "." & App.Minor & "." & App.Revision
Label1.Caption = "현제 버전 : " & O
c = Inet1.OpenURL(Q)
Label2.Caption = "최신 버전 : " & c
If Label2.Caption = "최신 버전 : " Then
Command1.Enabled = False
Else
If O = c Then
PB.Value = 100
Form2.Show
Unload Me
End If

If O <> c Then
Command1.Enabled = True
Form2.Show
Form2.Label1.Caption = " 최신 버전이 발견 됬습니다. 업데이트 버튼을 클릭하여 업데이트를 하십시오."
Timer1.Enabled = False

End If
End If
End If
End If
End Sub
