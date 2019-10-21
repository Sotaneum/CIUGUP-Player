VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "CIUGUP Player"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1875
   ScaleMode       =   0  '사용자
   ScaleWidth      =   6300
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "외부 IP 체크"
      Height          =   495
      Left            =   1680
      Picture         =   "Form1.frx":6852
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CPU 속도 체크"
      Height          =   495
      Left            =   1680
      Picture         =   "Form1.frx":D0A4
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "팩맨 Game"
      DisabledPicture =   "Form1.frx":138F6
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":1A148
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3600
      Top             =   0
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "현제 프로그램은 오프라인 상태입니다."
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   3150
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "시간 로드 중...."
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   1290
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6238.835
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "버전 로드 중..."
      Height          =   180
      Left            =   4995
      TabIndex        =   8
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◀ 창이 자동으로 종료 됩니다."
      Height          =   180
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   2460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◀ 창이 자동으로 종료 됩니다."
      Height          =   180
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   2460
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6361.165
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Label Label2 
      Caption         =   "50"
      Height          =   135
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6361.165
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image12 
      Height          =   240
      Left            =   1320
      Picture         =   "Form1.frx":2099A
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image11 
      Height          =   240
      Left            =   1080
      Picture         =   "Form1.frx":271EC
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image10 
      Height          =   240
      Left            =   3240
      Picture         =   "Form1.frx":2DA3E
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   240
      Left            =   2520
      Picture         =   "Form1.frx":34290
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   240
      Left            =   2280
      Picture         =   "Form1.frx":3AAE2
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   2040
      Picture         =   "Form1.frx":41334
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   3000
      Picture         =   "Form1.frx":47B86
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "상태 : "
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   570
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3201
      _cy             =   450
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1800
      Picture         =   "Form1.frx":4E3D8
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":54C2A
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   720
      Picture         =   "Form1.frx":5B47C
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":61CCE
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   240
      Picture         =   "Form1.frx":68520
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmPacMan.Show

End Sub

Private Sub Command2_Click()
Form5.Show

End Sub

Private Sub Command3_Click()
Form4.Show

End Sub

Private Sub Form_Load()
Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision
Shell "Updayt.exe", vbNormalNoFocus
If App.PrevInstance = True Then
    Call MsgBox("프로그램이 이미 실행중 입니다. 고객 센터 gnyontu39@gmail.com에 문의 하여 주십시오.", vbExclamation)
    End
    End If
    
End Sub

Private Sub Image1_Click()
MsgBox "현제 버전에서는 음원 다운로드 서비스를 지원 받으실 수 없습니다.", vbOKOnly, "음원 다운로드"

End Sub

Private Sub Image10_Click()
MsgBox " 현제 버전에서는 설정이 불가능합니다. ", vbOKOnly + vbInformation, " Tip - 설정 "

End Sub

Private Sub Image11_Click()
Dim V, P
V = wmp.settings.volume
P = Val(Label2.Caption)
If P >= 100 Then
MsgBox " 볼륨이 최대 입니다. ", vbInformation + vbOKOnly, " 에러 "
Else
Label2.Caption = P + 10
wmp.settings.volume = V + 10
Form3.Show
End If

End Sub

Private Sub Image12_Click()
Dim V, P
V = wmp.settings.volume
P = Val(Label2.Caption)
If P <= 0 Then
MsgBox " 볼륨이 최저 입니다. ", vbInformation + vbOKOnly, " 에러 "
Else
Label2.Caption = P - 10
wmp.settings.volume = V - 10
Form3.Show
End If

End Sub

Private Sub Image2_Click()
wmp.Controls.play

End Sub

Private Sub Image3_Click()
wmp.Controls.pause

End Sub

Private Sub Image4_Click()
wmp.Controls.stop

End Sub

Private Sub Image5_Click()
Dim A, U
                                      CD.DialogTitle = "Open File"
                                      CD.Filter = "MP3File(*.mp3)|*.mp3"
                                      CD.FileName = ""
                                      CD.Flags = cdlOFNHideReadOnly
                                      CD.InitDir = ""
                                      CD.ShowOpen
                                
                                      If CD.FileName = "" Then
                                         Exit Sub
                                      End If
                                
                                      A = CD.FileName
                                
                                      If A = "" Then
                                         
                                         MsgBox "MP3파일을 선택해 주십시오", vbOKOnly
                                         Exit Sub
                                         
                                      End If
                                      wmp.URL = A
                                      
End Sub

Private Sub Image6_Click()
Shell "Updayt.exe", vbNormalNoFocus

End Sub

Private Sub Image7_Click()
Form2.Show

End Sub

Private Sub Image8_Click()
MsgBox " 제작자 : gnyontu39@gmail.com / 이동건 (cyydo96@hotmail.com) ", vbInformation + vbOKOnly, " 제작자 "

End Sub

Private Sub Image9_Click()
MsgBox " 서버와 연결 할 수 없습니다. ", vbExclamation + vbOKOnly, " 서버에러 "
MsgBox " 사용자의 id로는 접속이 불가능합니다. ", vbInformation + vbOKOnly, " 설정 "

End Sub

Private Sub Timer1_Timer()
Label1.Caption = "상태 : " & wmp.Status
Label1.Left = 0
Line1.X2 = Form1.Width
Line2.X2 = Form1.Width
Line3.X2 = Form1.Width
Label7.Left = Label6.Width
Me.Label6.Caption = "현제시간 : " & Format(Time, "hh:mm:ss AMPM")
If Label1.Width >= 6300 Then
Form1.Width = Val(Label1.Width) + 500

End If

End Sub
