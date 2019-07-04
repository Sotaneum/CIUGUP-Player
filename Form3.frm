VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  '없음
   Caption         =   "50"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   LinkTopic       =   "Form3"
   ScaleHeight     =   225
   ScaleWidth      =   930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "볼륨 : "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "50"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Unload Me

End Sub

Private Sub Timer2_Timer()
Label1.Caption = Form1.Label2
End Sub
