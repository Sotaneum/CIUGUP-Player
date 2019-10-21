VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  '없음
   Caption         =   "Form2"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   LinkTopic       =   "Form2"
   ScaleHeight     =   225
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "최신 버전이 없습니다. "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1890
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Unload Me

End Sub

Private Sub Timer2_Timer()
Form2.Width = Label1.Width
Form2.Height = Label1.Height

End Sub
