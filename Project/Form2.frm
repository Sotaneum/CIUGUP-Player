VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  '없음
   Caption         =   "Form2"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6240
      Top             =   0
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   $"Form2.frx":0000
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   $"Form2.frx":00F7
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      Caption         =   "For get ! CIUGUP Player 1.0.2v"
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Form2.frx":02D5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
Unload Me

End Sub
