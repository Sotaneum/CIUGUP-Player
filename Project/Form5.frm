VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  '없음
   Caption         =   "Form5"
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   2400
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Text            =   "체크 중입니다..."
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Text            =   "체크 중입니다..."
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Declare Function rdtsc Lib "csRdtsc.dll" () As Currency
    Private Declare Function rdtsc2 Lib "csRdtsc.dll" () As Currency
    Private Declare Function rdtsc3 Lib "csRdtsc.dll" () As Currency

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()

    Dim startValue As Currency
    Dim endValue As Currency

    Dim resultValue As Double
    
    startValue = rdtsc

    Sleep (1000)

    endValue = rdtsc

    resultValue = CurrToText(endValue - startValue)
    Text2.Text = "CPU 속도 : " & (resultValue / 1000000) & "Mhz"
    Text1.Text = "타임 스텀프 : " & CurrToText(rdtsc)
    
    startValue = 0
    endValue = 0
    resultValue = 0

End Sub


Private Sub Form_Click()
Unload Me

End Sub
Private Function CurrToText(ByVal Value As Currency) As String

     Dim Temp As String, L As Long
     
     Temp = Format$(Value, "#.0000")
     
     L = Len(Temp)
     
     Temp = Left$(Temp, L - 5) & Right$(Temp, 4)
     
     Do While Len(Temp) > 1 And Left$(Temp, 1) = "0"
     
       Temp = Mid$(Temp, 2)
       
     Loop
     
     Do While Len(Temp) > 2 And Left$(Temp, 2) = "-0"
     
       Temp = "-" & Mid$(Temp, 3)
       
     Loop
     
     CurrToText = Temp
     
End Function







Private Sub Timer1_Timer()
Unload Me

End Sub


