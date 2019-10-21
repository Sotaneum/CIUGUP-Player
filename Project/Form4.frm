VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  '없음
   Caption         =   "Form4"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form4"
   ScaleHeight     =   300
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   2160
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Text            =   "외부 IP 체크 중......"
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Z As String
Z = Inet1.OpenURL("http://devileyes.hosting.paran.com/showip.php")
Text1.Text = "외부 IP : " + " " + Z
End Sub

Private Sub Timer1_Timer()
Unload Me

End Sub

Private Sub Form_Click()
Unload Me

End Sub
