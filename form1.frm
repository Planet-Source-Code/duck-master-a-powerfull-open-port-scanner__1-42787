VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Fast Open Port Scan"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   Picture         =   "form1.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear list"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   3000
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim port As Integer

Private Sub Command1_Click()
port = 0
Text1_DblClick
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.Clear
port = 0
End Sub

Private Sub Form_Load()
port = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form1.ScaleHeight = 4125
Form1.ScaleLeft = 4125
Form1.ScaleLeft = 0
Form1.ScaleTop = 0
Form1.Width = 4755
Form1.Height = 4635
End Sub

Private Sub Text1_DblClick()
On Error Resume Next
Form1.Show
For i = 1 To 20000
If port = f Then port = ""
Winsock1.Close
Winsock1.Connect Text1.Text, port
f = List1.List(0)
port = port + 1
DoEvents
Next i
End Sub



Private Sub Timer2_Timer()
On Error Resume Next
Form1.Show
For i = 1 To 20000
Winsock1.Close
Winsock1.Connect Text1.Text, port
port = port + 1
DoEvents
Next i
End Sub

Private Sub Winsock1_Connect()
List1.AddItem ("Open Port : " & Winsock1.RemotePort)
End Sub




