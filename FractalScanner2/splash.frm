VERSION 5.00
Begin VB.Form splash 
   Caption         =   "ShareWare notification"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Payment info"
      Height          =   405
      Left            =   3075
      TabIndex        =   3
      Top             =   1815
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   210
      TabIndex        =   1
      Top             =   2475
      Width           =   4245
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   4200
      Top             =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Unlock code:"
      Height          =   225
      Left            =   165
      TabIndex        =   2
      Top             =   2220
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "This is shareware, please pay 10$ if you use it more than 30 days, or 100$ if you use it for commercial purposes."
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   4320
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Send to:" & vbCrLf & _
"Guy Giesbrecht" & vbCrLf & "6939 Rogue River Hwy" & _
vbCrLf & "Grants Pass, Or" & vbCrLf & "97527"

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
FracSearch.Show
Unload Me

End Sub
