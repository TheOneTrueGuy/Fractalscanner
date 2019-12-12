VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm manual 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   4890
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "Grade Manual"
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Return to autoscore in:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1305
         TabIndex        =   2
         Top             =   0
         Width           =   2880
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5700
      Top             =   255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7260
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnt As Integer
Dim limit As Integer
Dim rank() As Integer
Dim pix() As pixter

Private Sub MDIForm_Activate()
Timer1.Enabled = True
End Sub

Private Sub MDIForm_Load()
limit = 20
End Sub
Public Sub putPic(pic As StdPicture)
ImageList1.ListImages.Add cnt, , pic
cnt = cnt + 1
End Sub
Public Sub galleria()
ReDim pixter(cnt + 1)
For tzl = 0 To cnt
pixter(tzl).Picture = ImageList1.ListImages.Item(tzl)
Next tzl

End Sub

Private Sub Timer1_Timer()
Static tymkount As Integer
tymkount = tymkount + 1
Label1.Caption = "Return to autoscore in:" & 20 - tymkount

If tymkount > limit Then tymkount = 0: Timer1.Enabled = False: Exit Sub

End Sub

