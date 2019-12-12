VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Grader 
   Caption         =   "Manual Grading System"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Pause at generation end"
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   6210
      Width           =   2490
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   270
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   21
      Cols            =   6
      ScrollTrack     =   -1  'True
      TextStyle       =   3
      FormatString    =   ""
   End
   Begin VB.Timer Timer1 
      Left            =   9690
      Top             =   6270
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   6795
      Width           =   10245
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   6195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Grader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnt As Integer
Public cohl As Integer, roe As Integer
Public kolsel As Integer, roesel As Integer

Public Sub putPic(pic As StdPicture, ndex As Integer)
If cohl = 0 Then cohl = 1
If roe = 0 Then roe = 1

cnt = ndex
'If roe = MSFlexGrid1.Rows-1 Then roe = 1
If cohl > 5 Then roe = roe + 1: cohl = 1
If roe > 20 Then roe = 1
MSFlexGrid1.col = cohl: MSFlexGrid1.row = roe
Set MSFlexGrid1.CellPicture = pic
MSFlexGrid1.Text = cnt
cohl = cohl + 1

End Sub

Private Sub Check1_Click()
FracSearch.pauseatgenend = Not FracSearch.pauseatgenend

End Sub

Private Sub Form_Load()
cnt = 1
cohl = 1
roe = 1
MSFlexGrid1.ColWidth(0) = 450
MSFlexGrid1.ColWidth(1) = FracSearch.Picture2.Width * 15
MSFlexGrid1.ColWidth(2) = FracSearch.Picture2.Width * 15
MSFlexGrid1.ColWidth(3) = FracSearch.Picture2.Width * 15
MSFlexGrid1.ColWidth(4) = FracSearch.Picture2.Width * 15
MSFlexGrid1.ColWidth(5) = FracSearch.Picture2.Width * 15
For tzl = 1 To 20
MSFlexGrid1.RowHeight(tzl) = FracSearch.Picture2.Height * 15
Next tzl



End Sub

Private Sub Form_Resize()
If Me.Width - 450 > 1 Then MSFlexGrid1.Width = Me.Width - 450
If Me.Height - 650 > 1 Then MSFlexGrid1.Height = Me.Height - 650
End Sub

Private Sub MSFlexGrid1_Click()
kolsel = MSFlexGrid1.col
roesel = MSFlexGrid1.row
If kolsel > 0 And roesel > 0 Then
Clipboard.SetText FracSearch.getParam(CInt(MSFlexGrid1.Text))
Debug.Print "msf_click"; kolsel, roesel

End If
End Sub
Public Sub flushGrid()
For x = 1 To 5
MSFlexGrid1.col = x
For y = 1 To 20
MSFlexGrid1.row = y
Set MSFlexGrid1.CellPicture = Nothing
Next y
Next x
cohl = 1: roe = 1
End Sub

Private Sub MSFlexGrid1_EnterCell()
Static entering As Boolean
If entering Then Exit Sub
entering = True


entering = False
End Sub
Function faIndex(row As Long, col As Long) As Long
faIndex = row * MSFlexGrid1.Cols + col
End Function
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error GoTo erhandl
If Button = 1 Then
' left
If Shift = 1 Then
' shift was held down
Dim kol As Long, roe As Long, olscorenum$
kol = MSFlexGrid1.col: roe = MSFlexGrid1.row
Debug.Print kol, roe, MSFlexGrid1.TextArray(faIndex(roe, kol))
olscorenum = FracSearch.getScore(CInt(MSFlexGrid1.TextArray(faIndex(roe, kol))))

scor = InputBox("Enter score for" & MSFlexGrid1.TextArray(faIndex(roe, kol)) _
& " p.s.:" & olscorenum)
If scor = "" Then Exit Sub
FracSearch.setScore CInt(MSFlexGrid1.TextMatrix(roe, kol)), CDbl(scor)
Else
'it wasn't
End If

End If
If Button = 2 Then
'right
If Shift = 4 Then
' shift was held down
Else
' it wasn't
End If
End If
Exit Sub
erhandl:
Stop
End Sub
