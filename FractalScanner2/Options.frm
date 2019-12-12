VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Options 
   Caption         =   "Options"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check8 
      Caption         =   "Evolve Skew"
      Height          =   315
      Left            =   2340
      TabIndex        =   35
      Top             =   2970
      Width           =   2070
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Evolve Rotation"
      Height          =   285
      Left            =   135
      TabIndex        =   34
      Top             =   2940
      Width           =   1935
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Evolve Xmag"
      Height          =   255
      Left            =   4830
      TabIndex        =   33
      Top             =   2565
      Width           =   1470
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Evolve Inversions"
      Height          =   285
      Left            =   2340
      TabIndex        =   32
      Top             =   2542
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Options.frx":0000
      Left            =   1335
      List            =   "Options.frx":0019
      TabIndex        =   30
      Text            =   "mod"
      Top             =   1755
      Width           =   2745
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Options.frx":0043
      Left            =   1335
      List            =   "Options.frx":0059
      TabIndex        =   29
      Text            =   "iter"
      Top             =   1365
      Width           =   2745
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Options.frx":0081
      Left            =   1335
      List            =   "Options.frx":009A
      TabIndex        =   28
      Text            =   "Maxiter"
      Top             =   960
      Width           =   2745
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Options.frx":00CE
      Left            =   1335
      List            =   "Options.frx":00E1
      TabIndex        =   27
      Text            =   "Mandelbrot"
      Top             =   570
      Width           =   2745
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   90
      ScaleHeight     =   3015
      ScaleWidth      =   5385
      TabIndex        =   26
      Top             =   3810
      Width           =   5415
      Begin VB.CheckBox Check10 
         Caption         =   "Cellular comparison"
         Height          =   300
         Index           =   8
         Left            =   30
         TabIndex        =   55
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox Check10 
         Caption         =   "2 neighbor comparison"
         Height          =   300
         Index           =   7
         Left            =   30
         TabIndex        =   54
         Top             =   2325
         Width           =   2040
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Testing Mode"
         Height          =   300
         Index           =   6
         Left            =   45
         TabIndex        =   53
         Top             =   2025
         Width           =   1620
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   3240
         TabIndex        =   50
         Text            =   "900"
         Top             =   30
         Width           =   690
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3450
         TabIndex        =   49
         Text            =   "10"
         Top             =   375
         Width           =   360
      End
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   4050
         TabIndex        =   48
         Text            =   "10"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3060
         TabIndex        =   47
         Text            =   "31500"
         Top             =   720
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Change"
         Height          =   255
         Left            =   2820
         TabIndex        =   46
         Top             =   1110
         Width           =   765
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Inside "
         Height          =   255
         Left            =   3630
         TabIndex        =   45
         Top             =   1110
         Width           =   720
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2925
         TabIndex        =   44
         Text            =   "32"
         Top             =   1440
         Width           =   420
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3525
         TabIndex        =   43
         Text            =   "20"
         Top             =   1440
         Width           =   450
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Random Scan ----- Number of test points:"
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   42
         Top             =   30
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Grid Scan ---- horizontal and vertical  steps:"
         Height          =   300
         Index           =   1
         Left            =   30
         TabIndex        =   41
         Top             =   360
         Width           =   3435
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Boundary Scan "
         Height          =   300
         Index           =   2
         Left            =   45
         TabIndex        =   40
         Top             =   705
         Width           =   1590
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Mask Scan   ---- Mask Color::"
         Height          =   300
         Index           =   3
         Left            =   45
         TabIndex        =   39
         Top             =   1032
         Width           =   2415
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Random Square  --- Width && Height"
         Height          =   300
         Index           =   4
         Left            =   45
         TabIndex        =   38
         Top             =   1395
         Width           =   2790
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Direct Equality"
         Height          =   300
         Index           =   5
         Left            =   30
         TabIndex        =   37
         Top             =   1710
         Width           =   1515
      End
      Begin VB.Label Label15 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3855
         TabIndex        =   52
         Top             =   390
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   2460
         Top             =   1110
         Width           =   270
      End
      Begin VB.Label Label15 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   51
         Top             =   1440
         Width           =   195
      End
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1785
      TabIndex        =   25
      Top             =   150
      Width           =   2910
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7335
      TabIndex        =   22
      Text            =   "5.9"
      Top             =   2955
      Width           =   690
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select Pallette"
      Height          =   285
      Left            =   6930
      TabIndex        =   20
      Top             =   5325
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8205
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Region Restriction"
      Height          =   1590
      Left            =   4335
      TabIndex        =   11
      Top             =   495
      Width           =   4110
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   990
         TabIndex        =   19
         Text            =   "2"
         Top             =   1215
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   990
         TabIndex        =   17
         Text            =   "-2"
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   990
         TabIndex        =   15
         Text            =   "2"
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   990
         TabIndex        =   13
         Text            =   "-2"
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Y maximum:"
         Height          =   225
         Left            =   90
         TabIndex        =   18
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Y minimum:"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   930
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "X maximum:"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   615
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "X minimum:"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   285
         Width           =   825
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Evolve Bailout value"
      Height          =   225
      Left            =   135
      TabIndex        =   9
      Top             =   2565
      Width           =   1920
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Evolve Decomposition"
      Height          =   225
      Left            =   4830
      TabIndex        =   8
      Top             =   2160
      Width           =   1980
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Evolve Maximum iterations"
      Height          =   210
      Left            =   2340
      TabIndex        =   7
      Top             =   2190
      Width           =   2340
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Evolve BioMorph color"
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   2175
      Width           =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   480
      Left            =   7965
      TabIndex        =   4
      Top             =   6345
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   6960
      TabIndex        =   3
      Top             =   6345
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   480
      Left            =   5865
      TabIndex        =   2
      Top             =   6345
      Width           =   1020
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Testing Devices:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   36
      Top             =   3510
      Width           =   1275
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "default.map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6870
      TabIndex        =   31
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Label Label14 
      Caption         =   "Alternate name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   24
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label13 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8055
      TabIndex        =   23
      Top             =   2940
      Width           =   240
   End
   Begin VB.Label Label12 
      Caption         =   "Mutation Rate"
      Height          =   240
      Left            =   6240
      TabIndex        =   21
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "Bailout test"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   1785
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Fractal Type"
      Height          =   210
      Left            =   105
      TabIndex        =   5
      Top             =   630
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Outside Coloring"
      Height          =   240
      Left            =   105
      TabIndex        =   1
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Inside Coloring"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   1035
      Width           =   1155
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cb1, cb2, cb3, cb4
Dim scanchanged As Boolean


Private Sub Check10_Click(index As Integer)
scanchanged = True
'FracSearch.setPlex index

End Sub


Private Sub Command1_Click()
' probs with changing scoring -- curbest needs to be flushed when score
' settings are changed.
Dim tzl As Integer
For tzl = 0 To 8
If Check10(tzl).Value = 1 Then
FracSearch.plexiter = tzl
Grader.roe = 1
FracSearch.setPlex tzl
Select Case tzl
Case 0
If Text9 <> "" Then FracSearch.randstep = CInt(Text9.Text)
FracSearch.scantype = 1
Case 1
If Text10 <> "" Then FracSearch.xstep = CInt(Text10.Text)
If Text11 <> "" Then FracSearch.ystep = CInt(Text11.Text)
FracSearch.scantype = 2
Case 2
FracSearch.scantype = 3
FracSearch.mindifcount = CLng(Text5)
Case 3
'mask scan
FracSearch.maskcolor = Shape1.FillColor
FracSearch.scantype = 4
Case 4
' random square
FracSearch.scantype = 5
If Text10 <> "" Then FracSearch.rswid = CInt(Text12.Text)
If Text11 <> "" Then FracSearch.rshyt = CInt(Text13.Text)
Case 5
FracSearch.scantype = 6
Case 6
FracSearch.scantype = 7
Case 7
' 2 nabor compare
FracSearch.scantype = 8
Case 8
FracSearch.scantype = 9
' random cellular (center pixel, 8 nabors) compare
End Select
Else
FracSearch.resetPlex tzl
End If
Next tzl

With FracSearch
.fractype = Combo1.Text
.icolor = Combo2.Text
.ocolor = Combo3.Text
.bailouttest = Combo4.Text
.curnum = 0
.eBail = Check4.Value
everyone.evbail = Check4.Value
.eInvert = Check5.Value
everyone.evinvert = Check5.Value
.eXmag = Check6.Value
everyone.evxmag = Check6.Value
.eBio = Check1.Value
everyone.evbio = Check1.Value
.eMaxit = Check2.Value
everyone.evmaxiter = .eMaxit
.eDecomp = Check3.Value
everyone.evdecomp = .eDecomp
.eRot = Check7.Value
everyone.evrot = .eRot
.eSkew = Check8.Value
everyone.evskew = .eSkew
End With

'If Combo1.ListIndex <> -1 Then FracSearch.fractype = Combo1.List(Combo1.ListIndex)
FracSearch.mutrate = CSng(Text7)


If scanchanged Then
Grader.flushGrid
With FracSearch
.erasescores
.bestscore = 88821610: .curbest = 88821610
'.bestscore = 0: .curbest = 0
.fractype = Combo1.Text
.icolor = Combo2.Text
.ocolor = Combo3.Text
.bailouttest = Combo4.Text
End With
End If

FracSearch.curnum = 0
FracSearch.palmap = Label19.Caption
FracSearch.makeParams
Me.Hide

End Sub

Private Sub Command2_Click()
Me.Hide

End Sub

Private Sub Command4_Click()
On Error GoTo errhandlr
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Pallette map (*.map)|*.map"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
Dim fyl As String
fyl = CommonDialog1.FileTitle
If fyl = "" Then Exit Sub
FracSearch.palmap = fyl
Label19.Caption = fyl
Exit Sub
errhandlr:
End Sub
Private Sub Combo1_Change()
'Combo1.Text = Combo1.List(cb1)
On Error Resume Next
'Command1.SetFocus
If Err.Number Then Err.Clear
'Debug.Print Combo1.Text
End Sub

Private Sub Combo1_Click()
If cb1 = Combo1.ListIndex Then Exit Sub
cb1 = Combo1.ListIndex
'Command1.SetFocus
'Debug.Print Combo1.Text
End Sub

Private Sub Combo2_Change()
'Combo2.Text = Combo2.List(cb1)
On Error Resume Next
'Command1.SetFocus
If Err.Number Then Err.Clear
'Debug.Print Combo2.Text
'If CInt(Combo2.Text) > 255 Or CInt(Combo2.Text) < 0 Then Combo2.Text = "0"

End Sub

Private Sub Combo2_Click()
If cb2 = Combo2.ListIndex Then Exit Sub
cb2 = Combo2.ListIndex
'Command1.SetFocus
'Debug.Print Combo2.Text
End Sub
Private Sub Combo3_Change()
'Combo3.Text = Combo3.List(cb1)
On Error Resume Next
'Command1.SetFocus
If Err.Number Then Err.Clear
'Debug.Print Combo3.Text
If CInt(Combo3.Text) > 255 Or CInt(Combo3.Text) < 0 Then Combo3.Text = "0"
End Sub

Private Sub Combo3_Click()
If cb3 = Combo3.ListIndex Then Exit Sub
cb3 = Combo3.ListIndex
Command1.SetFocus
'Debug.Print Combo3.Text
End Sub

Private Sub Combo4_Change()
'Combo4.Text = Combo4.List(cb1)
On Error Resume Next
Command1.SetFocus
If Err.Number Then Err.Clear
'Debug.Print Combo4.Text
End Sub

Private Sub Combo4_Click()
If cb4 = Combo4.ListIndex Then Exit Sub
cb4 = Combo4.ListIndex
Command1.SetFocus
'Debug.Print Combo4.Text
End Sub

Private Sub Command5_Click()
On Error GoTo erhandl
Dim colr As Long
CommonDialog1.ShowColor
colr = CommonDialog1.Color
Shape1.FillColor = colr
FracSearch.maskcolor = colr
erhandl:

End Sub

Private Sub Form_Activate()
Command1.SetFocus
scanchanged = False
If everyone.evbio Then Check1.Value = 1 Else Check1.Value = 0
If everyone.evmaxiter Then Check2.Value = 1 Else Check2.Value = 0
If everyone.evdecomp Then Check3.Value = 1 Else Check3.Value = 0
If everyone.evbail Then Check4.Value = 1 Else Check4.Value = 0
If everyone.evinvert Then Check5.Value = 1 Else Check5.Value = 0
If everyone.evxmag Then Check6.Value = 1 Else Check6.Value = 0
If everyone.evrot Then Check7.Value = 1 Else Check7.Value = 0
If everyone.evskew Then Check8.Value = 1 Else Check8.Value = 0

With FracSearch

Combo1.Text = .fractype
Combo2.Text = .icolor
Combo3.Text = .ocolor
Combo4.Text = .bailouttest

Text7 = .mutrate
Label19.Caption = .palmap
End With

End Sub

Private Sub Text10_Change()
scanchanged = True

End Sub

Private Sub Text11_Change()
scanchanged = True

End Sub

Private Sub Text12_Change()
scanchanged = True
End Sub

Private Sub Text13_Change()
scanchanged = True
End Sub

Private Sub Text9_Change()
scanchanged = True

End Sub
