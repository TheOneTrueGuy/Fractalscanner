VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FracSearch 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Ducer Fractint scanner/evolver"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   570
      Left            =   1185
      TabIndex        =   26
      Top             =   1635
      Width           =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Manip"
      Height          =   525
      Left            =   150
      TabIndex        =   25
      Top             =   1620
      Width           =   870
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Spratt/Pickover"
      Height          =   255
      Left            =   75
      TabIndex        =   24
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Best"
      Height          =   315
      Left            =   90
      TabIndex        =   15
      Top             =   810
      Width           =   1770
   End
   Begin PicClip.PictureClip PiClip1 
      Left            =   1635
      Top             =   5130
      _ExtentX        =   1191
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      Height          =   3060
      Left            =   7170
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   14
      Top             =   15
      Width           =   4860
      Begin VB.CommandButton Command5 
         Caption         =   "S"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   2730
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   7170
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   13
      Top             =   3075
      Width           =   4860
      Begin VB.CommandButton Command6 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4500
         TabIndex        =   22
         Top             =   330
         Width           =   300
      End
      Begin VB.CommandButton Command6 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4500
         TabIndex        =   21
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton Command5 
         Caption         =   "S"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   2295
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   12
      Top             =   3075
      Width           =   4860
      Begin VB.CommandButton Command5 
         Caption         =   "S"
         Height          =   240
         Index           =   1
         Left            =   4530
         TabIndex        =   19
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   6765
      Width           =   12015
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   16000
      Left            =   11085
      Top             =   6195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   390
      Width           =   2010
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   6150
      Width           =   12015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10605
      Top             =   6300
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11595
      Top             =   6165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   2295
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   15
      Width           =   4860
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   480
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1395
      Top             =   5085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label lblScan 
      Height          =   225
      Left            =   30
      TabIndex        =   23
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1935
      Top             =   840
      Width           =   285
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   795
      TabIndex        =   17
      Top             =   4365
      Width           =   1440
   End
   Begin VB.Label Label7 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   4365
      Width           =   750
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Best score this generation params:"
      Height          =   210
      Left            =   0
      TabIndex        =   11
      Top             =   6495
      Width           =   2715
   End
   Begin VB.Label lblGenkount 
      BackColor       =   &H0080FF80&
      Caption         =   "Gen: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   2445
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Best Current:"
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   3735
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Best Image Now->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   3300
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Best Score"
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   5370
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Best Score Parameters:"
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   5775
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2910
      Width           =   2295
   End
   Begin VB.Menu fyl 
      Caption         =   "File"
      Begin VB.Menu openImg 
         Caption         =   "Open Image "
      End
      Begin VB.Menu multimage 
         Caption         =   "Multiple images"
      End
      Begin VB.Menu nupop 
         Caption         =   "New population"
      End
      Begin VB.Menu savepop 
         Caption         =   "Save Population"
      End
      Begin VB.Menu interbreed 
         Caption         =   "InterBreed Population"
      End
      Begin VB.Menu loadpop 
         Caption         =   "Load Population"
      End
      Begin VB.Menu mutatepop 
         Caption         =   "Mutate entire pop"
      End
      Begin VB.Menu export 
         Caption         =   "Export par file"
      End
      Begin VB.Menu xit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu popedit 
         Caption         =   "Population Editor"
      End
      Begin VB.Menu mangrad 
         Caption         =   "Manual Grader"
      End
   End
   Begin VB.Menu opt 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "FracSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tik As Long
Dim jobnaym$
Dim editor As EditDB
Public initPath As String
Dim stillrunning As Boolean
Public running As Boolean
Dim population(101) As popmem
' prepare flexible pop size: make p()'s dynamic
Dim popsize As Integer
Dim param$(101), score(101) As Double
Public curnum As Integer
Public bestscore As Long
Public curbest As Long
Dim best As popmem, bestnum As Integer, bestparam As String, curbestparam$
Dim prevscore As Long
Dim kount As Integer
Dim pause As Boolean
Dim goalfylname As String
Dim loadedpop As Boolean
Dim xcorn As Integer, ycorn As Integer

Dim pixelarray(320, 200) As Long, colorarray(256, 1) As Long, colorcount As Integer
Dim colortable, mostcolor As Long
Public maskcolor As Long
Dim comparing As Boolean, numgens As Long
'Dim srcPic As New StdPicture

Public palmap As String, fractype As String, xstep As Integer, ystep As Integer
Public rswid As Integer, rshyt As Integer
Dim plex(10) As Boolean
Public plexiter As Integer
Public pauseatgenend As Boolean

Public randstep As Integer, scantype As Integer, mindifcount As Long
Public icolor$, ocolor$, bailouttest$, mutrate As Single
Public eBail As Boolean, eInvert As Boolean, eXmag As Boolean
Public eBio As Boolean, eMaxit As Boolean, eDecomp As Boolean
Public eRot As Boolean, eSkew As Boolean
Public totalmaxpoints As Long
Dim quiting As Boolean
Dim avpixel As Long
Dim multipixur() As StdPicture
Dim pixurcount As Integer, multiplmage As Boolean

Public Sub runFract(parameter As String)

'Shell "C:\Program Files\Microsoft Visual Studio\VB98\FractalScanner\fractint.exe batch=yes type=mandel video=F3 savename=scan params=-0.3925313/1.868591 center-mag=-1.71268958/-0.260074/8.799493", vbMaximizedFocus
'Stop
lblStatus.Caption = "Engaging"
If Dir(initPath & "\scan.gif", vbNormal) <> "" Then Kill initPath & "\scan.gif"
Open "scan.txt" For Output As #1
Print #1, parameter
Close 1
Dim par As String
par = initPath & "\fractint.exe @scan.txt"
Debug.Print curnum; ":"; parameter
Dim shell_res As Long
shell_res = Shell(par, vbHide)
stillrunning = (shell_res <> 0)
lblStatus.Caption = "Waiting"
End Sub

Public Sub makePop()
Dim t As Integer
For t = 1 To 100
population(t).centerx = (Rnd * 3.5) - 2
population(t).centery = (Rnd * 3.5) - 2
population(t).mag = (Rnd * 16)
population(t).param1 = (Rnd * 4) - 2
population(t).param2 = (Rnd * 4) - 2
population(t).invert = (Rnd * 4) - 2
population(t).icenterx = (Rnd * 4) - 2
population(t).icentery = (Rnd * 4) - 2
population(t).bailout = (Rnd * 64000)
population(t).biomorph = (Rnd * 256)
population(t).decomp = Rnd * 16000
population(t).maxiter = Rnd * 5120 + 1
population(t).rot = Rnd * 360
population(t).xmag = (Rnd * 4) - 2
population(t).skew = Rnd * 360
everyone.pop(t) = population(t)
Next t
loadedpop = True

End Sub
'Public Sub makePop2()
'Dim t As Integer
'For t = 1 To 100
'pop2(t).cornerx = Rnd
'pop2(t).cornery = Rnd
'pop2(t).increm = Rnd
'pop2(t).iters = Int(Rnd * 256)
'pop2(t).param1 = Rnd
'pop2(t).param2 = Rnd
'Next t
'End Sub
Public Sub makeParams()
Dim t As Integer, param1 As Double, param2 As Double
Dim centerx As Double, centery As Double, mag As Double
Dim icentx As Double, icenty As Double, inver As Double
Dim rota As Double, decom As Integer, maxit As Long, xmg As Double
Dim biom As Integer, bail As Long
Dim skw As Double
Dim map$, rmp As Integer
rmp = Int(4)
'Dim decomp As Integer
'decomp = 128 + (2 ^ Int(Rnd * 10))
For t = 1 To 100
param1 = population(t).param1
param2 = population(t).param2
centerx = population(t).centerx
centery = population(t).centery
mag = population(t).mag
inver = population(t).invert
icentx = population(t).icenterx: icenty = population(t).icentery
rota = population(t).rot: xmg = population(t).xmag
decom = population(t).decomp: bail = population(t).bailout
biom = population(t).biomorph
maxit = population(t).maxiter
skw = population(t).skew
param$(t) = ""
Select Case LCase(fractype)
Case "mandelbrot"
param$(t) = " batch=yes type=mandel video=F3 savename=scan sound=no map=" & palmap
'param$(t) = param$(t) & " params=" & CStr(param1) & "/" & CStr(param2) '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
'param$(t) = param$(t) & " center-mag=" & CStr(centerx) & "/" & CStr(centery) & "/" & CStr(mag)
Case "julia"
param$(t) = "batch=yes type=julia video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)
'param$(t) = param$(t) & " params=" & Format(param1, "###0.000000") & "/" & Format(param2, "###0.000000") '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
'param$(t) = param$(t) & " center-mag=" & Format(centerx, "###0.000000") & "/" & Format(centery, "###0.000000") & "/" & Format(mag, "###0.000000")

Case "newton"
param$(t) = "batch=yes type=newton video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)
'param$(t) = param$(t) & " params=" & Format(param1, "###0.000000") & "/" & Format(param2, "###0.000000") '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
'param$(t) = param$(t) & " center-mag=" & Format(centerx, "###0.000000") & "/" & Format(centery, "###0.000000") & "/" & Format(mag, "###0.000000")

'Case "frothy basin"
'param$(t) = "batch=yes type=frothybasin video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)
'param$(t) = param$(t) & " params=" & Format(param1, "###0.000000") & "/" & Format(param2, "###0.000000") '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
''param$(t) = param$(t) & " inside=atan outside=atan"
'param$(t) = param$(t) & " center-mag=" & Format(centerx, "###0.000000") & "/" & Format(centery, "###0.000000") & "/" & Format(mag, "###0.000000")
''corners=-0.739/-0.123/0.288/0.291"
Case "escher julia"
param$(t) = "batch=yes type=escher_julia video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)
'param$(t) = param$(t) & " params=" & Format(param1, "###0.000000") & "/" & Format(param2, "###0.000000") '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
'param$(t) = param$(t) & " center-mag=" & Format(centerx, "###0.000000") & "/" & Format(centery, "###0.000000") & "/" & Format(mag, "###0.000000")

Case "spider"
param$(t) = "batch=yes type=spider video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)
'param$(t) = param$(t) & " params=" & Format(param1, "###0.000000") & "/" & Format(param2, "###0.000000") '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
'param$(t) = param$(t) & " center-mag=" & Format(centerx, "###0.000000") & "/" & Format(centery, "###0.000000") & "/" & Format(mag, "###0.000000")
Case Else
param$(t) = "batch=yes type=" & fractype & " video=F3 savename=scan sound=no map=" & palmap 'decomp=" & CStr(decomp)

End Select

param$(t) = param$(t) & " params=" & param1 & "/" & param2 '-0.780/0.326/200 center-mag=0.1/0.1/0.5"
param$(t) = param$(t) & " center-mag=" & centerx & "/" & centery & "/" & mag
If eXmag Then
param$(t) = param$(t) & "/" & xmg '& "/" & rota
Else
param$(t) = param$(t) & "/1"
End If
If eRot Then
param$(t) = param$(t) & "/" & rota
Else
param$(t) = param$(t) & "/0"
End If
If eSkew Then param$(t) = param$(t) & "/" & skw
If eMaxit Then param$(t) = param$(t) & " maxiter=" & maxit Else param$(t) = param$(t) & " maxiter=256"
If eBail Then param$(t) = param$(t) & " bailout=" & bail

param$(t) = param$(t) & " bailoutest=" & bailouttest
param$(t) = param$(t) & " inside=" & icolor
param$(t) = param$(t) & " outside=" & ocolor
If eInvert Then param$(t) = param$(t) & " invert=" & inver & "/" & icentx & "/" & icenty
If eDecomp Then param$(t) = param$(t) & " decomp=" & decom
If eBio Then param$(t) = param$(t) & " biomorph=" & biom
'potential=255/1000/1 periodicity=0
'float=y



Next t
End Sub

Private Sub Command1_Click()
'runFract " batch=yes type=mandel video=F3 sound=no savename=scan"
'Exit Sub

If Not loadedpop Then makePop
Command1.Enabled = False
'Do While Not plex(plexiter)
'plexiter = plexiter + 1
'Loop
'plexiter = plexiter Mod 6
makeParams
curnum = 1
scantype = plexiter + 1
If scantype = 5 Then
Shape1.Visible = True
Shape1.Width = rswid: Shape1.Height = rshyt
Shape1.Top = ycorn: Shape1.Left = xcorn
Shape1.Refresh
Picture1.Refresh: Me.Refresh
End If

Timer1.Enabled = True
ChDir initPath
runFract param$(curnum)
running = True
'makePop2
End Sub

Public Sub compare2(num As Integer)
On Error GoTo errhandl
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long

Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
'For t = 1 To 4000
'X = Rnd * 320
'Y = Rnd * 200
For x = 1 To 30
For y = 1 To 30
colr1 = pixelarray(x + xcorn, y + ycorn) 'Picture1.Point(x + xcorn, y + ycorn)
colr2 = Picture2.Point(x + xcorn, y + ycorn)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
Next y
Next x
'Next t
score(num) = score(num) + (totdif + ((compixel2 - compixel1) * 100))

Debug.Print totdif
Exit Sub
errhandl:

End Sub
Public Sub compare2Xor(num As Integer)
On Error GoTo errhandl
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long


Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
'For t = 1 To 4000
'X = Rnd * 320
'Y = Rnd * 200
For x = 1 To 30
For y = 1 To 30
colr1 = pixelarray(x + xcorn, y + ycorn) 'Picture1.Point(x + xcorn, y + ycorn)
colr2 = Picture2.Point(x + xcorn, y + ycorn)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
dif = (rd1 Xor rd2) Xor (gren1 Xor gren2) Xor (blu1 Xor blu2)
'dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
'dif = dblu + dgren + drd
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
Next y
Next x
'Next t
score(num) = score(num) + totdif '+ ((compixel2 - compixel1) * 100)

Debug.Print "totdif:"; totdif
Exit Sub
errhandl:

End Sub

Public Sub compare3(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long


Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
'For t = 1 To 4000
'X = Rnd * 320
'Y = Rnd * 200
For x = 1 To 30
For y = 1 To 30
colr1 = pixelarray(x + xcorn, y + ycorn)
colr2 = Picture2.Point(x + xcorn, y + ycorn)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
Next y
Next x
'Next t
score(num) = score(num) + (totdif + Abs((compixel2 - compixel1) * 500))

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop
End Sub
Public Sub compare3b(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long


Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
For t = 1 To 900
x = Rnd * 320
y = Rnd * 200
'For X = 1 To 30
'For Y = 1 To 30
colr1 = pixelarray(x, y)
colr2 = Picture2.Point(x, y)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
'Next Y
'Next X
Next t
score(num) = score(num) + (totdif + Abs((compixel2 - compixel1) * 500))

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop
End Sub
Public Sub gridscan(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
DoEvents
On Error GoTo errhandl
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long

Picture2.Refresh
Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Long, totmax As Long, difcount As Long
Dim colr1f As Long, colr2f As Long
'totmax = (scanwid / xstep) * (scanhyt / ystep) * (255+255+255)
totdif = 0
For y = 0 To 199
For x = 0 To 319
colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr2f = -1 Then colr2f = Picture2.Point(0, y)
If colr1f <> colr2f Then difcount = difcount + 1
If difcount > 2 Then Exit For
Next x
If difcount > 2 Then Exit For
Next y
Debug.Print "dif"; difcount
If difcount = 0 Then score(num) = 88821605: GoTo doneit


For x = numgens Mod 3 To 320 Step xstep
DoEvents
For y = numgens Mod 3 To 200 Step ystep

colr1 = pixelarray(x, y)
colr2 = Picture2.Point(x, y)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
'lblStatus.BackColor = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = colr2 Then dif = -100 ': GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
Next y
'Debug.Print x;
Next x

score(num) = score(num) + (totdif + Abs((compixel2 - compixel1) * 500))
doneit:

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop
End Sub

Public Sub gridscan2(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True

Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double, totmax As Long
For x = numgens Mod 2 To 320 Step xstep
For y = numgens Mod 2 To 200 Step ystep
    colr1 = pixelarray(x, y)
    colr2 = Picture2.Point(x, y)
    n1 = colr1: n2 = colr2
'    n1 = colr1 And colr2
    rd1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    blu1 = n1
    rd2 = n2 Mod 256: n2 = Int(n2 / 256)
    gren2 = n2 Mod 256: n2 = Int(n2 / 256)
    blu2 = n2
'dif = (blu1 Xor blu2) + (gren1 Xor gren2) + (rd1 Xor rd2)

    dif = Abs((blu1 Or blu2) - blu1) + Abs((gren1 Or gren2) - gren1) + Abs((rd1 Or rd2) - rd1)
    totdif = totdif + dif
Next y
'Debug.Print x;

Next x
score(num) = score(num) + totdif

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'stop
End Sub
Public Sub gridscan3(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True
'Dim n1 As Long, n2 As Long
'Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
'Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
'Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim colr1f As Long, colr2f As Long, difcount As Long
Dim dif As Double, totdif As Double, totmax As Long
Dim pixrank As Long
'totdif = 88821610
totdif = 0
For y = 0 To 199
For x = 0 To 319
colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr2f = -1 Then colr2f = Picture2.Point(0, y)
If colr1f <> colr2f Then difcount = difcount + 1
If difcount > 20 Then Exit For
Next x
If difcount > 20 Then Exit For
Next y
Debug.Print "dif"; difcount
If difcount = 0 Then score(num) = 88821605: GoTo doneit

For x = 1 To 320
DoEvents
For y = 1 To 200
    colr1 = pixelarray(x, y)
    colr2 = Picture2.Point(x, y)
  '  If colr2 = 0 Then Stop
 ' pixrank = colortable.Item(colr2)
 ' Debug.Print pixrank
 'If pixrank < 0 Then Stop
  'If pixrank <= 0 Then pixrank = 256
    If colr1 <> colr2 Then totdif = totdif + 1 'pixrank
Next y
'Debug.Print x;

Next x
If totdif = 0 Then Stop
score(num) = score(num) + totdif
doneit:


comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'stop
End Sub
Public Sub maskscan(num As Integer)
' open image as mask -- use mask color to determine which pixels to scan
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
Debug.Print "rsquare"
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long

Picture2.Refresh
Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Long, totmax As Long, difcount As Long
Dim colr1f As Long, colr2f As Long
'totmax = (scanwid / xstep) * (scanhyt / ystep) * (255+255+255)
totdif = 0
For y = 0 To 199
For x = 0 To 319

colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr1f = colr2f Then difcount = difcount + 1
Next x
Next y
Debug.Print "dif"; difcount
If difcount = 63800 Then score(num) = 88821605: GoTo doneit

Dim xpos As Integer, ypos As Integer

For x = 0 To 319
DoEvents
For y = 0 To 200


colr1 = pixelarray(x, y)
If colr1 = maskcolor Then GoTo skipit
colr2 = Picture2.Point(x, y)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
'lblStatus.BackColor = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = colr2 Then dif = -100 ': GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
skipit:
Next y
'Debug.Print x;
Next x
score(num) = score(num) + (totdif + Abs((compixel2 - compixel1) * 500))
doneit:
comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop

End Sub

Public Sub numdifpixels(num As Integer)
' scan image and determine how many time each color is used
'On Error Resume Next
Dim x As Integer, y As Integer, tested As Boolean
Dim pixel As Long, ccount As Integer, ocount As Integer
Set colortable = Nothing
Set colortable = CreateObject("Scripting.Dictionary")
Picture1.Refresh
For x = 0 To 319
For y = 0 To 199
ocount = 0
pixel = Picture1.Point(x, y)
'If pixel = -1 Then Stop
'Debug.Print x, y, pixel
pixelarray(x, y) = pixel
'If pixel = -1 Then Debug.Print x, y: GoTo skipit
' count and score pixel values

If x = 0 And y = 0 Then colorarray(0, 0) = pixel: colorarray(0, 1) = 1: GoTo skipit
tested = False
For ocount = 0 To ccount
If colorarray(ocount, 0) = pixel Then colorarray(ocount, 1) = colorarray(ocount, 1) + 1: tested = True: Exit For
If colorarray(ocount, 1) > mostcolor Then mostcolor = colorarray(ocount, 1) ' keep worst pixel value
'If colorarray(ocount, 1) = 0 Then Stop
Next ocount
If ccount = 255 Then Exit For
'If ocount = 1 And ccount = 1 Then Stop: colorarray(ocount, 0) = pixel: colorarray(ocount, 1) = 1: ccount = ccount + 1
If Not tested Then colorarray(ocount, 0) = pixel: colorarray(ocount, 1) = 1: ccount = ccount + 1
skipit:
Next y
If ccount = 255 Then Exit For
Next x
colorcount = ccount
'*______________
Dim gap As Integer, doneflag As Integer, index As Integer
Dim tmpsc As Double, tmpcnt As Double
 gap = Int(ccount / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For index = 1 To ccount - gap
    
     If colorarray(index, 1) > colorarray(index + gap, 1) Then
     tmpsc = colorarray(index, 1): tmpcnt = colorarray(index, 0)
     colorarray(index, 1) = colorarray(index + gap, 1): colorarray(index, 0) = colorarray(index + gap, 0)
     colorarray(index + gap, 1) = tmpsc: colorarray(index + gap, 0) = tmpcnt
     doneflag = 0
     End If
    Next index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
'*_________________

Dim tzl, price
'Stop
price = 1
colorarray(1, 1) = 1
For tzl = 0 To ccount
Debug.Print colorarray(tzl, 0); " _"; colorarray(tzl, 1);
'If colorarray(tzl, 1) = 0 Then Stop
colortable.Add colorarray(tzl, 0), Int(price)
If colorarray(tzl, 1) <> colorarray(tzl + 1, 1) Then price = price + 0.5
Debug.Print "CT"; colortable.Item(colorarray(tzl, 0))
Next tzl

End Sub


Public Sub rsquarescan(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
Debug.Print "rsquare"
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long
' there's something wrong with scanning black pixels!!!!!
Shape1.Visible = True
Picture2.Refresh
Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Long, totmax As Long, difcount As Long
Dim colr1f As Long, colr2f As Long
'totmax = (scanwid / xstep) * (scanhyt / ystep) * (255+255+255)
totdif = 0
For y = 0 To 199
For x = 0 To 319
colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr2f = -1 Then colr2f = Picture2.Point(0, y)
If colr1f <> colr2f Then difcount = difcount + 1
If difcount > 1000 Then Exit For
Next x
If difcount > 1000 Then Exit For
Next y
Debug.Print "dif"; difcount
If difcount = 0 Then score(num) = 88821605: GoTo doneit

Dim xpos As Integer, ypos As Integer

For x = 1 To rswid
DoEvents
For y = 1 To rshyt

xpos = x + xcorn: ypos = y + ycorn
If xpos > 320 Then xpos = xpos - 320
If ypos > 200 Then ypos = ypos - 200

colr1 = pixelarray(xpos, ypos)
colr2 = Picture2.Point(xpos, ypos)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
'lblStatus.BackColor = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = colr2 Then dif = -100 ': GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + dif
Next y
'Debug.Print x;

Next x

score(num) = score(num) + (totdif + Abs((compixel2 - compixel1) * 500))
If score(num) = 0 Then Stop
doneit:

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop
End Sub


Public Sub randscan(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
DoEvents
On Error GoTo errhandl
comparing = True
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long

Dim n1 As Long, n2 As Long, t As Integer
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double, difcount As Long
Dim colr1f As Long, colr2f As Long
For y = 0 To 199
For x = 0 To 319
colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr2f = -1 Then colr2f = Picture2.Point(0, y)
If colr1f <> colr2f Then difcount = difcount + 1
If difcount > 20 Then Exit For
Next x
If difcount > 20 Then Exit For
Next y
Debug.Print "dif"; difcount
If difcount = 0 Then score(num) = 88821605: GoTo doneit

For t = 1 To randstep
x = Rnd * 320
y = Rnd * 200

colr1 = pixelarray(x, y)
colr2 = Picture2.Point(x, y)
If colr1 = comcolr1 Then compixel1 = compixel1 + 1
If colr2 = comcolr2 Then compixel2 = compixel2 + 1
comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'If colr1 = colr2 Then dif = -1000
'skipZero:
totdif = totdif + dif
Next t
score(num) = score(num) + totdif + Abs((compixel2 - compixel1) * 500)
doneit:

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'Stop
End Sub
Public Sub randscan2(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True
'Dim compixel1 As Long, compixel2 As Long
'Dim comcolr1 As Long, comcolr2 As Long
Dim n1 As Long, n2 As Long, t As Integer
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
For t = 1 To randstep
x = Rnd * 320
y = Rnd * 200

colr1 = pixelarray(x, y)
colr2 = Picture2.Point(x, y)
    n1 = colr1: n2 = colr2
'    n1 = colr1 And colr2
    rd1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    blu1 = n1
    rd2 = n2 Mod 256: n2 = Int(n2 / 256)
    gren2 = n2 Mod 256: n2 = Int(n2 / 256)
    blu2 = n2
    dif = Abs((blu1 Or blu2) - blu1) + Abs((gren1 Or gren2) - gren1) + Abs((rd1 Or rd2) - rd1)
    totdif = totdif + dif
Next t
score(num) = score(num) + totdif

comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'stop
End Sub
Public Sub randscan3(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
On Error GoTo errhandl
comparing = True
'Dim compixel1 As Long, compixel2 As Long
'Dim comcolr1 As Long, comcolr2 As Long
Dim n1 As Long, n2 As Long, t As Integer
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double
'totdif = 88821610
For t = 1 To randstep
x = Rnd * 320
y = Rnd * 200

colr1 = pixelarray(x, y)
colr2 = Picture2.Point(x, y)
    If colr1 = colr2 Then totdif = totdif + 1
Next t
score(num) = score(num) + totdif
comparing = False
Debug.Print score(num)
Exit Sub
errhandl:
'stop
End Sub


Public Sub compare2nabors(num As Integer)
' get the difference between 2 adjacent pixels in each image and compare their differences.
comparing = True
On Error GoTo errhandl
Dim compixel1 As Long, compixel2 As Long
Dim comcolr1 As Long, comcolr2 As Long
Dim n1 As Long, n2 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim dblu2 As Integer, dgren2 As Integer, drd2 As Integer
Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
Dim colr1f As Long, colr2f As Long
Dim dif As Double, dif2 As Double, totdif As Double
Dim difcount As Long, difcount2 As Long
'For t = 1 To 300
'x = Rnd * 317
'y = Rnd * 197
For y = 0 To 199
For x = 0 To 319 Step 2
colr1 = pixelarray(x, y)
colr2 = pixelarray(x + 1, y)
n1 = colr1: n2 = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
dif = dblu + dgren + drd

colr1f = Picture2.Point(x, y)
colr2f = Picture2.Point(x + 1, y)
n1 = colr1f: n2 = colr2f
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
dblu2 = Abs(blu1 - blu2): dgren2 = Abs(gren1 - gren2): drd2 = Abs(rd1 - rd2)

dif2 = dblu2 + dgren2 + drd2
If dif2 = 0 Then difcount = difcount + 1
'If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
'If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
'skipZero:
totdif = totdif + Abs(dif - dif2)

Next x
Next y
Debug.Print "difcount:"; difcount

If difcount >= mindifcount Then totdif = totdif + 50000

score(num) = score(num) + totdif '+ ((compixel2 - compixel1) * 100)

comparing = False
Debug.Print totdif
Exit Sub
errhandl:
'Stop
End Sub
Public Sub boundaryscan(num As Integer)
DoEvents
comparing = True
lblStatus.Caption = "Scanning": lblStatus.Refresh
Dim colr1 As Long, colr2 As Long, blu1 As Integer, blu2 As Integer
Dim gren1 As Integer, gren2 As Integer, rd1 As Integer, rd2 As Integer
Dim rd3 As Integer, rd4 As Integer, gren3 As Integer, gren4 As Integer
Dim blu3 As Integer, blu4 As Integer

Dim colr1f As Long, colr2f As Long, n1 As Long, n2 As Long, n3 As Long, n4 As Long
Dim dblu As Integer, dgren As Integer, drd As Integer, dif As Long
Dim dblu2 As Integer, dgren2 As Integer, drd2 As Integer
Dim difcount As Long
Dim dif2 As Long, totdif As Long
Dim x As Integer, y As Integer
totdif = 0
For y = 0 To 199
For x = 0 To 319
colr1f = Picture2.Point(x, y): colr2f = Picture2.Point(x + 1, y)
If colr2f = -1 Then colr2f = Picture2.Point(0, y)
If colr1f <> colr2f Then difcount = difcount + 1
If difcount > 1000 Then Exit For
Next x
If difcount > 1000 Then Exit For
Next y
Debug.Print "dif"; difcount
If difcount = 0 Then score(num) = 88821605: GoTo doneit

For y = 0 To 199
DoEvents
For x = 0 To 318

colr1 = pixelarray(x, y)
colr2 = pixelarray(x + 1, y)
If colr1 = colr2 Then GoTo skipit
n1 = colr1: n2 = colr2
rd1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
blu1 = n1
rd2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
blu2 = n2
colr1f = Picture2.Point(x, y)
colr2f = Picture2.Point(x + 1, y)
'If colr1F = colr2F Then difcount = difcount + 1
n3 = colr1f: n4 = colr2f
rd3 = n3 Mod 256: n3 = Int(n3 / 256)
gren3 = n3 Mod 256: n3 = Int(n3 / 256)
blu3 = n3
rd4 = n4 Mod 256: n4 = Int(n4 / 256)
gren4 = n4 Mod 256: n4 = Int(n4 / 256)
blu4 = n4
dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)

dblu2 = Abs(blu4 - blu2): dgren2 = Abs(gren4 - gren2): drd2 = Abs(rd4 - rd2)
dif2 = dblu2 + dgren2 + drd2
dif = dblu + dgren + drd
totdif = totdif + dif + dif2 'Abs(dif - dif2)
skipit:
Next x
Next y
'totdif = totdif ' + (difcount * 100)
score(num) = score(num) + totdif '+ ((compixel2 - compixel1) * 100)
doneit:

comparing = False

End Sub
Private Sub Command2_Click()
If pause Then
Command2.Caption = "Pause": pause = False: Timer1.Enabled = True: Exit Sub
End If
If pauseatgenend And Not pause Then
doner
End If
pause = True
If scantype = 5 Then
Shape1.Visible = True
Shape1.Width = rswid: Shape1.Height = rshyt
Shape1.Top = ycorn: Shape1.Left = xcorn
Shape1.Refresh
Picture1.Refresh: Me.Refresh
End If


Command2.Caption = "Resume"
Timer1.Enabled = False
End Sub



Private Sub Command3_Click()
Dim naym$, winder As Long
naym$ = InputBox("Enter name of output image")
If naym$ = "" Then Exit Sub
Timer1.Enabled = False
winder = FindWindow(CLng(0), "Finished - fractint")
'
'If winder <> 0 Then
End Sub

Private Sub Command4_Click()
Me.Hide
Artsy.Show

End Sub

Private Sub Command5_Click(index As Integer)
'Dim naym$
'naym$ = InputBox("Enter name to save as:")
'If naym$ = "" Then Exit Sub

Select Case index
Case 0
Clipboard.SetData Picture2.Picture
Case 1
Clipboard.SetData Picture3.Picture
Case 2
Clipboard.SetData Picture3.Picture
End Select
End Sub

Public Sub setPlex(num As Integer)
plex(num) = True
End Sub
Public Sub resetPlex(num As Integer)
plex(num) = False
End Sub
Public Sub erasescores()
Erase score
End Sub


Private Sub Command6_Click(index As Integer)
If index = 1 Then score(curnum) = 88812560

End Sub


Private Sub Command7_Click()
Manip.Show
End Sub

Private Sub Command8_Click()
Dim winder As Long

Debug.Print tik
Label1.Caption = tik
winder = FindWindow(CLng(0), "Ultra Fractal")
'SendMessage winder, 17, CLng(0), CLng(0)
SendMessage winder, 18, CLng(0), CLng(0)

SendMessage winder, 70, CLng(0), CLng(0)

SendMessage winder, 40, CLng(0), CLng(0)

'SendMessage winder, 17, CLng(0), CLng(0)
SendMessage winder, 69, CLng(0), CLng(0)
SendMessage winder, 5, CLng(0), CLng(0)
tik = tik + 1
End Sub

Private Sub export_Click()
Dim parfylname As String, tzl
parfylname = InputBox("Enter filename")
If parfylname = "" Then Exit Sub
If InStr(1, parfylname, ".", vbBinaryCompare) = 0 Then parfylname = parfylname & ".par"
Open parfylname For Output As #1
For tzl = 1 To 100
Print #1, Left(parfylname, 3) & tzl & "{"
Print #1, param$(tzl)
Print #1, "}"
Next tzl

Close 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii
Select Case Chr(KeyAscii)
Case "s"
End Select

End Sub

Private Sub Form_Load()
tik = 17
initPath = App.Path
fractype = "mandelbrot"
icolor = "1": ocolor = "iter": bailouttest = "mod"
scantype = 1
mutrate = 5.9
jobnaym$ = CStr(Int(Rnd * 100000))
popsize = 100
Randomize Timer
palmap = "default.map"
xstep = 10: ystep = 10: randstep = 900
ycorn = Int(Rnd * 320)
xcorn = Int(Rnd * 200)
rswid = 32: rshyt = 20
plex(0) = True
plexiter = 0
Dim winder As Long
ChDir App.Path
If Dir(App.Path & "\scan.gif", vbNormal) <> "" Then Kill App.Path & "\scan.gif"
winder = FindWindow(CLng(0), "Finished - fractint")
' might want to use a do loop or enumerate windows here to be sure no extra old fractints are running
If winder <> 0 Then SendMessage winder, WM_CLOSE, CLng(0), CLng(0)
Randomize
'bestscore = 0: curbest = 0
bestscore = 88821610: curbest = 88821610
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim winder As Long
Unload Options
winder = FindWindow(CLng(0), "fractint")
If winder <> 0 Then SendMessage winder, WM_CLOSE, CLng(0), CLng(0)
winder = FindWindow(CLng(0), "Finished - fractint")
If winder <> 0 Then SendMessage winder, WM_CLOSE, CLng(0), CLng(0)
End
End Sub

Private Sub Image1_Click()
Picture1.Visible = Not Picture1.Visible

End Sub

Private Sub interbreed_Click()
' open a pop file and breed to current pop

Dim crosspop(101) As popmem
Dim everyoneelse As everybody

Dim fylnam As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Fractal Evolver population files (*.pop)|*.pop"
CommonDialog1.ShowOpen
fylnam = CommonDialog1.FileName
'If Not pause Then Exit Sub
Open fylnam For Binary As #1
Dim t As Integer
Get #1, 1, everyoneelse
Close 1
For t = 1 To 100
 crosspop(t) = everyoneelse.pop(t)
Next t

Debug.Print "breed"
Dim md As Double
Dim fem(20) As popmem, mal(20) As popmem, fi As popmem, mi As popmem
Dim off1 As popmem
Dim k As Integer, q As Integer, ma As Integer, fe As Integer
Dim param1 As Double, param2 As Double
Dim centerx As Double, centery As Double, mag As Double

'Dim fparam1 As double, fparam2 As double
'Dim fcenterx As double, fcentery As double, fmag As double
'Dim mparam1 As double, mparam2 As double
'Dim mcenterx As double, mcentery As double, m_mag As double
k = 5
For q = 1 To 10
fem(q) = population(q)
'If score(q) = bestscore Then fem(q) = mutate(fem(q))
Next
For q = 1 To 5
mal(q) = population(q + 10)
'If score(q + 10) = bestscore Then mal(q) = mutate(mal(q))
Next
For ma = 1 To 5
For fe = 1 To 19
fi = fem(fe)
mi = mal(ma)
If (fi.centerx = mi.centerx) Or (fi.centery = mi.centery) Or (fi.mag = mi.mag) Or (fi.param1 = mi.param1) Or (fi.param2 = mi.param2) Then fi = mutate(fi)
' average routine here
param1 = (fi.param1 + mi.param1) / 2
param2 = (fi.param2 + mi.param2) / 2
mag = (fi.mag + mi.mag) / 2
centerx = (fi.centerx + mi.centerx) / 2
centery = (fi.centery + mi.centery) / 2
    population(k).centerx = centerx
    population(k).centery = centery
    population(k).mag = mag
    population(k).param1 = param1
    population(k).param2 = param2
'population(k) = off1
'population(k - 1) = off2
k = k + 1
brnxt: Next fe, ma
For t = 10 To 100
md = Rnd * 100
If md > 10 Then
   Debug.Print t;
'   off1 = population(t)
'   population(t) = mutate(off1)
population(t).centerx = population(t).centerx + ((Rnd * 2) - 1)
population(t).centery = population(t).centery + ((Rnd * 2) - 1)
population(t).mag = population(t).mag + ((Rnd * 2) - 1)
population(t).param1 = population(t).param1 + ((Rnd * 2) - 1)
population(t).param2 = population(t).param2 + ((Rnd * 2) - 1)
End If
Next t
Erase mal
Erase fem
curbest = 88821610
'curbest = 0
End Sub
Public Function getParam(num As Integer) As String
getParam = param(num)

End Function
Private Sub loadpop_Click()
On Error GoTo erhandl
'Dim everyone As everybody
Dim fylnam As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Fractal Evolver population files (*.pop)|*.pop"
CommonDialog1.ShowOpen
fylnam = CommonDialog1.FileName
'If Not pause Then Exit Sub
Open fylnam For Binary As #1
Dim t As Integer
Get #1, 1, Module1.everyone
Close 1
For t = 1 To 100
 population(t) = Module1.everyone.pop(t)
 'Debug.Print population(t).centerx
Next t

With Module1.everyone
fractype = .fractype
numgens = .gen
icolor = .icolor
ocolor = .ocolor
palmap = .palmap
scantype = .scantype
End With

If scantype = 0 Then scantype = 1: randstep = 900
loadedpop = True
makeParams

erhandl:
'makeParams
End Sub

Private Sub mangrad_Click()
Grader.Show

End Sub

Private Sub multimage_Click()
On Error GoTo errhandlr
'CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
Dim fyl As String
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub
'Image1.Picture = LoadPicture(fyl)
PiClip1.Picture = LoadPicture(fyl)
PiClip1.ClipX = 0: PiClip1.ClipY = 0
PiClip1.ClipHeight = PiClip1.Height
PiClip1.ClipWidth = PiClip1.Width

PiClip1.StretchX = 320
PiClip1.StretchY = 200
Picture1.Picture = PiClip1.Clip

Picture1.Picture = PiClip1.Clip 'Image1.Picture ' LoadPicture(fyl)
Picture1.Refresh
'Set srcPic = Picture1.Picture

'goalfylname = CommonDialog1.FileTitle
'FracSearch.Palette = Picture1.Picture
'FracSearch.PaletteMode = vbPaletteModeCustom
'ImageScan
ReDim Preserve multipixur(pixurcount)
Set multipixur(pixurcount) = PiClip1.Clip
Debug.Print multipixur(pixurcount)
pixurcount = pixurcount + 1
If pixurcount > 1 Then multiplmage = True



If Not Command2.Enabled Then Command1.Enabled = True
Command2.Enabled = True
Exit Sub
errhandlr:


End Sub

Private Sub mutatepop_Click()
Dim tzl As Integer, unpause As Boolean
If Command2.Caption <> "Pause" Then Command2_Click: unpause = True
For tzl = 1 To 100
population(tzl) = mutate(population(tzl))
Next tzl
makeParams
If unpause Then Command2_Click

End Sub

Private Sub nupop_Click()
curnum = 0
numgens = 0
lblGenkount.Caption = "Gen: 0"
makePop
makeParams
End Sub

Private Sub openImg_Click()
On Error GoTo errhandlr
'CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
Dim fyl As String
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub
'Image1.Picture = LoadPicture(fyl)
PiClip1.Picture = LoadPicture(fyl)
PiClip1.ClipX = 0: PiClip1.ClipY = 0
PiClip1.ClipHeight = PiClip1.Height
PiClip1.ClipWidth = PiClip1.Width

PiClip1.StretchX = 320
PiClip1.StretchY = 200
Picture1.Picture = PiClip1.Clip

Picture1.Picture = PiClip1.Clip 'Image1.Picture ' LoadPicture(fyl)
Picture1.Refresh
'Set srcPic = Picture1.Picture

goalfylname = CommonDialog1.FileTitle
'FracSearch.Palette = Picture1.Picture
'FracSearch.PaletteMode = vbPaletteModeCustom
ImageScan
If Not Command2.Enabled Then Command1.Enabled = True
Command2.Enabled = True
Exit Sub
errhandlr:

End Sub
Public Sub mapmaker(imagename$)
' decompose color list and write out map file
Dim tzl
Open imagename$ & ".map" For Output As #1
For tzl = 0 To 255

Next tzl

Close 1
End Sub


Private Sub opt_Click()
Options.Show
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo erhandl
If Button = 1 Then
Clipboard.SetText param$(curnum)
End If
If Button = 2 Then
    If comparing Then
    score(curnum) = curbest - 10 'score(curnum) - 12000
    'Label1.Caption = curnum - 1 & ": " & score(curnum - 1): Label1.Refresh
    Else
    score(curnum - 1) = curbest - 10 ' score(curnum - 1) - 12000
    Label1.Caption = curnum - 1 & ": " & score(curnum - 1): Label1.Refresh
    End If
End If
erhandl:
End Sub

Private Sub Picture3_Click()
Clipboard.SetText Text2
End Sub

Private Sub Picture4_Click()
Clipboard.SetText Text1
End Sub

Private Sub popedit_Click()
If Not loadedpop Then Exit Sub
EditDB.loadGrid

EditDB.Show
End Sub

Private Sub savepop_Click()
On Error GoTo errhandl
'Dim everyone As everybody
Dim fylnam As String
CommonDialog1.Filter = "Fractal Evolver population files (*.pop)|*.pop"
CommonDialog1.ShowSave
fylnam = CommonDialog1.FileName
If fylnam = "" Then Exit Sub
If LCase(Right(fylnam, 4)) <> ".pop" Then fylnam = fylnam & ".pop"
Dim t As Integer
For t = 1 To 100
everyone.pop(t) = population(t)
Next

everyone.fractype = fractype
everyone.gen = numgens
everyone.icolor = icolor
everyone.ocolor = ocolor
everyone.palmap = palmap
everyone.scantype = scantype
everyone.mutrate = mutrate
everyone.evbail = eBail
everyone.evinvert = eInvert
everyone.evxmag = eXmag
everyone.evbio = eBio
everyone.evmaxiter = eMaxit
everyone.evdecomp = eDecomp
everyone.evrot = eRot
everyone.evskew = eSkew

'Public eBail As Boolean, eInvert As Boolean, eXmag As Boolean
'Public eBio As Boolean, eMaxit As Boolean, eDecomp As Boolean
'Public eRot As Boolean, eSkew As Boolean


Debug.Print fylnam

Open fylnam For Binary As #1
Put #1, , everyone
Close 1
errhandl:
End Sub
Private Sub saveGen()
On Error GoTo errhandl
Dim everyone As everybody
Dim fylnam As String
fylnam = "tempGen.pop"
Dim t As Integer
For t = 1 To 100
everyone.pop(t) = population(t)
Next

everyone.fractype = fractype
everyone.gen = numgens
everyone.icolor = icolor
everyone.ocolor = ocolor
everyone.palmap = palmap
everyone.scantype = scantype
Debug.Print fylnam
Open fylnam For Binary As #1
Put #1, , everyone
Close 1
errhandl:

End Sub

Private Sub Timer1_Timer()
DoEvents
Static multim As Integer


If comparing Then Exit Sub
If quiting Then Exit Sub
If pause Then Exit Sub
If Not Timer1.Enabled Then Exit Sub
On Error GoTo erhand
Dim winder As Long

If Not stillrunning Then Exit Sub

winder = FindWindow(CLng(0), "Finished - fractint")

If winder <> 0 Then
If Dir(initPath & "\scan.gif") = "" Then Stop: Exit Sub
Picture2.Picture = LoadPicture(initPath & "\scan.gif")

lblStatus.Caption = "Scanning": lblStatus.Refresh
'If Dir(App.Path & "\scan.gif", vbNormal) = "" Then Stop
SendMessage winder, WM_CLOSE, CLng(0), CLng(0)
stillrunning = False
If multiplmage Then
    Picture1.Picture = multipixur(multim)
    ImageScan
    multim = multim + 1
    If multim > pixurcount - 1 Then pixurcount = 0
End If

'If plexiter > 5 Then plexiter = 0
    Debug.Print "P:"; plexiter
   
        Select Case plexiter
        Case 0
        randscan curnum
        Case 1
        gridscan curnum
        Case 2
        boundaryscan curnum
        Case 3
        maskscan curnum
        Case 4
        rsquarescan curnum
        Case 5
        gridscan3 curnum
'        Case 6
'        ' manual tic
'        Grader.putPic Picture2.Picture
        Case 8
        compare2nabors curnum
        Case Else
        randscan curnum
        End Select
    
'    If score(curnum) = 88821605 Then Exit For
DoEvents
    
'Else
'' alternative comparisons
'Select Case scantype
'Case 1
'randscan curnum
'Case 2
'gridscan curnum
'Case 3
'boundaryscan curnum
'Case 4
'maskscan curnum
'Case 5
'rsquarescan curnum
'Case 6
'gridscan3 curnum
'End Select
'End If
Kill initPath & "\scan.gif"
'lblStatus.Caption = "Counting"
' try to suppress average coalesence
If score(curnum) < bestscore Then Picture4.Picture = Picture2.Picture: bestscore = score(curnum): best = population(curnum): Text1.Text = param(curnum): Label3.Caption = "Best Score:" & CStr(bestscore) & " Gen:" & numgens: Picture4.Refresh
If score(curnum) < curbest Then curbest = score(curnum): Picture3.Picture = Picture2.Picture: Label5.Caption = "Best Current:" & CStr(score(curnum)):: Text2 = param(curnum): Picture3.Refresh
Label1.Caption = CStr(curnum) & ": " & CStr(score(curnum))

If score(curnum) = prevscore Then population(curnum) = mutate(population(curnum))
prevscore = score(curnum)
Grader.putPic Picture2.Picture, curnum
curnum = curnum + 1
If curnum > 100 Then Timer1.Enabled = False: Timer2.Enabled = True: kount = kount + 1: Exit Sub
'lblStatus.Caption = "Fracting"
runFract param$(curnum)
End If

If winder = 0 Then
 winder = FindWindow(CLng(0), "fractint")
 If winder <> 0 Then
 stillrunning = True: ' Label1.Caption = "still running: " & CStr(curnum)
 Else:
' If uploading <> "true" And Not downloading Then getJob
Label1.Caption = "not running" ' not done not running, what the fuck?
 End If
End If
Exit Sub
erhand:
'Dim reserror As String
'Stop
'process = process & "(error " & Err.Description & ")"
'reserror = Inet1.OpenURL("http://ducer.ducer.net/asp/poverrors.asp?error=" & "timer:" & Err.Description & "job" & frame_number & "name" & job2do, icString)
'Text1.Text = Text1.Text & "error in the timer event: " & Err.Description

End Sub

Public Sub doner()
lblStatus.Caption = "Saving"
saveGen
Open jobnaym$ For Append As #1
Write #1, Text2
Close 1


lblStatus.Caption = "Sorting"
sort
' put top score saver routine here
' clean up 0 index pop array problems!~>
lblStatus.Caption = "Breeding"
breed2
numgens = numgens + 1
curnum = 1
'Grader.roe = 1
Grader.flushGrid

lblStatus.Caption = "Converting Params"
makeParams
'ycorn = Int(Rnd * 170) ' for scanning one square area.
'xcorn = Int(Rnd * 290) ' the top and left coords of a 30x30 square

ycorn = Int(Rnd * 320)
xcorn = Int(Rnd * 200)
plexiter = plexiter + 1
If plexiter > 8 Then plexiter = 0
Do While Not plex(plexiter)
If plexiter > 8 Then plexiter = 0
plexiter = plexiter + 1
Loop
If plexiter > 8 Then plexiter = 0
 scantype = plexiter + 1
If scantype = 5 Then Shape1.Visible = True: Shape1.Left = xcorn: Shape1.Top = ycorn Else Shape1.Visible = False
If scantype = 7 Then Grader.Show: Exit Sub
runFract param(curnum)
Timer1.Enabled = True
lblStatus.Caption = "Waiting"

'Stop
End Sub
Public Sub returnfromGrader()
Do While Not plex(plexiter)
plexiter = plexiter + 1
Loop
If plexiter > 5 Then plexiter = 0

runFract param(curnum)
Timer1.Enabled = True
lblStatus.Caption = "Waiting"
End Sub
Public Sub setscore(num As Integer, scor As Double)
score(num) = scor

End Sub
Public Function getScore(num As Integer) As Double
getScore = score(num)
End Function
Sub sort()
Debug.Print "sort"
Dim gap As Integer, doneflag As Integer, index As Integer
Dim tmpsc As Double
Dim tmpb As popmem
 gap = Int(100 / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For index = 1 To 100 - gap
    
     If score(index) > score(index + gap) Then
     tmpsc = score(index)
     score(index) = score(index + gap)
     score(index + gap) = tmpsc
     tmpb = population(index)
     population(index) = population(index + gap)
     population(index + gap) = tmpb
           doneflag = 0
     End If
    Next index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
'  Dim ztl As Integer
'  For ztl = 1 To 30
'  Debug.Print ztl; ":"; score(ztl); " #";
'  Next ztl
Debug.Print score(1), score(2)


End Sub
Public Sub breed()

Debug.Print "breed"
Dim t As Integer, md As Double
Dim fem(20) As popmem, mal(20) As popmem, fi As popmem, mi As popmem

Dim off1 As popmem
Dim k As Integer, q As Integer, ma As Integer, fe As Integer
Dim param1 As Double, param2 As Double
Dim centerx As Double, centery As Double, mag As Double
'Dim fparam1 As double, fparam2 As double
'Dim fcenterx As double, fcentery As double, fmag As double
'Dim mparam1 As double, mparam2 As double
'Dim mcenterx As double, mcentery As double, m_mag As double
k = 0
For q = 1 To 5
mal(q) = population(q)
'If score(q) = curbest Then mal(q) = mutate(fem(q))
Next
For q = 1 To 20
fem(q) = population(q + 5)
'If score(q + 10) = curbest Then mal(q) = mutate(mal(q))
Next
curbest = 88821610
For ma = 1 To 5
For fe = 1 To 20
fi = fem(fe)
mi = mal(ma)
If (fi.centerx = mi.centerx) Or (fi.centery = mi.centery) Or (fi.mag = mi.mag) Or (fi.param1 = mi.param1) Or (fi.param2 = mi.param2) Then fi = mutate(fi)

' average routine here
param1 = (fi.param1 + mi.param1 * 2) / 3
param2 = (fi.param2 + mi.param2 * 2) / 3
mag = (fi.mag + mi.mag * 2) / 3
centerx = (fi.centerx + mi.centerx * 2) / 3
centery = (fi.centery + mi.centery * 2) / 3
population(k).centerx = centerx
population(k).centery = centery
population(k).mag = mag
population(k).param1 = param1
population(k).param2 = param2
'population(k) = off1
'population(k - 1) = off2
k = k + 1
brnxt: Next fe, ma
For t = 1 To 100
md = Rnd * 100
If md < mutrate Then
   Debug.Print t;
'   off1 = population(t)
'   population(t) = mutate(off1)
population(t).centerx = population(t).centerx + ((Rnd * 2) - 1)
population(t).centery = population(t).centery + ((Rnd * 2) - 1)
population(t).mag = population(t).mag + ((Rnd * 2) - 1)
population(t).param1 = population(t).param1 + ((Rnd * 2) - 1)
population(t).param2 = population(t).param2 + ((Rnd * 2) - 1)

   End If
Next t
Erase mal
Erase fem

End Sub

Public Sub breed2()
Debug.Print "breed2"
lblStatus.Caption = "Breeding"
Dim t As Integer, md As Double
Dim fem(20) As popmem, mal(20) As popmem, fi As popmem, mi As popmem
Dim k As Integer, q As Integer, ma As Integer, fe As Integer
Dim param1 As Double, param2 As Double
Dim centx As Double, centy As Double, mag As Double, ivert As Double
Dim icentx As Double, icenty As Double, rota As Double
Dim biom As Integer, maxit As Long, decmp As Integer, bail As Long, xmg As Double
Dim skw As Double
k = 0
For q = 1 To 20
fem(q) = population(q)
If score(q) = curbest Then fem(q) = mutate(fem(q))
Next
For q = 1 To 5
mal(q) = population(q + 20)
If score(q + 10) = curbest Then mal(q) = mutate(mal(q))
Next
curbest = 88821610
For ma = 1 To 5
For fe = 1 To 20
fi = fem(fe): mi = mal(ma)
If (fi.centerx = mi.centerx) Or (fi.centery = mi.centery) Or (fi.mag = mi.mag) Or (fi.param1 = mi.param1) Or (fi.param2 = mi.param2) Then fi = mutate(fi)
' average routine here
param1 = (fi.param1 + mi.param1) / 2
param2 = (fi.param2 + mi.param2) / 2
mag = (fi.mag + mi.mag) / 2
centx = (fi.centerx + mi.centerx) / 2
centy = (fi.centery + mi.centery) / 2
ivert = (fi.invert + mi.invert) / 2
icentx = (fi.icenterx + mi.icenterx) / 2
icenty = (fi.icentery + mi.icentery) / 2
rota = (fi.rot + mi.rot) / 2
biom = (fi.biomorph + mi.biomorph) / 2
maxit = (fi.maxiter + mi.maxiter) / 2
decmp = (fi.decomp + mi.decomp) / 2
bail = (fi.bailout + mi.bailout) / 2
xmg = (fi.xmag + mi.xmag) / 2
skw = (fi.skew + mi.skew) / 2
population(k).centerx = centx
population(k).centery = centy
population(k).mag = mag
population(k).param1 = param1
population(k).param2 = param2
population(k).invert = ivert
population(k).icenterx = icentx
population(k).rot = rota
population(k).xmag = xmg
population(k).biomorph = biom
population(k).maxiter = maxit
population(k).decomp = decmp
population(k).bailout = bail
population(k).skew = skw
k = k + 1
brnxt: Next fe, ma
For t = 1 To 100
md = Rnd * 100
If md < mutrate Then
   Debug.Print t;
'   off1 = population(t)
'   population(t) = mutate(off1)
population(t) = mutate(population(t))
   End If
Next t
Erase mal
Erase fem
Erase score
For t = 0 To 100
everyone.pop(t) = population(t)
Next t

End Sub
Public Sub tourney()
Dim mas(20) As popmem
Dim das(5) As popmem
Dim beenpicked(100) As Boolean
Dim tzl, rpick, rpick2
' select 20 mas and 5 pas
For tzl = 1 To 20
rpick = Int(Rnd * 101)
rpick2 = Int(Rnd * 101)
If score(rpick) > score(rpick2) And Not beenpicked(rpick) Then
beenpicked(rpick) = True
End If
If score(rpick2) > score(rpick) And Not beenpicked(rpick) Then
beenpicked(rpick2) = True
End If
Next tzl
' setup tenths ranging breeding
' compare odd/even, keep best
' repeat until breeding pool diminishes to 1/10 concentration
' breed each pair with neighbor with 1/10 increment averaging across
' parents value ranges.



'Dim ma, fe
'For ma = 1 To 5
'For fe = 1 To 20
'fi = fem(fe): mi = mal(ma)
'If (fi.centerx = mi.centerx) Or (fi.centery = mi.centery) Or (fi.mag = mi.mag) Or (fi.param1 = mi.param1) Or (fi.param2 = mi.param2) Then fi = mutate(fi)
'' average routine here
'param1 = (fi.param1 + mi.param1) / 2
'param2 = (fi.param2 + mi.param2) / 2
'mag = (fi.mag + mi.mag) / 2
'centx = (fi.centerx + mi.centerx) / 2
'centy = (fi.centery + mi.centery) / 2
'ivert = (fi.invert + mi.invert) / 2
'icentx = (fi.icenterx + mi.icenterx) / 2
'icenty = (fi.icentery + mi.icentery) / 2
'rota = (fi.rot + mi.rot) / 2
'biom = (fi.biomorph + mi.biomorph) / 2
'maxit = (fi.maxiter + mi.maxiter) / 2
'decmp = (fi.decomp + mi.decomp) / 2
'bail = (fi.bailout + mi.bailout) / 2
'xmg = (fi.xmag + mi.xmag) / 2
'skw = (fi.skew + mi.skew) / 2
'population(k).centerx = centx
'population(k).centery = centy
'population(k).mag = mag
'population(k).param1 = param1
'population(k).param2 = param2
'population(k).invert = ivert
'population(k).icenterx = icentx
'population(k).rot = rota
'population(k).xmag = xmg
'population(k).biomorph = biom
'population(k).maxiter = maxit
'population(k).decomp = decmp
'population(k).bailout = bail
'population(k).skew = skw
'k = k + 1
'brnxt: Next fe, ma
End Sub

Private Function mutate(member As popmem) As popmem
member.centerx = member.centerx + ((Rnd * 2) - 1)
member.centery = member.centery + ((Rnd * 2) - 1)
member.mag = member.mag + ((Rnd * 2) - 1)
member.param1 = member.param1 + ((Rnd * 2) - 1)
member.param2 = member.param2 + ((Rnd * 2) - 1)
If eBail Then member.bailout = member.bailout + ((Rnd * 4) - 2): member.bailout = Abs(member.bailout)
If eInvert Then
member.icenterx = member.icenterx + ((Rnd * 2) - 1)
member.icentery = member.icentery + ((Rnd * 2) - 1)
member.invert = member.invert + (Rnd - 0.5)
End If
If eBio Then member.biomorph = member.biomorph + ((Rnd * 4) - 2): member.biomorph = Abs(member.biomorph)
If eRot Then member.rot = member.rot + ((Rnd * 4) - 2)
If eXmag Then member.xmag = member.xmag + ((Rnd * 4) - 2)
If eDecomp Then member.decomp = member.decomp + ((Rnd * 4) - 2): member.decomp = Abs(member.decomp)
If eMaxit Then member.maxiter = member.maxiter + ((Rnd * 4) - 2): member.maxiter = Abs(member.maxiter)
If eSkew Then member.skew = member.skew + ((Rnd * 4) - 2)
mutate = member
End Function

Private Sub Timer2_Timer()
If Not Timer2.Enabled Then Exit Sub
If quiting Then Exit Sub
Timer2.Enabled = False
lblGenkount.Caption = "Gen " & CStr(kount)
If Not pauseatgenend Then doner Else Command2_Click

End Sub

Private Sub xit_Click()
quiting = True
On Error Resume Next
 Kill initPath & "\scan.gif"
Dim winder As Long
winder = FindWindow(CLng(0), "Finished - fractint")
If winder <> 0 Then SendMessage winder, WM_CLOSE, CLng(0), CLng(0)
Erase population
Unload Me

End Sub
 
 Private Sub FracScan()
 Dim tzl As Integer
 ' ideas for criss-cross and other forms of scanning
 ' find the corner pixels first ( lines which have the start and end pixels
 ' the same as the corners of the desired image.
 ' Mark off regions.
 ' Sub scan regions for additional pixel correctness.

' Xor is an instant test for commonality at each bit location.
' common values will Xor closer to 0. Identical values=0


 End Sub

'Private Function fractalpick(member As popmem2, typer As Integer) As Long 'param1 As double, param2 As double, x As double, y As double, increm As double) As double
'Dim gzl As Integer, xi As Double, yi As Double, param1 As Double, param2 As Double, increm As Double
'' get a fractal point and it's neighbor
'Dim wfactor, hfactor As Double, xr As Integer, yr As Integer, F1 As Double
'wfactor = 1 / 320
'hfactor = 1 / 200
'
'param1 = member.param1
'param2 = member.param2
'increm = member.increm
'xi = member.cornerx
'yi = member.cornery
'
'Select Case typer
'Case 1 'mandel
'xr = Int(Rnd * 320)
'yr = Int(Rnd * 200)
'increm = member.increm
'xi = member.cornerx
'yi = member.cornery
'xi = xi + (increm) ' fix here
'yi = yi + (increm)
'F1 = mandelpixel(xi, yi, increm)
'Case 2 ' julia
'xi = xi + (Int(Rnd * 320) * increm)
'yi = yi + (Int(Rnd * 200) * increm)
'F1 = juliapixel(param1, param2, xi, yi, increm)
'End Select
'' calc the delta
'
'' get a pixel and it's neighbor
'' calc the delta RGB
'' calc the delta of dF and dP and add to score
'For gzl = 1 To 20
'
'Next gzl
'End Function
Public Function getRGB(colr As Long, rgb123 As Integer) As Integer
Dim rd, gren, blu As Integer
Dim n1 As Long
blu = n1 Mod 256: n1 = Int(n1 / 256)
gren = n1 Mod 256: n1 = Int(n1 / 256)
rd = n1
Select Case rgb123
Case 1
getRGB = rd
Case 2
getRGB = gren
Case 3
getRGB = blu
End Select
End Function
Public Sub ImageScan()
'On Error Resume Next
Dim x As Integer, y As Integer, tested As Boolean
Dim pixel As Long, ccount As Integer, ocount As Integer
Set colortable = Nothing
Set colortable = CreateObject("Scripting.Dictionary")
Picture1.Refresh
For x = 0 To 319
For y = 0 To 199
ocount = 0
pixel = Picture1.Point(x, y)
'If pixel = -1 Then Stop
'Debug.Print x, y, pixel
pixelarray(x, y) = pixel
'If pixel = -1 Then Debug.Print x, y: GoTo skipit
' count and score pixel values

If x = 0 And y = 0 Then colorarray(0, 0) = pixel: colorarray(0, 1) = 1: GoTo skipit
tested = False
For ocount = 0 To ccount
If colorarray(ocount, 0) = pixel Then colorarray(ocount, 1) = colorarray(ocount, 1) + 1: tested = True: Exit For
If colorarray(ocount, 1) > mostcolor Then mostcolor = colorarray(ocount, 1) ' keep worst pixel value
'If colorarray(ocount, 1) = 0 Then Stop
Next ocount
If ccount = 255 Then Exit For
'If ocount = 1 And ccount = 1 Then Stop: colorarray(ocount, 0) = pixel: colorarray(ocount, 1) = 1: ccount = ccount + 1
If Not tested Then colorarray(ocount, 0) = pixel: colorarray(ocount, 1) = 1: ccount = ccount + 1
skipit:
Next y
If ccount = 255 Then Exit For
Next x
colorcount = ccount
'*______________
Dim gap As Integer, doneflag As Integer, index As Integer
Dim tmpsc As Double, tmpcnt As Double
 gap = Int(ccount / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For index = 1 To ccount - gap
    
     If colorarray(index, 1) > colorarray(index + gap, 1) Then
     tmpsc = colorarray(index, 1): tmpcnt = colorarray(index, 0)
     colorarray(index, 1) = colorarray(index + gap, 1): colorarray(index, 0) = colorarray(index + gap, 0)
     colorarray(index + gap, 1) = tmpsc: colorarray(index + gap, 0) = tmpcnt
     doneflag = 0
     End If
    Next index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
'*_________________

Dim tzl, price
'Stop
price = 1
colorarray(1, 1) = 1
For tzl = 0 To ccount
Debug.Print colorarray(tzl, 0); " _"; colorarray(tzl, 1);
'If colorarray(tzl, 1) = 0 Then Stop
colortable.Add colorarray(tzl, 0), Int(price)
If colorarray(tzl, 1) <> colorarray(tzl + 1, 1) Then price = price + 0.5
Debug.Print "CT"; colortable.Item(colorarray(tzl, 0))
Next tzl
End Sub
