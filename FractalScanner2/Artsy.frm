VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Artsy 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   270
      Left            =   5715
      TabIndex        =   4
      Top             =   60
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   315
      Left            =   4245
      TabIndex        =   3
      Top             =   60
      Width           =   1080
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   4500
      Left            =   6120
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   2
      Top             =   420
      Width           =   6000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9285
      Top             =   -75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4500
      Left            =   60
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   1
      Top             =   405
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   345
      Left            =   2250
      TabIndex        =   0
      Top             =   15
      Width           =   945
   End
   Begin VB.Menu fyl 
      Caption         =   "File"
      Begin VB.Menu load_image 
         Caption         =   "Load Image"
      End
      Begin VB.Menu savepop 
         Caption         =   "Save Pop"
      End
      Begin VB.Menu loadpop 
         Caption         =   "Load Population"
      End
      Begin VB.Menu xit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Artsy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' most of the code in this module from
' "The Computer Artist and Art Critic" by J.C. Sprott
' in Clifford Pickover's book "Fractal Horizons"

DefDbl A-Z 'Use double precision
Dim a(12) 'Array of coefficients
Dim qpop(100) As quadpop
Dim n%
Dim inkey$, x, xnew, y, dx, dy, popnum
Dim keepgoin As Boolean
Const w% = 400
Const h% = 300
Dim poploaded As Boolean

Sub doodle()
'Program code 4.1. BASIC Program to Search for Strange Attractors from
'Two-Dimensional Quadratic Maps

Randomize Timer 'Reseed random numbers
'Screen 12 'Assume VGA graphics
n% = 0
While inkey$ = "" 'Loop until a key is pressed
If n% = 0 Then Call setparams(x, y)
Call advancexy(x, y, n%) 'Advance the solution
Call display(x, y, n%) 'Display the results
Call testsoln(x, y, n%) 'Test the solution
Wend
End
End Sub
Sub advancexy(x, y, n%)  'Advance (x, y) at step n%
'SHARED a ( )
xnew = a(1) + x * (a(2) + a(3) * x + a(4) * y) + y * (a(5) + a(6) * y)
y = a(7) + x * (a(8) + a(9) * x + a(10) * y) + y * (a(11) + a(12) * y)
x = xnew
n% = n% + 1
End Sub
Sub dancexy(x, y, n%)  'Advance (x, y) at step n%
'SHARED a ( )
xnew = a(1) + x * (a(2) * x + a(3) * y) + y
y = a(4) + x * (a(5) * x + a(6) * y)
x = xnew
n% = n% + 1
End Sub



Sub display(x, y, n%)  'Display (x, y) at step n%
DoEvents
On Error GoTo skoop
Static xmin As Double, xmax As Double, ymin As Double, ymax As Double, w1 As Long, h1 As Long, xp%, yp%
Dim poynt1 As Long, poynt2 As Long, samelasttime As Boolean, sltcount As Integer
Dim prevx%, prevy%
Select Case n%
Case 1 'Initialize min and MaX x and y
 xmin = 1000: ymin = xmin: xmax = -xmin: ymax = -ymin
 Case 2 To 99 'Skip these
 Case 100 To 999 'Update min and max x and y
If x < xmin Then xmin = x
If x > xmax Then xmax = x
If y < ymin Then ymin = y
If y > ymax Then ymax = y
Case 1000 'Clear the screen and resca1e
Cls
If CSng(xmax) = CSng(xmin) Then xmax = xmin + 1
If CSng(ymax) = CSng(ymin) Then ymax = ymin + 1
dx = (xmax - xmin) / 10: xmin = xmin - dx: xmax = xmax + dx
dy = (ymax - ymin) / 10: ymin = ymin - dy: ymax = ymax + dy
w1 = w% / (xmax - xmin): h1 = h% / (ymin - ymax)
Case Else 'Plot the data
xp% = w1 * (x - xmin)
yp% = h1 * (y - ymax)
'If xp = prevx And yp = prevy Then qpop(popnum).score = qpop(popnum).score - 1
'
'poynt1 = Picture1.Point(xp%, yp%): poynt2 = Picture2.Point(xp%, yp%)
'If poynt1 > RGB(0, 0, 0) And poynt2 = 0 Then
'qpop(popnum).score = qpop(popnum).score + comparePixel(xp, yp)
'End If

prevx = xp: prevy = yp
Picture1.PSet (xp%, yp%), QBColor(n% Mod 16) 'Illuminate screen pixel
qpop(popnum).score = qpop(popnum).score + comparePixel(xp, yp)
End Select
Exit Sub
skoop:
xmin = 1000: ymin = xmin: xmax = -xmin: ymax = -ymin
End Sub

Sub lyapunov(x, y, n%, ByRef l)   'Calculate Lyapunov Exp (1)
Static xe, ye, lsum, xsave, ysave, d2, df, rs
If n% = 1 Then lsum = 0: xe = 0.000001: ye = 0
xsave = x: ysave = y: x = xe: y = ye
Call advancexy(x, y, n% - 1) 'Reiterate equations
'Call dancexy(x, y, n% - 1)
dx = x - xsave: dy = y - ysave: d2 = dx * dx + dy * dy
df = 100000000000# * d2: rs = 1 / Sqr(df)
xe = xsave + rs * (x - xsave): x = xsave
ye = ysave + rs * (y - ysave): y = ysave
lsum = lsum + Log(df)
l = 0.721348 * lsum / n% 'Convert to bits per iteration
End Sub
Sub setparams(x, y)  'Set a() and initialize (x, y)
'SHARED a()
Picture1.Cls
Dim i%
x = 0: y = 0
For i% = 1 To 12: a(i%) = (Int(25 * Rnd) - 12) / 10: Next i%
End Sub
Sub makePop()
If poploaded Then Exit Sub
Debug.Print "making pop"
Dim i%, pn, mzl, nzl
x = 0: y = 0: n% = 0
For pn = 1 To 100
doitagain:
popnum = pn
For i% = 1 To 12
'a(i%) = (Int(25 * Rnd) - 12) / 10
qpop(pn).a(i%) = (Int(25 * Rnd) - 12) / 10
Next i%
'If pn < 25 Then
'    For mzl = 1 To 12
'    a(mzl) = qpop(pn).a(mzl)
'    Next mzl
'    x = 0: y = 0
'    For nzl = 1 To 10000
'    Call advancexy(x, y, n%) 'Advance the solution
'    'Call dancexy(x, y, n%)
'    Call display(x, y, n%) 'Display the results
'    Call testsoln(x, y, n%) 'Test the solution
'    If n% = 0 Then Exit For
'    Next nzl
'    Picture1.Cls
'    n% = 0
'    If qpop(pn).score < 5 Then GoTo doitagain
'    qpop(pn).score = 0
'Debug.Print pn; "M";
'    If pn / 10 = Int(pn / 10) Then Debug.Print
'End If
Next pn

End Sub
Sub testpop()
Dim tzl As Integer, mzl, nzl
n% = 0
For tzl = 1 To 100
popnum = tzl
For mzl = 1 To 12
a(mzl) = qpop(tzl).a(mzl)
Next mzl
x = 0: y = 0
For nzl = 1 To 10000
Call advancexy(x, y, n%) 'Advance the solution
'Call dancexy(x, y, n%)
Call display(x, y, n%) 'Display the results
Call testsoln(x, y, n%) 'Test the solution
If n% = 0 Then Exit For
Next nzl
If nzl < 1000 Then qpop(tzl).score = 999999999

Picture1.Cls
n% = 0
Debug.Print tzl; "-"; qpop(tzl).score; ",";
If tzl / 10 = Int(tzl / 10) Then Debug.Print
Next tzl

End Sub
Public Sub runEm()
makePop
While keepgoin
testpop
sort
breedEm
Wend
End Sub
Public Sub breedEm()
Dim mals(10) As quadpop, fems(20) As quadpop, mzl, fzl, gin
Dim tempb As quadpop, bzl, bodycount
For mzl = 1 To 5
mals(mzl) = qpop(mzl)
Next mzl
For fzl = 6 To 25
fems(fzl - 5) = qpop(fzl)
Next fzl
For mzl = 1 To 5
For fzl = 1 To 20
Debug.Print bodycount + 1; "#";
For bzl = 1 To 12
'If mals(mzl).a(bzl) > fems(fzl).a(bzl) Then
tempb.a(bzl) = (mals(mzl).a(bzl) + fems(fzl).a(bzl)) / 2
'tempb.a(bzl) = (mals(mzl).a(bzl) - fems(fzl).a(bzl)) / 2
If Rnd < 0.2 Then tempb.a(bzl) = tempb.a(bzl) + (Rnd - 0.5): Debug.Print "m";
Debug.Print tempb.a(bzl); " ";
Next bzl
Debug.Print
bodycount = bodycount + 1
qpop(bodycount) = tempb

Next fzl
Next mzl
End Sub

Sub sort()
Debug.Print "sort"
Dim gap As Integer, doneflag As Integer, index As Integer
Dim tmpsc As Double
Dim tmpb As quadpop
 gap = Int(100 / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For index = 1 To 100 - gap
'     If qpop(index).score < qpop(index + gap).score Then 'largest first
     If qpop(index).score > qpop(index + gap).score Then ' smallest
     'tmpsc = qpop(index).score
     'qpop(index).score = qpop(index + gap).score
     'qpop(index + gap).score = tmpsc
     tmpb = qpop(index)
     qpop(index) = qpop(index + gap)
     qpop(index + gap) = tmpb
    doneflag = 0
     End If
    Next index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
  For index = 1 To 100
  Debug.Print qpop(index).score; " ";
  Next index
  Debug.Print
End Sub
Public Sub savEm()


End Sub

Sub scoreIt(popnum As Integer)
If Not Picture1.Point(x, y) Then
qpop(popnum).score = qpop(popnum).score + 1
End If
End Sub




Sub testsoln(x, y, n%)  'Test solution at (x, y. n%)
Dim nmax%, l
nmax% = 10000 'Bailout value
Call lyapunov(x, y, n%, l) 'Get Lyapunov exponent (l)
If n% = nmax% Then n% = 0 'Bailout value reached
If n% > 100 And l < 0.005 Then n% = 0 'Solution is not chaotic
If Abs(x) + Abs(y) > 1000000 Then n% = 0 'Solution is unbounded
End Sub

Private Sub Command1_Click()
'doodle
'makePop
'testpop
If keepgoin Then Exit Sub
keepgoin = True
runEm

End Sub


'Program code 4.2 BASIC Subroutine to Estimate the Correlation Dimension and Its Uncertainty
'.
Sub corrdim(x, y, n%, f, df)  'Returns correlation dim (f~df) STATIC nl, n2, xs()', ys(), rsqm, newt, newdf .
Static n1, n2, xs(), ys(), rsqm, newf, newdf, ns%
Dim i%, j%, rsq
ns% = 1000 'Number of previous points saved
If n% = 1 Then 'Initialize variables
ReDim xs(999), ys(999)
n1 = 0: n2 = 0: rsqm = 0: newf = 0: newdf = 0
End If
i% = n% Mod ns%
j% = (i% + Int(ns% * Rnd / 2)) Mod ns% 'Choose a random reference point
dx = x - xs(j%): dy = y - ys(j%)
rsq = dx * dx + dy * dy 'Calculate square of separation
If n% < ns% Then
If rsq > rsqm Then rsqm = rsq 'Save maximum rsq in rsqm ELSE
If rsq < 0.0003 * rsqm Then 'Point was inside large sphere
n2 = n2 + 1
If rsq < 0.00003 * rsqm Then 'Point was inside small sphere
n1 = n1 + 1
newf = 0.868589 * Log(n2 / n1)
newdf = 0.868589 * Sqr(1 / n1 - 1 / n2)
End If
End If
End If
xs(i%) = x: ys(i%) = y 'Repl~ce oldest point with new
f = newf: df = newdf
End Sub


'Program code 4.3 BASIC Subroutine to Plot Escape-Time Contours
Sub basin(xmin, xmax, ymin, ymax)  'Plot escape-time contours nmax% = 1040 'Bailout value
Dim nc%, dx, dy, nmax%, c%
Dim i%, j%
nc% = 16 'Screen size and colors
dx = (xmax - xmin) / w%: dy = (ymax - ymin) / h%
For i% = 0 To w% - 1
For j% = 0 To h% - 1
x = xmin + dx * i%: y = ymax - dy * j%
n% = 0
While n% < nmax% And x * x + y * y < 1000000
Call advancexy(x, y, n%)
Wend
c% = n% Mod (2 * nc%)
If c% > nc% - 1 Then c% = 2 * nc% - 1 - c%
PSet (i%, j%), c%
Next j%
Next i%
End Sub

Sub ifsxy(x, y, n%)
Dim r%, xnew
r% = 6 * Int(2 * Rnd)
xnew = a(1 + r%) * x + a(2 + r%) * y + a(5 + r%)
y = a(3 + r%) * x + a(4 + r%) * y + a(6 + r%)
x = xnew
n% = n% + 1

End Sub

'
'Program code 4.5 BASIC Subroutines to Calculate Capacity Dimension and Unpredictability from a Screen Display
Sub capdim(ByRef f, ByRef df)  'Returns capacity
Dim n1, n2, n2old, j%, i%, di%, dj%
n1 = 0: n2 = 0: n2old = 0
For j% = 0 To h% - 2 Step 2: For i% = 0 To w% - 2 Step 2
For di% = 0 To 1: For dj% = 0 To 1
If Point(i% + di%, j% + dj%) Then n2 = n2 + 1:
Next dj%: Next di%
If n2 > n2old Then n1 = n1 + 1: n2old = n2
Next i%, j%
f = 1.442695 * Log(n2 / n1)
df = 1.442695 * Sqr(1 / n1 - 1 / n2)
End Sub

Function unpredict() As Double   'Returns unpredictabi1ity (u)
Dim a%(32766), e%, cont%, usum, pold%
Dim j%, i%, p%, k%, dsqm, dsq, d, im%, jm%
    'Big integer array
e% = 5          'Embedding dimension

pold% = Picture2.Point(0, 0): cont% = 0: k% = 0: usum = 0
For j% = 0 To h% - 1: For i% = 0 To w% - 1
    p% = Picture2.Point(i%, j%)
    If p% = pold% Then 'Pixel is same as previous
        If cont% < 32766 Then cont% = cont% + 1
Else 'Pixel is different
        pold% = p%
        a%(k%) = cont%
        If k% < 32766 Then k% = k% + 1
        cont% = 0
End If
Next i%, j% 'Data now stored in array a%(k%)
For i% = 1 To k% - 1 - e%
    dsqm = 1E+37
    For j% = 1 To k% - 1 - e% 'Find closest different point
        dsq = 0
        For n% = 0 To e% - 1
            d = a%(i% + n%) - a%(j% + n%)
            dsq = dsq + d * d
        Next n%
        If dsq > 0 And dsq < dsqm Then dsqm = dsq: im% = i%: jm% = j%
    Next j%
    dsq = 0
    For n% = 1 To e% 'Find separation of next points
        d = a%(im% + n%) - a%(jm% + n%)
        dsq = dsq + d * d
    Next n%
    If dsqm * dsq > 0 Then usum = usum + Log(dsq / dsqm)
Next i%
unpredict = 0.721348 * usum / i% 'Convert to bits per iteration
End Function
'

Private Sub Command2_Click()
keepgoin = False

End Sub

Private Sub Command3_Click()
Grader.Show
End Sub
Public Sub setscore(index As Integer, scor As Long)
qpop(index).score = scor

End Sub
Private Sub Form_Load()
'Me.ScaleWidth = 1280
Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
Picture2.Top = Picture1.Top
Picture2.Left = Picture1.Left + 420

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub load_image_Click()
CommonDialog1.Filter = "gif,jpg or bmp | *.gif;*.jpg;*.bmp;"
CommonDialog1.ShowOpen
Dim fyl As String
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub
Picture2.Picture = LoadPicture(fyl)
End Sub

Public Function comparePixel(xp As Integer, yp As Integer) As Long
Dim colr1, colr2, compixel1, compixel2, comcolr1, comcolr2, n1, n2
Dim blu1, gren1, rd1, blu2, gren2, rd2, dblu
Dim dgren, drd, dif

colr1 = Picture2.Point(xp, yp) 'Picture1.Point(x + xcorn, y + ycorn)
colr2 = Picture1.Point(xp, yp)
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
comparePixel = dif
End Function

Private Sub loadpop_Click()
Dim fyl$
fyl$ = InputBox("Filename")
If fyl$ = "" Then Exit Sub
Dim tzl, qp As qPoper
Open fyl$ For Binary As #1
Get #1, , qp
Close #1
For tzl = 1 To 100
qpop(tzl) = qp.qpop(tzl)
Next tzl
poploaded = True
End Sub

Private Sub savepop_Click()
Dim fyl$
fyl$ = InputBox("Filename")
If fyl$ = "" Then Exit Sub
Dim tzl, qp As qPoper
For tzl = 1 To 100
qp.qpop(tzl) = qpop(tzl)
Next tzl
Open fyl$ For Binary As #1
Put #1, , qp
Close #1


End Sub

Private Sub xit_Click()
Unload Me
Unload FracSearch
End
End Sub
