VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Manip 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Destroy Image List"
      Height          =   315
      Left            =   8490
      TabIndex        =   9
      Top             =   3345
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ImgPicker"
      Height          =   405
      Left            =   8385
      TabIndex        =   8
      Top             =   2820
      Width           =   1920
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Choose Target Pic"
      Height          =   450
      Left            =   8370
      TabIndex        =   6
      Top             =   1650
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9915
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   510
      Left            =   9480
      TabIndex        =   4
      Top             =   4485
      Width           =   885
   End
   Begin PicClip.PictureClip PiClip1 
      Left            =   9375
      Top             =   6870
      _ExtentX        =   1667
      _ExtentY        =   1085
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   420
      Left            =   8385
      TabIndex        =   3
      Top             =   2205
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build Image List"
      Height          =   345
      Left            =   8355
      TabIndex        =   2
      Top             =   1215
      Width           =   1875
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9675
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   524800
      MaxFileSize     =   32000
      Orientation     =   2
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7200
      Left            =   30
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   75
      Width           =   9600
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   7200
      Left            =   285
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   1
      Top             =   585
      Width           =   9600
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   7200
      Left            =   510
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   5
      Top             =   870
      Width           =   9600
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   7200
      Left            =   720
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   7
      Top             =   1395
      Width           =   9600
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu buildList 
         Caption         =   "Build Image List"
      End
      Begin VB.Menu chooseTarget 
         Caption         =   "Choose Target Image"
      End
      Begin VB.Menu sav 
         Caption         =   "Save"
      End
      Begin VB.Menu destroyList 
         Caption         =   "Destroy image list"
      End
   End
   Begin VB.Menu process 
      Caption         =   "Process"
      Begin VB.Menu rando 
         Caption         =   "Random Combine"
      End
      Begin VB.Menu aver 
         Caption         =   "Random Average"
      End
      Begin VB.Menu subtract 
         Caption         =   "Random Subtract"
      End
      Begin VB.Menu randxor 
         Caption         =   "Random Xor"
      End
      Begin VB.Menu pick2 
         Caption         =   "Pick2 Cell compare"
      End
      Begin VB.Menu allimgCombo 
         Caption         =   "All images 8Cell Combo"
      End
      Begin VB.Menu all9cell 
         Caption         =   "All images 9cell"
      End
      Begin VB.Menu dirPix 
         Caption         =   "Direct Pixel"
      End
      Begin VB.Menu BinProbDist 
         Caption         =   "Binary Prob. Distrib."
      End
      Begin VB.Menu Cen9 
         Caption         =   "Center takes 9"
      End
   End
End
Attribute VB_Name = "Manip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim multipixur() As StdPicture
Dim pixurcount

Private Sub allimgCombo_Click()
multicellComp
End Sub
Private Sub aver_Click()
aveRand
End Sub
Private Sub buildList_Click()
Command1_Click
End Sub

Private Sub Cen9_Click()
'Error 1
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2, tzl
Set Picture1.Picture = Nothing
Dim pickit, x1, y1, bestit
bestit = 1000000
For x = 1 To 639 Step 3
    For y = 1 To 479 Step 3
    scor1 = 0: bestit = 1000000
        For tzl = 1 To pixurcount
            scor1 = 0
            pick1 = tzl
            Set Picture2.Picture = multipixur(pick1 - 1)
            pix1 = Picture2.Point(x, y)
            pix3 = Picture4.Point(x, y) ' target image
            n1 = pix1: 'n2 = pix2
            blu1 = n1 Mod 256: n1 = Int(n1 / 256)
            gren1 = n1 Mod 256: n1 = Int(n1 / 256)
            rd1 = n1
        '    blu2 = n2 Mod 256: n2 = Int(n2 / 256)
        '    gren2 = n2 Mod 256: n2 = Int(n2 / 256)
        '    rd2 = n2
            n2 = pix3
            blu3 = n2 Mod 256: n2 = Int(n2 / 256)
            gren3 = n2 Mod 256: n2 = Int(n2 / 256)
            rd3 = n2
            dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
            scor1 = dblu + dgren + drd
            'dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
            'scor2 = scor2 + dblu + dgren + drd
            'dif = dblu + dgren + drd
            'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
        If scor1 < bestit Then pickit = tzl - 1: bestit = scor1
        Next tzl
        Set Picture3.Picture = multipixur(pickit)
        Debug.Print pickit
        For x1 = -1 To 1
            For y1 = -1 To 1
            Picture1.PSet (x + x1, y + y1), Picture3.Point(x + x1, y + y1)
            Next y1
           Set Picture1.Picture = Picture1.Image
        Next x1
        Picture1.Refresh
        Me.Refresh
        DoEvents
        Next y
     
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x


End Sub

Private Sub chooseTarget_Click()
Command4_Click
End Sub

Private Sub Command1_Click()
On Error Resume Next 'GoTo errhandlr
'CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
Dim fyl As String
fyl = CommonDialog1.FileName
Debug.Print fyl
If fyl = "" Then Exit Sub
'Image1.Picture = LoadPicture(fyl)
Dim pars, tzl As Integer, parsed() As String
pars = Split(fyl, Chr(0)): parsed = pars
If UBound(parsed) = 0 Then
    PiClip1.Picture = LoadPicture(fyl)
    If Err.Number = 481 Then tzl = tzl + 1: GoTo erd
    PiClip1.ClipX = 0: PiClip1.ClipY = 0
    PiClip1.ClipHeight = PiClip1.Height
    PiClip1.ClipWidth = PiClip1.Width
    PiClip1.StretchX = 640
    PiClip1.StretchY = 480
    Picture1.Picture = PiClip1.Clip
    Picture1.Refresh
    ImgPicker.putPic Picture1.Picture, tzl
    ReDim Preserve multipixur(pixurcount)
    Set multipixur(pixurcount) = PiClip1.Clip
    Debug.Print multipixur(pixurcount)
    pixurcount = pixurcount + 1
    Exit Sub
End If
For tzl = 1 To UBound(parsed)
erd:
PiClip1.Picture = LoadPicture(parsed(0) & "\" & parsed(tzl))
If Err.Number = 481 Then tzl = tzl + 1: GoTo erd
PiClip1.ClipX = 0: PiClip1.ClipY = 0
PiClip1.ClipHeight = PiClip1.Height
PiClip1.ClipWidth = PiClip1.Width

PiClip1.StretchX = 640
PiClip1.StretchY = 480
Picture1.Picture = PiClip1.Clip 'Image1.Picture ' LoadPicture(fyl)
Picture1.Refresh
ImgPicker.putPic Picture1.Picture, tzl

'Set srcPic = Picture1.Picture

'goalfylname = CommonDialog1.FileTitle
'FracSearch.Palette = Picture1.Picture
'FracSearch.PaletteMode = vbPaletteModeCustom
'ImageScan
ReDim Preserve multipixur(pixurcount)
Set multipixur(pixurcount) = PiClip1.Clip
Debug.Print multipixur(pixurcount)
pixurcount = pixurcount + 1
'If pixurcount > 1 Then multiplmage = True
Next tzl
'If Not Command2.Enabled Then Command1.Enabled = True
'Command2.Enabled = True
Exit Sub
errhandlr:

End Sub

Public Sub combotron()
' combine most frequent rgb pixel values from every image
' in picture array

End Sub
Public Sub cellComp()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2
Set Picture1.Picture = Nothing
For x = 2 To 639
For y = 2 To 479
pick1 = Rnd * UBound(multipixur)
pick2 = Rnd * UBound(multipixur)

Set Picture2.Picture = multipixur(pick1)
Set Picture3.Picture = multipixur(pick2)
scor1 = 0: scor2 = 0
 For xo = -1 To 1
 For yo = -1 To 1
 If xo <> 0 And yo <> 0 Then
    pix1 = Picture2.Point(x, y)
    pix2 = Picture3.Point(x, y)
    pix3 = Picture4.Point(x, y)
    'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
    'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
    'comcolr1 = colr1: comcolr2 = colr2
    n1 = pix1: n2 = pix2
    blu1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    rd1 = n1
    blu2 = n2 Mod 256: n2 = Int(n2 / 256)
    gren2 = n2 Mod 256: n2 = Int(n2 / 256)
    rd2 = n2
    n2 = pix3
    blu3 = n2 Mod 256: n2 = Int(n2 / 256)
    gren3 = n2 Mod 256: n2 = Int(n2 / 256)
    rd3 = n2
    dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
    scor1 = scor1 + dblu + dgren + drd
    dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
    scor2 = scor2 + dblu + dgren + drd
    'dif = dblu + dgren + drd
    'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
End If
'If scor1 < scor2 Then
'pix4 = Picture2.Point(x, y)
'Else
'pix4 = Picture3.Point(x, y)
'End If
Next yo
Next xo
If scor1 < scor2 Then
pix4 = Picture2.Point(x, y)
Else
pix4 = Picture3.Point(x, y)
End If
Picture1.PSet (x, y), pix4
DoEvents
Next y
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"


End Sub
Public Sub multicellComp()
Dim scors()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2
ReDim scors(UBound(multipixur))
Dim best As Long
Set Picture1.Picture = Nothing
For x = 2 To 639
For y = 2 To 479
'Erase scors
best = 77777777
For pick1 = 1 To UBound(multipixur)
Set Picture2.Picture = multipixur(pick1)
'Set Picture3.Picture = multipixur(pick2)
scor1 = 0: scor2 = 0
 For xo = -1 To 1
 For yo = -1 To 1
 If xo <> 0 And yo <> 0 Then
    pix1 = Picture2.Point(x, y)
    'pix2 = Picture3.Point(x, y)
    pix3 = Picture4.Point(x, y)
    'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
    'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
    'comcolr1 = colr1: comcolr2 = colr2
    n1 = pix1: ' n2 = pix2
    blu1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    rd1 = n1
 '   blu2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   gren2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   rd2 = n2
    n2 = pix3
    blu3 = n2 Mod 256: n2 = Int(n2 / 256)
    gren3 = n2 Mod 256: n2 = Int(n2 / 256)
    rd3 = n2
    dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
    scors(pick1) = scors(pick1) + dblu + dgren + drd
 '   dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
 '   scor2 = scor2 + dblu + dgren + drd
    'dif = dblu + dgren + drd
    'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
End If
'pix3 = RGB(drd, dgren, dblu)
'If scor1 < scor2 Then
'pix4 = Picture2.Point(x, y)
'Else
'pix4 = Picture3.Point(x, y)
'End If
Next yo
Next xo
If scors(pick1) < best Then best = scors(pick1): pix4 = Picture2.Point(x, y)
scors(pick1) = 0
Next pick1
Picture1.PSet (x, y), pix4
DoEvents
Next y
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"
End Sub
Public Sub cell9()
Dim scors()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2
ReDim scors(UBound(multipixur))
Dim best As Long
Set Picture1.Picture = Nothing
For x = 2 To 639
For y = 2 To 479
'Erase scors
best = 77777777
For pick1 = 1 To UBound(multipixur)
Set Picture2.Picture = multipixur(pick1)
'Set Picture3.Picture = multipixur(pick2)
scor1 = 0: scor2 = 0
 For xo = -1 To 1
 For yo = -1 To 1
' If xo <> 0 And yo <> 0 Then
    pix1 = Picture2.Point(x, y)
    'pix2 = Picture3.Point(x, y)
    pix3 = Picture4.Point(x, y)
    'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
    'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
    'comcolr1 = colr1: comcolr2 = colr2
    n1 = pix1: ' n2 = pix2
    blu1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    rd1 = n1
 '   blu2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   gren2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   rd2 = n2
    n2 = pix3
    blu3 = n2 Mod 256: n2 = Int(n2 / 256)
    gren3 = n2 Mod 256: n2 = Int(n2 / 256)
    rd3 = n2
    dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
    scors(pick1) = scors(pick1) + dblu + dgren + drd
 '   dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
 '   scor2 = scor2 + dblu + dgren + drd
    'dif = dblu + dgren + drd
    'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
'End If
'pix3 = RGB(drd, dgren, dblu)
'If scor1 < scor2 Then
'pix4 = Picture2.Point(x, y)
'Else
'pix4 = Picture3.Point(x, y)
'End If
Next yo
Next xo
If scors(pick1) < best Then best = scors(pick1): pix4 = Picture2.Point(x, y)
scors(pick1) = 0
Next pick1
Picture1.PSet (x, y), pix4
DoEvents
Next y
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"

End Sub
Public Sub directPixel()
Dim scors()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2
ReDim scors(UBound(multipixur))
Dim best As Long
Set Picture1.Picture = Nothing
For x = 2 To 639
For y = 2 To 479
'Erase scors
best = 97777777
For pick1 = 1 To UBound(multipixur)
Set Picture2.Picture = multipixur(pick1)
'Set Picture3.Picture = multipixur(pick2)
scor1 = 0
     pix1 = Picture2.Point(x, y)
    'pix2 = Picture3.Point(x, y)
    pix3 = Picture4.Point(x, y)
    'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
    'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
    'comcolr1 = colr1: comcolr2 = colr2
    n1 = pix1: ' n2 = pix2
    blu1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    rd1 = n1
 '   blu2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   gren2 = n2 Mod 256: n2 = Int(n2 / 256)
 '   rd2 = n2
    n2 = pix3
    blu3 = n2 Mod 256: n2 = Int(n2 / 256)
    gren3 = n2 Mod 256: n2 = Int(n2 / 256)
    rd3 = n2
    dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
    scor1 = dblu + dgren + drd
 '   dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
 '   scor2 = scor2 + dblu + dgren + drd
    'dif = dblu + dgren + drd
    'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
If scor1 < best Then best = scors(pick1): pix4 = Picture2.Point(x, y)

Next pick1
Picture1.PSet (x, y), pix4
DoEvents
Next y
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"

End Sub

Public Sub analyse()
' call different analysis routines and display their results
End Sub
Public Sub randomoid()
' take a pixel from a randomly chosen image for each pixel
' in the output image
Dim x, y, pick, pixl As Long
Set Picture1.Picture = Nothing
For x = 1 To 640
For y = 1 To 480
pick = Rnd * UBound(multipixur)
Set Picture2.Picture = multipixur(pick)
 pixl = Picture2.Point(x, y)
Picture1.PSet (x, y), pixl
Next y
DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x

End Sub
Public Sub aveRand()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double

Set Picture1.Picture = Nothing
For x = 1 To 639
For y = 1 To 479
pick1 = Rnd * UBound(multipixur)
pick2 = Rnd * UBound(multipixur)

Set Picture2.Picture = multipixur(pick1)
Set Picture3.Picture = multipixur(pick2)
 pix1 = Picture2.Point(x, y)
 pix2 = Picture3.Point(x, y)

colr1 = pix1
colr2 = pix2
'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
'comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
'dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
'dif = dblu + dgren + drd
dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
pix3 = RGB(drd, dgren, dblu)
Picture1.PSet (x, y), pix3
DoEvents
Next y
DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"


End Sub
Public Sub subrand()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double

Set Picture1.Picture = Nothing
For x = 1 To 639
For y = 1 To 479
pick1 = Rnd * UBound(multipixur)
pick2 = Rnd * UBound(multipixur)

Set Picture2.Picture = multipixur(pick1)
Set Picture3.Picture = multipixur(pick2)
 pix1 = Picture2.Point(x, y)
 pix2 = Picture3.Point(x, y)

colr1 = pix1
colr2 = pix2
'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
'comcolr1 = colr1: comcolr2 = colr2
n1 = colr1: n2 = colr2
blu1 = n1 Mod 256: n1 = Int(n1 / 256)
gren1 = n1 Mod 256: n1 = Int(n1 / 256)
rd1 = n1
blu2 = n2 Mod 256: n2 = Int(n2 / 256)
gren2 = n2 Mod 256: n2 = Int(n2 / 256)
rd2 = n2
dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
'dif = dblu + dgren + drd
'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
pix3 = RGB(drd, dgren, dblu)
Picture1.PSet (x, y), pix3
DoEvents
Next y
DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
'SavePicture Picture1.Picture, "doogle1.bmp"
End Sub
Public Sub XorRand()
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long
Dim dif As Double, totdif As Double

Set Picture1.Picture = Nothing
For x = 1 To 639
For y = 1 To 479
pick1 = Rnd * UBound(multipixur)
pick2 = Rnd * UBound(multipixur)

Set Picture2.Picture = multipixur(pick1)
Set Picture3.Picture = multipixur(pick2)
 pix1 = Picture2.Point(x, y)
 pix2 = Picture3.Point(x, y)
'
'colr1 = pix1
'colr2 = pix2
''If colr1 = comcolr1 Then compixel1 = compixel1 + 1
''If colr2 = comcolr2 Then compixel2 = compixel2 + 1
''comcolr1 = colr1: comcolr2 = colr2
'n1 = colr1: n2 = colr2
'blu1 = n1 Mod 256: n1 = Int(n1 / 256)
'gren1 = n1 Mod 256: n1 = Int(n1 / 256)
'rd1 = n1
'blu2 = n2 Mod 256: n2 = Int(n2 / 256)
'gren2 = n2 Mod 256: n2 = Int(n2 / 256)
'rd2 = n2
'dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
'dif = dblu + dgren + drd
'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2


pix3 = pix1 Xor pix2 'RGB(drd, dgren, dblu)
Picture1.PSet (x, y), pix3
DoEvents
Next y
DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Set Picture1.Picture = Picture1.Image
SavePicture Picture1.Picture, "doogle2.bmp"

End Sub


Private Sub Command2_Click()
'randomoid
'averand
' subrand
'XorRand
'cellComp
multicellComp

End Sub

Private Sub Command3_Click()
Dim fyl
'fyl = InputBox("Name")
CommonDialog2.Filter = "Pictures (*.bmp)|*.bmp"

CommonDialog2.ShowSave
fyl = CommonDialog2.FileName
If fyl = "" Then Exit Sub
If LCase(Right(fyl, 3)) <> "bmp" Then fyl = fyl & ".bmp"
Set Picture1.Picture = Picture1.Image
SavePicture Picture1.Picture, fyl

End Sub

Private Sub Command4_Click()
On Error GoTo errhandlr
'CommonDialog1.InitDir = App.Path
CommonDialog2.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
CommonDialog2.CancelError = True
CommonDialog2.ShowOpen
Dim fyl As String
fyl = CommonDialog2.FileName
If fyl = "" Then Exit Sub
'Image1.Picture = LoadPicture(fyl)
PiClip1.Picture = LoadPicture(fyl)
PiClip1.ClipX = 0: PiClip1.ClipY = 0
PiClip1.ClipHeight = PiClip1.Height
PiClip1.ClipWidth = PiClip1.Width

PiClip1.StretchX = 640
PiClip1.StretchY = 480
'Picture1.Picture = PiClip1.Clip

Picture4.Picture = PiClip1.Clip 'Image1.Picture ' LoadPicture(fyl)
Picture4.Refresh

'Set srcPic = Picture1.Picture
'goalfylname = CommonDialog2.FileTitle
'FracSearch.Palette = Picture1.Picture
'FracSearch.PaletteMode = vbPaletteModeCustom
'ImageScan
'If Not Command2.Enabled Then Command1.Enabled = True
'Command2.Enabled = True
Exit Sub
errhandlr:
End Sub

Private Sub Command5_Click()
ImgPicker.Show
End Sub

Private Sub Command6_Click()
ReDim multipixur(1)
pixurcount = 0
ImgPicker.flushGrid
End Sub

Private Sub dirPix_Click()
' chopped from pick2, not transformed
'Error 1
Dim x, y, pick1, pick2, pix1 As Long, pix2 As Long, pix3 As Long, pix4 As Long
Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
Dim blu3 As Integer, gren3 As Integer, rd3 As Integer
Dim dblu As Integer, dgren As Integer, drd As Integer
Dim t As Integer, colr1 As Long, colr2 As Long, colr3 As Long
Dim dif As Double, totdif As Double
Dim xo, yo, scor1, scor2, tzl
Set Picture1.Picture = Nothing
For tzl = 1 To pixurcount Step 2
For x = 2 To 639
For y = 2 To 479
pick1 = tzl 'Rnd * UBound(multipixur)
pick2 = tzl + 1 'Rnd * UBound(multipixur)
Set Picture2.Picture = multipixur(pick1)
Set Picture3.Picture = multipixur(pick2)
scor1 = 0: scor2 = 0
    pix1 = Picture2.Point(x, y)
    pix2 = Picture3.Point(x, y)
    pix3 = Picture4.Point(x, y)
    'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
    'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
    'comcolr1 = colr1: comcolr2 = colr2
    n1 = pix1: n2 = pix2
    blu1 = n1 Mod 256: n1 = Int(n1 / 256)
    gren1 = n1 Mod 256: n1 = Int(n1 / 256)
    rd1 = n1
    blu2 = n2 Mod 256: n2 = Int(n2 / 256)
    gren2 = n2 Mod 256: n2 = Int(n2 / 256)
    rd2 = n2
    n2 = pix3
    blu3 = n2 Mod 256: n2 = Int(n2 / 256)
    gren3 = n2 Mod 256: n2 = Int(n2 / 256)
    rd3 = n2
    dblu = Abs(blu1 - blu3): dgren = Abs(gren1 - gren3): drd = Abs(rd1 - rd3)
    scor1 = scor1 + dblu + dgren + drd
    dblu = Abs(blu2 - blu3): dgren = Abs(gren2 - gren3): drd = Abs(rd2 - rd3)
    scor2 = scor2 + dblu + dgren + drd
    'dif = dblu + dgren + drd
    'dblu = (blu1 + blu2) / 2: dgren = (gren1 + gren2) / 2: drd = (rd1 + rd2) / 2
If scor1 < scor2 Then
pix4 = Picture2.Point(x, y)
Else
pix4 = Picture3.Point(x, y)
End If
Picture1.PSet (x, y), pix4
DoEvents
Next y
'DoEvents
'Picture1.Picture = Picture1.Image
'Me.Refresh
Next x
Next tzl
Set Picture1.Picture = Picture1.Image
End Sub

Private Sub pick2_Click()
cellComp
End Sub

Private Sub rando_Click()
randomoid
End Sub

Private Sub randxor_Click()
XorRand
End Sub

Private Sub sav_Click()
Command3_Click
End Sub

Private Sub subtract_Click()
subrand
End Sub
