VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MapForm 
   Caption         =   "MapMaster"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   8025
      Left            =   165
      TabIndex        =   4
      Top             =   90
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   14155
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"MapForm.frx":0000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   510
      Left            =   5295
      TabIndex        =   3
      Top             =   2235
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   585
      Left            =   5160
      TabIndex        =   2
      Top             =   1305
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5205
      TabIndex        =   1
      Top             =   795
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   5220
      TabIndex        =   0
      Top             =   45
      Width           =   660
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu open 
         Caption         =   "open"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu sort 
         Caption         =   "Sort"
      End
   End
End
Attribute VB_Name = "MapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tzl, yzl
Debug.Print App.Path
Open App.Path & "\bnw2.map" For Output As #1
For tzl = 255 To -1 Step -8
If tzl = -1 Then Exit For
For yzl = 0 To 6

Print #1, tzl & "   " & tzl & "   " & tzl
Next yzl
Next tzl
For yzl = 0 To 6
Print #1, "0   0   0"
Next yzl

Close 1
End Sub

Private Sub Command2_Click()
Open App.Path & "\bnw3.map" For Output As #1
For tzl = 0 To 255
If tzl Mod 2 Then
Print #1, "0   0   0"
Else
Print #1, "255 255 255"
End If
Next tzl
End Sub

Private Sub Command3_Click()
Open App.Path & "\bnw4.map" For Output As #1
For yzl = 0 To 3
For tzl = 0 To 31
Print #1, "0   0   0"
Next tzl
For tzl = 0 To 31
Print #1, "255 255 255"
Next tzl
Next yzl
Close 1
End Sub

Private Sub Command4_Click()
Open App.Path & "\bnw4.map" For Output As #1
Print #1, "0   0   0"
Print #1, "255   255   255"
For tzl = 0 To 255
Next tzl

Close 1
End Sub

Private Sub open_Click()
CommonDialog1.ShowOpen
Dim fyl, pit, allofit$
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub
'Open fyl For Input As #1
'Do While Not EOF(1)
'Get #1, , pit
'allofit = allofit & pit
'Loop
rtf1.LoadFile fyl, 1

End Sub

'Sub sortit()
'Debug.Print "sort"
' gap = 128 'Int(300 / 2)
'  Do While gap >= 1
'   Do
'   doneflag = 1
'    For Index = 1 To 128 '300 - gap
'     If sc(Index) < sc(Index + gap) Then
'     tmpsc = sc(Index)
'     sc(Index) = sc(Index + gap)
'     sc(Index + gap) = tmpsc
'     tmpb$ = a5$(Index)
'     a5$(Index) = a5$(Index + gap)
'     a5$(Index + gap) = tmpb$
'
'
'      doneflag = 0
'     End If
'    Next Index
'   Loop Until doneflag = 1
'   gap = Int(gap / 2)
'  Loop
'Beep
'
'End Sub
'
'Sub redsort()
'Debug.Print "sort"
' gap = 128 'Int(300 / 2)
'  Do While gap >= 1
'   Do
'   doneflag = 1
'    For Index = 1 To 128 '300 - gap
'     If sc(Index) < sc(Index + gap) Then
'     tmpsc = sc(Index)
'     sc(Index) = sc(Index + gap)
'     sc(Index + gap) = tmpsc
'     tmpb$ = a5$(Index)
'     a5$(Index) = a5$(Index + gap)
'     a5$(Index + gap) = tmpb$
'
'
'      doneflag = 0
'     End If
'    Next Index
'   Loop Until doneflag = 1
'   gap = Int(gap / 2)
'  Loop
'Beep
'
'End Sub
'
'
