VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EditDB 
   Caption         =   "Editor"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2475
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7005
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Pop stats"
      Height          =   600
      Left            =   105
      TabIndex        =   1
      Top             =   6645
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6525
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   11509
      _Version        =   393216
      Rows            =   100
      Cols            =   16
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   16761024
      BackColorSel    =   8438015
      GridColor       =   8421631
      GridLines       =   3
      FormatString    =   $"edit2.frx":0000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6525
      Index           =   1
      Left            =   450
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   11509
      _Version        =   393216
      Rows            =   100
      Cols            =   15
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   16761024
      BackColorSel    =   8438015
      GridColor       =   8421631
      GridLines       =   3
      FormatString    =   $"edit2.frx":00D9
   End
End
Attribute VB_Name = "EditDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bDoNotEdit   As Boolean
Dim bOnFixedPart As Boolean
Dim fritz As everybody
Dim pops(100) As popmem

Private Sub Command1_Click()
Static editingcurrent As Boolean
If editingcurrent Then
Me.Caption = "Editing current population"
MSFlexGrid1(0).Visible = True
MSFlexGrid1(1).Visible = False
Command1.Caption = "Edit previous pop"
Else
Me.Caption = "Editing previous population"
MSFlexGrid1(0).Visible = False
MSFlexGrid1(1).Visible = True
Command1.Caption = "Edit current pop"
End If

editingcurrent = Not editingcurrent
End Sub

Private Sub Form_Click()
Call pSetCellValue
End Sub

Private Sub Form_Load()
'MSFlexGrid1.Cols = 15
'MSFlexGrid1.Rows = 100
Dim i As Long
'
' Set the grid and textbox
' to the same font.
'
With txtEdit.Font
    .Name = MSFlexGrid1(0).Font.Name
    .Size = MSFlexGrid1(0).Font.Size
    .Weight = MSFlexGrid1(0).Font.Weight
End With
txtEdit.BackColor = vb3DLight
'
' Add some rows and columns to the grid so we
' have something to start with.
'
With MSFlexGrid1(0)
    .RowHeightMin = txtEdit.Height

    ' Size the first fixed column.
    .ColWidth(0) = .ColWidth(0) / 2
    .ColAlignment(0) = 1   ' Center center.
    
'    ' Label the rows.
    For i = .FixedRows To .Rows - 1
         .TextArray(fLabel(i, 0)) = i
    Next
'
'    ' Label the columns.
'    For i = .FixedCols To .Cols - 1
'        .TextArray(fLabel(0, i)) = i
'    Next
    
    ' Right align data.
    For i = .FixedCols To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With

txtEdit = ""
bDoNotEdit = False
End Sub

Public Sub loadmember(index As Integer, p1 As Double, p2 As Double, _
mg As Double, cx As Double, cy As Double _
, inv As Double, icx As Double, icy As Double, bo As Long, bi As Long _
, dc As Long, mx As Double, rt As Double, sk As Double, xm As Double)



End Sub

Private Function fLabel(lRow As Long, lCol As Long) As Long
    fLabel = lCol + MSFlexGrid1(0).Cols * lRow
End Function
Public Sub loadGrid()
Dim member As popmem
Dim y As Integer
For y = 1 To 99
member = everyone.pop(y)
MSFlexGrid1(0).row = y
'For x = 1 To 15
MSFlexGrid1(0).col = 1

MSFlexGrid1(0).Text = member.param1
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.param2
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.mag
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.centerx
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.centery
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.invert
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.icenterx
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.icentery
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.bailout
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.decomp
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.biomorph
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.maxiter
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.rot
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.skew
MSFlexGrid1(0).col = MSFlexGrid1(0).col + 1
MSFlexGrid1(0).Text = member.xmag
'Next x
Next y
End Sub

Private Sub MSFlexGrid1_Click(index As Integer)
Dim query
If FracSearch.running Then query = MsgBox("Changing these values will force a restart" & vbCrLf & "Do You wish to proceed?", vbYesNoCancel)
If query = 1 Then Exit Sub
If bOnFixedPart Then Exit Sub
Call pEditGrid(32)

End Sub

Private Sub MSFlexGrid1_GotFocus(index As Integer)
If bDoNotEdit Then Exit Sub
'
' Copy the textbox's value to the grid
' and hide the textbox.
'
Call pSetCellValue
End Sub

Private Sub MSFlexGrid1_KeyPress(index As Integer, KeyAscii As Integer)
'
' Display the textbox.
'
Call pEditGrid(KeyAscii)
End Sub

Private Sub MSFlexGrid1_LeaveCell(index As Integer)
If bDoNotEdit Then Exit Sub
Call MSFlexGrid1_GotFocus(0)

End Sub

Private Sub MSFlexGrid1_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim l      As Long
Dim lWidth As Long
Dim col2change
With MSFlexGrid1(0)
    For l = 0 To .Cols - 1
        If .ColIsVisible(l) Then
            lWidth = lWidth + .ColWidth(l)
        End If
    Next
    '
    ' See if we are on the fixed part of the grid.
    '
    bOnFixedPart = (x < .ColWidth(0)) Or _
                   (x > lWidth) Or _
                   (y < .RowHeight(0)) Or _
                   (y > .Rows * .RowHeightMin)
End With
If MSFlexGrid1(0).row = 0 Then
' place total column control here
col2change = MSFlexGrid1(0).col
Dim query
If Shift Then
'query = InputBox("amount to randomize")
'Dim tzl
'For tzl = 0 To MSFlexGrid1(0).Rows
'MSFlexGrid1(0).Text = CStr(CDbl(MSFlexGrid1(0).Text + (Rnd * query)))
'With everyone
'Select Case MSFlexGrid1(0).TextArray(fLabel(0, MSFlexGrid1(0).col))
'Case "param1"
'.pop.param1 = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "param2"
'.pop.param2 = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "mag"
'.pop.mag = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "centerx"
'.pop.centerx = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "centery"
'.pop.centery = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "invert"
'.pop.invert = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "icenterx"
'.pop.icenterx = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "icentery"
'.pop.icentery = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "bailout"
'.pop.bailout = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "decomp"
'.pop.decomp = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "maxiter"
'.pop.maxiter = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "biomorph"
'.pop.biomorph = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'Case "rotation"
'.pop.rot = CDbl(MSFlexGrid1(0).Text + (Rnd * query))
'End Select
'
End If
End If



End Sub

Private Sub MSFlexGrid1_RowColChange(index As Integer)
Dim colum As Integer, roe As Integer
colum = MSFlexGrid1(0).col
roe = MSFlexGrid1(0).row

End Sub

Private Sub MSFlexGrid1_Scroll(index As Integer)
Call MSFlexGrid1_GotFocus(0)

End Sub

Private Sub MSFlexGrid1_SelChange(index As Integer)
If MSFlexGrid1(0).row = 0 Then
' allow handling of entire columns
End If
End Sub
Private Sub pEditGrid(KeyAscii As Integer)
'
' Populate the textbox and position it.
'
With txtEdit
    Select Case KeyAscii
        Case 0 To 32
            '
            ' Edit the current text.
            '
            .Text = MSFlexGrid1(0)
            .SelStart = 0
            .SelLength = 1000
            
        Case 8, 46, 48 To 57
            '
            ' Replace the current text but only
            ' if the user entered a number.
            '
            .Text = Chr(KeyAscii)
            .SelStart = 1
        Case Else
            '
            ' If an alpha character was entered,
            ' use a zero instead.
            '
            .Text = "0"
    End Select
End With
'
' Show the textbox at the right place.
'
With MSFlexGrid1(0)
    If .CellWidth < 0 Then Exit Sub
    txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
    '
    ' NOTE:
    '   Depending on the style of the Grid Lines that you set, you
    '   may need to adjust the textbox position slightly. For example
    '   if you use raised grid lines use the following:
    '
    'txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
End With

txtEdit.Visible = True
txtEdit.SetFocus
End Sub
Private Sub pSetCellValue()
'
' NOTE:
'       This code should be called anytime
'       the grid loses focus and the grid's
'       contents may change.  Otherwise, the
'       cell's new value may be lost and the
'       textbox may not line up correctly.
'
If txtEdit.Visible Then
    MSFlexGrid1(0).Text = txtEdit.Text
    txtEdit.Visible = False
End If


End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
With MSFlexGrid1(0)
    Select Case KeyCode
        Case 13   'ENTER
            .SetFocus
        Case 27   'ESC
             txtEdit.Visible = False
            .SetFocus
        Case 38   'Up arrow
            .SetFocus
            DoEvents
            If .row > .FixedRows Then
                bDoNotEdit = True
                .row = .row - 1
                bDoNotEdit = False
            End If
        Case 40   'Down arrow
            .SetFocus
            DoEvents
            If .row < .Rows - 1 Then
                bDoNotEdit = True
                .row = .row + 1
                bDoNotEdit = False
            End If
    End Select
End With
End Sub
