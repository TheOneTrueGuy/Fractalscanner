Attribute VB_Name = "Module3"
Public Function mandelpixel(x As Double, y As Double, stip As Double) As Long

  cx = -2 * x / stip
  cy = 2 - y / stip
    x0 = 0
    y0 = 0
    it = 0
    For I = 1 To 20
      X1 = x0 ^ 2 - y0 ^ 2 + cx
      Y1 = 2 * x0 * y0 + cy
      If X1 ^ 2 + Y1 ^ 2 > 4 Then
        it = 1
        GoTo out
      End If
      x0 = X1
      y0 = Y1
    Next I
    col = (X1 * X1 + Y1 * Y1)
'    PSet (a * 6.4, b * 4.8), col + 1
mandelpixel = col + 1
out:
'If it = 1 Then PSet (a * 6.4, b * 4.8), 12
mandelpixel = col + 1

End Function
'Public Sub compare1(num As Integer)
'On Error GoTo errhandl
'Dim compixel1 As Long, compixel2 As Long
'Dim comcolr1 As Long, comcolr2 As Long
'Dim colr1a As Long, colr2a As Long
'If Dir(App.Path & "\scan.gif") = "" Then Exit Sub
''Picture2.Picture = Null
'Picture2.Picture = LoadPicture(App.Path & "\scan.gif")
'Dim n1 As Long, n2 As Long
'Dim blu1 As Integer, gren1 As Integer, rd1 As Integer
'Dim blu2 As Integer, gren2 As Integer, rd2 As Integer
'Dim dblu As Integer, dgren As Integer, drd As Integer
'Dim t As Integer, x As Integer, y As Integer, colr1 As Long, colr2 As Long
'Dim dif As Double, totdif As Double
'For t = 1 To 4000
'x = Rnd * 320
'y = Rnd * 200
''For X = 1 To 320 Step 2
''For Y = 1 To 200 Step 2
'colr1 = Picture1.Point(x, y)
'colr2 = Picture2.Point(x, y)
'If colr1 = comcolr1 Then compixel1 = compixel1 + 1
'If colr2 = comcolr2 Then compixel2 = compixel2 + 1
'comcolr1 = colr1: comcolr2 = colr2
'n1 = colr1: n2 = colr2
'blu1 = n1 Mod 256: n1 = Int(n1 / 256)
'gren1 = n1 Mod 256: n1 = Int(n1 / 256)
'rd1 = n1
'blu2 = n2 Mod 256: n2 = Int(n2 / 256)
'gren2 = n2 Mod 256: n2 = Int(n2 / 256)
'rd2 = n2
'dblu = Abs(blu1 - blu2): dgren = Abs(gren1 - gren2): drd = Abs(rd1 - rd2)
'dif = dblu + dgren + drd
''If colr1 = 0 And colr2 = 0 Then dif = 1: GoTo skipZero
''If colr1 > colr2 Then dif = colr2 / colr1 Else dif = colr1 / colr2
''skipZero:
'totdif = totdif + dif
''Next Y
''Next X
'Next t
'score(num) = totdif + ((compixel2 - compixel1) * 100)
'If score(num) < bestscore Then bestscore = score(num): best = population(num): Text1.Text = param(num): Label3.Caption = "Best Score:" & CStr(bestscore)
'If score(num) < curbest Then curbest = score(num): Picture3.Picture = Picture2.Picture: Label5.Caption = "Best Current:" & CStr(score(num)): Text2.Text = param(num)
'
'Label1.Caption = CStr(curnum) & ": " & CStr(score(num))
'Kill App.Path & "\scan.gif"
'Debug.Print totdif
'Exit Sub
'errhandl:
'
'End Sub
