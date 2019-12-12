Attribute VB_Name = "Module2"
'INPUT "Real Constant ? ", cx
'INPUT "Imaginary Constant ? ", cy
'Screen 9
'For a = 0 To 100 Step 0.5
'  For b = 0 To 100 Step 0.5
Public Function juliapixel(param1 As Double, param2 As Double, x As Double, y As Double, stip As Double) As Long

    x0 = -2 + x / stip
    y0 = 2 - y / stip
    
    For I = 1 To 20
      X1 = x0 * x0 - y0 * y0 + cx
      Y1 = 2 * x0 * y0 + cy
      If X1 ^ 2 + Y1 ^ 2 > 4 Then GoTo 180
      x0 = X1
      y0 = Y1
    Next I
  col = (X1 ^ 2 + Y1 ^ 2) * (255 / 4)
 
180:

juliapixel = col
End Function
