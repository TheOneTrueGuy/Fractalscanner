Attribute VB_Name = "Module1"
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long


Public Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End Type
Public Declare Sub GlobalMemoryStatus _
Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Type quadpop
a(12) As Double
score As Long
End Type
Type qPoper
qpop(100) As quadpop
End Type



Public Type popmem
param1 As Double
param2 As Double
centerx As Double
centery As Double
mag As Double
invert As Double
icenterx As Double
icentery As Double
rot As Double
xmag As Double
biomorph As Integer ' 0-255
maxiter As Long ' 0-64k
decomp As Integer '0-16k
bailout As Long ' 0-64k
skew As Double

End Type


'Public Type popmem2
'param1 As Double
'param2 As Double
'cornerx As Double
'cornery As Double
'increm As Double
'iters As Long
'End Type
'
'Public Type fractal
'fractype As String
'params() As Double
'numparams As Integer
'xmin As Double
'xmax As Double
'ymin As Double
'ymax As Double
'biomorph As Integer
'decomp As Integer
'bailout As Long
'maxiter As Long
'functions() As String
'beenTested As Boolean
'score As Double
'End Type
'
'Public Type fracpop
'fractal(101) As fractal
'biomorph As Boolean
'decomp As Boolean
'bailout As Boolean
'maxiter As Boolean
'rxmax As Double
'rxmin As Double
'rymax As Double
'rymin As Double
'
'End Type

'Now, add this code to get the values:
'Dim MS As MEMORYSTATUS
'MS.dwLength = Len(MS)
'GlobalMemoryStatus MS


'MS.dwMemoryLoad contains percentage memory used
'MS.dwTotalPhys contains total amount of physical memory in bytes
'MS.dwAvailPhys contains available physical memory
'MS.dwTotalPageFile contains total amount of memory in the page file
'MS.dwAvailPageFile contains available amount of memory in the page file
'MS.dwTotalVirtual contains total amount of virtual memory
'MS.dwAvailVirtual contains available virtual memory

Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Public Type everybody
pop(101) As popmem
gen As Long
fractype As String
icolor As String
ocolor As String
palmap As String
scantype As String
mutrate As Double
evbail As Boolean
evinvert As Boolean
evxmag As Boolean
evbio As Boolean
evmaxiter As Boolean
evdecomp As Boolean
evrot As Boolean
evskew As Boolean
numparams As Integer
region As Variant
End Type
Public everyone As everybody


Public Sub Main()

End Sub
