VERSION 5.00
Begin VB.Form pixter 
   BorderStyle     =   0  'None
   Caption         =   "Manual Scan"
   ClientHeight    =   3000
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "pixter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rank As Integer
Private Sub Form_Click()
rank = CInt(InputBox("Enter rank:"))
End Sub

Private Sub Form_Load()

End Sub
