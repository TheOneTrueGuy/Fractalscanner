VERSION 5.00
Begin VB.Form Edit 
   Caption         =   "Edit"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1710
      TabIndex        =   31
      Text            =   "Text13"
      Top             =   5505
      Width           =   1485
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2130
      TabIndex        =   29
      Text            =   "Text12"
      Top             =   5040
      Width           =   1170
   End
   Begin VB.TextBox Text11 
      Height          =   330
      Left            =   1320
      TabIndex        =   27
      Text            =   "Text11"
      Top             =   4575
      Width           =   1380
   End
   Begin VB.TextBox Text10 
      Height          =   300
      Left            =   1440
      TabIndex        =   25
      Text            =   "Text10"
      Top             =   4170
      Width           =   1290
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   2145
      TabIndex        =   23
      Text            =   "Text9"
      Top             =   3675
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2145
      TabIndex        =   22
      Text            =   "Text8"
      Top             =   3225
      Width           =   1755
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1740
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   2730
      Width           =   2010
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   5265
      TabIndex        =   17
      Top             =   7395
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   3870
      TabIndex        =   16
      Top             =   7365
      Width           =   1380
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   4
      Left            =   4065
      TabIndex        =   15
      Text            =   "1"
      Top             =   2295
      Width           =   795
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   3
      Left            =   4065
      TabIndex        =   14
      Text            =   "1"
      Top             =   1830
      Width           =   795
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   2
      Left            =   4065
      TabIndex        =   13
      Text            =   "1"
      Top             =   1365
      Width           =   795
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   1
      Left            =   4065
      TabIndex        =   12
      Text            =   "1"
      Top             =   885
      Width           =   795
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   0
      Left            =   4065
      TabIndex        =   11
      Text            =   "1"
      Top             =   435
      Width           =   795
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1065
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   2262
      Width           =   2595
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1065
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1779
      Width           =   2595
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1065
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1296
      Width           =   2595
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1065
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   858
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1065
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   420
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "Maxiter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   90
      TabIndex        =   30
      Top             =   5535
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Biomorph color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   60
      TabIndex        =   28
      Top             =   5025
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "X mag:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   150
      TabIndex        =   26
      Top             =   4590
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Rotation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Invert Center Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   105
      TabIndex        =   21
      Top             =   3705
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Invert Center X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   90
      TabIndex        =   20
      Top             =   3240
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Invert Ratio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   135
      TabIndex        =   18
      Top             =   2715
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "Number of pixels to move (- for left + for right)"
      Height          =   210
      Left            =   2910
      TabIndex        =   10
      Top             =   45
      Width           =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "CenterY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   2265
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "CenterX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Mag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Param2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Param1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   930
   End
End
Attribute VB_Name = "Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
