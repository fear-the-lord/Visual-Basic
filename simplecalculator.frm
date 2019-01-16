VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      TabIndex        =   12
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton RESET 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      Height          =   735
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton DIVIDE 
      Caption         =   "&DIVIDE"
      Height          =   735
      Left            =   10920
      TabIndex        =   9
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton MULTIPLY 
      Caption         =   "&MULTIPLY"
      Height          =   735
      Left            =   8760
      TabIndex        =   8
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton SUBTRACT 
      Caption         =   "&SUBTRACT"
      Height          =   735
      Left            =   6480
      TabIndex        =   7
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OPERATIONS"
      Height          =   2175
      Left            =   3600
      TabIndex        =   5
      Top             =   8760
      Width           =   11655
      Begin VB.CommandButton ADD 
         Caption         =   "&ADD"
         Height          =   735
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "RESULT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   11
      Top             =   7200
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER SECOND NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   5160
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER FIRST NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SIMPLE CALCULATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5760
      TabIndex        =   0
      Top             =   720
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADD_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)

End Sub

Private Sub DIVIDE_Click()
Text3.Text = Val(Text1.Text) / Val(Text2.Text)

End Sub

Private Sub MULTIPLY_Click()
Text3.Text = Val(Text1.Text) * Val(Text2.Text)

End Sub

Private Sub RESET_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub SUBTRACT_Click()
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub
