VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      Height          =   855
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   11040
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&GENERATE"
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   11040
      Width           =   2535
   End
   Begin VB.TextBox ls 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14520
      TabIndex        =   28
      Top             =   9480
      Width           =   3735
   End
   Begin VB.TextBox lm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   27
      Top             =   9480
      Width           =   1695
   End
   Begin VB.TextBox hs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14520
      TabIndex        =   24
      Top             =   8160
      Width           =   3735
   End
   Begin VB.TextBox hm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   22
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox avg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14520
      TabIndex        =   20
      Top             =   6600
      Width           =   3735
   End
   Begin VB.TextBox total 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   18
      Top             =   6600
      Width           =   3735
   End
   Begin VB.TextBox marks6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16560
      TabIndex        =   16
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox marks5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   13
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox marks4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox marks3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16560
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox marks2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox marks1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox roll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14520
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox name1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "LOWEST SUBJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   26
      Top             =   9480
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "LOWEST MARKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   25
      Top             =   9480
      Width           =   3375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "HIGHEST SUBJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   23
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "HIGHEST MARKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   21
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "AVERAGE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   19
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   17
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label sub6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "HISTORY:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      TabIndex        =   15
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label sub5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "ENGLISH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   14
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label sub4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "COMP SCI. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   11
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label sub3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "CHEM:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label sub2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "PHYSICS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label sub1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "MATHS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "ROLL NUMBER:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "MARK SHEET GENERATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()
Dim marks(5) As Integer
Dim subj(5) As String

total.Text = Val(marks1.Text) + Val(marks2.Text) + Val(marks3.Text) + Val(marks4.Text) + Val(marks5.Text) + Val(marks6.Text)
avg.Text = Val(total.Text) / 6

marks(0) = marks1.Text
marks(1) = marks2.Text
marks(2) = marks3.Text
marks(3) = marks4.Text
marks(4) = marks5.Text
marks(5) = marks6.Text

subj(0) = "Maths"
subj(1) = "Physics"
subj(2) = "Chemistry"
subj(3) = "Computer"
subj(4) = "Englsih"
subj(5) = "History"


highest = marks1.Text
For i = 1 To 5
    If marks(i) > highest Then
        highest = marks(i)
        p = i
    End If
Next
hm.Text = highest
hs.Text = subj(p)





lowest = marks1.Text
For i = 1 To 5
    If marks(i) < lowest Then
        lowest = marks(i)
        q = i
    End If
Next
lm.Text = lowest
ls.Text = subj(q)



End Sub

Private Sub Command2_Click()
name1.Text = ""
roll.Text = ""
marks1.Text = ""
marks2.Text = ""
marks3.Text = ""
marks4.Text = ""
marks5.Text = ""
marks6.Text = ""
total.Text = ""
avg.Text = ""
hm.Text = ""
hs.Text = ""
lm.Text = ""
ls.Text = ""
End Sub
