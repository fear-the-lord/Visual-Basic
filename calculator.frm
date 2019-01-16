VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   30
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   29
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command28 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   28
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command27 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   27
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   26
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   25
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   23
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   20
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   19
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b, sign, result As Double
Dim mem As Double

Private Sub Command1_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 9

End Sub

Private Sub Command10_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 0

End Sub

Private Sub Command11_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 1

End Sub

Private Sub Command12_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 2

End Sub

Private Sub Command13_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 3

End Sub

Private Sub Command14_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 4

End Sub

Private Sub Command15_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 5

End Sub

Private Sub Command16_Click()
Dim c As Double
b = Val(Label1.Caption)
If sign = 1 Then
    c = a + b
ElseIf sign = 2 Then
    c = a - b
ElseIf sign = 3 Then
    c = a / b
ElseIf sign = 4 Then
    c = a * b
ElseIf sign = 5 Then
    c = a ^ b
Else
    c = (a / 100) * b
End If
Label1.Caption = c

End Sub

Private Sub Command17_Click()
Label1.Caption = ""
a = 0
b = 0
result = 0
End Sub

Private Sub Command18_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
result = Sin(a)
Label1.Caption = result

End Sub

Private Sub Command19_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
result = Cos(a)
Label1.Caption = result

End Sub

Private Sub Command2_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 8

End Sub

Private Sub Command20_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
sign = 6

End Sub

Private Sub Command21_Click()
a = Val(Label1.Caption)
Label1.Caption = ""
Label1.Caption = Sqr(a)

End Sub

Private Sub Command23_Click()

Label1.Caption = Label1.Caption & "."
End Sub

Private Sub Command24_Click()
mem = 0
Text1.Text = ""
Label1.Caption = ""

End Sub

Private Sub Command25_Click()
Label1.Caption = mem

End Sub

Private Sub Command26_Click()
mem = Val(Label1.Caption)
Text1.Text = "M"
Label1.Caption = ""

End Sub

Private Sub Command27_Click()
Dim s1 As Double
a = Val(Label1.Caption)
Label1.Caption = ""
s1 = a + mem

Label1.Caption = s1

End Sub

Private Sub Command28_Click()
Dim s1 As Double
a = Val(Label1.Caption)
Label1.Caption = ""
s1 = mem - a
Label1.Caption = s1

End Sub

Private Sub Command29_Click()
Dim n As Integer

n = Len(Label1.Caption)
Label1.Caption = Left(Label1.Caption, n - 1)

End Sub

Private Sub Command3_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 7

End Sub

Private Sub Command4_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 4

End Sub

Private Sub Command5_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 5

End Sub

Private Sub Command6_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 6

End Sub

Private Sub Command7_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 3

End Sub

Private Sub Command8_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 2

End Sub

Private Sub Command9_Click()
If Label1.Caption = "0" Then
    Label1.Caption = ""
End If
Label1.Caption = Label1.Caption & 1

End Sub

Private Sub Form_Load()
Label1.Caption = 0
End Sub
