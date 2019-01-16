VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   600
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&GENERATE"
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   2775
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
      Height          =   615
      Left            =   3960
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4560
      Width           =   2295
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
      Left            =   12360
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER A NUMBER N(4 TO 50)"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "PRIME NUMBER GENERATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, i, j, fact, x As Integer
Dim disp As String
Dim a(100) As Integer


Private Sub Command1_Click()
n = Val(Text1.Text)
If Not IsNumeric(n) Or n < 3 Or n > 50 Then
    MsgBox "Enter an integer between 4 and 50!!", vbOKOnly, "Invalid input"
    Text1.Text = ""
    Text1.SetFocus
    Exit Sub
End If

disp = ""
x = 0
For i = 1 To n
    fact = 0
    For j = 1 To n
    If ((i Mod j) = 0) Then
        fact = fact + 1
        End If
     Next
        
    If fact = 2 Then
        disp = disp & i & " "
        a(x) = i
        x = x + 1
        End If

Next i
n = x
x = 0
Timer1.Enabled = True



    
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Timer1_Timer()

If x <= (n - 1) Then
    Text2.Text = Text2.Text & a(x) & " "
    x = x + 1
Else
    Timer1.Enabled = False
End If
    

End Sub
