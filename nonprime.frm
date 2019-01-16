VERSION 5.00
Begin VB.Form Form4 
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
      Interval        =   1000
      Left            =   720
      Top             =   1200
   End
   Begin VB.TextBox Text2 
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
      Left            =   2520
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   3360
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GENERATE"
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
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
      Height          =   615
      Left            =   9120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ENTER AN INTEGER(<=20)"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, i, j, a(100), x As Integer

Private Sub Command1_Click()
Dim count As Integer
n = Val(Text1.Text)
For i = 1 To n
    count = 0
    For j = 1 To n
        If ((i Mod j) = 0) Then
        count = count + 1
        End If
    Next
    If count <> 2 Then
    a(x) = i
    x = x + 1
    End If
Next

n = x
x = 0
Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
If x <= (n - 1) Then
    Text2.Text = Text2.Text & a(x) & " "
    x = x + 1
Else
    Timer1.Enabled = False
End If


End Sub
