VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   1920
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
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   3600
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GENERATE"
      Height          =   1215
      Left            =   5280
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
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
      Height          =   735
      Left            =   10800
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ENTER NO:OF TERMS(<20)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(100), n, x, y, i, subc, sum As Integer
Dim disp As String


Private Sub Command1_Click()
n = Val(Text1.Text)
x = 1
y = 1
sum = 0
disp = "1 1 "
Text2.Text = disp

For i = 0 To n - 3
    sum = x + y
    x = y
    y = sum
    a(i) = sum
    'Text2.Text = Text2.Text & a(i) & " "
Next
'Text2.Text = Text2.Text & disp & a(0) & " "
i = 0
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
If i <= (n - 3) Then
    Text2.Text = Text2.Text & a(i) & " "
    i = i + 1
Else
    Timer1.Enabled = False
End If


End Sub
