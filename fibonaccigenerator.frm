VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&GENERATE"
      Height          =   1215
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   3135
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
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   4800
      Width           =   10695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1800
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
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER NUMBER OF TERMS N (4 TO 20)"
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
      Top             =   2880
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "FIBONACCI GENERATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   0
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x, y, n, sum, counter As Integer
Dim disp As String

Private Sub Command1_Click()
n = Val(Text1.Text)
If Not IsNumeric(n) Or n < 3 And n > 20 Then

    MsgBox "Enter an integer between 4 and 20!!"
    Text1.Text = ""
    Exit Sub
    End If
    

    disp = "1 1 "
    x = 1
    y = 1
    Text2.Text = disp
    counter = 3
    Timer1.Enabled = True

    
End Sub


Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()
Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()
sum = x + y
disp = disp & sum & " "
x = y
y = sum
Text2.Text = disp
counter = counter + 1
If counter > n Or counter > 20 Then
Timer1.Enabled = False
End If

End Sub
