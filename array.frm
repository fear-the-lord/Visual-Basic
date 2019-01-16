VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
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
      Left            =   14400
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AVERAGE"
      Height          =   855
      Left            =   11400
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   9495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SORT"
      Height          =   855
      Left            =   8520
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE "
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4800
      TabIndex        =   0
      Top             =   1320
      Width           =   9375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(10), result(10), i, j As Integer

Private Sub Command1_Click()
Dim str As String

For i = 0 To 5
    a(i) = InputBox("Enter (" & i & ") element")
    str = str & " " & a(i)
    List1.AddItem str
    str = ""
    
Next

End Sub

Private Sub Command2_Click()
Dim str As String
Dim tmp As Integer
For i = 0 To 5
    For j = 0 To 4 - i
        If (Val(a(j)) > Val(a(j + 1))) Then
            tmp = Val(a(j))
            a(j) = a(j + 1)
            a(j + 1) = tmp
        End If
    Next
Next
For i = 0 To 5
    result(i) = a(i)
    str = str & " " & result(i)
    List2.AddItem str
    str = ""
Next

        
        
End Sub

Private Sub Command3_Click()
Dim sum As Integer
Dim avg As Double
For i = 0 To 5
    sum = sum + a(i)
Next
avg = sum / 5
Text1.Text = avg

End Sub

