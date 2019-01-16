VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "EXIT"
      Height          =   1095
      Left            =   9720
      TabIndex        =   9
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TRANSPOSE"
      Height          =   1095
      Left            =   6240
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MULTIPLY"
      Height          =   1095
      Left            =   14520
      TabIndex        =   7
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SUBTRACT"
      Height          =   1095
      Left            =   11040
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD"
      Height          =   1095
      Left            =   7920
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CREATE 2"
      Height          =   1095
      Left            =   5040
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE 1"
      Height          =   1095
      Left            =   1920
      TabIndex        =   3
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   12960
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
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
      Height          =   1560
      Left            =   7200
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
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
      Height          =   1560
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(3, 3), b(3, 3), result(3, 3), i, j, k As Integer

Private Sub Command1_Click()
List1.Clear
Dim str As String

For i = 0 To 2
    str = ""
    For j = 0 To 2
     a(i, j) = InputBox("Enter (" & i & "," & j & ") element")
     str = str & Space(3) & a(i, j)
    Next
    List1.AddItem str
Next
        
        
End Sub

Private Sub Command2_Click()
List2.Clear
Dim str As String
For i = 0 To 2
    str = ""
    For j = 0 To 2
        b(i, j) = InputBox("Enter (" & i & "," & j & ")element")
        str = str & Space(3) & b(i, j)
    Next
    List2.AddItem str
Next
        
End Sub

Private Sub Command3_Click()
List3.Clear
Dim str As String

For i = 0 To 2
    str = ""
    For j = 0 To 2
        result(i, j) = Val(a(i, j)) + Val(b(i, j))
        str = str & Space(3) & result(i, j)
    Next
    List3.AddItem str
Next
End Sub

Private Sub Command4_Click()
List3.Clear
Dim str As String

For i = 0 To 2
    str = ""
    For j = 0 To 2
        result(i, j) = Val(a(i, j)) - Val(b(i, j))
        str = str & Space(3) & result(i, j)
    Next
    List3.AddItem str
Next

End Sub

Private Sub Command5_Click()
List3.Clear
Dim str As String
For i = 0 To 2
    For j = 0 To 2
        result(i, j) = 0
        For k = 0 To 2
            result(i, j) = result(i, j) + (Val(a(i, k)) * Val(b(k, j)))
        Next
    Next
Next
For i = 0 To 2
    str = ""
    For j = 0 To 2
        str = str & Space(3) & result(i, j)
    Next
    List3.AddItem str
Next

    
            
End Sub

Private Sub Command6_Click()
List2.Clear
Dim str As String
For i = 0 To 2
    str = ""
    For j = 0 To 2
        b(i, j) = a(j, i)
        str = str & Space(3) & b(i, j)
    Next
    List2.AddItem str
Next

End Sub

Private Sub Command7_Click()
List1.Clear
List2.Clear
List3.Clear

End Sub

