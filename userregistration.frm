VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form9"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9120
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9120
      Width           =   3855
   End
   Begin VB.TextBox repass 
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
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   7440
      Width           =   3615
   End
   Begin VB.TextBox pass 
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
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   6000
      Width           =   3615
   End
   Begin VB.TextBox email1 
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
      Left            =   10200
      TabIndex        =   8
      Top             =   4560
      Width           =   3615
   End
   Begin VB.TextBox dob 
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
      Left            =   10200
      TabIndex        =   7
      Top             =   3120
      Width           =   3615
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
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   13
      Top             =   10800
      Width           =   8055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "RE-ENTERED PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   5
      Top             =   7440
      Width           =   3735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "EMAIL ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "DATE OF BIRTH:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "USER REGISTRATION FORM"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = InStr(1, name1.Text, 0)
b = InStr(1, name1.Text, 1)
c = InStr(1, name1.Text, 2)
d = InStr(1, name1.Text, 3)
e = InStr(1, name1.Text, 4)
f = InStr(1, name1.Text, 5)
g = InStr(1, name1.Text, 6)
h = InStr(1, name1.Text, 7)
i = InStr(1, name1.Text, 8)
j = InStr(1, name1.Text, 9)
If a = 0 And b = 0 And c = 0 And d = 0 And e = 0 And f = 0 And g = 0 And h = 0 And i = 0 And j = 0 Then
    dob.Enabled = True
Else
    MsgBox "Enter name correctly!!"
    name1.Text = ""
    Exit Sub
End If

a = InStr(1, email1.Text, "@")
b = InStr(1, email1.Text, ".")
If a = 0 Or b = 0 Then
  MsgBox "Enter valid email!!"
    email1.Text = ""
    email1.Enabled = True
    Exit Sub
Else
    pass.Enabled = True
End If


a = Len(pass.Text)
b = Len(repass.Text)
If a < 8 Or b < 8 Then
    MsgBox "Password should be minimum of 8 characters "
    pass.Text = ""
    repass.Text = ""
    Exit Sub
End If

If pass.Text = repass.Text Then
    Command1.Enabled = True
Else
    MsgBox "Re-typed password does not match!!"
    repass.Text = ""
End If

a = Left(dob.Text, 2)
b = Mid(dob.Text, 4, 2)
If a > 31 Or b > 12 Then
    MsgBox "Date is invalid.Enter valid date!!"
    dob.Text = ""
End If

Label7.Caption = "User Registered Successfully"

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Command2_Click()
name1.Text = ""
dob.Text = ""
email1.Text = ""
pass.Text = ""
repass.Text = ""
Label7.Caption = ""

End Sub
