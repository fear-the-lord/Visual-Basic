VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form13"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15105
   LinkTopic       =   "Form13"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "new"
      Height          =   615
      Left            =   960
      TabIndex        =   24
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   25
      Top             =   8400
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "sex"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   14040
      TabIndex        =   8
      Text            =   "Select Sex"
      Top             =   4680
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "dept"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4920
      TabIndex        =   7
      Text            =   "Select Department"
      Top             =   4680
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&SUBMIT"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9120
      Width           =   2535
   End
   Begin VB.TextBox yophisec 
      DataField       =   "yearofpassinghisec"
      DataSource      =   "Adodc1"
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
      Left            =   14040
      TabIndex        =   12
      Top             =   7440
      Width           =   3975
   End
   Begin VB.TextBox hisecsch 
      DataField       =   "school hisec"
      DataSource      =   "Adodc1"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   7320
      Width           =   3975
   End
   Begin VB.TextBox yopsec 
      DataField       =   "yearofpassingsec"
      DataSource      =   "Adodc1"
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
      Left            =   14040
      TabIndex        =   10
      Top             =   6000
      Width           =   3975
   End
   Begin VB.TextBox secsch 
      DataField       =   "school sec"
      DataSource      =   "Adodc1"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   6000
      Width           =   3975
   End
   Begin VB.TextBox mobile 
      DataField       =   "mobile no"
      DataSource      =   "Adodc1"
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
      Left            =   14040
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox dob 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox address 
      DataField       =   "address"
      DataSource      =   "Adodc1"
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
      Left            =   14040
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox name1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label12 
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
      Left            =   4560
      TabIndex        =   23
      Top             =   10680
      Width           =   10095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "YEAR OF PASSING(HI. SEC.)"
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
      Left            =   9480
      TabIndex        =   20
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SCHOOL(HI. SEC.)"
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
      Left            =   360
      TabIndex        =   19
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "YEAR OF PASSING(SEC.)"
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
      Left            =   9480
      TabIndex        =   18
      Top             =   6000
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SCHOOL(SEC.)"
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
      Left            =   360
      TabIndex        =   17
      Top             =   6000
      Width           =   3975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "SEX"
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
      Left            =   9480
      TabIndex        =   16
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEPARTMENT"
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
      Left            =   360
      TabIndex        =   15
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "MOBILE NUMBER"
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
      Left            =   9480
      TabIndex        =   14
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "DATE OF BIRTH"
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
      Left            =   360
      TabIndex        =   13
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADDRESS"
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
      Left            =   9480
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "NAME"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "STUDENT INFORMATION"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "Form13"
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
    address.Enabled = True
Else
    MsgBox "Enter name correctly!!"
    name1.Text = ""
    name1.Enabled = True
    Exit Sub
End If


a = Left(dob.Text, 2)
b = Mid(dob.Text, 4, 2)
If a > 31 Or b > 12 Then
    MsgBox "Wrong date inserted!!"
    dob.Text = ""
    dob.Enabled = True
    Exit Sub
End If



If dob.Text = "" Then
    MsgBox "Please enter date"
    Exit Sub
End If


n = Len(yopsec.Text)
m = Len(yophisec.Text)
If m <> 4 Or n <> 4 Then
    MsgBox "Wrong year inserted"
    yopsec.Text = ""
    yophisec.Text = ""
    Exit Sub
End If

a = InStr(1, mobile.Text, 0)
b = InStr(1, mobile.Text, 1)
c = InStr(1, mobile.Text, 2)
d = InStr(1, mobile.Text, 3)
e = InStr(1, mobile.Text, 4)
f = InStr(1, mobile.Text, 5)
g = InStr(1, mobile.Text, 6)
h = InStr(1, mobile.Text, 7)
i = InStr(1, mobile.Text, 8)
j = InStr(1, mobile.Text, 9)


n = Len(mobile.Text)
If n <> 10 Then
    MsgBox "Mobile number should contain 10 digits!!"
    mobile.Text = ""
    Exit Sub
End If

If a = 0 And b = 0 And c = 0 And d = 0 And e = 0 And f = 0 And g = 0 And h = 0 And i = 0 And j = 0 Then
    MsgBox "Enter proper phone number!!"
    mobile.Text = ""
    Exit Sub
End If


    
Label12.Caption = "Student Record Submitted Successfully!!"

Adodc1.Recordset.Update

End Sub

Private Sub Text10_Change()

End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text6_Change()

End Sub

Private Sub Command2_Click()
name1.Text = ""
mobile.Text = ""
address.Text = ""
secsch.Text = ""
hisecsch.Text = ""
yopsec.Text = ""
yophisec.Text = ""
dob.Text = ""
Label12.Caption = ""
Combo1.Text = "Select Department"
Combo2.Text = "Select Sex"


End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Form_Load()
Combo1.AddItem "Computer Science"
Combo1.AddItem "Physics"
Combo1.AddItem "Maths"
Combo1.AddItem "Chemistry"
Combo1.AddItem "Economics"
Combo1.AddItem "English"
Combo1.AddItem "Political Science"

Combo2.AddItem "Male"
Combo2.AddItem "Female"
Combo2.AddItem "Others"
End Sub
