VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   360
      Top             =   1560
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "&NEXT"
      Height          =   855
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&RESET"
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SCORE"
      Height          =   855
      Left            =   4800
      TabIndex        =   31
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 3"
      Height          =   2055
      Left            =   1080
      TabIndex        =   20
      Top             =   6600
      Width           =   16575
      Begin VB.OptionButton Option12 
         Caption         =   "Option4"
         Height          =   195
         Left            =   12840
         TabIndex        =   24
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Option3"
         Height          =   195
         Left            =   9000
         TabIndex        =   23
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Option2"
         Height          =   195
         Left            =   5040
         TabIndex        =   22
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Boxing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12720
         TabIndex        =   29
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Badminton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   28
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Weightlifting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   27
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Wrestling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   26
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "THE FIRST MEDAL WON BY AN INDIAN WOMAN WAS IN WHICH GAME?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   14415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 2"
      Height          =   2055
      Left            =   1080
      TabIndex        =   10
      Top             =   3960
      Width           =   16575
      Begin VB.OptionButton Option5 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option2"
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   13
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option3"
         Height          =   195
         Index           =   0
         Left            =   9000
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option4"
         Height          =   195
         Index           =   0
         Left            =   12840
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "THE STATUE OF WHICH FREEDOM FIGHTER WAS RECENTLY BUILT IN GUJARAT?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   14415
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vallabhbhai Patel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   18
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mahatma Gandhi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Subhas Chandra Bose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   16
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bhagat Singh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12720
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 1"
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   16575
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   195
         Left            =   12840
         TabIndex        =   9
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   9000
         TabIndex        =   8
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   5040
         TabIndex        =   7
         Top             =   1440
         Width           =   135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ricky Ponting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12720
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Virat Kohli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sourav Ganguly"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sachin Tendulkar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHO IS THE FASTEST BATSMAN TO REACH 10000 RUNS IN ODIS?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   14415
      End
   End
   Begin VB.Label score 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
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
      Left            =   1080
      TabIndex        =   30
      Top             =   9240
      Width           =   2415
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim marks As Integer

Private Sub Command2_Click()
score.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5(1).Value = False
Option6(1).Value = False
Option7(0).Value = False
Option8(0).Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False

End Sub

Private Sub Command1_Click()
marks = 0
If Option3.Value = True Then
marks = marks + 4
Else
marks = marks - 2
End If

If Option5(1).Value = True Then
marks = marks + 4
Else
marks = marks - 2
End If

If Option10.Value = True Then
marks = marks + 4
Else
marks = marks - 2
End If
score.Visible = True
score.Caption = marks
Timer1.Enabled = True

End Sub

Private Sub Command3_Click()
Me.Hide
'Form7.Show

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
score.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5(1).Value = False
Option6(1).Value = False
Option7(0).Value = False
Option8(0).Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False


End Sub

Private Sub Timer1_Timer()
If Timer1.Interval >= 10000 Then
    Command1_Click
    'Command3_Click
End If
End Sub
