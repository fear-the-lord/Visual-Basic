VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "PREVIOUS"
      Height          =   1095
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "RESET"
      Height          =   1095
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SUBMIT"
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 6"
      Height          =   2175
      Left            =   1560
      TabIndex        =   21
      Top             =   6360
      Width           =   15015
      Begin VB.OptionButton Option11 
         Caption         =   "Option4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   1560
         Width           =   225
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Option3"
         Height          =   195
         Left            =   4560
         TabIndex        =   25
         Top             =   1560
         Width           =   195
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option2"
         Height          =   195
         Left            =   8280
         TabIndex        =   24
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option9"
         Height          =   195
         Index           =   2
         Left            =   12000
         TabIndex        =   22
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Linus Torvalis"
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
         Left            =   11880
         TabIndex        =   31
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Elon Musk"
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
         Left            =   8160
         TabIndex        =   30
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bill Gates"
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
         Left            =   4320
         TabIndex        =   29
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Jeff Bezos"
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
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHO IS THE FOUNDER OF TESLA?"
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
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   14175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 5"
      Height          =   2175
      Left            =   1560
      TabIndex        =   10
      Top             =   3600
      Width           =   15015
      Begin VB.OptionButton Option8 
         Caption         =   "Option9"
         Height          =   195
         Index           =   1
         Left            =   12000
         TabIndex        =   20
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option2"
         Height          =   195
         Left            =   8280
         TabIndex        =   13
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option3"
         Height          =   195
         Left            =   4560
         TabIndex        =   12
         Top             =   1560
         Width           =   195
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHO WAS THE HIGHEST RUN GETTER FOR INDIA IN 2011 WORLD CUP?"
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
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   14175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gautam Gambhir"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "MS Dhoni"
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
         Left            =   4320
         TabIndex        =   17
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label7 
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
         Left            =   8160
         TabIndex        =   16
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label6 
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
         Left            =   11880
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "QUESTION 4"
      Height          =   2175
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   15015
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12000
         TabIndex        =   9
         Top             =   1560
         Width           =   135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   8280
         TabIndex        =   7
         Top             =   1560
         Width           =   195
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   4560
         TabIndex        =   5
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "ARGENTINA"
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
         Left            =   11880
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "INDIA"
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
         Left            =   8160
         TabIndex        =   6
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "USA && MEXICO"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "UAE"
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
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHERE WILL THE FIFA WORLD CUP 2022 BE HELD?"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   14175
      End
   End
   Begin VB.Label finalscore 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   1680
      TabIndex        =   34
      Top             =   9240
      Width           =   3735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
marks1 = 0
marks2 = 0

If Option2.Value = True Then
marks1 = marks1 + 4
Else
marks1 = marks1 - 2
End If

If Option5.Value = True Then
marks1 = marks1 + 4
Else
marks1 = marks1 - 2
End If

If Option9.Value = True Then
marks1 = marks1 + 4
Else
marks1 = marks1 - 2
End If

marks2 = marks1 + Val(Form6.score.Caption)
finalscore.Visible = True
finalscore.Caption = "YOUR FINAL SCORE IS:" & Val(Form6.score.Caption) + marks1


End Sub

Private Sub Command2_Click()
finalscore.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8(1).Value = False
Option11.Value = False
Option10.Value = False
Option9.Value = False
Option8(2).Value = False

End Sub

Private Sub Command3_Click()
Me.Hide
Form6.Show

End Sub

Private Sub Form_Load()
finalscore.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8(1).Value = False
Option11.Value = False
Option10.Value = False
Option9.Value = False
Option8(2).Value = False

End Sub

