VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H000080FF&
   Caption         =   "Form16"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Caption         =   "&EXIT"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "&BILL"
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BILL"
      Height          =   1695
      Left            =   5520
      TabIndex        =   22
      Top             =   8400
      Width           =   7095
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   5280
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "&RESET"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   23
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check6"
      Height          =   195
      Left            =   10680
      TabIndex        =   11
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check5"
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check4"
      Height          =   195
      Left            =   10680
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   6
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   4
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check3"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check2"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check1"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "      EGG CHICKEN      ROLL 80/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10440
      TabIndex        =   21
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "            CHICKEN             ROLL 60/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10440
      TabIndex        =   20
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "        MUTTON         REZALA 180/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10440
      TabIndex        =   19
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "           CHICKEN            REZALA 150/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   18
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "            MUTTON             BIRYANI 200/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   17
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "            CHICKEN             BIRYANI 180/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "THE TAJ"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check1.Value = Unchecked Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
End If


End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check2.Value = Unchecked Then
        Check2.Value = Checked
    Else
        Check2.Value = Unchecked
    End If
End If
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check3.Value = Unchecked Then
        Check3.Value = Checked
    Else
        Check3.Value = Unchecked
    End If
End If
End Sub



Private Sub Check4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check4.Value = Unchecked Then
        Check4.Value = Checked
    Else
        Check4.Value = Unchecked
    End If
End If
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check5.Value = Unchecked Then
        Check5.Value = Checked
    Else
        Check5.Value = Unchecked
    End If
End If
End Sub

Private Sub Check6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check6.Value = Unchecked Then
        Check6.Value = Checked
    Else
        Check6.Value = Unchecked
    End If
End If
End Sub

Private Sub Command1_Click()
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
Check5.Value = Unchecked
Check6.Value = Unchecked
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""



End Sub

Private Sub Command2_Click()
Dim x As Integer
If Check1.Value = Checked Then
    x = x + Val(Text1.Text) * 180
End If

If Check2.Value = Checked Then
    x = x + Val(Text2.Text) * 200
End If

If Check3.Value = Checked Then
    x = x + Val(Text3.Text) * 150
End If

If Check4.Value = Checked Then
    x = x + Val(Text4.Text) * 180
End If

If Check5.Value = Checked Then
    x = x + Val(Text5.Text) * 60
End If

If Check6.Value = Checked Then
    x = x + Val(Text6.Text) * 80
End If

If x > 1000 Then
    MsgBox "You are eligible for 10% discount on purchase above 1000/-", vbOKOnly, "Discount Applied"
    x = 0.9 * x
End If

Text7.Text = Val(x)



End Sub

Private Sub Command3_Click()
Form16.Hide


End Sub
