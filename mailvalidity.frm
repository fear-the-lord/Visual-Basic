VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      Height          =   975
      Left            =   7560
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   9480
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   4560
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENTER AN EMAIL ID"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = InStr(1, Text1.Text, ".com")
m = InStr(1, Text1.Text, "@")
If n = 0 Or m = 0 Then
    Label2.Caption = "IT IS NOT A VALID MAIL ID"
Else
    Label2.Caption = "IT IS A VALID MAIL ID"
End If

    


End Sub

Private Sub Form_Load()

End Sub
