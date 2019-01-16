VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   2520
      Top             =   6960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2143
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "INSERT"
      Height          =   735
      Left            =   11520
      TabIndex        =   15
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
      Height          =   735
      Left            =   11520
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UPDATE"
      Height          =   735
      Left            =   11520
      TabIndex        =   13
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAST"
      Height          =   735
      Left            =   11520
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FIRST "
      Height          =   735
      Left            =   11520
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PREVIOUS"
      Height          =   735
      Left            =   11520
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   735
      Left            =   11520
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "COPIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "BOOK NAME "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "BOOK_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LIBRARY MANAGEMENT SYSTEM"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
    MsgBox "This is the last record!!"
    Adodc1.Recordset.MovePrevious
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
    MsgBox "This is the first record"
    Adodc1.Recordset.MoveNext
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveFirst
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False
End Sub

Private Sub Command6_Click()
Dim x As Integer
x = MsgBox("Delete a record?", vbYesNo, "Confirm Deletion?")
If x = vbYes Then
    Adodc1.Recordset.Delete
    MsgBox "Record Deleted"
    Adodc1.Recordset.MoveFirst
Else
    Adodc1.Recordset.Cancel
End If

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command7.Enabled = False
Command7.Visible = False

End Sub

Private Sub Command7_Click()
Adodc1.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command8.Enabled = True
Command8.Visible = True

End Sub

Private Sub Label2_Click(Index As Integer)

End Sub
