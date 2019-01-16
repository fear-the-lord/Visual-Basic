VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form17 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form17"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form17"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      TabIndex        =   24
      Top             =   10680
      Width           =   2295
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
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10680
      Width           =   2295
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
      Height          =   1095
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10680
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "QUESTION 2"
      Height          =   2295
      Left            =   480
      TabIndex        =   12
      Top             =   7680
      Width           =   17535
      Begin VB.OptionButton Option8 
         Caption         =   "Option1"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   1680
         Width           =   135
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option2"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option3"
         Height          =   255
         Left            =   9360
         TabIndex        =   14
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option4"
         Height          =   195
         Left            =   13800
         TabIndex        =   13
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "ZAYN MALIK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   21
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHO IS THE MALE LEAD SINGER SHOWN IN THIS CLIP?"
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
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   16335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "JUSTIN BIEBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   19
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "SHAWN MENDES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9240
         TabIndex        =   18
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "HARRY STYLES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13680
         TabIndex        =   17
         Top             =   1440
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "QUESTION 1"
      Height          =   2295
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   17535
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   195
         Left            =   13800
         TabIndex        =   10
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   9360
         TabIndex        =   8
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   5040
         TabIndex        =   6
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "MONKEY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13680
         TabIndex        =   9
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "LION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9240
         TabIndex        =   7
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "HORSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "WHICH ANIMALS ARE SHOWN IN THE CLIP?"
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
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   16335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "CHEETAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   3000
      Left            =   5400
      TabIndex        =   11
      Top             =   5160
      Width           =   4500
      URL             =   "C:\Users\Souvik\Downloads\ZAYN, Taylor Swift - I Don’t Wanna Live Forever (Fifty Shades Darker).mp4"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7938
      _cy             =   5292
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3000
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      URL             =   "C:\Users\Souvik\Downloads\wildlife.wmv"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7938
      _cy             =   5292
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim marks As Integer
If Option2.Value = True Then
    marks = marks + 4
Else
    marks = marks - 2
End If

If Option8.Value = True Then
    marks = marks + 4
Else
    marks = marks - 2
End If
WindowsMediaPlayer1.Controls.stop
WindowsMediaPlayer2.Controls.stop

Text1.Visible = True

Text1.Text = "SCORE: " & marks



End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Text1.Visible = False
WindowsMediaPlayer1.Controls.stop
WindowsMediaPlayer2.Controls.stop


End Sub

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Text1.Visible = False
WindowsMediaPlayer2.Controls.stop


End Sub
