VERSION 5.00
Begin VB.Form CAL 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CALCULATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   3480
      TabIndex        =   0
      Top             =   1920
      Width           =   12615
      Begin VB.CommandButton MINUS 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   17
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton MULTI 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   16
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton DIV 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   15
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton EQU 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   14
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton PLUS 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   13
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton AC 
         Caption         =   "AC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   12
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   4800
         TabIndex        =   11
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   5760
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   6720
         TabIndex        =   9
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   7680
         TabIndex        =   8
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   4800
         TabIndex        =   7
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   5760
         TabIndex        =   6
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   7680
         TabIndex        =   5
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   6720
         TabIndex        =   4
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   5760
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CMD 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   4800
         TabIndex        =   2
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         TabIndex        =   1
         Top             =   1440
         Width           =   3855
      End
   End
End
Attribute VB_Name = "CAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Single
Dim P As Single
Dim R As Single
Dim CH As String
Private Sub AC_Click()
P = C = 0
Text1.Text = ""
Text1.SetFocus
End Sub
Private Sub CMD_Click(Index As Integer)
Text1.Text = Text1.Text & CMD(Index).Caption
C = Val(Text1.Text)
End Sub
Private Sub DIV_Click()
Text1.Text = ""
P = C
C = 0
CH = "/"
End Sub
Private Sub EQU_Click()
Select Case CH
Case "+"
R = P + C
Text1.Text = R
Case "-"
R = P - C
Text1.Text = R
Case "*"
R = P * C
Text1.Text = R
Case "/"
R = P / C
Text1.Text = R
End Select
C = R
End Sub
Private Sub MINUS_Click()
Text1.Text = ""
P = C
C = 0
CH = "-"
End Sub
Private Sub MULTI_Click()
Text1.Text = ""
P = C
C = 0
CH = "*"
End Sub
Private Sub PLUS_Click()
Text1.Text = ""
P = C
C = 0
CH = "+"
End Sub
