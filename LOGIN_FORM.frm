VERSION 5.00
Begin VB.Form LOGIN_FORM 
   BackColor       =   &H00FFC0FF&
   Caption         =   "USER AUTHENTICATION"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15690
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "LOGIN_FORM.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   11880
      MaxLength       =   3
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "LOGIN"
      Height          =   3975
      Left            =   6240
      TabIndex        =   0
      Top             =   3240
      Width           =   8895
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         Picture         =   "LOGIN_FORM.frx":125F2F
         ScaleHeight     =   2835
         ScaleWidth      =   2835
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "OK"
         Height          =   615
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "CANCEL"
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         FillColor       =   &H00FF0000&
         Height          =   855
         Left            =   4200
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USER ID"
         Height          =   735
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PASSWORD"
         Height          =   735
         Left            =   3240
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
   End
End
Attribute VB_Name = "LOGIN_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Enter all fields", vbInformation, "MESSAGE"
Text1.SetFocus
Exit Sub
Else
If Text1.Text = "MEDICINE" And Text2.Text = "mmp" Then
    Unload Me
    MDIForm1.Show
Else
    MsgBox "PLEASE TRY AGAIN", vbInformation, "MESSAGE"
End If
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub Form_Load()
path
End Sub

