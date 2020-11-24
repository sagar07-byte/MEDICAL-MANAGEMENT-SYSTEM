VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form STK_DTL 
   BackColor       =   &H00C0C0FF&
   Caption         =   "STOCK DETAILS"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   14850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   13
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "STOCK DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   9360
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
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
         Connect         =   "Provider=MSDAORA.1;User ID=medicine/mmp;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=medicine/mmp;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from stk_dtl"
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
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5280
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   5280
         TabIndex        =   7
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   5280
         TabIndex        =   6
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ADD NEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   5
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   4
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5760
         TabIndex        =   3
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7560
         TabIndex        =   2
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9360
         TabIndex        =   1
         Top             =   6720
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "STK_DEL.frx":0000
         Height          =   2295
         Left            =   1560
         TabIndex        =   12
         Top             =   3960
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4048
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   10
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAPACITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   9
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   1440
         Width           =   2175
      End
   End
End
Attribute VB_Name = "STK_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text <> "" Then
'If KeyAscii = 13 Then
path
S = "select cpt,qty from stk_dtl where p_id= '" + Combo1.Text + " '"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
Exit Sub
Else
Command5.Enabled = True
Combo1.Enabled = False
Text1.Text = R.Fields("cpt")
Text2.Text = R.Fields("qty")
Text1.Enabled = True
Text2.Enabled = True
End If
End If
End Sub
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo1.SetFocus
Text2.Text = ""
End Sub
Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into stk_dtl values('" + Combo1.Text + "'," + Text1.Text + "," + Text2.Text + ")"
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Adodc1.Refresh
Command2.Enabled = False
Command1.Enabled = True
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Command3.Enabled = True
End If
End Sub
Private Sub Command3_Click()
Dim id As String
id = InputBox("Enter Purchase_Return_Id to be Deleted ! ")
path
S = "select pr_id from pur_rt"
Set R2 = CC.Execute(S)
While Not R2.EOF
If id = Trim(R2.Fields(0)) Then
S = "delete from pur_rt where pr_id='" + id + "'"
Set R = CC.Execute(S)
MsgBox "RECORD DELETED"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Adodc1.Refresh
Exit Sub
Else
R2.MoveNext
End If
Wend
MsgBox "Record Not Found ! ", vbCritical, "Message"
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Command4_Click()
Command3.Enabled = True
Combo1.Enabled = True
MsgBox "select product id to be searched", vbInformation, "MESSAGE"
Combo1.SetFocus
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
Combo1.Enabled = False
Text1.Text = False
Text2.Text = False
path
S = "update stk_dtl set cpt=" + Text1.Text + ",qty=" + Text2.Text + " where p_id='" + Combo1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !", vbOKCancel, "Message"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
Private Sub Form_Load()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
path
S = "select p_id from prd_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub

