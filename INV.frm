VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form INV_DTL 
   BackColor       =   &H00FFFF80&
   Caption         =   "INVOICE DETAILS"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14475
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
      Left            =   11040
      TabIndex        =   19
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "INVOICE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   9480
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   615
         Left            =   3120
         TabIndex        =   18
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   42714
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
         Left            =   2880
         TabIndex        =   17
         Top             =   7320
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1080
         TabIndex        =   16
         Top             =   7320
         Width           =   1815
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
         Left            =   3120
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   2280
         Width           =   2775
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
         Left            =   8280
         TabIndex        =   13
         Top             =   7320
         Width           =   1815
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
         Height          =   615
         Left            =   3120
         TabIndex        =   12
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text2 
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
         Left            =   8880
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text3 
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
         Left            =   8880
         TabIndex        =   4
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox Text4 
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
         Left            =   8880
         TabIndex        =   3
         Top             =   2880
         Width           =   2775
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
         Left            =   4680
         TabIndex        =   2
         Top             =   7320
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
         Left            =   6480
         TabIndex        =   1
         Top             =   7320
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "INV.frx":0000
         Height          =   2535
         Left            =   1080
         TabIndex        =   14
         Top             =   4440
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4471
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
         Caption         =   "DISCOUNT"
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
         Left            =   6720
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT"
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
         Left            =   6720
         TabIndex        =   10
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VAT"
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
         Left            =   6720
         TabIndex        =   9
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INVOICE ID"
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
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE"
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
         Left            =   960
         TabIndex        =   7
         Top             =   2880
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
         Left            =   960
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
   End
End
Attribute VB_Name = "INV_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
DTP1.Enabled = True
Combo1.SetFocus
Text2.Text = ""
Text3.Text = ""
path
S = "select max(to_number(substr(pu_id,4,length(pu_id)))) from pur_dtl"
Set R = CC.Execute(S)
code = "PUR"
If IsNull(R.Fields(0)) Then
I = 1
Text1.Text = code & I
Else
I = R.Fields(0) + 1
Text1.Text = code & I
End If
End Sub
Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
If Text1.Text = "" Or Combo1.Text = "" Or Text2.Text = "" Or DTP1.Value = "" Or Combo2.Text = "" Or Text3.Text = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into pur_dtl values('" + Text1.Text + "','" + Combo1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "','" + Combo2.Text + "'," + Text2.Text + "," + Text3.Text + ")"
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Adodc1.Refresh
Command2.Enabled = False
Command1.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
DTP1.Enabled = False
Command3.Enabled = True
End If
End Sub

Private Sub Command3_Click()
Dim id As String
id = InputBox("Enter Purchase_Id to be Deleted ! ")
path
S = "select pu_id from pur_dtl"
Set R2 = CC.Execute(S)
While Not R2.EOF
If id = Trim(R2.Fields(0)) Then
S = "delete from pur_dtl where pu_id='" + id + "'"
Set R = CC.Execute(S)
MsgBox "RECORD DELETED"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DTP1.Value = ""
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
DTP1.Enabled = False
End Sub

Private Sub Command4_Click()
Command3.Enabled = True
Text1.Enabled = True
MsgBox "Enter product id to be searched"
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
Text1.Enabled = True
Text1.SetFocus
path
S = "update pur_dtl set p_id='" + Combo1.Text + "',p_dt='" + Format(DTP1.Value, "dd-mmm-yyyy") + "',s_id='" + Combo2.Text + "',qty=" + Text2.Text + " ,amt=" + Text3.Text + " where pu_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !", vbOKCancel, "Message"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
DTP1.Value = ""
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
DTP1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
DTP1.Value = Date
path
S = "select p_id from prd_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
path
S = "select s_id from sup_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
path
S = "select p_id,p_dt,s_id,qty,amt from pur_dtl where pu_id= '" + Text1.Text + " '"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    DTP1.Value = ""
Exit Sub
Else
Command5.Enabled = True
Text2.Text = R.Fields("qty")
Text3.Text = R.Fields("amt")
DTP1.Value = R.Fields("p_dt")
Combo1.Text = R.Fields("p_id")
Combo2.Text = R.Fields("s_id")
Text2.Enabled = True
Text3.Enabled = True
DTP1.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
End If
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
