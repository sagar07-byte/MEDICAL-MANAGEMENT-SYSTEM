VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SUP_DTL 
   BackColor       =   &H0000FF00&
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15930
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF80FF&
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20415
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080FF80&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "SUP_DTL.frx":0000
         Height          =   2775
         Left            =   4320
         TabIndex        =   15
         Top             =   5040
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   65280
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   24
         FormatLocked    =   -1  'True
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "S_ID"
            Caption         =   "SUPPLIER_ID"
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
            DataField       =   "S_NM"
            Caption         =   "SUPPLIER_NM"
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
         BeginProperty Column02 
            DataField       =   "S_ADD"
            Caption         =   "SUPPLIER_ADDRESS"
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
         BeginProperty Column03 
            DataField       =   "MBL_NO"
            Caption         =   "MOBILE_NUMBER"
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
               ColumnWidth     =   2039.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3119.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3119.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2310.236
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   4320
         Top             =   4680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   "Provider=MSDAORA.1;User ID=medicine/mmp;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=medicine/mmp;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM SUP_DTL ORDER BY S_ID"
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
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF80FF&
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8520
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF80FF&
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8520
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF80FF&
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8520
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF80FF&
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
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8520
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF80FF&
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8520
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
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
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
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
         Left            =   9600
         TabIndex        =   3
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   9600
         TabIndex        =   2
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Left            =   9600
         TabIndex        =   1
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H008080FF&
         BorderWidth     =   6
         Height          =   975
         Left            =   5040
         Top             =   8400
         Width           =   9855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H008080FF&
         BorderWidth     =   6
         Height          =   3375
         Left            =   6000
         Top             =   1440
         Width           =   7935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUPPLIER ADDRESS"
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
         Left            =   7440
         TabIndex        =   14
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUPPLIER NAME"
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
         Left            =   7440
         TabIndex        =   13
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUPPLIER ID"
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
         Left            =   7440
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOBILE NUMBER"
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
         Left            =   7440
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "SUP_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into sup_dtl values('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "'," + Text4.Text + ")"
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Adodc1.Refresh
Command2.Enabled = False
Command1.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command3.Enabled = True
End If
End Sub
Private Sub Command3_Click()
Dim id As String
id = InputBox("enter supplier id to delete the record")
path
S = "select s_id from sup_dtl"
Set R2 = CC.Execute(S)
While Not R2.EOF
If id = Trim(R2.Fields(0)) Then
S = "delete from sup_dtl where s_id='" + id + "'"
Set R = CC.Execute(S)
MsgBox "RECORD DELETED"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Adodc1.Refresh
Exit Sub
Else
R2.MoveNext
End If
Wend
MsgBox "record not found"
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
Private Sub Command4_Click()
Command3.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
MsgBox "Enter Supplier_Id or Supplier_Name To Be Searched !", vbInformation, "MESSAGE"
Text1.SetFocus
End Sub
Private Sub Command5_Click()
Command5.Enabled = False
Text1.Enabled = True
Text1.SetFocus
path
S = "update sup_dtl set s_nm='" + Text2.Text + "',s_add='" + Text3.Text + "',mbl_no=" + Text4.Text + " where s_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
Private Sub Command6_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Text1.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text2.SetFocus
path
S = "select max(to_number(substr(s_id,4,length(s_id)))) from sup_dtl"
Set R = CC.Execute(S)
code = "S00"
If IsNull(R.Fields(0)) Then
I = 1
Text1.Text = code & I
Else
I = R.Fields(0) + 1
Text1.Text = code & I
End If
End Sub

Private Sub Command9_Click()
Unload Me
SUP_DTL.Show
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
Text3.Enabled = False
Text4.Enabled = False
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
path
S = "select s_nm,s_add,mbl_no from sup_dtl where s_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
Exit Sub
Else
Command5.Enabled = True
Text1.Enabled = False
Text2.Text = R.Fields("s_nm")
Text3.Text = R.Fields("s_add")
Text4.Text = R.Fields("mbl_no")
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End If
End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.Text <> "" Then
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
path
S = "select * from sup_dtl where s_nm='" & Text2.Text & "'"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
Exit Sub
Else
Command5.Enabled = True
Text1.Enabled = False
Text1.Text = R.Fields("s_id")
Text3.Text = R.Fields("s_add")
Text4.Text = R.Fields("mbl_no")
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End If
End If
End If
If KeyAscii >= 48 And KeyAscii <= 57 And KeyAscii <> 8 And KeyAscii <> 32 Then
MsgBox "PLEASE ENTER ONLY A TO Z"
KeyAscii = 0
Text2.SetFocus
Else
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
