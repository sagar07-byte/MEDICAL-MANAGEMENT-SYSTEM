VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form STK_DTL 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
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
      Height          =   15735
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "STK_DTL.frx":0000
         Height          =   3135
         Left            =   2880
         TabIndex        =   17
         Top             =   4320
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   5530
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "STK_ID"
            Caption         =   "STOCK_ID"
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
            DataField       =   "P_ID"
            Caption         =   "PRODUCT_ID"
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
            DataField       =   "CPT"
            Caption         =   "CAPACITY"
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
            DataField       =   "REM_QTY"
            Caption         =   "REMAINING_QUANTITY"
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
         BeginProperty Column04 
            DataField       =   "P_NM"
            Caption         =   "PRODUCT_NAME"
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
               ColumnWidth     =   2039.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2984.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2984.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3119.811
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text4 
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
         Left            =   12600
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox Text3 
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
         Left            =   12600
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
      End
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
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8160
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2880
         Top             =   3960
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         RecordSource    =   "select * from stk_dtl order by STK_ID"
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
         TabIndex        =   10
         Top             =   8160
         Width           =   1815
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
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8160
         Width           =   1815
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8160
         Width           =   1815
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8160
         Width           =   1815
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8160
         Width           =   1935
      End
      Begin VB.TextBox Text2 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox Text1 
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
         Left            =   6000
         TabIndex        =   1
         Top             =   1560
         Width           =   2775
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
         Left            =   6000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         Height          =   975
         Left            =   4200
         Top             =   8040
         Width           =   11175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         Height          =   2535
         Left            =   2520
         Top             =   1080
         Width           =   14175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT NAME"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REMAINING QUANTITY"
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
         Left            =   10440
         TabIndex        =   15
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STOCK ID"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         Left            =   10440
         TabIndex        =   12
         Top             =   1560
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
S = "SELECT P_NM FROM PRD_dTL WHERE P_ID='" & Trim(Combo1.Text) & "'"
Set R = CC.Execute(S)
Text2.Text = R.Fields("P_NM")
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
Combo1.Enabled = True
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.SetFocus
path
S = "select max(to_number(substr(stk_id,4,length(stk_id)))) from stk_dtl"
Set R = CC.Execute(S)
code = "STK"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into stk_dtl values('" + Text1.Text + "','" + Combo1.Text + "'," + Text3.Text + ",'" + Text4.Text + "','" + Text2.Text + "')"
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.Enabled = False
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
id = InputBox("Enter Stock_Id to be Deleted ! ")
id = UCase(id)
path
S = "select stk_id from stk_dtl"
Set R2 = CC.Execute(S)
While Not R2.EOF
If id = Trim(R2.Fields(0)) Then
S = "delete from stk_dtl where stk_id='" & id & "'"
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
MsgBox "Record Not Found ! ", vbCritical, "Message"
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.Enabled = False
End Sub
Private Sub Command4_Click()
Command3.Enabled = True
Combo1.Enabled = True
MsgBox "Enter Stock Id to be Searched", vbInformation, "MESSAGE"
Text1.Enabled = True
Text1.SetFocus
End Sub
Private Sub Command5_Click()
Command5.Enabled = False
Combo1.Enabled = False
'Text1.Text = False
'Text2.Text = False
path
S = "update stk_dtl set cpt=" & Text3.Text & ",rem_qty=" & Text4.Text & " where stk_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !", vbOKCancel, "Message"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
End Sub
Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command9_Click()
Unload Me
STK_DTL.Show
End Sub

Private Sub Form_Load()
Adodc1.Refresh
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
Combo1.Enabled = False
path
S = "select p_id from prd_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo1.AddItem R.Fields("p_id")
R.MoveNext
Wend
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
path
S = "select p_id,p_nm,cpt,rem_qty from stk_dtl where stk_id='" + Text1.Text + "'"
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
Text2.Text = R.Fields("p_nm")
Text3.Text = R.Fields("cpt")
Text4.Text = R.Fields("rem_qty")
Combo1.Text = R.Fields("p_id")
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End If
End If
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

