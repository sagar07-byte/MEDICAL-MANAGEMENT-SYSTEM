VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PRD_DTL 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10920
   ScaleWidth      =   16770
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "PRODUCT DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "PRD_DLT.frx":0000
         Height          =   4095
         Left            =   960
         TabIndex        =   28
         Top             =   4320
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   65280
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "BTH_NO"
            Caption         =   "BATCH_NO"
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
            DataField       =   "EXP_DT"
            Caption         =   "EXPIRY_DATE"
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
            DataField       =   "PACK"
            Caption         =   "PACK"
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
         BeginProperty Column05 
            DataField       =   "MNF_NM"
            Caption         =   "MNF_NAME"
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
         BeginProperty Column06 
            DataField       =   "MNF_DT"
            Caption         =   "MNF_DATE"
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
         BeginProperty Column07 
            DataField       =   "WT"
            Caption         =   "WEIGHT"
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
         BeginProperty Column08 
            DataField       =   "TYPE"
            Caption         =   "TYPE"
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
         BeginProperty Column09 
            DataField       =   "MRP"
            Caption         =   "MRP"
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
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1995.024
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text8 
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
         Left            =   16320
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text7 
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
         Left            =   15720
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15720
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1200
         Top             =   3600
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
         RecordSource    =   "select * from prd_dtl order by p_id asc"
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
         Height          =   735
         Left            =   4560
         TabIndex        =   4
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1296
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   42714
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF8080&
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
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   9120
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
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
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   9120
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   9120
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   9120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   9120
         Width           =   1935
      End
      Begin VB.TextBox Text5 
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
         Left            =   10080
         TabIndex        =   7
         Top             =   2280
         Width           =   2535
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
         Left            =   10080
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
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
         Left            =   4560
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
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
         Left            =   4560
         TabIndex        =   2
         ToolTipText     =   "ENTER PRODUCT NAME"
         Top             =   1560
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   615
         Left            =   10080
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   42714
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF80FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         Height          =   975
         Left            =   4080
         Top             =   9000
         Width           =   11175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         Height          =   3615
         Left            =   960
         Top             =   480
         Width           =   18375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15720
         TabIndex        =   27
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WEIGHT"
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
         Left            =   7920
         TabIndex        =   26
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M.R.P."
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
         Left            =   13440
         TabIndex        =   25
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MANUFACTURING DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   24
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TYPE"
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
         Left            =   13440
         TabIndex        =   23
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MANUFACTURER NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13440
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXPIRY DATE"
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
         Left            =   7920
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BATCH NO"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label2 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PACK"
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
         Left            =   7920
         TabIndex        =   17
         Top             =   1680
         Width           =   2175
      End
   End
End
Attribute VB_Name = "PRD_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
DTP1.Enabled = True
DTP2.Enabled = True
Text2.SetFocus
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
path
S = "select max(to_number(substr(p_id,4,length(p_id)))) from prd_dtl"
Set R = CC.Execute(S)
code = "P00"
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
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or DTP1.Value = "" Or DTP2.Value = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into prd_dtl values('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "','" + Text4.Text + "','" + Text5.Text + "','" + Format(DTP2.Value, "dd-mmm-yyyy") + "','" + Text6.Text + "','" + Text7.Text + "'," + Text8.Text + ")"
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
DTP1.Enabled = False
DTP2.Enabled = False
Adodc1.Refresh
Command2.Enabled = False
Command1.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
DTP1.Enabled = False
Command3.Enabled = True
End If
End Sub
Private Sub Command3_Click()
Dim id As String
id = InputBox("Enter Product_Id to be Deleted ! ")
path
S = "select p_id from prd_dtl"
Set R2 = CC.Execute(S)
While Not R2.EOF
If id = Trim(R2.Fields(0)) Then
S = "delete from prd_dtl where p_id='" + id + "'"
Set R = CC.Execute(S)
MsgBox "RECORD DELETED"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
DTP1.Value = ""
DTP2.Value = ""
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
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
DTP1.Enabled = False
DTP2.Enabled = False
End Sub
Private Sub Command4_Click()
Command3.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
MsgBox "Enter Product Id or Product Name To Be Searched", vbInformation, "MESSAGE"
Text1.SetFocus
End Sub
Private Sub Command5_Click()
Command5.Enabled = False
Text1.Enabled = True
Text1.SetFocus
path
S = "update prd_dtl set p_nm='" + Text2.Text + "',bth_no='" + Text3.Text + "',mnf_dt='" + Format(DTP1.Value, "dd-mmm-yyyy") + "',exp_dt='" + Format(DTP2.Value, "dd-mmm-yyyy") + "',pack='" + Text4.Text + "',wt='" + Text5.Text + "'  ,mnf_nm='" + Text6.Text + "',type='" + Text7.Text + "',mrp=" + Text8.Text + " where p_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !", vbOKCancel, "MESSAGE"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub
Private Sub Command6_Click()
Unload Me
End Sub
Private Sub Command7_Click()
DataReport1.Show
End Sub

Private Sub Command9_Click()
Unload Me
PRD_DTL.Show
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
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
DTP1.Enabled = False
DTP2.Enabled = False
DTP1.Value = Date
DTP2.Value = Date
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
path
S = "select * from prd_dtl where p_id= '" + Text1.Text + " '"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    DTP1.Value = ""
Exit Sub
Else
Command5.Enabled = True
Text2.Text = R.Fields("p_nm")
Text3.Text = R.Fields("bth_no")
DTP2.Value = R.Fields("exp_dt")
DTP1.Value = R.Fields("exp_dt")
Text4.Text = R.Fields("pack")
Text5.Text = R.Fields("wt")
Text6.Text = R.Fields("mnf_nm")
Text7.Text = R.Fields("type")
Text8.Text = R.Fields("mrp")
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
DTP1.Enabled = True
DTP2.Enabled = True
End If
End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.Text <> "" Then
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
path
S = "select * from prd_dtl where p_nm= '" & Text2.Text & "'"
Set R = CC.Execute(S)
If R.EOF Or R.BOF = True Then
    MsgBox "Record Not Found ! Try Again", vbCritical, "Message"
    Command5.Enabled = False
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    DTP1.Value = ""
Exit Sub
Else
Command5.Enabled = True
Text1.Text = R.Fields("p_id")
Text3.Text = R.Fields("bth_no")
DTP2.Value = R.Fields("exp_dt")
DTP1.Value = R.Fields("exp_dt")
Text4.Text = R.Fields("pack")
Text5.Text = R.Fields("wt")
Text6.Text = R.Fields("mnf_nm")
Text7.Text = R.Fields("type")
Text8.Text = R.Fields("mrp")
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
DTP1.Enabled = True
DTP2.Enabled = True
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 And KeyAscii <> 8 And KeyAscii <> 32 Then
MsgBox "PLEASE ENTER ONLY A TO Z"
KeyAscii = 0
Text2.SetFocus
Else
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 And KeyAscii <> 8 And KeyAscii <> 32 Then
MsgBox "PLEASE ENTER ONLY A TO Z"
KeyAscii = 0
Text2.SetFocus
Else
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If

End Sub
