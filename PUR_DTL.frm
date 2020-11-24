VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PUR_DTL 
   BackColor       =   &H0080FFFF&
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text22 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text21 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text20 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080FF80&
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
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "ADD MORE"
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
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      Caption         =   "MULTIPLE ENTRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   15000
      TabIndex        =   73
      Top             =   8760
      Width           =   3855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FF80&
      Caption         =   "REPORT"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "CONFIRM INVOICE"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "NEW INVOICE"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PURCHASE INVOICE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1.90125e5
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2.45745e5
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FFFF&
         Caption         =   "CONTROL'S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1680
         TabIndex        =   72
         Top             =   8760
         Width           =   9255
      End
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "PUR_DTL.frx":0000
         Left            =   12360
         List            =   "PUR_DTL.frx":000A
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox Text18 
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
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4320
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox Text17 
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
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text16 
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
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text14 
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
         Left            =   17040
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text13 
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
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text12 
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
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text11 
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
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text10 
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox Text9 
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   2655
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   2655
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1440
         Width           =   2655
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3000
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
         Height          =   615
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3600
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
         Height          =   735
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1560
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
         Height          =   735
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
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
         Height          =   735
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   2535
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
         Left            =   2880
         TabIndex        =   1
         Top             =   960
         Width           =   2655
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
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Text15 
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
         Left            =   17040
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   23
         Text            =   "15"
         Top             =   2520
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   5880
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
         RecordSource    =   "select * from pur_dtl order by inv_id"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "PUR_DTL.frx":001C
         Height          =   3255
         Left            =   0
         TabIndex        =   34
         Top             =   5280
         Width           =   20655
         _ExtentX        =   36433
         _ExtentY        =   5741
         _Version        =   393216
         BackColor       =   65280
         Enabled         =   -1  'True
         ForeColor       =   -2147483630
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   32
         BeginProperty Column00 
            DataField       =   "INV_ID"
            Caption         =   "INVOICE_ID"
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
         BeginProperty Column03 
            DataField       =   "PUR_DT"
            Caption         =   "PURCHASE_DATE"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "S_NM"
            Caption         =   "SUPPLIER_NAME"
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
            DataField       =   "QTY"
            Caption         =   "QUANTITY"
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
         BeginProperty Column12 
            DataField       =   "RT"
            Caption         =   "RATE"
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
         BeginProperty Column13 
            DataField       =   "AMT"
            Caption         =   "AMOUNT"
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
         BeginProperty Column14 
            DataField       =   "DISC"
            Caption         =   "DISCOUNT"
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
         BeginProperty Column15 
            DataField       =   "VAT"
            Caption         =   "VAT"
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
         BeginProperty Column16 
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
         BeginProperty Column17 
            DataField       =   "TOT_AMT"
            Caption         =   "TOTAL_AMOUNT"
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
         BeginProperty Column18 
            DataField       =   "ORD_ID"
            Caption         =   "ORDER_ID"
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
         BeginProperty Column19 
            DataField       =   "ORD_DT"
            Caption         =   "ORDER_DATE"
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
         BeginProperty Column20 
            DataField       =   "ADV_PRC"
            Caption         =   "ADVANCE_PRICE"
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
         BeginProperty Column21 
            DataField       =   "REM_AMT"
            Caption         =   "REMAINING_AMOUNT"
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
         BeginProperty Column22 
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
         BeginProperty Column23 
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
         BeginProperty Column24 
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
         BeginProperty Column25 
            DataField       =   "MOD_PTM"
            Caption         =   "MODE_OF_PTM"
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
         BeginProperty Column26 
            DataField       =   "BNK_NM"
            Caption         =   "BANK_NAME"
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
         BeginProperty Column27 
            DataField       =   "BH_NM"
            Caption         =   "BRANCH_NAME"
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
         BeginProperty Column28 
            DataField       =   "AH_NM"
            Caption         =   "ACCOUNT_HOLDER_NAME"
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
         BeginProperty Column29 
            DataField       =   "ACC_NO"
            Caption         =   "ACCOUNT_NO"
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
         BeginProperty Column30 
            DataField       =   "CHQ_NO"
            Caption         =   "CHEQUE_NO"
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
         BeginProperty Column31 
            DataField       =   "IFSC"
            Caption         =   "IFSC"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column26 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column27 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column28 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column29 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column30 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column31 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   615
         Left            =   2880
         TabIndex        =   3
         Top             =   2160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   42714
      End
      Begin MSComCtl2.DTPicker DTP3 
         Height          =   615
         Left            =   7680
         TabIndex        =   12
         Top             =   4200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   42714
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   615
         Left            =   2880
         TabIndex        =   5
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   42714
      End
      Begin MSComCtl2.DTPicker DTP4 
         Height          =   615
         Left            =   12360
         TabIndex        =   13
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   42714
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   6
         Height          =   1455
         Left            =   14880
         Top             =   8640
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   6
         Height          =   1455
         Left            =   1560
         Top             =   8640
         Width           =   9495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   6
         Height          =   4695
         Left            =   600
         Top             =   480
         Width           =   19215
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REMAINING QUANTITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   70
         Top             =   4560
         Width           =   1665
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Left            =   19080
         TabIndex        =   68
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Left            =   19080
         TabIndex        =   67
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   17040
         TabIndex        =   66
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   17040
         TabIndex        =   65
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   17040
         TabIndex        =   64
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   17040
         TabIndex        =   63
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   17040
         TabIndex        =   62
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R.S."
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
         Left            =   12360
         TabIndex        =   61
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MODE OF PAYMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10200
         TabIndex        =   60
         Top             =   3240
         Width           =   2145
      End
      Begin VB.Label Label26 
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
         Height          =   600
         Left            =   10200
         TabIndex        =   59
         Top             =   2640
         Width           =   2145
      End
      Begin VB.Label Label25 
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
         Height          =   600
         Left            =   10200
         TabIndex        =   58
         Top             =   2040
         Width           =   2145
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MFG DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5520
         TabIndex        =   57
         Top             =   4200
         Width           =   2145
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ORDER ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   720
         TabIndex        =   56
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ORDER DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   720
         TabIndex        =   55
         Top             =   2160
         Width           =   2145
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REMAINING AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   15000
         TabIndex        =   54
         Top             =   4320
         Width           =   2025
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   15000
         TabIndex        =   53
         Top             =   3120
         Width           =   2025
      End
      Begin VB.Label Label17 
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
         Height          =   600
         Left            =   10200
         TabIndex        =   52
         Top             =   4440
         Width           =   2145
      End
      Begin VB.Label Label16 
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
         Height          =   600
         Left            =   15000
         TabIndex        =   51
         Top             =   2520
         Width           =   2025
      End
      Begin VB.Label Label15 
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
         Height          =   600
         Left            =   15000
         TabIndex        =   50
         Top             =   1920
         Width           =   2025
      End
      Begin VB.Label Label14 
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
         Height          =   600
         Left            =   15000
         TabIndex        =   49
         Top             =   1320
         Width           =   2025
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   15000
         TabIndex        =   48
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label Label12 
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
         Height          =   600
         Left            =   10200
         TabIndex        =   47
         Top             =   3840
         Width           =   2145
      End
      Begin VB.Label Label6 
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
         Height          =   720
         Left            =   5520
         TabIndex        =   46
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Label Label7 
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
         Height          =   720
         Left            =   5520
         TabIndex        =   45
         Top             =   2280
         Width           =   2145
      End
      Begin VB.Label Label8 
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
         Height          =   600
         Left            =   5520
         TabIndex        =   44
         Top             =   3000
         Width           =   2145
      End
      Begin VB.Label Label9 
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
         Height          =   600
         Left            =   5520
         TabIndex        =   43
         Top             =   3600
         Width           =   2145
      End
      Begin VB.Label Label10 
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
         Height          =   600
         Left            =   10200
         TabIndex        =   42
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MFG NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10200
         TabIndex        =   41
         Top             =   1440
         Width           =   2145
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
         Height          =   600
         Left            =   720
         TabIndex        =   40
         Top             =   960
         Width           =   2145
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
         Height          =   600
         Left            =   720
         TabIndex        =   39
         Top             =   3960
         Width           =   2145
      End
      Begin VB.Label Label3 
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
         Height          =   600
         Left            =   720
         TabIndex        =   38
         Top             =   2760
         Width           =   2145
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PURCHASE DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   720
         TabIndex        =   37
         Top             =   3360
         Width           =   2145
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT  NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5520
         TabIndex        =   36
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADVANCE PRICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   15000
         TabIndex        =   35
         Top             =   3720
         Width           =   2025
      End
   End
End
Attribute VB_Name = "PUR_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qty As Integer
Private Sub Combo1_Click()
Combo1.Enabled = False
If Command1.Enabled = False Then
path
S = "select ord_Dt,s_id from ord_Dtl where ord_id='" + Trim(Combo1.Text) + "' "
Set R = CC.Execute(S)
Combo2.AddItem R.Fields("s_id")
DTP1.Value = R.Fields("ord_dt")
path
T = "SELECT P_ID FROM ORD_dTL WHERE ORD_ID='" + Trim(Combo1.Text) + "'"
Set R2 = CC.Execute(T)
While R2.EOF = False
Combo3.AddItem R2.Fields("p_id")
R2.MoveNext
Wend
End If
If Command3.Enabled = True And Command1.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from PUR_dtl where ORD_id='" & Combo1.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
'Command4.Enabled = True
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox " Record Not FOund ! ", vbCritical, "MESSAGE"
Command3.Enabled = True
End If
End If
End Sub
Private Sub Combo3_Click()
S = "SELECT REM_QTY FROM STK_DTL WHERE P_ID='" & Trim(Combo3.Text) & "' "
Set R = CC.Execute(S)
Dim RQ As Integer
RQ = R.Fields("REM_QTY")
Text19.Text = RQ
If Command1.Enabled = False Then
path
S = "select * from prd_dtl where p_id='" + Trim(Combo3.Text) + "'  "
Set R = CC.Execute(S)
Text2.Text = R.Fields("p_nm")
Text5.Text = R.Fields("pack")
Text6.Text = R.Fields("bth_no")
DTP3.Value = R.Fields("MNF_dt")
DTP4.Value = R.Fields("exp_dt")
Text7.Text = R.Fields("mnf_nm")
Text8.Text = R.Fields("TYPE")
Text9.Text = R.Fields("wt")
path
S = "select * from ORD_dtl where p_id='" + Trim(Combo3.Text) + "' "
Set R = CC.Execute(S)
Text10.Text = R.Fields("QTY")
Text11.Text = R.Fields("mrp")
Text12.Text = R.Fields("Rt")
Text14.Text = R.Fields("DISC")
Text15.Text = R.Fields("VAT")
Text16.Text = R.Fields("TOT_AMT")
Text17.Text = R.Fields("ADV_PRC")
Text18.Text = R.Fields("REM_AMT")
Combo4.Text = R.Fields("mod_ptm")
qty = Val(Text10.Text)
End If
If Val(Text17.Text) = 0 Then
    Text18.Text = Val(Text16.Text)
End If
End Sub
Private Sub Combo2_Click()
If Command1.Enabled = False Then
path
S = "select S_nm from SUP_dtl where  S_id='" + Trim(Combo2.Text) + "'"
Set R = CC.Execute(S)
Text4.Text = R.Fields("S_nm")
path
S = "select S_add from SUP_dtl where S_id='" + Trim(Combo2.Text) + "'"
Set R = CC.Execute(S)
Text3.Text = R.Fields("S_add")
End If
If Command3.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from PUR_dtl where S_id='" & Combo2.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "record not found ! ", vbCritical, "ERROR"
Command3.Enabled = True
End If
End If
End Sub
Private Sub Combo4_Click()
If Combo4.ListIndex = 1 And Val(Text18.Text) <> 0 Then
c1 = 2
PUR_DTL.Visible = False
CHQ_DTL.Show
End If
End Sub
Private Sub Command1_Click()
Text20.Visible = False
Text21.Visible = False
Text22.Visible = False
Text17.Text = 0
Text18.Text = 0
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command5.Enabled = True
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
Combo1.SetFocus
path
S = "select max(to_number(substr(inv_id,4,length(inv_id)))) from PUR_dtl"
Set R = CC.Execute(S)
code = "INV"
If IsNull(R.Fields(0)) Then
I = 1
Text1.Text = code & I
Else
I = R.Fields(0) + 1
Text1.Text = code & I
End If
End Sub
Private Sub Command10_Click()
Command10.Enabled = False
Text1.Text = ""
If Text1.Text = "" Then
MsgBox "ENTER INVOICE_ID IN TEXTBOX"
Text1.Enabled = True
Text1.SetFocus
End If
End Sub
Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command5.Enabled = True
If Command8.Enabled = False Then
S = "select tot_amt,rem_amt,adv_prc from temp_ord_dtl where ord_id='" & Trim(Combo1.Text) & "'"
Set R = CC.Execute(S)
'Text16.Text = tot_amt
Dim tot_amt As Double
tot_amt = R.Fields("tot_amt")
MsgBox "The total amount of INVOICE_ID " & Text1.Text & " is " & R.Fields("TOT_AMT")
MsgBox "The ADVANCE PRICE of INVOICE_ID " & Text1.Text & " is " & R.Fields("ADV_PRC")
MsgBox "The REMAINING amount of INVOICE_ID " & Text1.Text & " is " & R.Fields("REM_AMT")
U = "insert into temp_pur_Dtl values('" & Text1.Text & "'," & tot_amt & ")"
Set R3 = CC.Execute(U)
Exit Sub
Else
If Text1.Text = "" Or Combo1.Text = "" Or Text2.Text = "" Or DTP1.Value = "" Or DTP2.Value = "" Or DTP3.Value = "" Or DTP4.Value = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Then
MsgBox "ENTER ALL FIELDS"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command5.Enabled = False
Exit Sub
Else
path
S = "insert into PUR_dtl values('" + Text1.Text + "','" + Combo3.Text + "','" + Combo2.Text + "','" + Format(DTP2.Value, "dd-mmm-yyyy") + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text5.Text + "','" + Text6.Text + "','" + Format(DTP4.Value, "dd-mmm-yyyy") + "','" + Text7.Text + "'," + Text10.Text + "," + Text12.Text + "," + Text13.Text + "," + Text14.Text + "," + Text15.Text + "," + Text11.Text + "," + Text16.Text + ",'" + Combo1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "'," + Text17.Text + "," + Text18.Text + ",'" + Format(DTP3.Value, "dd-mmm-yyyy") + "','" + Text8.Text + "','" + Text9.Text + "','" + Combo4.Text + "','" + CHQ_DTL.Text1.Text + "','" + CHQ_DTL.Text2.Text + "','" + CHQ_DTL.Text3.Text + "','" + CHQ_DTL.Text4.Text + "','" + CHQ_DTL.Text5.Text + "','" + CHQ_DTL.Text6.Text + "' ) "
Set R = CC.Execute(S)
Unload CHQ_DTL
MsgBox "RECORD SAVED"
DTP1.Enabled = False
DTP2.Enabled = False
DTP3.Enabled = False
DTP4.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Adodc1.Refresh
Command2.Enabled = False
Command1.Enabled = True
Text2.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
DTP1.Enabled = False
DTP2.Enabled = False
DTP3.Enabled = False
DTP4.Enabled = False
Command3.Enabled = True
S = "SELECT REM_QTY FROM STK_DTL WHERE P_ID='" & Trim(Combo3.Text) & "'"
Set R = CC.Execute(S)
Dim QTY2, TQ As Integer
QTY2 = R.Fields("REM_QTY")
TQ = qty + QTY2
T = "UPDATE STK_DTL SET REM_QTY=" & TQ & " WHERE P_ID='" & Trim(Combo3.Text) & "' "
Set R2 = CC.Execute(T)
Text19.Text = Val(Text19.Text) + qty
MsgBox "STOCK UPDATED"
If MsgBox("do you want to continue for Multiple PURCHASE", vbYesNo, "message") = vbYes Then
MsgBox "Click Add More Button ", vbInformation, "MESSAGE"
Command7.SetFocus
Command1.Enabled = False
Command2.Enabled = True
Else
S = "SELECT TOT_AMT FROM PUR_DTL WHERE INV_ID='" & Text1.Text & "'"
Set R = CC.Execute(S)
'Text16.Text = R.Fields(" tot_amt")
U = "insert into temp_Pur_Dtl values('" & Text1.Text & "'," & R.Fields("tot_amt") & ")"
Set R3 = CC.Execute(U)
End If
End If
End If
End Sub
Private Sub Command3_Click()
Command3.Enabled = True
DTP1.Enabled = False
DTP2.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Text2.Enabled = False
MsgBox "Enter  Invoice_Id OR Order_Id or Purchase_Date for searching", vbInformation, "MESSAGE"
Text1.SetFocus
End Sub
Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Command7_Click()
Command7.Enabled = False
Command8.Enabled = True
Text1.Enabled = False
Combo1.Enabled = False
DTP1.Enabled = False
Text2.Enabled = False
DTP2.Enabled = False
Combo3.Enabled = True
End Sub
Private Sub Command8_Click()
If Combo2.Text <> "" And Combo3.Text <> "" Then
CHQ_DTL.Visible = False
PUR_DTL.WindowState = 2 - MAXIMIZED
Command8.Enabled = False
Command7.Enabled = True
path
S = "insert into PUR_dtl values('" + Text1.Text + "','" + Combo3.Text + "','" + Combo2.Text + "','" + Format(DTP2.Value, "dd-mmm-yyyy") + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text5.Text + "','" + Text6.Text + "','" + Format(DTP4.Value, "dd-mmm-yyyy") + "','" + Text7.Text + "'," + Text10.Text + "," + Text12.Text + "," + Text13.Text + "," + Text14.Text + "," + Text15.Text + "," + Text11.Text + "," + Text16.Text + ",'" + Combo1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "'," + Text17.Text + "," + Text18.Text + ",'" + Format(DTP3.Value, "dd-mmm-yyyy") + "','" + Text8.Text + "','" + Text9.Text + "','" + Combo4.Text + "','" + CHQ_DTL.Text1.Text + "','" + CHQ_DTL.Text2.Text + "','" + CHQ_DTL.Text3.Text + "','" + CHQ_DTL.Text4.Text + "','" + CHQ_DTL.Text5.Text + "','" + CHQ_DTL.Text6.Text + "' ) "
Set R = CC.Execute(S)
MsgBox "RECORD SAVED"
Adodc1.Refresh
S = "SELECT REM_QTY FROM STK_DTL WHERE P_ID='" & Trim(Combo3.Text) & "'"
Set R = CC.Execute(S)
Dim QTY2, TQ As Integer
QTY2 = R.Fields("REM_QTY")
TQ = qty + QTY2
T = "UPDATE STK_DTL SET REM_QTY=" & TQ & " WHERE P_ID='" & Trim(Combo3.Text) & "' "
Set R2 = CC.Execute(T)
Text19.Text = Val(Text19.Text) + qty
MsgBox "STOCK UPDATED"
Else
MsgBox "Enter All Fields", vbInformation, "Message"
End If
End Sub
Private Sub Command9_Click()
Unload PUR_DTL
PUR_DTL.Show
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT *  FROM pur_dtl ORDER BY inv_ID"
Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command5.Enabled = True
DTP1.Enabled = False
DTP2.Enabled = False
DTP3.Enabled = False
DTP4.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text20.Visible = False
Text21.Visible = False
Text22.Visible = False
DTP1.Value = Date
DTP2.Value = Date
DTP3.Value = Date
DTP4.Value = Date
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Command1.Enabled = True
Text20.Visible = True
Text21.Visible = True
Text22.Visible = True
Text1.Text = DataGrid1.Columns(0).Text
Text22.Text = DataGrid1.Columns(1).Text
Text21.Text = DataGrid1.Columns(2).Text
DTP2.Value = DataGrid1.Columns(3).Text
Text2.Text = DataGrid1.Columns(4).Text
Text3.Text = DataGrid1.Columns(5).Text
Text4.Text = DataGrid1.Columns(6).Text
Text5.Text = DataGrid1.Columns(7).Text
Text6.Text = DataGrid1.Columns(8).Text
DTP4.Value = DataGrid1.Columns(9).Text
Text7.Text = DataGrid1.Columns(10).Text
Text10.Text = DataGrid1.Columns(11).Text
Text12.Text = DataGrid1.Columns(12).Text
Text13.Text = DataGrid1.Columns(13).Text
Text14.Text = DataGrid1.Columns(14).Text
Text15.Text = DataGrid1.Columns(15).Text
Text11.Text = DataGrid1.Columns(16).Text
Text16.Text = DataGrid1.Columns(17).Text
Text20.Text = DataGrid1.Columns(18).Text
DTP1.Value = DataGrid1.Columns(19).Text
Text17.Text = DataGrid1.Columns(20).Text
Text18.Text = DataGrid1.Columns(21).Text
DTP3.Value = DataGrid1.Columns(22).Text
Text8.Text = DataGrid1.Columns(23).Text
Text9.Text = DataGrid1.Columns(24).Text
Combo4.Text = DataGrid1.Columns(25).Text
Command5.Enabled = True
If Command5.Enabled = True Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
DTP1.Enabled = True
DTP2.Enabled = True
DTP3.Enabled = True
DTP4.Enabled = True
End If
Unload CHQ_DTL
PUR_DTL.Show
End Sub
Private Sub DTP1_Click()
If Command3.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from inv_dtl where sl_dt='" & Format(DTP1.Value, "dd-mmm-yyyy") & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox " Record Not FOund ! ", vbCritical, "MESSAGE"
Exit Sub
End If
Set DataGrid1.DataSource = Adodc1
End If
Command3.Enabled = True
End Sub
Private Sub DTP2_Click()
If Command3.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from PUR_dtl where PUR_DT='" & Format(DTP2.Value, "DD-MMM-YYYY") & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "record not found ! ", vbCritical, "ERROR"
Command3.Enabled = True
End If
End If
End Sub
Private Sub Form_Load()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command5.Enabled = True
DTP1.Enabled = False
DTP2.Enabled = False
DTP3.Enabled = False
DTP4.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
DTP1.Value = Date
DTP2.Value = Date
DTP3.Value = Date
DTP4.Value = Date
path
S = "select distinct(ORD_id) from ORD_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from PUR_dtl where inv_id='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "record not found"
Exit Sub
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
Command5.Enabled = True
End If
End If
End If
If Command10.Enabled = False Then
If Text1.Text <> "" Then
If KeyAscii = 13 Then
If DataEnvironment1.rsCommand6.State = 1 Then DataEnvironment1.rsCommand6.Close
DataEnvironment1.Command6 Text1.Text
Set R = New ADODB.Recordset
S = "select net_prc from temp_pur_dtl where inv_id='" & Text1.Text & "'"
Set R = CC.Execute(S)
DataReport6.Sections("section5").Controls("label10").Caption = R.Fields(0)
T = "SELECT S_NM,S_ADD,MOD_PTM,INV_ID FROM PUR_dTL WHERE INV_ID='" & Text1.Text & "'"
Set R2 = CC.Execute(T)
DataReport6.Sections("section2").Controls("label2").Caption = R2.Fields("S_NM")
DataReport6.Sections("section2").Controls("label4").Caption = R2.Fields("S_ADD")
DataReport6.Sections("section2").Controls("label5").Caption = R2.Fields("MOD_PTM")
DataReport6.Sections("section2").Controls("label8").Caption = R2.Fields("INV_ID")
DataReport6.Show
Command10.Enabled = True
End If
End If
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 1 TO 9"
End If
If KeyAscii = 46 Or KeyAscii = 48 Then
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 1 TO 9"
End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text12_Change()
Dim q, R, a As Double
q = Val(Text10.Text)
R = Val(Text12.Text)
a = q * R
Text13.Text = a
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text15_Change()
Dim d, a, v, TA As Double
a = Val(Text13.Text)
d = Val(Text14.Text)
v = Val(Text15.Text)
TA = a + a * (v / 100) - a * (d / 100)
Text16.Text = Round(TA, 2)
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text17_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.Text <> "" Then
If KeyAscii = 13 Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from cus_ord_dtl where  c_nm='" + Text2.Text + "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
Command5.Enabled = True
End If
End If
End Sub
Private Sub Text22_Change()
If Text22.Text <> "" Then
S = "select rem_qty from stk_dtl where p_id='" & Text22.Text & "'"
Set R = CC.Execute(S)
Text19.Text = R.Fields("rem_qty")
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
