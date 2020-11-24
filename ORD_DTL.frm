VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ORD_DTL 
   BackColor       =   &H00FFC0FF&
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "CONFIRM ORDER"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "NEW ORDER"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Frame Frame4 
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
      Left            =   360
      TabIndex        =   58
      Top             =   8880
      Width           =   11055
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
      Left            =   15720
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FF80&
      Caption         =   "COLLECTIVE"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
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
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
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
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9240
      Width           =   1935
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
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SUPPLIER ORDER DETAILS"
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
         Left            =   15720
         TabIndex        =   60
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080FF80&
         Caption         =   "SELECTIVE"
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
         Left            =   13800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   9240
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000FFFF&
         Caption         =   "REPORT"
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
         Left            =   11880
         TabIndex        =   57
         Top             =   8880
         Width           =   3735
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3720
         Width           =   2055
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
         Left            =   15000
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
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
         Left            =   4200
         TabIndex        =   54
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   4200
         TabIndex        =   51
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   15720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3120
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
         Left            =   15720
         TabIndex        =   15
         Top             =   2520
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "ORD_DTL.frx":0000
         Height          =   3975
         Left            =   0
         TabIndex        =   34
         Top             =   4680
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   7011
         _Version        =   393216
         BackColor       =   65280
         Enabled         =   -1  'True
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
         ColumnCount     =   23
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column15 
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
         BeginProperty Column16 
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
         BeginProperty Column17 
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
         BeginProperty Column18 
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
         BeginProperty Column19 
            DataField       =   "AH_NM"
            Caption         =   "ACCOUNT_HOLDER_NM"
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
         BeginProperty Column21 
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
         BeginProperty Column22 
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column22 
            EndProperty
         EndProperty
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
         ItemData        =   "ORD_DTL.frx":0015
         Left            =   15720
         List            =   "ORD_DTL.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3720
         Width           =   2055
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
         Left            =   15000
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "15"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame Frame2 
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
         Left            =   16080
         TabIndex        =   39
         Top             =   8880
         Width           =   3975
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
         Left            =   10320
         TabIndex        =   10
         Top             =   3120
         Width           =   2055
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
         Left            =   9600
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2520
         Width           =   2775
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   4200
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
         RecordSource    =   "SELECT * FROM ORD_DTL ORDER BY ORD_ID"
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
         Left            =   4200
         TabIndex        =   2
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   54788097
         CurrentDate     =   42714
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3360
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2775
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
         Left            =   4200
         TabIndex        =   4
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
         Left            =   4200
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         Height          =   1455
         Left            =   11760
         Top             =   8760
         Width           =   3975
      End
      Begin VB.Label Label25 
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
         Height          =   495
         Left            =   9600
         TabIndex        =   56
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label24 
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
         Left            =   7440
         TabIndex        =   55
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         Height          =   3975
         Left            =   1560
         Top             =   480
         Width           =   16935
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         Height          =   1455
         Left            =   15960
         Top             =   8760
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         Height          =   1455
         Left            =   240
         Top             =   8760
         Width           =   11295
      End
      Begin VB.Label Label23 
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
         Left            =   17040
         TabIndex        =   53
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label22 
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
         Left            =   17040
         TabIndex        =   52
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label21 
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
         Left            =   9600
         TabIndex        =   50
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label20 
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
         Height          =   615
         Left            =   7440
         TabIndex        =   49
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label19 
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
         Left            =   12840
         TabIndex        =   48
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label18 
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
         Left            =   12840
         TabIndex        =   47
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label11 
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
         Height          =   615
         Left            =   12840
         TabIndex        =   46
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label17 
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
         Height          =   615
         Left            =   12840
         TabIndex        =   45
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label12 
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
         Height          =   615
         Left            =   12840
         TabIndex        =   44
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label10 
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
         Height          =   615
         Left            =   12840
         TabIndex        =   43
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label9 
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
         Left            =   15000
         TabIndex        =   42
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label15 
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
         Left            =   15000
         TabIndex        =   41
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label16 
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
         Left            =   15000
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label14 
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
         Left            =   9600
         TabIndex        =   38
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label13 
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
         Left            =   7440
         TabIndex        =   37
         Top             =   1920
         Width           =   2175
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
         Height          =   615
         Left            =   7440
         TabIndex        =   36
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
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
         Left            =   7440
         TabIndex        =   35
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         Height          =   615
         Left            =   2040
         TabIndex        =   33
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Left            =   2040
         TabIndex        =   32
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         Left            =   2040
         TabIndex        =   31
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         Left            =   7440
         TabIndex        =   30
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         Height          =   615
         Left            =   2040
         TabIndex        =   29
         Top             =   960
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
         Left            =   2040
         TabIndex        =   28
         Top             =   3360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ORD_DTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C  As Integer
Private Sub Combo2_Click()
path
S = "select * from prd_dtl where p_id='" + Trim(Combo2.Text) + "'"
Set R = CC.Execute(S)
Text3.Text = R.Fields("p_nm")
Text4.Text = R.Fields("pack")
Text5.Text = R.Fields("mrp")
End Sub
Private Sub Combo1_Click()
If Command1.Enabled = False Then
path
S = "select s_nm from sup_dtl where s_id='" + Trim(Combo1.Text) + "'"
Set R = CC.Execute(S)
Text2.Text = R.Fields("s_nm")
End If
If Command4.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ord_dtl where s_id='" & Combo1.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
Command5.Enabled = True
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox " Record Not FOund ! ", vbCritical, "MESSAGE"
Command4.Enabled = True
End If
End If
End Sub
Private Sub Combo3_Click()
If Combo3.ListIndex = 1 Then
c1 = 1
ORD_DTL.Visible = False
CHQ_DTL.Show
End If
End Sub
Private Sub Command1_Click()
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
Text14.Visible = False
Text15.Visible = False
Text8.Text = 15
Text12.Text = 0
Text13.Text = 0
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Text1.SetFocus
path
S = "select max(to_number(substr(ord_id,4,length(ord_id)))) from ord_dtl"
Set R = CC.Execute(S)
code = "ORD"
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
MsgBox "ENTER ORDER_ID IN TEXTBOX"
Text1.Enabled = True
Text1.SetFocus
End If
End Sub
Private Sub Command11_Click()
Unload ORD_DTL
ORD_DTL.Show
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
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT *  FROM ord_dtl ORDER BY ORD_ID"
Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
DTP1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
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
Text14.Visible = False
Text15.Visible = False
DTP1.Value = Date
Command6.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
End Sub
Private Sub Command2_Click()
Dim amt, VAT, DISC As Double
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
If Command9.Enabled = False Then
    path
    S = "select TOT_amt from ord_dtl where ord_id='" + Text1.Text + "'"
    Set R = CC.Execute(S)
    While R.EOF = False
    amt = amt + R.Fields(0)
    R.MoveNext
    Wend
    Dim TA, ADV, ra As Double
    MsgBox "total amount of " & Text1.Text & " is " & amt
    ADV = InputBox("Now enter advance price ")
    If ADV = "" Then
        ADV = 0
    End If
    ra = amt - ADV
    Text8.Text = DISC
    Text9.Text = VAT
    Text11.Text = amt
    Text12.Text = ADV
    Text13.Text = ra
    path
    T = "insert into temp_ord_dtl values('" + Text1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy ") + "','" + Combo1.Text + "','" + Text2.Text + "'," + Text11.Text + "," + Text12.Text + "," + Text13.Text + ")"
    Set R2 = CC.Execute(T)
    MsgBox "record saved"
    MsgBox "Remaining amount of " & Text1.Text & "  is  " & ra
Exit Sub
Else
If Text1.Text = "" Or Combo1.Text = "" Or Text2.Text = "" Or DTP1.Value = "" Or Combo2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Combo3.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Then
    MsgBox "ENTER ALL FIELDS"
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Exit Sub
Else
    path
    T = "select rem_qty,cpt from stk_dtl where p_id='" & Trim(Combo2.Text) & "'"
    Set R2 = CC.Execute(T)
    Dim c1, r1, c2, q1 As Integer
    r1 = R2.Fields("rem_qty")
    c1 = R2.Fields("cpt")
    q1 = Val(Text6.Text)
    c2 = q1 + r1
If (c2 < c1) Then
    S = "insert into ord_dtl values('" + Text1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "','" + Combo1.Text + "','" + Text2.Text + "','" + Combo2.Text + "','" + Text3.Text + "','" + Text4.Text + "'," + Text5.Text + "," + Text6.Text + "," + Text7.Text + ",'" + Combo3.Text + "'," + Text8.Text + "," + Text9.Text + "," + Text10.Text + "," + Text11.Text + "," + Text12.Text + "," + Text13.Text + ",'" + CHQ_DTL.Text1.Text + "','" + CHQ_DTL.Text2.Text + "','" + CHQ_DTL.Text3.Text + "','" + CHQ_DTL.Text4.Text + "','" + CHQ_DTL.Text5.Text + "','" + CHQ_DTL.Text6.Text + "')"
    Set R = CC.Execute(S)
    Text16.Text = Val(Text11.Text)
    Text13.Text = Val(Text13.Text)
    Unload CHQ_DTL
    ORD_DTL.WindowState = 2 - MAXIMIZED
    MsgBox "RECORD SAVED"
    DTP1.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    Combo3.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Text9.Enabled = False
    Adodc1.Refresh
    Command2.Enabled = False
    Command1.Enabled = True
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Command3.Enabled = True
Else
    MsgBox "You can't give any order as Stock is overflowing ! ", vbCritical, "Message"
    Command1.Enabled = False
    Command2.Enabled = True
Exit Sub
End If
If MsgBox("do you want to continue for Multiple Entry", vbYesNo, "message") = vbYes Then
    MsgBox "Click Add More Button ", vbInformation, "MESSAGE"
    Command8.SetFocus
Else
    T = "insert into temp_ord_dtl values('" + Text1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy ") + "','" + Combo1.Text + "','" + Text2.Text + "'," + Text16.Text + "," + Text12.Text + "," + Text13.Text + ")"
    Set R2 = CC.Execute(T)
End If
End If
End If
End Sub
Private Sub Command3_Click()
DataGrid1.Enabled = True
If MsgBox("do you want to continue for deleting", vbYesNo, "message") = vbYes Then
If Text1.Text <> "" Then
path
S = "delete from ord_dtl where ord_id='" & Text1.Text & "'"
Set R = CC.Execute(S)
Adodc1.Refresh
MsgBox "record deleted"
End If
If Text1.Text = "" Then
path
S = "delete from ord_dtl where ord_dt='" & Format(DTP1.Value, "dd-mmm-yyyy") & "'"
Set R = CC.Execute(S)
Adodc1.Refresh
MsgBox "record deleted"
End If
End If
End Sub
Private Sub Command4_Click()
Command3.Enabled = False
Command4.Enabled = True
DTP1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = False
MsgBox "Enter  Order_Id or Order_Date or Supplier_Id or Supplier_Name for searching", vbInformation, "MESSAGE"
Text1.SetFocus
End Sub
Private Sub Command5_Click()
Command5.Enabled = False
path
S = "update ord_dtl set ord_dt='" + Format(DTP1.Value, "dd-mmm-yyyy") + "',s_id='" + Text14.Text + "',s_nm='" + Text2.Text + "',p_id='" + Text15.Text + "',p_nm='" + Text3.Text + "',pack='" + Text4.Text + "',mrp=" + Text5.Text + ",qty=" + Text6.Text + ",RT=" + Text7.Text + ",mod_ptm='" + Combo3.Text + "',DISC=" + Text8.Text + ",VAT=" + Text9.Text + ",amt=" + Text10.Text + " ,TOT_AMT=" + Text11.Text + ",ADV_PRC=" + Text12.Text + ",REM_AMT=" + Text13.Text + " where ord_id='" + Text1.Text + "'"
Set R = CC.Execute(S)
MsgBox "Record Updated !", vbOKCancel, "Message"
Adodc1.Refresh
DataGrid1.Enabled = True
End Sub
Private Sub Command6_Click()
If DataEnvironment1.rsCommand2.State = 1 Then DataEnvironment1.rsCommand2.Close
DataReport2.Show
End Sub
Private Sub Command7_Click()
Unload Me
End Sub
Private Sub Command8_Click()
Text14.Enabled = False
Text15.Enabled = False
Text14.Visible = False
Text15.Visible = False
Command8.Enabled = False
Command9.Enabled = True
Text8.Text = 15
Text9.Text = 0
Command7.Enabled = False
Command2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Combo2.Enabled = True
End Sub
Private Sub Command9_Click()
CHQ_DTL.Visible = False
C = 1
If Text6.Text <> "" And Text7.Text <> "" Then
path
S = "insert into ord_dtl values('" + Text1.Text + "','" + Format(DTP1.Value, "dd-mmm-yyyy") + "','" + Combo1.Text + "','" + Text2.Text + "','" + Combo2.Text + "','" + Text3.Text + "','" + Text4.Text + "'," + Text5.Text + "," + Text6.Text + "," + Text7.Text + ",'" + Combo3.Text + "'," + Text8.Text + "," + Text9.Text + "," + Text10.Text + "," + Text11.Text + "," + Text12.Text + "," + Text13.Text + ",'" + CHQ_DTL.Text1.Text + "','" + CHQ_DTL.Text2.Text + "','" + CHQ_DTL.Text3.Text + "','" + CHQ_DTL.Text4.Text + "','" + CHQ_DTL.Text5.Text + "','" + CHQ_DTL.Text6.Text + "')"
Set R = CC.Execute(S)
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ord_dtl ORDER BY ORD_ID"
MsgBox "RECORD SAVED"
Unload CHQ_DTL
ORD_DTL.WindowState = 2 - MAXIMIZED
Adodc1.Refresh
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text10.Text = ""
Command8.Enabled = True
Command9.Enabled = False
Else
MsgBox "Enter All Fields", vbInformation, "Message"
End If
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Command1.Enabled = True
Text14.Visible = True
Text15.Visible = True
Text1.Text = DataGrid1.Columns(0).Text
DTP1.Value = DataGrid1.Columns(1).Text
Text14.Text = DataGrid1.Columns(2).Text
Text2.Text = DataGrid1.Columns(3).Text
Text15.Text = DataGrid1.Columns(4).Text
Text3.Text = DataGrid1.Columns(5).Text
Text4.Text = DataGrid1.Columns(6).Text
Text5.Text = DataGrid1.Columns(7).Text
Text6.Text = DataGrid1.Columns(8).Text
Text7.Text = DataGrid1.Columns(9).Text
Combo3.Text = DataGrid1.Columns(10).Text
Text8.Text = DataGrid1.Columns(11).Text
Text9.Text = DataGrid1.Columns(12).Text
Text10.Text = DataGrid1.Columns(13).Text
Text11.Text = DataGrid1.Columns(14).Text
Text12.Text = DataGrid1.Columns(15).Text
Text13.Text = DataGrid1.Columns(16).Text
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
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
DTP1.Enabled = True
End If
Unload CHQ_DTL
ORD_DTL.Show
End Sub
Private Sub DTP1_Click()
If Command4.Enabled = True Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ord_dtl where ord_dt='" & Format(DTP1.Value, "dd-mmm-yyyy") & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox " Record Not FOund ! ", vbCritical, "MESSAGE"
Exit Sub
End If
Set DataGrid1.DataSource = Adodc1
End If
Command3.Enabled = True
Command5.Enabled = True
End Sub
Private Sub Form_Load()
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT *  FROM ORD_dTL ORDER BY ORD_ID"
Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = True
DTP1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
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
DTP1.Value = Date
path
S = "select p_id from prd_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo2.AddItem R.Fields("p_id")
R.MoveNext
Wend
path
S = "select s_id from sup_dtl"
Set R = CC.Execute(S)
While R.EOF = False
Combo1.AddItem R.Fields("s_id")
R.MoveNext
Wend
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text <> "" Then
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ord_dtl where ord_id='" + Text1.Text + "'"
Adodc1.Refresh
Text10.Visible = True
Text11.Visible = True
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "RECORD NOT FOUND !", vbCritical, "MESSAGE"
Exit Sub
End If
Set DataGrid1.DataSource = Adodc1
Command3.Enabled = True
Command5.Enabled = True
End If
End If
If Command10.Enabled = False Then
If Text1.Text <> "" Then
If KeyAscii = 13 Then
If DataEnvironment1.rsCommand3.State = 1 Then DataEnvironment1.rsCommand3.Close
DataEnvironment1.Command3 Text1.Text
DataReport3.Show
Command10.Enabled = True
End If
End If
End If
End Sub
Private Sub Text11_Change()
If Val(Text12.Text) = 0 Then
    Text13.Text = Val(Text11.Text)
End If
End Sub
Private Sub Text12_Change()
Dim TA, a, ra As Double
TA = Val(Text11.Text)
a = Val(Text12.Text)
ra = TA - a
Text13.Text = ra
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.Text <> "" Then
If KeyAscii = 13 Then
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ord_dtl where  s_nm='" + Text2.Text + "'"
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
Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
If KeyAscii = 46 Then
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
End If
End Sub
Private Sub Text7_Change()
Dim QT, a, R As Double
QT = Val(Text6.Text)
R = Val(Text7.Text)
a = QT * R
Text10.Text = a
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
If KeyAscii = 13 Then
Text2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "PLEASE ENTER ONLY 0 TO 9"
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
Private Sub Text9_Change()
Dim d, a, v, TA As Double
a = Val(Text10.Text)
d = Val(Text9.Text)
v = Val(Text8.Text)
TA = a + a * (v / 100)
TA = TA - (d / 100 * TA)
Text11.Text = Round(TA, 2)
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

