VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSQL 
   Caption         =   "090 SQL Search and Review"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 SQL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Frame fraDateSelect 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   38
      Top             =   6480
      Width           =   4455
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   300
         Left            =   2250
         TabIndex        =   48
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
         Height          =   300
         Left            =   1245
         TabIndex        =   47
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
         Height          =   300
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Refresh"
         Height          =   300
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1080
         Width           =   1000
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day"
         Height          =   300
         Left            =   1680
         TabIndex        =   40
         Top             =   735
         Width           =   945
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Week"
         Height          =   300
         Left            =   360
         TabIndex        =   39
         Top             =   735
         Value           =   -1  'True
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48889857
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48889857
         CurrentDate     =   38117
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select DEPT_ID then refresh"
         Height          =   495
         Left            =   2880
         TabIndex        =   50
         Top             =   480
         Width           =   1305
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 FROM [DEPT CODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.CommandButton CommandData3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CommandData3"
      Height          =   300
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Lot Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   36
      Top             =   360
      Width           =   1400
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ATC Part "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   35
      Top             =   360
      Width           =   1400
   End
   Begin VB.OptionButton Option1 
      Caption         =   "W.O./Lot No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   34
      Top             =   360
      Value           =   -1  'True
      Width           =   1400
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8880
      TabIndex        =   6
      Top             =   4080
      Width           =   5535
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 5"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   1335
         Width           =   840
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 6"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1335
         Width           =   840
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 7"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   1335
         Width           =   840
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 8"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   1335
         Width           =   840
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FN HEAD 1"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   1710
         Width           =   840
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FN HEAD 2"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   1710
         Width           =   840
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FN HEAD 3"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   1710
         Width           =   840
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFFF&
         DataField       =   "FN HEAD 4"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   1710
         Width           =   840
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
         DataField       =   "HEAD 4"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   960
         Width           =   840
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         DataField       =   "HEAD 3"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   840
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         DataField       =   "HEAD 2"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   840
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         DataField       =   "HEAD 1"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   840
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 4"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   585
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 3"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   585
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED 2"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   585
         Width           =   840
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         DataField       =   "SPEED"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   585
         Width           =   840
      End
      Begin VB.Label lblSet 
         Caption         =   "Finish Speed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   1290
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Caption         =   "Finish Serial #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   1635
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "2NG/H4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   4440
         TabIndex        =   28
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "2G/H3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "1NG/H2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblSet 
         Caption         =   "Base Serial #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   945
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "1G/H1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblSet 
         Caption         =   "Base Speed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT],[BARCODE] "
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [SCHEDULE SETS],[GROUPING]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Search SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtWorkOrder 
      BackColor       =   &H00FFFFC0&
      DataField       =   "WORK ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      MaxLength       =   12
      TabIndex        =   2
      Top             =   360
      Width           =   1600
   End
   Begin VB.TextBox txtATCPart 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ATC PART"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   1600
   End
   Begin VB.TextBox txtLot 
      BackColor       =   &H00FFFFC0&
      DataField       =   "LOT NUM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   0
      Top             =   360
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 SQL.frx":0CCA
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "FROM [SCHEDULE SETS],[GROUPING]"
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "090 SQL.frx":0CDE
      Height          =   2295
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT],[BARCODE] "
      Top             =   4200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 SQL.frx":0CF2
      Height          =   3615
      Left            =   4560
      TabIndex        =   43
      ToolTipText     =   "FROM [SCHEDULE SETS]"
      Top             =   6600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6376
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
      Bindings        =   "090 SQL.frx":0D06
      Height          =   2295
      Left            =   360
      TabIndex        =   44
      ToolTipText     =   "FROM [DEPT CODE]"
      Top             =   8160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4048
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [SCHEDULE SETS] WHERE [DEPT_ID] = DEPT_ID"
      Height          =   270
      Left            =   4800
      TabIndex        =   49
      Top             =   10320
      Width           =   4665
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT],[BARCODE] "
      Height          =   270
      Left            =   480
      TabIndex        =   33
      Top             =   3840
      Width           =   5985
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT]"
      Height          =   270
      Left            =   8880
      TabIndex        =   32
      Top             =   3840
      Width           =   4665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [SCHEDULE SETS],[GROUPING]"
      Height          =   270
      Left            =   360
      TabIndex        =   31
      Top             =   840
      Width           =   3225
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRefresh1_Click()

End Sub

Private Sub cmdNext_Click()
If (optWeek.Value = True) Then
    DTPicker1.Value = DateAdd("WW", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("WW", 1, DTPicker2.Value)
End If

If (optDay.Value = True) Then
    DTPicker1.Value = DateAdd("D", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("D", 1, DTPicker2.Value)
End If

cmdRefresh1_Click
End Sub

Private Sub cmdPrevious_Click()

If (optWeek.Value = True) Then
    DTPicker1.Value = DateAdd("WW", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("WW", -1, DTPicker2.Value)
End If

If (optDay.Value = True) Then
    DTPicker1.Value = DateAdd("D", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("D", -1, DTPicker2.Value)
End If

cmdRefresh1_Click

End Sub

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],[DATE_ID],[TYPE_ID],[SERIES_ID]," & _
              "[EQ BASE],format([BASE AMP],'0.0'),format([BASE AMP MIN],'0.0')," & _
              "[EQ FINISH],format([FINISH AMP],'0.0'),format([FINISH AMP MIN],'0.0')  " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DEPT_ID]=" & DEPT_ID & " " & _
        "ORDER BY [DATE_ID] DESC"
                                   

sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^                     |^Series|Base|Amp      |Amp Min  |Finish|Amp    |Amp Min   "

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdReset_Click()

DTPicker1.Value = Date

If (optWeek.Value = True) Then
    DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
    DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
   DTPicker1.Value = DTPicker1.Value
   DTPicker2.Value = DTPicker1.Value
End If

cmdRefresh1_Click

End Sub

Private Sub cmdSearch_Click()

Dim sGROUPING As String
If (Option1.Value = True) Then
        WO_ID = txtWorkOrder.Text
        sGROUPING = "[GROUPING].[WORK ORDER]='" & WO_ID & "'"
End If
If (Option2.Value = True) Then
       WO_ID = txtATCPart.Text
       sGROUPING = "[GROUPING].[ATC PART]='" & WO_ID & "'"
End If
If (Option3.Value = True) Then
       WO_ID = txtLot.Text
       sGROUPING = "[GROUPING].[LOT NUM]='" & WO_ID & "'"
End If

Dim sSQL As String
Dim sSQLF As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [SCHEDULE SETS].[DATE_ID]," & _
              "[SCHEDULE SETS].[TYPE_ID] AS [SQL TYPE_ID] " & _
       "FROM [SCHEDULE SETS],[GROUPING]" & _
       "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND "
sSQL = sSQL & sGROUPING
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    TYPE_ID = FR_Table.Fields("[SQL TYPE_ID]")
Else
    MsgBox "Not Found", vbCritical, "ATC Plating"
    Exit Sub
End If

Select Case TYPE_ID
Case "SBE"
             sSQL = "SELECT [SCHEDULE SETS].[SET_ID]," & _
                           "[SCHEDULE SETS].[DATE_ID]," & _
                                "[GROUPING].[WORK ORDER]," & _
                                "[GROUPING].[ATC PART]," & _
                                "[GROUPING].[LOT NUM]," & _
                                "[GROUPING].[QTY]," & _
                           "[SCHEDULE SETS].[DEPT_ID]," & _
                           "[SCHEDULE SETS].[SET NUMBER]," & _
                           "[SCHEDULE SETS].[TYPE_ID]," & _
                           "[SCHEDULE SETS].[EQ BASE]," & _
                    "format([SCHEDULE SETS].[BASE AMP MIN],'0')," & _
                           "[SCHEDULE SETS].[EQ FINISH]," & _
                    "format([SCHEDULE SETS].[FINISH AMP MIN],'0') " & _
               "FROM [SCHEDULE SETS],[GROUPING]" & _
               "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND "
                                          
            sSQL = sSQL & sGROUPING
            
Case "BARREL"   'MIN

        sSQL = "SELECT [SCHEDULE SETS].[SET_ID]," & _
                      "[SCHEDULE SETS].[DATE_ID]," & _
                      "[GROUPING].[WORK ORDER]," & _
                      "[GROUPING].[ATC PART]," & _
                      "[GROUPING].[LOT NUM]," & _
                      "[GROUPING].[QTY]," & _
                      "[SCHEDULE SETS].[DEPT_ID]," & _
                      "[SCHEDULE SETS].[SET NUMBER]," & _
                      "[SCHEDULE SETS].[TYPE_ID]," & _
                      "[SCHEDULE SETS].[EQ BASE]," & _
                      "format([SCHEDULE SETS].[BASE AMP MIN]/60,'0')," & _
                      "[SCHEDULE SETS].[EQ FINISH]," & _
                      "format([SCHEDULE SETS].[FINISH AMP MIN]/60,'0') " & _
               "FROM [SCHEDULE SETS],[GROUPING]" & _
               "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND "
                     
        sSQL = sSQL & sGROUPING
                                                                         
End Select

Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    ||^Create Date    |^Work Order           |^ATC Part                     |^Lot Number        |Quantity     "
sSQLF = sSQLF & "|^Dept|^Set#|^Type        |^Base EQ|Time (Hr)|^Finish EQ|Time (Hr) "
 
MSFlexGrid2.FormatString = sSQLF

MSFlexGrid2_Click

End Sub



Private Sub CommandData3_Click()
 
Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
              "[WORK SHEET PT].[CODE_ID]," & _
              "[WORK SHEET PT].[DATE_ID]," & _
              "format([WORK SHEET PT].[DATE_ID],'dddd')," & _
              "[BARCODE].[FIRST] & ' ' & [BARCODE].[LAST]," & _
              "format([START TIME],'h:mm AM/PM')," & _
              "format([STOP TIME],'h:mm AM/PM')," & _
              "[WORK SHEET PT].[TOTAL TIME] " & _
       "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT],[BARCODE] " & _
       "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
             "[WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
             "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
             "[SCHEDULE SETS].[SET_ID]=" & SET_ID & " " & _
        "ORDER BY [WORK SHEET PT].[CODE_ID]"
                                                            
sSQLF = "    ||^CODE|^Work Date  |^Day of Week|Plating Operator              |^Start Time     |^Stop Time      |Time "
                                                            
Data3.RecordSource = sSQL
Data3.Refresh
                              
MSFlexGrid3.FormatString = sSQLF

End Sub


Private Sub Form_Load()

Caption = "SQL Search and Review     " & ATC_DWG & "    " & ATC_VERSION


Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION
Data5.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid2.Left = 0
MSFlexGrid2.Width = 14000
MSFlexGrid3.Left = 0
MSFlexGrid3.Width = 8000
 
MSFlexGrid1.Width = 10000
 
Dim sSQL As String
Dim sSQLF As String
    
Select Case LOCATION_ID
Case "JR"
        sSQL = "SELECT [DEPT_JR_ID],[DESCRIPTION] " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "

        sSQL = sSQL & "[LOC_JR]='" & LOCATION_ID & "' ORDER BY [DEPT_ID]"
Case "NY"
        sSQL = "SELECT [DEPT_ID],[DESCRIPTION]  " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "
        sSQL = sSQL & "[LOC_NY]='" & LOCATION_ID & "' ORDER BY [DEPT_ID]"
End Select

sSQLF = "    |^DEPT_ID|<Base / Finish              "

Data5.RecordSource = sSQL
Data5.Refresh

MSFlexGrid5.FormatString = sSQLF

MSFlexGrid5_Click

DTPicker1.Value = Date
If (optWeek.Value = True) Then
    cmdPrevious.Caption = "Week  <<"
    cmdNext.Caption = "Week  >>"
    DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
    DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
    cmdPrevious.Caption = "Day  <<"
    cmdNext.Caption = "Day  >>"
    DTPicker1.Value = DTPicker1.Value
    DTPicker2.Value = DTPicker1.Value
End If

cmdRefresh_Click
 
 
 

'Dim sSQLF As String

sSQLF = "    ||^Create Date    |^Work Order           |^ATC Part                     |^Lot Number        |Quantity     "
sSQLF = sSQLF & "|^Dept|^Set#|^Type        |^Base EQ|Time (Hr)|^Finish EQ|Time (Hr) "
 
MSFlexGrid2.FormatString = sSQLF

sSQLF = "    ||^CODE|^Work Date  |^Day of Week|Plating Operator              |^Start Time     |^Stop Time      |Time "
                                                            
MSFlexGrid3.FormatString = sSQLF

txtWorkOrder.Text = "272064003104"

txtWorkOrder.Text = "971114001001"

End Sub

Private Sub MSFlexGrid1_Click()
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
SET_ID = Val(MSFlexGrid2.Text)
  
MSFlexGrid2.Col = 5
txtLot.Text = MSFlexGrid2.Text
  
MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1

Dim sSQL As String
sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
Data4.RecordSource = sSQL
Data4.Refresh

CommandData3_Click

End Sub

Private Sub MSFlexGrid5_Click()

MSFlexGrid5.Col = 1
DEPT_ID = Val(MSFlexGrid5.Text)
  
MSFlexGrid5.Col = 0
MSFlexGrid5.ColSel = MSFlexGrid5.Cols - 1
 
End Sub

Private Sub optDay_Click()

DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Day  <<"
cmdNext.Caption = "Day  >>"

End Sub

Private Sub optWeek_Click()

DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Week  <<"
cmdNext.Caption = "Week  >>"

End Sub

Private Sub txtWorkOrder_GotFocus()
txtWorkOrder.SelStart = 0
txtWorkOrder.SelLength = Len(txtWorkOrder)
End Sub
