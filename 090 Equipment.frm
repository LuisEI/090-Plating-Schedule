VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEquipment 
   Caption         =   "ATC Equipment Termination and Plating"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Equipment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option8 
      Caption         =   "[8]  [TF]  Firing"
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
      Left            =   240
      TabIndex        =   58
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Frame Frame3 
      Caption         =   " OEE Parameters "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4800
      TabIndex        =   49
      Top             =   8280
      Width           =   2655
      Begin VB.TextBox txtStandard 
         BackColor       =   &H00C0FFC0&
         DataField       =   "STANDARD"
         DataSource      =   "Data1"
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
         Left            =   1680
         TabIndex        =   53
         Text            =   "ST"
         ToolTipText     =   "STANDARD"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtAT 
         BackColor       =   &H00C0FFC0&
         DataField       =   "AVAILABLE TIME"
         DataSource      =   "Data1"
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
         Left            =   1680
         TabIndex        =   52
         Text            =   "AT"
         ToolTipText     =   "AVAILABLE TIME"
         Top             =   900
         Width           =   800
      End
      Begin VB.TextBox txtST 
         BackColor       =   &H00C0FFC0&
         DataField       =   "SETUP TIME"
         DataSource      =   "Data1"
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
         Left            =   1680
         TabIndex        =   51
         Text            =   "ST"
         ToolTipText     =   "SETUP TIME"
         Top             =   1320
         Width           =   800
      End
      Begin VB.TextBox txtVA 
         BackColor       =   &H00C0FFC0&
         DataField       =   "NON VA"
         DataSource      =   "Data1"
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
         Left            =   1680
         TabIndex        =   50
         Text            =   "VA"
         ToolTipText     =   "NON VA"
         Top             =   1680
         Width           =   800
      End
      Begin VB.Label lblInfo 
         Caption         =   "Stand Cycle: "
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
         Index           =   4
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblInfo 
         Caption         =   "Available (m) : "
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
         Index           =   3
         Left            =   240
         TabIndex        =   56
         Top             =   900
         Width           =   1605
      End
      Begin VB.Label lblInfo 
         Caption         =   "Setup (m) : "
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
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblInfo 
         Caption         =   "Non Value (m) : "
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
         Index           =   7
         Left            =   240
         TabIndex        =   54
         Top             =   1740
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Plating Specific Parameters "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   33
      Top             =   8280
      Width           =   4575
      Begin VB.TextBox Text2 
         DataField       =   "PROCESS"
         DataSource      =   "Data1"
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
         Left            =   1800
         TabIndex        =   39
         Text            =   "PROCESS"
         ToolTipText     =   "PROCESS"
         Top             =   1590
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         DataField       =   "TYPE"
         DataSource      =   "Data1"
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
         Left            =   1800
         TabIndex        =   38
         Text            =   "TYPE"
         ToolTipText     =   "TYPE"
         Top             =   1230
         Width           =   1000
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         DataField       =   "BF_ID"
         DataSource      =   "Data1"
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
         Left            =   1800
         TabIndex        =   37
         Text            =   "BF_ID"
         ToolTipText     =   "BF_ID"
         Top             =   855
         Width           =   1000
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         DataField       =   "NAME"
         DataSource      =   "Data1"
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
         Left            =   1800
         TabIndex        =   36
         Text            =   "NAME"
         ToolTipText     =   "NAME"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         DataField       =   "GALLONS"
         DataSource      =   "Data1"
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
         Left            =   3720
         TabIndex        =   35
         Text            =   "GALLONS"
         ToolTipText     =   "GALLONS"
         Top             =   2160
         Width           =   600
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "RECTIFIER"
         DataSource      =   "Data1"
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   34
         Text            =   "1"
         ToolTipText     =   "RECTIFIER"
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label lblInfo 
         Caption         =   "Process ID : "
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
         Index           =   5
         Left            =   240
         TabIndex        =   48
         Top             =   1590
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         Caption         =   "Process Type : "
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
         Index           =   10
         Left            =   240
         TabIndex        =   47
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label lblInfo 
         Caption         =   "BF_ID"
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
         Index           =   11
         Left            =   240
         TabIndex        =   46
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         Caption         =   "Name"
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
         Index           =   12
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         Caption         =   "Capacity (g): "
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
         Index           =   13
         Left            =   2520
         TabIndex        =   44
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         Caption         =   "Mult Rectifier [0/1]: "
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
         Index           =   15
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "[Barrel:SBE]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   3000
         TabIndex        =   42
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "[Base:Finish]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   3000
         TabIndex        =   41
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nickel,Copper"
         Height          =   300
         Index           =   21
         Left            =   3000
         TabIndex        =   40
         Top             =   840
         Width           =   1395
      End
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataField       =   "ID ORDER"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1440
      TabIndex        =   32
      Text            =   "0"
      ToolTipText     =   "ID ORDER"
      Top             =   5280
      Width           =   480
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   "SERIES"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   31
      Text            =   "SERIE"
      ToolTipText     =   "SERIES"
      Top             =   4920
      Width           =   720
   End
   Begin VB.CommandButton CommandTBL_CALCULATION_EQ 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TBL CALCULATION EQ"
      Height          =   250
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8640
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Active [1] /Reset [0]"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin VB.OptionButton Option7 
      Caption         =   "[7]  [PT] Plating BARREL"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2985
   End
   Begin VB.OptionButton Option6 
      Caption         =   "[6]  [PT] Plating SBE"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2145
   End
   Begin VB.TextBox Text11 
      DataField       =   "MACHINE"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   2280
      TabIndex        =   8
      Text            =   "MACHINE"
      ToolTipText     =   "MACHINE"
      Top             =   4080
      Width           =   720
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5]  All"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   2025
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4]  [TM] Termination"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2265
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3]  [PT] Plating"
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Caption         =   " Option View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "[1] Normal"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         Caption         =   "[2] OEE "
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
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "LOCATION_ID"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "NY"
      ToolTipText     =   "LOCATION_ID"
      Top             =   3255
      Width           =   480
   End
   Begin VB.CommandButton cmdAdds 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adds"
      Height          =   300
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9000
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "DEPT_ID"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   3120
      TabIndex        =   9
      Text            =   "PT"
      ToolTipText     =   "DEPT_ID"
      Top             =   3600
      Width           =   480
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   19
      Text            =   "XXXX"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "CASE"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "ABCER"
      ToolTipText     =   "CASE"
      Top             =   4920
      Width           =   720
   End
   Begin VB.CommandButton cmdUpdateRecord 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtActive 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "ACTIVE"
      DataSource      =   "Data1"
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
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      ToolTipText     =   "ACTIVE"
      Top             =   5880
      Width           =   480
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Equipment"
      Enabled         =   0   'False
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   5880
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtDesc 
      DataField       =   "DESCRIPTION"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Text            =   "DESCRIPTION"
      ToolTipText     =   "DESCRIPTION"
      Top             =   4560
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      DataField       =   "NUMBER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Text            =   "NUMBER"
      ToolTipText     =   "NUMBER"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MACHINE"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      DataField       =   "MACHINE_ID"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Text            =   "MACHINE_ID"
      ToolTipText     =   "MACHINE_ID"
      Top             =   3600
      Width           =   480
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   3300
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Equipment.frx":0CCA
      Height          =   1095
      Left            =   4320
      TabIndex        =   18
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1931
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
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
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Copies selected parameters"
      Height          =   195
      Index           =   20
      Left            =   2160
      TabIndex        =   62
      Top             =   6720
      Width           =   1950
   End
   Begin VB.Label lblInfo 
      Caption         =   "LOCATION_ID"
      Height          =   300
      Index           =   16
      Left            =   240
      TabIndex        =   61
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Label lblInfo 
      Caption         =   "Order"
      Height          =   300
      Index           =   19
      Left            =   240
      TabIndex        =   60
      Top             =   5280
      Width           =   885
   End
   Begin VB.Label lblInfo 
      Caption         =   "SERIES_ID"
      Height          =   300
      Index           =   8
      Left            =   2280
      TabIndex        =   59
      Top             =   4920
      Width           =   885
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   5940
      Left            =   10080
      Picture         =   "090 Equipment.frx":0CDE
      Top             =   960
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label LabelDB 
      AutoSize        =   -1  'True
      Caption         =   "DB"
      Height          =   195
      Left            =   8640
      TabIndex        =   30
      Top             =   8280
      Width           =   225
   End
   Begin VB.Label LabelLOCATION_ID 
      Caption         =   "NY/JR"
      Height          =   300
      Left            =   7680
      TabIndex        =   29
      Top             =   8280
      Width           =   600
   End
   Begin VB.Label lblInfo 
      Caption         =   "DEPT_ID :"
      Height          =   300
      Index           =   9
      Left            =   2280
      TabIndex        =   26
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label lblInfo 
      Caption         =   "Case Sizes: "
      Height          =   300
      Index           =   14
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "090 Equipment.frx":81C0
      Top             =   0
      Width           =   4170
   End
   Begin VB.Label lblInfo 
      Caption         =   "Description"
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   1245
   End
   Begin VB.Label lblInfo 
      Caption         =   "Machine #:"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "NUMBER"
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label lblInfo 
      Caption         =   "Machine ID :"
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   3600
      Width           =   1365
   End
End
Attribute VB_Name = "frmEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()


Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
  
sSQL = "SELECT * FROM [MACHINE]"
Set TO_Table = FR_Database.OpenRecordset(sSQL)
  
TO_Table.AddNew

TO_Table.Fields("[MACHINE]") = "NA"

TO_Table.Fields("[NUMBER]") = 0
TO_Table.Fields("[NAME]") = "NEW"
TO_Table.Fields("[DEPARTMENT]") = FR_Table.Fields("[DEPARTMENT]")
TO_Table.Fields("[DEPT_ID]") = FR_Table.Fields("[DEPT_ID]")

TO_Table.Fields("[BF_ID]") = FR_Table.Fields("[BF_ID]")
TO_Table.Fields("[TYPE]") = FR_Table.Fields("[TYPE]")
TO_Table.Fields("[DESCRIPTION]") = FR_Table.Fields("[DESCRIPTION]")

TO_Table.Fields("[PROCESS]") = FR_Table.Fields("[PROCESS]")

TO_Table.Fields("[STANDARD]") = FR_Table.Fields("[STANDARD]")
TO_Table.Fields("[AVAILABLE TIME]") = FR_Table.Fields("[AVAILABLE TIME]")
TO_Table.Fields("[SETUP TIME]") = FR_Table.Fields("[SETUP TIME]")
TO_Table.Fields("[NON VA]") = FR_Table.Fields("[NON VA]")
TO_Table.Fields("[SERIES]") = FR_Table.Fields("[SERIES]")
TO_Table.Fields("[WASTE]") = FR_Table.Fields("[WASTE]")
TO_Table.Fields("[CASE]") = FR_Table.Fields("[CASE]")
TO_Table.Fields("[CAPACITY]") = FR_Table.Fields("[CAPACITY]")
TO_Table.Fields("[RECTIFIER]") = FR_Table.Fields("[RECTIFIER]")

TO_Table.Fields("[LOCATION_ID]") = FR_Table.Fields("[LOCATION_ID]")
TO_Table.Fields("[GALLONS]") = FR_Table.Fields("[GALLONS]")
TO_Table.Fields("[ID ORDER]") = FR_Table.Fields("[ID ORDER]")
TO_Table.Fields("[ACTIVE]") = FR_Table.Fields("[ACTIVE]")

TO_Table.Update

TO_Database.Close
FR_Database.Close

cmdRefresh_Click

End Sub

Private Sub cmdAdds_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [MACHINE] ORDER BY [MACHINE_ID]"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
Set TO_Table = TO_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF

    TO_Table.AddNew
    
    TO_Table.Fields("[MACHINE_ID]") = FR_Table.Fields("[MACHINE_ID]")
    TO_Table.Fields("[MACHINE]") = FR_Table.Fields("[MACHINE]")
    TO_Table.Fields("[DESCRIPTION]") = FR_Table.Fields("[DESCRIPTION]")
    TO_Table.Fields("[DEPARTMENT]") = FR_Table.Fields("[DEPARTMENT]")
    TO_Table.Fields("[DEPT_ID]") = FR_Table.Fields("[DEPT_ID]")
    TO_Table.Fields("[STANDARD]") = FR_Table.Fields("[STANDARD]")
    TO_Table.Fields("[AVAILABLE TIME]") = FR_Table.Fields("[AVAILABLE TIME]")
     
    TO_Table.Fields("[ACTIVE]") = FR_Table.Fields("[ACTIVE]")
    TO_Table.Fields("[SETUP TIME]") = FR_Table.Fields("[SETUP TIME]")
    TO_Table.Fields("[NON VA]") = FR_Table.Fields("[NON VA]")
    
    TO_Table.Fields("[SERIES]") = FR_Table.Fields("[SERIES]")
    TO_Table.Fields("[WASTE]") = FR_Table.Fields("[WASTE]")
    TO_Table.Fields("[CASE]") = FR_Table.Fields("[CASE]")
    TO_Table.Fields("[ID ORDER]") = FR_Table.Fields("[ID ORDER]")
                     
    TO_Table.Update
    FR_Table.MoveNext
Loop

TO_Database.Close
FR_Database.Close

MsgBox "Complete", vbInformation, "ATC EBD System"

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

If (Option1.VALUE = True) Then
    sSQL = "SELECT [MACHINE_ID],[ID ORDER],[DEPT_ID],[NUMBER],[NAME],[BF_ID],[TYPE],[DESCRIPTION],[PROCESS],[SERIES],[CASE],[ACTIVE],[LOCATION_ID] " & _
           "FROM [MACHINE] "
           
    sSQLF = "    |^        |^Order|^Dept  |^M#  |<Name        "
    sSQLF = sSQLF & "|<BF_ID       |^TYPE            |<Description                 |<Process  |^SERIES|^Case     |^Active|^L_ID"
End If
       
If (Option2.VALUE = True) Then
    sSQL = "SELECT [MACHINE_ID],[ID ORDER],[DEPT_ID],[NUMBER],[NAME],[RECTIFIER],[STANDARD],[AVAILABLE TIME],[SETUP TIME],[NON VA],[GALLONS],[LOCATION_ID] " & _
           "FROM [MACHINE] "

    sSQLF = "    |^|^Order|^Dept  |^M#  |<Name        |^Rect"
    sSQLF = sSQLF & "|<Standard|^Available (m)|<Setup (m)|Non Val(m)|>Capacity (g)|^L_ID"
End If
 
If (Option3.VALUE = True) Then
    sSQL = sSQL & " WHERE [DEPT_ID] IN ('PT') AND [LOCATION_ID]= '" & LOCATION_ID & "' ORDER BY [MACHINE_ID] ASC"
End If

If (Option4.VALUE = True) Then
    sSQL = sSQL & " WHERE [DEPT_ID] IN ('TM','IP')AND [LOCATION_ID]= '" & LOCATION_ID & "' ORDER BY [PROCESS],[MACHINE_ID] ASC"
End If
If (Option8.VALUE = True) Then
    sSQL = sSQL & " WHERE [DEPT_ID] IN ('TF')AND [LOCATION_ID]= '" & LOCATION_ID & "' ORDER BY [PROCESS],[MACHINE_ID] ASC"
End If

If (Option6.VALUE = True) Then
    sSQL = sSQL & " WHERE [TYPE] IN ('SBE') AND [LOCATION_ID] ='" & LOCATION_ID & "' ORDER BY [NUMBER] ASC"
End If
If (Option7.VALUE = True) Then
    sSQL = sSQL & " WHERE [TYPE] IN ('BARREL') AND [LOCATION_ID] ='" & LOCATION_ID & "' ORDER BY [NUMBER] ASC"
End If

'ALL
If (Option5.VALUE = True) Then
    sSQL = sSQL & " ORDER BY [MACHINE_ID] ASC"
End If
 
Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub cmdUpdateRecord_Click()
Data1.UpdateRecord
cmdRefresh_Click
End Sub

Private Sub Command1_Click()

If Val(txtActive.Text) = 1 Then
        txtActive.Text = 0
Else
        txtActive.Text = 1
End If

End Sub




Private Sub CommandTBL_CALCULATION_EQ_Click()

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [MACHINE] WHERE [LOCATION_ID]='" & LOCATION_ID & "' AND [DEPT_ID]='PT' AND [ACTIVE]=1"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [TBL CALCULATION EQ]"
               
Set TO_Table = TO_Database.OpenRecordset(sSQL)

Do Until FR_Table.EOF

        If (TO_Table.EOF = False) Then
            TO_Table.Edit
            TO_Table.Fields("[NUMBER]") = FR_Table.Fields("[NUMBER]")
            TO_Table.Update
            TO_Table.MoveNext
        Else
            TO_Table.AddNew
            TO_Table.Fields("[NUMBER]") = FR_Table.Fields("[NUMBER]")
            TO_Table.Update
        End If
        
        FR_Table.MoveNext
Loop

TO_Database.Close
FR_Database.Close

MsgBox "Complete", vbInformation, "Plating Schedule"

End Sub

Private Sub Form_Load()

Caption = "ATC Equipment Termination and Plating       " & ATC_DWG & "    " & ATC_VERSION

MSFlexGrid2.Top = 0
'MSFlexGrid2.Left = 0
MSFlexGrid2.Width = 11400
MSFlexGrid2.Height = 8000

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION

LabelLOCATION_ID.Caption = LOCATION_ID
LabelDB.Caption = DB_PLATING_TERMINATION

cmdRefresh_Click
MSFlexGrid2_Click

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then
    Select Case UCase(txtPassword.Text)
    Case "ERIK"
                cmdAdd.Enabled = True
                Text1.Enabled = True
    Case Else
                cmdAdd.Enabled = False
                Text1.Enabled = False
    End Select
End If
txtPassword.Text = "XXXX"

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
MACHINE_ID = Val(MSFlexGrid2.Text)

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

Dim sSQL As String
sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID

Data1.RecordSource = sSQL
Data1.Refresh

End Sub

Private Sub Option1_Click()
cmdRefresh_Click
End Sub

Private Sub Option2_Click()
cmdRefresh_Click
End Sub

Private Sub Option3_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option4_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option5_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option6_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option7_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option8_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option9_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10)
End Sub



Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13)
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3)
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4)
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5)
End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6)
End Sub

Private Sub Text7_GotFocus()
Text7.SelStart = 0
Text7.SelLength = Len(Text7)
End Sub

Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8)
End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9)
End Sub

Private Sub txtActive_GotFocus()
txtActive.SelStart = 0
txtActive.SelLength = Len(txtActive)
End Sub

Private Sub txtAT_GotFocus()
txtAT.SelStart = 0
txtAT.SelLength = Len(txtAT)
End Sub

Private Sub txtDept_ID_GotFocus()
txtDept_ID.SelStart = 0
txtDept_ID.SelLength = Len(txtDept_ID)
End Sub

Private Sub txtDesc_GotFocus()
txtDesc.SelStart = 0
txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtST_GotFocus()
txtST.SelStart = 0
txtST.SelLength = Len(txtST)
End Sub

Private Sub txtStandard_GotFocus()
txtStandard.SelStart = 0
txtStandard.SelLength = Len(txtStandard)
End Sub

Private Sub txtWaste_GotFocus()
txtWaste.SelStart = 0
txtWaste.SelLength = Len(txtWaste)
End Sub
