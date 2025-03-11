VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDept 
   Caption         =   "090 ATC Plating  Dept Codes"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Dept Code.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   45
      Text            =   "XXXX"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Active : 1"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset : 0"
      Height          =   300
      Left            =   1440
      TabIndex        =   43
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      DataField       =   "DEPT_JR_ID"
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
      Left            =   3600
      TabIndex        =   42
      Text            =   "XXX"
      ToolTipText     =   "DEPT_JR_ID"
      Top             =   720
      Width           =   720
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "LOC_JR"
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
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   40
      Text            =   "JR"
      ToolTipText     =   "LOC_JR"
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      DataField       =   "LOC_NY"
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
      MaxLength       =   2
      TabIndex        =   38
      Text            =   "NY"
      ToolTipText     =   "LOC_NY"
      Top             =   1200
      Width           =   720
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Option  3"
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
      Left            =   2640
      TabIndex        =   37
      Top             =   8640
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Option  2"
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
      Left            =   1320
      TabIndex        =   36
      Top             =   8640
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option 1"
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
      TabIndex        =   35
      Top             =   8640
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00C0FFFF&
      DataField       =   "FINISH VALID TEST"
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
      Left            =   120
      TabIndex        =   32
      Text            =   "FINISH VALID TEST"
      ToolTipText     =   "FINISH VALID TEST"
      Top             =   7920
      Width           =   3600
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00C0FFFF&
      DataField       =   "BASE VALID TEST"
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
      Left            =   120
      TabIndex        =   31
      Text            =   "BASE VALID TEST"
      ToolTipText     =   "BASE VALID TEST"
      Top             =   6960
      Width           =   3600
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00C0FFFF&
      DataField       =   "FINISH COL"
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
      Left            =   960
      TabIndex        =   29
      Text            =   "BASE COL"
      ToolTipText     =   "BASE COL"
      Top             =   7440
      Width           =   600
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C0FFFF&
      DataField       =   "BASE COL"
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
      Left            =   960
      TabIndex        =   27
      Text            =   "BASE COL"
      ToolTipText     =   "BASE COL"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text12 
      DataField       =   "STRIKE2_ID"
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
      Left            =   2520
      TabIndex        =   26
      Text            =   "Y"
      ToolTipText     =   "STRIKE2_ID"
      Top             =   4200
      Width           =   960
   End
   Begin VB.TextBox Text11 
      DataField       =   "STRIKE1_ID"
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
      Left            =   2520
      TabIndex        =   25
      Text            =   "Y"
      ToolTipText     =   "STRIKE1_ID"
      Top             =   3240
      Width           =   960
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "Overplate"
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
      TabIndex        =   23
      Text            =   "Y"
      ToolTipText     =   "Overplate"
      Top             =   5160
      Width           =   480
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "STRIKE1"
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
      TabIndex        =   20
      Text            =   "Y"
      ToolTipText     =   "STRIKE1"
      Top             =   3240
      Width           =   480
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "STRIKE2"
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
      TabIndex        =   19
      Text            =   "Y"
      ToolTipText     =   "STRIKE2"
      Top             =   4200
      Width           =   480
   End
   Begin VB.TextBox Text7 
      DataField       =   "TANK DWG"
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
      TabIndex        =   17
      Text            =   "TANK DWG"
      ToolTipText     =   "TANK DWG"
      Top             =   2160
      Width           =   2640
   End
   Begin VB.TextBox Text6 
      DataField       =   "BASE_ID"
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
      TabIndex        =   15
      Text            =   "ID"
      ToolTipText     =   "BASE_ID"
      Top             =   3720
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      DataField       =   "FINISH_ID"
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
      TabIndex        =   13
      Text            =   "ID"
      ToolTipText     =   "FINISH_ID"
      Top             =   4680
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "TANK"
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
      Left            =   3360
      TabIndex        =   11
      Text            =   "Y"
      ToolTipText     =   "TANK"
      Top             =   2760
      Width           =   480
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "DEPT_ID"
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
      TabIndex        =   8
      Text            =   "XXX"
      ToolTipText     =   "DEPT_ID"
      Top             =   720
      Width           =   720
   End
   Begin VB.TextBox Text4 
      DataField       =   "DESCRIPTION"
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
      TabIndex        =   7
      Text            =   "DESCRIPTION"
      ToolTipText     =   "DESCRIPTION"
      Top             =   1680
      Width           =   2640
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "SBE"
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
      TabIndex        =   5
      Text            =   "Y"
      ToolTipText     =   "SBE"
      Top             =   2760
      Width           =   480
   End
   Begin VB.CommandButton cmdR 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update Record"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9240
      Width           =   1800
   End
   Begin VB.TextBox txtActive 
      Alignment       =   2  'Center
      DataField       =   "ACTIVE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   300
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "A"
      ToolTipText     =   "ACTIVE"
      Top             =   5760
      Width           =   480
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Item"
      Enabled         =   0   'False
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
      Left            =   960
      TabIndex        =   2
      Top             =   9720
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [DEPT CODE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DEPT CODE"
      Top             =   1560
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [DEPT CODE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DEPT CODE"
      Top             =   2040
      Visible         =   0   'False
      Width           =   4380
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Dept Code.frx":0CCA
      Height          =   1095
      Left            =   4440
      TabIndex        =   0
      ToolTipText     =   "FROM [DEPT CODE]"
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1931
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      Caption         =   "LOC_JR:"
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
      Index           =   15
      Left            =   2640
      TabIndex        =   41
      ToolTipText     =   "LOC_JR"
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label lblInfo 
      Caption         =   "LOC_NY:"
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
      Index           =   16
      Left            =   120
      TabIndex        =   39
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblInfo 
      Caption         =   "Col # :"
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
      Index           =   14
      Left            =   120
      TabIndex        =   34
      Top             =   7440
      Width           =   840
   End
   Begin VB.Label lblInfo 
      Caption         =   "Col # :"
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
      Index           =   13
      Left            =   120
      TabIndex        =   33
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Finish Validation"
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
      Index           =   9
      Left            =   1560
      TabIndex        =   30
      Top             =   7440
      Width           =   2160
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Base Validation"
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
      Index           =   2
      Left            =   1680
      TabIndex        =   28
      Top             =   6480
      Width           =   2040
   End
   Begin VB.Label lblInfo 
      Caption         =   "Overplate:"
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
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      Caption         =   "Strike 1:"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      Caption         =   "Strike 2:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      Caption         =   "Tank DWG"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label lblInfo 
      Caption         =   "Base ID"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      Caption         =   "Finish ID"
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
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      Caption         =   "Tank"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label lblInfo 
      Caption         =   "Dept Code"
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
      Index           =   12
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lblInfo 
      Caption         =   "Description"
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
      Index           =   11
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label lblInfo 
      Caption         =   "SBE"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "090 Dept Code.frx":0CDE
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [DEPT CODE]"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
  
FR_Table.AddNew
 
FR_Table.Fields("[DEPT_ID]") = 123
FR_Table.Update

FR_Table.Close
FR_Database.Close

cmdRefresh_Click

End Sub


Private Sub cmdR_Click()

Data1.UpdateRecord
cmdRefresh_Click

End Sub

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

If (Option1.Value = True) Then
sSQL = "SELECT [DEPT_ID],[DEPT_JR_ID],[DESCRIPTION],[BASE_ID],[FINISH_ID]," & _
              "[TANK DWG],[SBE],[TANK],[STRIKE1],[STRIKE1_ID],[STRIKE2],[STRIKE2_ID],[OVERPLATE],[ACTIVE] " & _
       "FROM [DEPT CODE] ORDER BY [DEPT_ID]"

sSQLF = "    |^Dept|^JR   |<Base / Finish             |<Base ID  |<Finish ID||^SBE|^Tank|^STK1|^_ID|^STK2|^_ID|^    |^Active"

End If

If (Option2.Value = True) Then
sSQL = "SELECT [DEPT_ID],[DEPT_JR_ID],[BASE_ID],[BASE COL],[BASE VALID TEST]," & _
              "[FINISH_ID],[FINISH COL],[FINISH VALID TEST] " & _
       "FROM [DEPT CODE] ORDER BY [DEPT_ID]"

sSQLF = "    |^Dept|^JR   |<Base ID|^Col|Valid Test                                                    |<Finish ID|^Col|Valid Test                "

End If

If (Option3.Value = True) Then
sSQL = "SELECT [DEPT_ID],[DEPT_JR_ID],[DESCRIPTION],[BASE_ID],[FINISH_ID],[SBE],[TANK]," & _
              "[ACTIVE],[LOC_NY],[LOC_JR] " & _
       "FROM [DEPT CODE] ORDER BY [DEPT_ID]"

sSQLF = "    |^Dept|^JR   |<Base / Finish             |<Base ID  |<Finish ID|^SBE|^Tank|^Active|^LOC_NY |^LOC_JR "

End If


Data2.RecordSource = sSQL
Data2.Refresh
 
MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub Command1_Click()
txtActive.Text = 1
End Sub

Private Sub Command2_Click()
txtActive.Text = 0
End Sub

Private Sub Form_Load()

Caption = "ATC Plating Dept Codes    " & ATC_DWG & "    " & ATC_VERSION

MSFlexGrid2.Top = 0
MSFlexGrid2.Width = 10800
MSFlexGrid2.Height = Me.Height - 800

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION

cmdRefresh_Click
MSFlexGrid2_Click

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then

    Select Case UCase(txtPassword.Text)
    Case "ERIK"
                cmdAdd.Enabled = True
    Case Else
                cmdAdd.Enabled = False
    End Select
End If

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
TABLE_ID = Val(MSFlexGrid2.Text)

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

Dim sSQL As String

sSQL = "SELECT * FROM  [DEPT CODE] WHERE [DEPT_ID]=" & TABLE_ID

Data1.RecordSource = sSQL
Data1.Refresh

End Sub

Private Sub Option1_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option2_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Option3_Click()
cmdRefresh_Click
MSFlexGrid2_Click
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub



Private Sub Text17_GotFocus()
Text17.SelStart = 0
Text17.SelLength = Len(Text17)
End Sub



Private Sub Text18_GotFocus()
Text18.SelStart = 0
Text18.SelLength = Len(Text18)
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
