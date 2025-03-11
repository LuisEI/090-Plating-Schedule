VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOperator 
   Caption         =   "ATC Operators Termination and Plating "
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "090 Operator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "XXXX"
      ToolTipText     =   "Password"
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame fraDB 
      Caption         =   " Operator Managment "
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
      Height          =   1695
      Left            =   120
      TabIndex        =   34
      Top             =   8640
      Width           =   5055
      Begin VB.TextBox textFrom 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2760
         TabIndex        =   19
         ToolTipText     =   "OP_ID"
         Top             =   480
         Width           =   705
      End
      Begin VB.CommandButton cmdTest2 
         Caption         =   "Test 2 OP_ID"
         Height          =   300
         Left            =   2760
         TabIndex        =   22
         ToolTipText     =   "[TERMINATION]"
         Top             =   840
         Width           =   1755
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test 1 OP_ID"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "[WORK SHEET]"
         Top             =   840
         Width           =   1755
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Spare Z Operator"
         Height          =   300
         Left            =   2760
         TabIndex        =   24
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete "
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1755
      End
      Begin VB.TextBox textTo 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4080
         TabIndex        =   20
         ToolTipText     =   "OP_ID"
         Top             =   480
         Width           =   705
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change OP_ID"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "[TERMINATION]"
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         Caption         =   "To :"
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   36
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Caption         =   "From:"
         Height          =   300
         Index           =   9
         Left            =   2160
         TabIndex        =   35
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.Frame fraOperator 
      Caption         =   " Operator "
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
      Height          =   4815
      Left            =   240
      TabIndex        =   28
      Top             =   3720
      Width           =   5055
      Begin VB.CommandButton CommandReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Reset"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DP"
         Height          =   300
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1560
         Width           =   400
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "IP"
         Height          =   300
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1560
         Width           =   400
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TM"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1560
         Width           =   400
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PT"
         Height          =   300
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1560
         Width           =   400
      End
      Begin VB.CommandButton CommandActive 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Active [1] /Reset [0]"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         DataField       =   "MASTER_ID"
         DataSource      =   "Data2"
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
         TabIndex        =   43
         ToolTipText     =   "Operator ID"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "LEVEL"
         DataSource      =   "Data2"
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
         MaxLength       =   2
         TabIndex        =   41
         ToolTipText     =   "LEVEL"
         Top             =   3480
         Width           =   480
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "PASSWORD"
         DataSource      =   "Data2"
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
         ToolTipText     =   "PASSWORD"
         Top             =   3120
         Width           =   1680
      End
      Begin VB.TextBox TextLocation 
         Alignment       =   2  'Center
         DataField       =   "LOCATION_ID"
         DataSource      =   "Data2"
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
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   9
         ToolTipText     =   "LOCATION_ID"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Update Record"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4200
         Width           =   1755
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "SHIFT_ID"
         DataSource      =   "Data2"
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
         Left            =   4080
         MaxLength       =   1
         TabIndex        =   16
         ToolTipText     =   "SHIFT_ID"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtDEPT 
         Alignment       =   2  'Center
         DataField       =   "DEPT_ID"
         DataSource      =   "Data2"
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
         MaxLength       =   2
         TabIndex        =   11
         ToolTipText     =   "Dept ID"
         Top             =   1560
         Width           =   600
      End
      Begin VB.TextBox txtActive 
         Alignment       =   2  'Center
         DataField       =   "ACTIVE"
         DataSource      =   "Data2"
         Height          =   300
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   10
         ToolTipText     =   "Active"
         Top             =   4200
         Width           =   480
      End
      Begin VB.TextBox Text6 
         DataField       =   "BARCODE"
         DataSource      =   "Data2"
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
         TabIndex        =   14
         ToolTipText     =   "Bar Code"
         Top             =   2760
         Width           =   2280
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         DataField       =   "OP_ID"
         DataSource      =   "Data2"
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
         TabIndex        =   8
         ToolTipText     =   "Operator ID"
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox txtFirst 
         DataField       =   "FIRST"
         DataSource      =   "Data2"
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
         TabIndex        =   12
         ToolTipText     =   "First Name"
         Top             =   2040
         Width           =   2280
      End
      Begin VB.TextBox txtLast 
         DataField       =   "LAST"
         DataSource      =   "Data2"
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
         TabIndex        =   13
         ToolTipText     =   "Last Name"
         Top             =   2400
         Width           =   2280
      End
      Begin VB.Label lblInfo 
         Caption         =   "MASTER_ID :"
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
         Left            =   360
         TabIndex        =   44
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         Caption         =   "Level:"
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
         Left            =   360
         TabIndex        =   42
         Top             =   3480
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Caption         =   "Password :"
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
         Left            =   360
         TabIndex        =   40
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Caption         =   "LOCATION_ID :"
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
         Left            =   2640
         TabIndex        =   39
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label lblInfo 
         Caption         =   "SHIFT_ID:"
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
         Index           =   8
         Left            =   2640
         TabIndex        =   37
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Caption         =   "DEPT_ID :"
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
         Left            =   360
         TabIndex        =   33
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Bar Code :"
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
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Caption         =   "OP ID :"
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
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblInfo 
         Caption         =   "First:"
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
         Left            =   360
         TabIndex        =   30
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Caption         =   "Last :"
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
         Left            =   360
         TabIndex        =   29
         Top             =   2400
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Department "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   5055
      Begin VB.OptionButton OptionNot 
         Caption         =   "Not"
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
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "[4] [DP] DPA "
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
         Left            =   360
         TabIndex        =   46
         Top             =   1440
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "[3] [IP] Inspection"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "[2] [TM] Termination && Firing"
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
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "[1] [PT] Plating"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Shift "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   5055
      Begin VB.OptionButton OptCombined 
         Caption         =   "Combined"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optShift_Evening 
         Caption         =   "[E] Evening"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optShift_Day 
         Caption         =   "[D] Day"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BARCODE"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.CommandButton cmdRefreshDisplay 
      Caption         =   "Refresh Display"
      Height          =   300
      Left            =   6240
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Operator.frx":0CCA
      Height          =   1215
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "FROM [BARCODE]"
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2143
      _Version        =   393216
      ScrollBars      =   2
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "090 Operator.frx":0CDE
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

DEPT_ID = "PT"
If (Option1.VALUE = True) Then
     DEPT_ID = "PT"
End If
If (Option2.VALUE = True) Then
     DEPT_ID = "TM"
End If
If (Option3.VALUE = True) Then
     DEPT_ID = "IP"
End If
If (Option4.VALUE = True) Then
     DEPT_ID = "DP"
End If

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [BARCODE] ORDER BY [OP_ID] DESC"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Dim OP_ID As Long

OP_ID = FR_Table.Fields("[OP_ID]") + 1

FR_Table.AddNew
FR_Table.Fields("[OP_ID]") = OP_ID
FR_Table.Fields("[LOCATION_ID]") = LOCATION_ID
FR_Table.Fields("[FIRST]") = "Spare"
FR_Table.Fields("[LAST]") = "Z Operator"
FR_Table.Fields("[BARCODE]") = "BARCODE"
FR_Table.Fields("[DEPT_ID]") = DEPT_ID
FR_Table.Fields("[SHIFT_ID]") = "D"
FR_Table.Fields("[ACTIVE]") = 0
FR_Table.Update

FR_Table.Close
FR_Database.Close

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdChange_Click()

Screen.MousePointer = vbHourglass

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

OP_ID = Val(textFrom.Text)

Dim sSQL As String

Dim iCount As Integer

sSQL = "SELECT * FROM [TERMINATION] WHERE [OPERATOR ID TERM]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OPERATOR ID TERM]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop
 
sSQL = "SELECT * FROM [TERMINATION] WHERE [OPERATOR ID TERM DD]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OPERATOR ID TERM DD]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop
sSQL = "SELECT * FROM [TERMINATION] WHERE [OPERATOR ID FIRING]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OPERATOR ID FIRING]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop
sSQL = "SELECT * FROM [TERMINATION] WHERE [OPERATOR ID FIRING DD]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OPERATOR ID FIRING DD]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop

sSQL = "SELECT * FROM [TERMINATION] WHERE [OPERATOR ID INSP]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OPERATOR ID INSP]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop

sSQL = "SELECT * FROM [WORK SHEET] WHERE [OP_ID]=" & OP_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Do Until FR_Table.EOF
    FR_Table.Edit
    FR_Table.Fields("[OP_ID]") = Val(textTo.Text)
    FR_Table.Update
    iCount = iCount + 1
    FR_Table.MoveNext
Loop

FR_Table.Close
FR_Database.Close

Screen.MousePointer = vbDefault

MsgBox "Complete Change " & iCount, vbInformation, "ATC EBD System"

End Sub

Private Sub cmdDelete_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID] = " & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
 
    Dim iAns As Integer
    iAns = MsgBox("Delete Operator ID " & OP_ID, vbQuestion + vbYesNo, "Equipment Tracking System")
    If (iAns = vbYes) Then
        FR_Table.Delete
    End If
End If

FR_Table.Close
FR_Database.Close

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub


Private Sub cmdRefresh_Click()
Data2.UpdateRecord
cmdRefreshDisplay_Click
End Sub

Private Sub cmdRefreshDisplay_Click()

Dim DEPT_ID As String
Dim SHIFT_ID As String
Dim LOCATION_ID As String

If (optShift_Day.VALUE = True) Then
    SHIFT_ID = "'D'"
End If
If (optShift_Evening.VALUE = True) Then
    SHIFT_ID = "'E'"
End If
If (OptCombined.VALUE = True) Then
    SHIFT_ID = "'D','E'"
End If
         
If (Option1.VALUE = True) Then
     DEPT_ID = "IN('PT')"
End If
If (Option2.VALUE = True) Then
     DEPT_ID = "IN('TM')"
End If
If (Option3.VALUE = True) Then
     DEPT_ID = "IN('IP')"
End If
If (Option4.VALUE = True) Then
     DEPT_ID = "IN('DP')"
End If
If (optAll.VALUE = True) Then
     DEPT_ID = "IN('PT','TM','IP','DP')"
End If
                                
If (OptionNot.VALUE = True) Then
     DEPT_ID = "NOT IN('PT','TM','IP','DP')"
End If
                                
Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [MASTER_ID],[OP_ID],[FIRST] & ' ' & [LAST]," & _
              "[BARCODE],[DEPT_ID],[SHIFT_ID]," & _
              "[ACTIVE],[LOCATION_ID] " & _
        "FROM [BARCODE] " & _
        "WHERE [SHIFT_ID] IN (" & SHIFT_ID & ") AND " & _
               "[DEPT_ID] " & DEPT_ID & "  " & _
        "ORDER BY [DEPT_ID],[ACTIVE] DESC,[LAST]"
                                                                                                                        
                                                                                                                        
If (OptionNot.VALUE = True) Then
sSQL = "SELECT [MASTER_ID],[OP_ID],[FIRST] & ' ' & [LAST]," & _
              "[BARCODE],[DEPT_ID],[SHIFT_ID]," & _
              "[ACTIVE],[LOCATION_ID] " & _
        "FROM [BARCODE] " & _
        "WHERE [DEPT_ID] " & DEPT_ID & "  " & _
        "ORDER BY [DEPT_ID],[ACTIVE] DESC,[LAST]"
                                                     
End If
                                                                                                                        
                                                                                                                        

sSQLF = "      ||^OP_ID|<Operator                            |^Bar Code                          "
sSQLF = sSQLF & "|^Dept|^Shift|^Active|^L_ID"
                                              
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdTest_Click()

Dim iCount As Long
        
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] WHERE [OP_ID]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
        
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        iCount = iCount + 1
        FR_Table.MoveNext
    Loop
    MsgBox "Operator Active " & iCount, vbInformation, "ATC"
Else
    MsgBox "Operator Not Active", vbInformation, "ATC"
End If

 
FR_Table.Close
FR_Database.Close

End Sub


Private Sub cmdTest2_Click()

Dim iCount As Long
        
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
 
sSQL = "SELECT * FROM [TERMINATION] " & _
       "WHERE [OPERATOR ID FIRING]=" & OP_ID & " OR [OPERATOR ID FIRING DD]=" & OP_ID & " OR " & _
             "[OPERATOR ID TERM]=" & OP_ID & " OR [OPERATOR ID TERM DD]=" & OP_ID & " OR " & _
             "[OPERATOR ID INSP]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        iCount = iCount + 1
        FR_Table.MoveNext
    Loop
    MsgBox "Operator Active " & iCount, vbInformation, "ATC"
Else
    MsgBox "Operator Not Active", vbInformation, "ATC"
End If
 
FR_Table.Close
FR_Database.Close

End Sub

Private Sub Command1_Click()
txtDEPT.Text = "PT"
End Sub

Private Sub Command2_Click()
txtDEPT.Text = "TM"
End Sub

Private Sub Command3_Click()
txtDEPT.Text = "IP"
End Sub

Private Sub Command4_Click()
txtDEPT.Text = "DP"
End Sub

Private Sub CommandActive_Click()

If txtActive.Text = "1" Then
    txtActive.Text = "0"
Else
    txtActive.Text = "1"
End If

End Sub

Private Sub CommandReset_Click()

TextLocation.Text = LOCATION_ID
txtFirst.Text = "New"
txtLast.Text = "Operator"
Text2.Text = "D"

End Sub

Private Sub Form_Load()

Caption = "ATC Operators Termination and Plating       " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 10000
MSFlexGrid1.Height = Me.Height - 800

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then
    Select Case UCase(txtPassword.Text)
    Case "ERIK", "GOLD"
                fraDB.Enabled = True
                CommandReset.Visible = True
    Case Else
                fraDB.Enabled = False
    End Select
    txtPassword.Text = "XXXX"
End If

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 2
OP_ID = Val(MSFlexGrid1.Text)

Select Case OP_ID
Case 0
        fraOperator.Enabled = False
Case Else
        fraOperator.Enabled = True
End Select

MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

fraOperator.Enabled = True

Dim sSQL As String
sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
Data2.RecordSource = sSQL
Data2.Refresh

End Sub

Private Sub optAll_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub OptCombined_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub optInspection_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Option1_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Option2_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Option3_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Option4_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub Option5_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub Option6_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub OptionAll_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub OptionJR_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub OptionNY_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click
End Sub

Private Sub OptionNot_Click()
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub optShift_Day_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub optShift_Evening_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub optTF_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub optTM_Click()

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6)
End Sub


Private Sub TextLocation_GotFocus()
TextLocation.SelStart = 0
TextLocation.SelLength = Len(TextLocation)
End Sub

Private Sub txtActive_GotFocus()
txtActive.SelStart = 0
txtActive.SelLength = Len(txtActive)
End Sub

Private Sub txtDEPT_GotFocus()
txtDEPT.SelStart = 0
txtDEPT.SelLength = Len(txtDEPT)
End Sub

Private Sub txtFirst_GotFocus()
txtFirst.SelStart = 0
txtFirst.SelLength = Len(txtFirst)
End Sub

Private Sub txtFR_GotFocus()
txtFR.SelStart = 0
txtFR.SelLength = Len(txtFR)
End Sub

Private Sub txtLast_GotFocus()
txtLast.SelStart = 0
txtLast.SelLength = Len(txtLast)
End Sub

Private Sub txtOP_ID_GotFocus()
txtOP_ID.SelStart = 0
txtOP_ID.SelLength = Len(txtOP_ID)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub
