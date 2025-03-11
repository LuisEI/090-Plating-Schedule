VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfiguration 
   Caption         =   "090 Configuration Plating"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15465
   Icon            =   "090 Configuration.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15457.19
   ScaleMode       =   0  'User
   ScaleWidth      =   15576.96
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandSET_ID2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SET_ID"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   10560
      Width           =   1400
   End
   Begin VB.CommandButton CommandSET_ID 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SET_ID"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   10080
      Width           =   1400
   End
   Begin VB.CommandButton CommandDetail2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detail"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9600
      Width           =   1400
   End
   Begin VB.CommandButton CommandDetail1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detail"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8760
      Width           =   1400
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Set Compute"
      Height          =   300
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CommandSets 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Schedule Sets"
      Height          =   300
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7920
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.Frame fraDateSelect 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   40
      Top             =   5640
      Width           =   4335
      Begin VB.OptionButton OptionYear 
         Caption         =   "Year"
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
         Left            =   3240
         TabIndex        =   57
         Top             =   1440
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
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
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
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
         TabIndex        =   46
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
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
         TabIndex        =   45
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Refresh"
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
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day"
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
         Left            =   3240
         TabIndex        =   43
         Top             =   360
         Width           =   945
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Week"
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
         Left            =   3240
         TabIndex        =   42
         Top             =   720
         Width           =   1065
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "Month"
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
         Left            =   3240
         TabIndex        =   41
         Top             =   1080
         Width           =   945
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   16384001
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1680
         TabIndex        =   49
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   16384001
         CurrentDate     =   38117
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "On Selection Changes click Refresh Command Button"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   1800
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Group DATE_ID"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9240
      Width           =   1400
   End
   Begin VB.CommandButton CommandGroup 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Group DATE_ID"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8400
      Width           =   1400
   End
   Begin VB.CheckBox CheckDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3360
      TabIndex        =   35
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton CommandCountAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Count All"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7920
      Width           =   1400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [WORK SHEET PT], [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   5460
   End
   Begin VB.CommandButton cmdTranfer 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Tranfer Tables"
      Height          =   260
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3960
      Width           =   1965
   End
   Begin VB.TextBox txtDBID 
      Alignment       =   2  'Center
      DataField       =   " "
      DataSource      =   " "
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
      MaxLength       =   1
      TabIndex        =   24
      Text            =   "9"
      Top             =   2880
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Yes"
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
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2880
      Width           =   675
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "No"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2880
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "JR"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NY"
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
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   675
   End
   Begin VB.Frame Frame5 
      Caption         =   " Enables "
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
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   3015
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
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
         MaxLength       =   1
         TabIndex        =   31
         Text            =   "1"
         ToolTipText     =   "ENABLE 1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
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
         MaxLength       =   1
         TabIndex        =   30
         Text            =   "1"
         ToolTipText     =   "ENABLE 1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         MaxLength       =   1
         TabIndex        =   29
         Text            =   "6"
         ToolTipText     =   "ENABLE 1"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
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
         MaxLength       =   1
         TabIndex        =   28
         Text            =   "4"
         ToolTipText     =   "ENABLE 1"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "ATC Tables : "
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
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Department Codes: "
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
         TabIndex        =   13
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Plating Chemistry : "
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Equipment : "
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
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Save 090 Configuration"
      Height          =   260
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1965
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   " "
      DataSource      =   " "
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
      TabIndex        =   5
      Text            =   "12"
      Top             =   2400
      Width           =   600
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "XXXX"
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame fraDB 
      Caption         =   " Data Base Mode "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "JR: 4"
         Height          =   250
         Left            =   1200
         TabIndex        =   21
         Top             =   480
         Width           =   800
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "FIL : 2"
         Height          =   250
         Left            =   1200
         TabIndex        =   20
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Text            =   "8"
         Top             =   480
         Width           =   480
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "NY : 0"
         Height          =   250
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   800
      End
      Begin VB.CommandButton cmdLocal 
         Caption         =   "LCL : 1"
         Height          =   250
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   800
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Configuration.frx":0CCA
      Height          =   4575
      Left            =   5040
      TabIndex        =   32
      ToolTipText     =   "FROM [MACHINE]"
      Top             =   6240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
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
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   3930
      Left            =   5880
      Picture         =   "090 Configuration.frx":0CDE
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   4590
      Left            =   8160
      Picture         =   "090 Configuration.frx":1A850
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label LabelCounter1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3720
      TabIndex        =   62
      ToolTipText     =   "LabelCounter1"
      Top             =   10080
      Width           =   945
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Delete unmatched"
      Height          =   195
      Left            =   2040
      TabIndex        =   61
      Top             =   10560
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [GROUPING]"
      Height          =   255
      Left            =   2040
      TabIndex        =   59
      Top             =   10080
      Width           =   1545
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   2490
      Left            =   12840
      Picture         =   "090 Configuration.frx":38E9A
      Top             =   3720
      Width           =   1830
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLATING JR.MDB"
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
      Left            =   3480
      TabIndex        =   55
      Top             =   4800
      Width           =   1710
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected Date Range"
      Height          =   255
      Left            =   13680
      TabIndex        =   54
      Top             =   8280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "<< DB_SOURCE_ID"
      Height          =   195
      Left            =   4200
      TabIndex        =   53
      Top             =   3000
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [SCHEDULE SETS]"
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   9240
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [WORK SHEET PT]"
      Height          =   255
      Left            =   2040
      TabIndex        =   38
      Top             =   8400
      Width           =   2025
   End
   Begin VB.Label LabelCount1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2040
      TabIndex        =   34
      Top             =   7920
      Width           =   1065
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ATC Plating Tables.MDB"
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
      Left            =   3480
      TabIndex        =   27
      Top             =   4440
      Width           =   2325
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Remote Control NY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   25
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2940
      Left            =   10440
      Picture         =   "090 Configuration.frx":47024
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4260
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2940
      Left            =   10440
      Picture         =   "090 Configuration.frx":6C523
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   4260
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "C:\ATC\Plating\ SET_ID && LETTER_ID.TXT"
      Height          =   195
      Index           =   15
      Left            =   3360
      TabIndex        =   19
      Top             =   1320
      Width           =   3210
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "C:\ATC\090 Configuration.TXT"
      Height          =   195
      Index           =   14
      Left            =   3360
      TabIndex        =   16
      Top             =   720
      Width           =   2220
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Plating Log Sheet Master.xls"
      Height          =   195
      Index           =   13
      Left            =   3360
      TabIndex        =   15
      Top             =   1980
      Width           =   2010
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Excel Setup Sheet : SET_ID"
      Height          =   195
      Index           =   6
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   2040
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Plating Report : SET_ID && LETTER_ID.TXT"
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   8
      Top             =   1020
      Width           =   3150
   End
   Begin VB.Label Label6 
      Caption         =   "LOCATION_ID:"
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
      Top             =   2400
      Width           =   1635
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "090 Configuration.frx":8EC18
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFile_Click()
Text8.Text = 2
End Sub

Private Sub cmdLocal_Click()
Text8.Text = 1
End Sub



Private Sub cmdNext_Click()
If OptionYear.Value = vbTrue Then
    DTPicker1.Value = DateAdd("YYYY", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("YYYY", 1, DTPicker2.Value)
End If
End Sub

Private Sub cmdPrevious_Click()

If OptionYear.Value = vbTrue Then
    DTPicker1.Value = DateAdd("YYYY", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("YYYY", -1, DTPicker2.Value)
End If

End Sub

Private Sub cmdRemote_Click()
Text8.Text = 0
End Sub

Private Sub cmdSave_Click()

Dim iAns As Integer
iAns = MsgBox("Save Changes to Configuration 090 Configuration.TXT", vbYesNo, "ATC Plating System")
If (iAns = vbYes) Then
        DataBase_MODE = Text8.Text
        LOCATION_ID = Text12.Text
        DB_SOURCE_ID = txtDBID.Text
        
        ENABLE1_ATC_TABLES = Text4.Text
        ENABLE2_DEPT_CODES = Text6.Text
        ENABLE3_CHEMISTRY = Text10.Text
        ENABLE4_EQ = Text14.Text
                      
        Configuration (FWRITE)
        
        DataBase_Address
End If

End Sub

Private Sub cmdTranfer_Click()

Screen.MousePointer = vbHourglass
 
Dim sFileTo As String
sFileTo = SERVER_DB_JR & "ATC Plating Tables.MDB"

Dim sFileFR As String
sFileFR = SERVER_DB_NY & "ATC Plating Tables.MDB"

 
MsgBox "Table Transfer Complete", vbInformation, "Data Base Operation"

End Sub

Private Sub Command1_Click()
Text12.Text = "NY"
End Sub

Private Sub Command2_Click()
Text12.Text = "JR"
End Sub

Private Sub Command3_Click()
Text8.Text = 4
End Sub

Private Sub Command4_Click()
txtDBID.Text = 0
End Sub

Private Sub Command5_Click()
txtDBID.Text = 1
End Sub

Private Sub Command6_Click()
Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT format(first([DATE_ID]),'YYYY'), " & _
               "count([DATE_ID])     " & _
        "FROM [SCHEDULE SETS] " & _
        "GROUP BY  format([DATE_ID],'YYYY') "
                       
sSQLF = "   |[DATE_ID]|   Count"
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF
Screen.MousePointer = vbDefault

End Sub

Private Sub Command7_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],[DATE_ID],[TYPE_ID],[SERIES_ID]," & _
              "[EQ BASE]," & _
              "[EQ FINISH],[PART_SA],[SA] " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "ORDER BY [DATE_ID] DESC"
                                   

Dim COUNT As Long

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Table = TO_Database.OpenRecordset(sSQL)

If (TO_Table.RecordCount <> 0) Then
    Do Until TO_Table.EOF
    
        Select Case Mid(TO_Table.Fields("[SERIES_ID]"), 1, 3)
        Case "300"
        
        Case Else
                SET_ID = TO_Table.Fields("[SET_ID]")
                
                Set_Calculation
                
                TO_Table.Edit
                TO_Table.Fields("[PART_SA]") = Format(PART_SA, "0.0")          'PART SA
                TO_Table.Fields("[SA]") = Format(SA, "0.0")                    'SQ FT
                TO_Table.Update
                COUNT = COUNT + 1
                
                CommandSets.Caption = COUNT
                CommandSets.Refresh
                DoEvents
        End Select
        TO_Table.MoveNext
    Loop
End If
TO_Database.Close

MsgBox "Ok Count" & COUNT, vbInformation, "ATC"


End Sub

Private Sub CommandCountAll_Click()

Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT [WS_ID], " & _
               "[DATE_ID]    AS [SQL 5 FR]," & _
               "[WORK ORDER] AS [SQL 1 FR]," & _
               "[LOT NUM]    AS [SQL 2 FR]," & _
               "[CODE_ID]    AS [SQL 3 FR]," & _
               "[QUANTITY]   AS [SQL 4 FR] " & _
        "FROM [WORK SHEET] " & _
        "ORDER BY [LOT NUM] DESC"
                       
sSQLF = "   ||^DATE_ID     |<WORK ORDER           |<LOT NUM              |^CODE_ID|>QUANTITY    "
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

Dim COUNT As Long
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        FR_Table.MoveNext
        DoEvents
    Loop
End If
FR_Database.Close

Screen.MousePointer = vbDefault
 
DoEvents

LabelCount1.Caption = Format(COUNT, "###,###")

MsgBox "Complete " & COUNT, vbInformation, "ATC"

End Sub

Private Sub CommandDetail1_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT format([DATE_ID],'mm-dd-yy'), " & _
               "[SET_ID],[MACHINE_ID]      " & _
        "FROM [WORK SHEET PT] " & _
        "WHERE [DATE_ID]BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "ORDER BY [DATE_ID] DESC"
                       
sSQLF = "   |[DATE_ID]    |^[SET_ID]|^[MACHINE_ID]"
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF


Dim COUNT As Long
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        If CheckDelete.Value = vbChecked Then
            FR_Table.Delete
        End If
        
        FR_Table.MoveNext
        DoEvents
    Loop
End If
FR_Database.Close

Screen.MousePointer = vbDefault
 
DoEvents

LabelCount1.Caption = Format(COUNT, "###,###")

Screen.MousePointer = vbDefault

MsgBox "Complete " & COUNT, vbInformation, "ATC"

End Sub

Private Sub CommandDetail2_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String


 sSQL = "SELECT format([DATE_ID],'mm-dd-yy'), " & _
               "[SET_ID],[SERIES_ID]      " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "ORDER BY [DATE_ID] DESC"
                       
sSQLF = "   |[DATE_ID]    |^[SET_ID]|^[SERIES_ID]"
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF


Dim COUNT As Long
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        If CheckDelete.Value = vbChecked Then
            FR_Table.Delete
        End If
        
        FR_Table.MoveNext
        DoEvents
    Loop
End If
FR_Database.Close

Screen.MousePointer = vbDefault
 
DoEvents

LabelCount1.Caption = Format(COUNT, "###,###")

Screen.MousePointer = vbDefault

MsgBox "Complete " & COUNT, vbInformation, "ATC"
End Sub

Private Sub CommandGroup_Click()

Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT format(first([DATE_ID]),'YYYY'), " & _
               "count([DATE_ID])     " & _
        "FROM [WORK SHEET PT] " & _
        "GROUP BY  format([DATE_ID],'YYYY') "
                       
sSQLF = "   |[DATE_ID]|   Count"
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF
Screen.MousePointer = vbDefault

End Sub

Private Sub CommandSET_ID_Click()


Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT [SET_ID], " & _
               "[DATE_ID]     " & _
        "FROM [SCHEDULE SETS] " & _
        "ORDER BY [SET_ID] ASC "
                       
sSQLF = "   |^[SET_ID]|^[DATE_ID]   "
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF


Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
            LabelCounter1.Caption = FR_Table.Fields("[SET_ID]")
End If
FR_Database.Close

Screen.MousePointer = vbDefault

End Sub

Private Sub CommandSET_ID2_Click()
 

Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQLF As String

 sSQL = "SELECT [SET_ID], " & _
               "[GP_ID]     " & _
        "FROM [GROUPING] " & _
        "WHERE [SET_ID]< " & Val(LabelCounter1.Caption)
                       
sSQLF = "   |^[SET_ID]|^[GP_ID]   "
         
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF


Dim COUNT As Long
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        If CheckDelete.Value = vbChecked Then
            FR_Table.Delete
        End If
        
        FR_Table.MoveNext
        DoEvents
    Loop
End If
FR_Database.Close

Screen.MousePointer = vbDefault
 
DoEvents

LabelCount1.Caption = Format(COUNT, "###,###")

Screen.MousePointer = vbDefault

MsgBox "Complete " & COUNT, vbInformation, "ATC"

End Sub

Private Sub Form_Load()

Caption = ATC_DWG & "     Configuration Plating             " & ATC_VERSION

Text8.Text = DataBase_MODE
Text12.Text = LOCATION_ID
txtDBID.Text = DB_SOURCE_ID
 
Text4.Text = ENABLE1_ATC_TABLES
Text6.Text = ENABLE2_DEPT_CODES
Text10.Text = ENABLE3_CHEMISTRY
Text14.Text = ENABLE4_EQ
 
OptionYear_Click
 
Data1.DatabaseName = DB_PLATING_TERMINATION
 
End Sub

 
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then

    Select Case UCase(txtPassword.Text)
    Case "ERIK"
                MSFlexGrid2.Visible = True
    Case Else
                MSFlexGrid2.Visible = False
    End Select
    txtPassword.Text = "XXXX"
    
End If

End Sub

Private Sub LabelTO_Click()

End Sub

Private Sub OptionYear_Click()

DTPicker1.Value = Format(Date, "1/01/yyyy")
DTPicker2.Value = Format(DateAdd("YYYY", 1, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", -1, DTPicker2.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Year  <<"
cmdNext.Caption = "Year  >>"

End Sub

Private Sub optMonth_Click()
cmdPrevious.Caption = "Month  <<"
cmdNext.Caption = "Month  >>"

DTPicker1.Value = Format(DTPicker1.Value, "mm/01/yyyy")
DTPicker2.Value = Format(DateAdd("m", 1, DTPicker1.Value), "mm/dd/yyyy")
End Sub

Private Sub optWeek_Click()
DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Week  <<"
cmdNext.Caption = "Week  >>"

End Sub
