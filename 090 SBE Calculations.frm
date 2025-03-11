VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSBECalculation 
   Caption         =   "090 SBE Plating Calculations"
   ClientHeight    =   12825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16860
   Icon            =   "090 SBE Calculations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12825
   ScaleWidth      =   16860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandMedia_SA 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Media_SA"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox TextDV_ID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   69
      Text            =   "1000"
      Top             =   7200
      Width           =   1200
   End
   Begin VB.CommandButton cmdAM 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Calculate Amps/Amp Hr"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   10680
      Width           =   2415
   End
   Begin VB.Frame Frame3 
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
      Left            =   7320
      TabIndex        =   39
      Top             =   11160
      Width           =   7695
      Begin VB.Label Label18 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3960
         TabIndex        =   49
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label19 
         Caption         =   "Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   48
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label20 
         Caption         =   "Amps Hr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   47
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblBaseAmpMin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   46
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblBaseAmp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   45
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   44
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblFinishAmp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4920
         TabIndex        =   43
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblFinishAmpMin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   42
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label24 
         Caption         =   "Amps Hr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   41
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label27 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4920
         TabIndex        =   40
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Heads "
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   33
      Top             =   6000
      Width           =   3135
      Begin VB.OptionButton OptionG4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2085
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton OptionG3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1470
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton OptionG2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   855
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton OptionG1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 FROM [TBL SBE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.CommandButton cmdTBL_SBE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TBL SBE"
      Height          =   300
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [SERIES CASE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [DEPT CODE] "
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.TextBox txtQTY 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   6
      Top             =   7920
      Width           =   1200
   End
   Begin VB.CommandButton cmdCal 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate SF"
      Height          =   250
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   " Series ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton Option600SFL 
         Caption         =   "600 S/F/L   JAX"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   66
         Top             =   2040
         Width           =   2415
      End
      Begin VB.OptionButton Option800AB 
         Caption         =   "800 A/B/R   JAX"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   65
         Top             =   2460
         Width           =   2415
      End
      Begin VB.OptionButton Option7 
         Caption         =   "700"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   38
         Top             =   780
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "900 C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "200 A/B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   4
         Top             =   1620
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "100/710/800"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " Case Size "
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
      Begin VB.OptionButton OptionCaseR 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   67
         Top             =   1200
         Width           =   495
      End
      Begin VB.OptionButton OptionCaseS 
         Caption         =   "S"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   61
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton OptionCaseL 
         Caption         =   "L"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   60
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton OptionCaseF 
         Caption         =   "F"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   59
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optCaseA 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   58
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optCaseB 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   57
         Top             =   780
         Width           =   495
      End
      Begin VB.OptionButton optCaseC 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   56
         Top             =   1620
         Width           =   495
      End
      Begin VB.OptionButton optCaseE 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   55
         Top             =   2040
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [PCS PER SIDE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   3780
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 SBE Calculations.frx":0CCA
      Height          =   3495
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "[PCS PER SIDE]"
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6165
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 SBE Calculations.frx":0CDE
      Height          =   1575
      Left            =   9360
      TabIndex        =   12
      ToolTipText     =   "[DEPT CODE]"
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2778
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "090 SBE Calculations.frx":0CF2
      Height          =   5415
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "[SERIES CASE]"
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9551
      _Version        =   393216
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
   Begin VB.Frame Frame2 
      Caption         =   " Table Lookups "
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
      Left            =   360
      TabIndex        =   15
      Top             =   11160
      Width           =   6735
      Begin VB.Label lblASF2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4560
         TabIndex        =   18
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblMIN1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   22
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3480
         TabIndex        =   28
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label15 
         Caption         =   "Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "MIN1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblASF1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "ASF1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   20
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "ASF2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4560
         TabIndex        =   19
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblMIN2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5520
         TabIndex        =   17
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "MIN2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5520
         TabIndex        =   16
         Top             =   480
         Width           =   840
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
      Bindings        =   "090 SBE Calculations.frx":0D06
      Height          =   1215
      Left            =   3960
      TabIndex        =   32
      ToolTipText     =   "[TBL SBE]"
      Top             =   9600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "090 SBE Calculations.frx":0D1A
      Height          =   3495
      Left            =   10560
      TabIndex        =   68
      ToolTipText     =   "[PCS PER SIDE]"
      Top             =   4200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
      _Version        =   393216
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
   Begin VB.Label Label25 
      Caption         =   "FROM [TBL SBE 144]  WHERE CASE_ID AND [QTY]<="
      Height          =   465
      Left            =   7200
      TabIndex        =   74
      Top             =   5040
      Width           =   2280
   End
   Begin VB.Label LabelDEPT_ID 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "532"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7800
      TabIndex        =   72
      Top             =   9240
      Width           =   480
   End
   Begin VB.Label Label22 
      Caption         =   "DEPT_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   71
      Top             =   9240
      Width           =   1020
   End
   Begin VB.Label Label21 
      Caption         =   "DV_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   70
      Top             =   7200
      Width           =   1020
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "RP 553,554,555,556  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12600
      TabIndex        =   64
      Top             =   8760
      Width           =   2310
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "LW 285,286,525,526,530,532,533,534 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10320
      TabIndex        =   63
      Top             =   8400
      Width           =   4140
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "TN 287,288,528,529"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8640
      TabIndex        =   62
      Top             =   8040
      Width           =   2160
   End
   Begin VB.Label Label7 
      Caption         =   "FROM [SERIES CASE]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   2280
   End
   Begin VB.Label Label3 
      Caption         =   "FROM [TBL SBE]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   53
      Top             =   9240
      Width           =   2520
   End
   Begin VB.Label Label17 
      Caption         =   "Media : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   51
      Top             =   9000
      Width           =   780
   End
   Begin VB.Label lblShot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   50
      Top             =   9000
      Width           =   1200
   End
   Begin VB.Label lblMediaSA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   30
      Top             =   9600
      Width           =   1200
   End
   Begin VB.Label lblMedia 
      Caption         =   "Media SA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   29
      Top             =   9600
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Dept Code Table 2 FROM [DEPT CODE] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9480
      TabIndex        =   26
      Top             =   120
      Width           =   4560
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sq. In. * Quantity / 144"
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
      Left            =   1155
      TabIndex        =   25
      Top             =   7680
      Width           =   1890
   End
   Begin VB.Label Label9 
      Caption         =   "Sq. In. per Case Size Table 1 FROM [PCS PER SIDE]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5880
      TabIndex        =   24
      Top             =   4080
      Width           =   3000
   End
   Begin VB.Label Label10 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   11
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Part SA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   10
      Top             =   8460
      Width           =   1500
   End
   Begin VB.Label lblPartSA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   9
      Top             =   8460
      Width           =   1200
   End
   Begin VB.Label lblSumSA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   8
      Top             =   10200
      Width           =   1200
   End
   Begin VB.Label lblSum 
      Caption         =   "Sum SA (sq ft)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   7
      Top             =   10200
      Width           =   1620
   End
End
Attribute VB_Name = "frmSBECalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAM_Click()


DV_ID = Val(TextDV_ID.Text)

TOTAL_QTY = Val(txtQTY.Text)

If (TOTAL_QTY = 0) Then
    Exit Sub
End If

If (OptionG1.Value = True) Then
        NUMBER_HEADS = 1
End If
If (OptionG2.Value = True) Then
        NUMBER_HEADS = 2
End If
If (OptionG3.Value = True) Then
        NUMBER_HEADS = 3
End If
If (OptionG4.Value = True) Then
        NUMBER_HEADS = 4
End If

If (optCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optCaseC.Value = True) Then
    CASE_SIZE_ID = "C"
End If
If (optCaseE.Value = True) Then
    CASE_SIZE_ID = "E"
End If

If (OptionCaseF.Value = True) Then
    CASE_SIZE_ID = "F"
End If
If (OptionCaseS.Value = True) Then
    CASE_SIZE_ID = "S"
End If
If (OptionCaseL.Value = True) Then
    CASE_SIZE_ID = "L"
End If
If (optionCaseR.Value = True) Then
    CASE_SIZE_ID = "r"
End If

If (Option1.Value = True) Then
    SERIES_ID = "100"
End If
If (Option2.Value = True) Then
    SERIES_ID = "200"
End If
If (Option7.Value = True) Then
    SERIES_ID = "700"
End If
If (Option9.Value = True) Then
    SERIES_ID = "900"
End If

If (Option600SFL.Value = True) Then
    SERIES_ID = "600"
End If
If (Option800AB.Value = True) Then
    SERIES_ID = "810"
End If

SBE_Calculation

'=========================================================
'
'=========================================================

lblShot.Caption = SHOT_ID

lblASF1.Caption = ASF1
lblMIN1.Caption = MIN1
  
lblASF2.Caption = ASF2
lblMIN2.Caption = MIN2
    
lblMediaSA.Caption = Media_SA
lblPartSA.Caption = Format(PART_SA, "0.000")
lblSumSA.Caption = Format(SA, "0.0")
    
lblBaseAmp.Caption = Format(ASF1 * SA, "0.0")
lblBaseAmpMin.Caption = Format(MIN1 * ASF1 * SA / 60, "0.0")

lblFinishAmp.Caption = Format(ASF2 * SA, "0.0")
lblFinishAmpMin.Caption = Format(MIN2 * ASF2 * SA / 60, "0.0")
 
End Sub

Private Sub cmdCal_Click()

Dim Qty As Long

Qty = Val(txtQTY.Text)

If (Qty = 0) Then
    Exit Sub
End If

If (OptionG1.Value = True) Then
Qty = Val(txtQTY.Text)
End If
If (OptionG2.Value = True) Then
Qty = Val(txtQTY.Text) / 2
End If
If (OptionG3.Value = True) Then
Qty = Val(txtQTY.Text) / 3
End If
If (OptionG4.Value = True) Then
Qty = Val(txtQTY.Text) / 4
End If

If (optCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optCaseC.Value = True) Then
    CASE_SIZE_ID = "C"
End If
If (optCaseE.Value = True) Then
    CASE_SIZE_ID = "E"
End If

If (Option1.Value = True) Then
    SERIES_ID = "100"
End If
If (Option2.Value = True) Then
    SERIES_ID = "200"
End If
If (Option7.Value = True) Then
    SERIES_ID = "700"
End If
If (Option9.Value = True) Then
    SERIES_ID = "900"
End If

'================================================================================
'   [1]  SURFACE AREA PART
'================================================================================
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [CASE],[SF] FROM [PCS PER SIDE] " & _
       "WHERE [CASE] ='" & CASE_SIZE_ID & "'"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim PART_SA As Double
Dim Media_SA As Double
Dim SA_Sum As Double

PART_SA = FR_Table.Fields("[SF]") * Qty / 144

'================================================================================
'   [2]  MEDIA SURFACE AREA
'================================================================================

sSQL = "SELECT [QTY],[MEDIA SF] " & _
       "FROM [TBL SBE 144] " & _
       "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & Qty & " ORDER BY [QTY] DESC"
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Select Case DEPT_ID
Case 555, 556, 553, 554
        
        Select Case SERIES_ID
        Case "100", "700"
                    Media_SA = 7.125
        Case "200", "900"
                    Media_SA = 9.5
        End Select

Case Else
        Media_SA = FR_Table.Fields("[MEDIA SF]")
End Select

SA_Sum = PART_SA + Media_SA

lblMediaSA.Caption = Media_SA
lblPartSA.Caption = Format(PART_SA, "0.000")
lblSumSA.Caption = Format(SA_Sum, "0.0")
 
FR_Table.Close
FR_Database.Close

End Sub


Private Sub cmdTBL_SBE_Click()

If (optCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optCaseC.Value = True) Then
    CASE_SIZE_ID = "C"
End If
If (optCaseE.Value = True) Then
    CASE_SIZE_ID = "E"
End If

If (OptionCaseS.Value = True) Then
    CASE_SIZE_ID = "S"
End If
If (OptionCaseF.Value = True) Then
    CASE_SIZE_ID = "F"
End If
If (OptionCaseL.Value = True) Then
    CASE_SIZE_ID = "L"
End If

If (optionCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
End If

If (Option1.Value = True) Then
    SERIES_ID = "100"
End If

If (Option2.Value = True) Then
    SERIES_ID = "200"
End If
If (Option7.Value = True) Then
    SERIES_ID = "700"
End If
If (Option9.Value = True) Then
    SERIES_ID = "900"
End If

If (Option600SFL.Value = True) Then
    SERIES_ID = "600"
End If
If (Option800AB.Value = True) Then
    SERIES_ID = "810"
End If

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [CASE],[SERIES]," & _
              "format([ASF1],'0.00'),[MIN1]," & _
              "format([ASF2],'0.00'),[MIN2]," & _
              "format([ASF3],'0.00'),[MIN3], " & _
              "format([ASF4],'0.00'),[MIN4],[MODE],[DV MIN],[DV MAX]," & _
              "format([QTY MIN],'#,###,##0'),format([QTY MAX],'#,###,##0') " & _
        "FROM [TBL SBE] WHERE [CASE]='" & CASE_SIZE_ID & "' "
                                   
sSQL = sSQL & " AND [SERIES_TYPE] =" & SERIES_ID & " "

Data5.RecordSource = sSQL
Data5.Refresh
 
sSQLF = "    |^Case||^NI ASF 1|^NI Min|^TN ASF 2|^TN Min|^LW ASF 2|^LW Min|^RP ASF 2|^RP Min|^Mode "

sSQLF = sSQLF & "|>DV MIN   |>DV  MAX  |>QTY MIN  |>QTY MAX  "
MSFlexGrid5.FormatString = sSQLF
MSFlexGrid5.Height = 1600
MSFlexGrid5.Width = 12500

End Sub

 
Private Sub CommandMedia_SA_Click()

TOTAL_QTY = Val(txtQTY.Text)

If (TOTAL_QTY = 0) Then
    Exit Sub
End If

If (optCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optCaseC.Value = True) Then
    CASE_SIZE_ID = "C"
End If
If (optCaseE.Value = True) Then
    CASE_SIZE_ID = "E"
End If

If (OptionCaseF.Value = True) Then
    CASE_SIZE_ID = "F"
End If
If (OptionCaseS.Value = True) Then
    CASE_SIZE_ID = "S"
End If
If (OptionCaseL.Value = True) Then
    CASE_SIZE_ID = "L"
End If
If (optionCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
End If

'================================================================================
'   [2]  MEDIA SURFACE AREA
'================================================================================
Set FR_Database = OpenDatabase(DB_PLATING_TABLES)

Dim sSQL As String

sSQL = "SELECT [QTY],[MEDIA VOL],[MEDIA SF] " & _
       "FROM [TBL SBE 144] " & _
       "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & TOTAL_QTY & " " & _
       "ORDER BY [QTY] DESC"

Select Case CASE_SIZE_ID
Case "A", "B", "R"
    Select Case SERIES_ID
    Case 810
            sSQL = "SELECT [QTY],[MEDIA VOL],[MEDIA SF] " & _
                   "FROM [TBL SBE ABR JAX] " & _
                   "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & TOTAL_QTY & " " & _
                   "ORDER BY [QTY] DESC"
    End Select
End Select
               
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
        MsgBox "Quantity not Available", vbInformation, "ATC Data Base System"
        Exit Sub
End If

Media_SA = FR_Table.Fields("[MEDIA SF]")

SHOT_ID = FR_Table.Fields("[MEDIA VOL]")

Select Case DEPT_ID
Case 553, 554, 555, 556 'REPLATE
        Select Case SERIES_ID
        Case "100", "700"
                    Media_SA = 7.125
        Case "200", "900"
                    Media_SA = 9.5
        Case "600", "810"
                    Media_SA = 9.5
        End Select

Case Else
        Media_SA = FR_Table.Fields("[MEDIA SF]")
End Select

FR_Database.Close

lblMediaSA.Caption = Media_SA
lblShot.Caption = SHOT_ID

End Sub

Private Sub Form_Load()

Caption = "SBE Plating Calculations     " & ATC_DWG & "    " & ATC_VERSION
 
Select Case LOCATION_ID
Case "NY"
         OptionCaseS.Enabled = True
         OptionCaseF.Enabled = True
         OptionCaseL.Enabled = True
         
         Option600SFL.Enabled = True
         Option800AB.Enabled = True
Case "JR"
         OptionCaseS.Enabled = True
         OptionCaseF.Enabled = True
         OptionCaseL.Enabled = True
         
         Option600SFL.Enabled = True
         Option800AB.Enabled = True
End Select
 
 
Data1.DatabaseName = DB_PLATING_TABLES
Data2.DatabaseName = DB_PLATING_TABLES
Data3.DatabaseName = DB_PLATING_TABLES
Data4.DatabaseName = DB_PLATING_TABLES
Data5.DatabaseName = DB_PLATING_TABLES


Dim sSQL As String
Dim sSQLF As String


sSQL = "SELECT FIRST([CASE]),FIRST([SERIES_TYPE])," & _
              "FIRST([MODE]),FIRST([MANUF]) " & _
        "FROM [TBL SBE] GROUP BY [CASE],[SERIES_TYPE]"
                                   
Data3.RecordSource = sSQL
Data3.Refresh

sSQLF = "    |^CASE|^Series|^MODE |^MANUF    "
MSFlexGrid3.FormatString = sSQLF



sSQL = "SELECT [SERIES CASE] " & _
       "FROM [SERIES CASE] " & _
       "WHERE mid([SERIES CASE],1,3) IN ('100','200','600','700','900','10E','70E','71E',800) " & _
       "ORDER BY mid([SERIES CASE],1,3)"
                                   
Data4.RecordSource = sSQL
Data4.Refresh

sSQLF = "    |^Valid Series"
MSFlexGrid4.FormatString = sSQLF

sSQL = "SELECT [DEPT_ID],[DEPT_JR_ID],[DESCRIPTION],[TANK DWG] " & _
       "FROM [DEPT CODE] " & _
       "WHERE [ACTIVE] = 1 AND [SBE]='Y' ORDER BY [DEPT_ID]"
                                   
Data1.RecordSource = sSQL
Data1.Refresh

sSQLF = "    |^Dept ID|^Dept JR|Plating  Process            |^Table                           "
MSFlexGrid1.FormatString = sSQLF
MSFlexGrid1.Height = 3620

sSQL = "SELECT first([CASE]),first(format([PCS PER SIDE MAX],'###,###')),first(format([SF],'0.0000')) " & _
       "FROM [PCS PER SIDE] " & _
       "GROUP BY [CASE] ORDER BY [CASE]"
                                   
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    |^Case|>Pcs/Sd   |^Sq In.        "
MSFlexGrid2.FormatString = sSQLF

MSFlexGrid2.Width = 3600

MSFlexGrid1_Click

cmdTBL_SBE_Click

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
DEPT_ID = Val(MSFlexGrid1.Text)
 
LabelDEPT_ID.Caption = DEPT_ID
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1

End Sub

Private Sub optCaseA_Click()
cmdTBL_SBE_Click
End Sub

Private Sub optCaseB_Click()
cmdTBL_SBE_Click
End Sub

Private Sub optCaseC_Click()
cmdTBL_SBE_Click
End Sub

Private Sub optCaseE_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option1_Click()

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

cmdTBL_SBE_Click

End Sub

Private Sub Option600SFL_Click()

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

OptionCaseS.Enabled = True
OptionCaseF.Enabled = True
OptionCaseL.Enabled = True

cmdTBL_SBE_Click

End Sub

Private Sub Option7_Click()

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

OptionCaseS.Enabled = True
OptionCaseF.Enabled = True
OptionCaseL.Enabled = True

cmdTBL_SBE_Click

End Sub

Private Sub Option2_Click()

optCaseA.Enabled = True
optCaseA.Value = True

optCaseB.Enabled = True

optCaseE.Enabled = True
optCaseC.Enabled = True

OptionCaseS.Enabled = True
OptionCaseF.Enabled = True
OptionCaseL.Enabled = True

cmdTBL_SBE_Click

End Sub

Private Sub Option800AB_Click()

optCaseA.Enabled = True
optCaseA.Value = True

optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

OptionCaseS.Enabled = True
OptionCaseF.Enabled = True
OptionCaseL.Enabled = True

cmdTBL_SBE_Click

End Sub

Private Sub Option9_Click()

optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = False

optCaseC.Enabled = True
optCaseC.Value = True

OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

cmdTBL_SBE_Click

End Sub

Private Sub OptionCaseF_Click()
cmdTBL_SBE_Click
End Sub

Private Sub OptionCaseL_Click()
cmdTBL_SBE_Click
End Sub

Private Sub OptionCaseR_Click()
cmdTBL_SBE_Click
End Sub

Private Sub OptionCaseS_Click()
cmdTBL_SBE_Click
End Sub

Private Sub TextDV_ID_GotFocus()
TextDV_ID.SelStart = 0
TextDV_ID.SelLength = Len(TextDV_ID)
End Sub

