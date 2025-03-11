VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBarrelCalculation 
   Caption         =   "090 Barrel Calculations"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Barrel Calculations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptionMSA 
      Caption         =   "MSA"
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
      Left            =   4800
      TabIndex        =   76
      Top             =   11640
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton OptionPYRO 
      Caption         =   "Pyro"
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
      TabIndex        =   75
      Top             =   11640
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [DEPT CODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8040
      TabIndex        =   57
      Top             =   9960
      Width           =   6735
      Begin VB.CommandButton cmdAM 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Calculate"
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
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   240
         Width           =   2895
      End
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
         Left            =   3480
         TabIndex        =   74
         Top             =   1680
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
         Left            =   120
         TabIndex        =   73
         Top             =   1680
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
         Left            =   2160
         TabIndex        =   72
         Top             =   720
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
         Left            =   2160
         TabIndex        =   71
         Top             =   1680
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
         TabIndex        =   70
         Top             =   1680
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
         TabIndex        =   69
         Top             =   720
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
         Left            =   4560
         TabIndex        =   68
         Top             =   1680
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
         Left            =   5520
         TabIndex        =   67
         Top             =   1680
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
         Left            =   5520
         TabIndex        =   66
         Top             =   720
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
         Left            =   4560
         TabIndex        =   65
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblStrikeAmp2 
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
         Left            =   4560
         TabIndex        =   64
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblStrikeAmpMin2 
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
         Left            =   5520
         TabIndex        =   63
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label30 
         Caption         =   "Strike 2"
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
         TabIndex        =   62
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label25 
         Caption         =   "Strike 1"
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
         Left            =   120
         TabIndex        =   61
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblStrikeAmpMin1 
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
         Left            =   2160
         TabIndex        =   60
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblStrikeAmp1 
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
         TabIndex        =   59
         Top             =   1200
         Width           =   795
      End
   End
   Begin VB.TextBox txtG1 
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
      Left            =   2280
      TabIndex        =   51
      Top             =   4920
      Width           =   1600
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [SERIES CASE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   3540
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
      Left            =   2280
      TabIndex        =   10
      Top             =   5520
      Width           =   1600
   End
   Begin VB.CommandButton cmdCal 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate SF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
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
      Height          =   1455
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   5535
      Begin VB.OptionButton Option800AB 
         Caption         =   "800 A/B"
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
         Left            =   2280
         TabIndex        =   79
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option600SFL 
         Caption         =   "600 S/F/L"
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
         TabIndex        =   78
         Top             =   840
         Width           =   1455
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
         Left            =   2280
         TabIndex        =   56
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option9 
         Caption         =   "900"
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
         Left            =   4440
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "200"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   855
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
         TabIndex        =   7
         Top             =   360
         Width           =   1815
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
      Height          =   1455
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3615
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
         Left            =   360
         TabIndex        =   82
         Top             =   840
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
         Left            =   1160
         TabIndex        =   81
         Top             =   840
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
         Left            =   1960
         TabIndex        =   80
         Top             =   840
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
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
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
         Left            =   1960
         TabIndex        =   4
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
         Left            =   1160
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [SHOT]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [PCS PER SIDE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Barrel Calculations.frx":0CCA
      Height          =   3615
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "[PCS PER SIDE]"
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
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
      Bindings        =   "090 Barrel Calculations.frx":0CDE
      Height          =   4575
      Left            =   360
      TabIndex        =   17
      ToolTipText     =   "[SERIES CASE]"
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   8070
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
      Height          =   2175
      Left            =   8040
      TabIndex        =   19
      Top             =   7680
      Width           =   6735
      Begin VB.Label Label28 
         Caption         =   "Strike 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   55
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblSKASF1 
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
         Height          =   405
         Left            =   1320
         TabIndex        =   54
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblSKMIN1 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   53
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblSKMIN2 
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
         Height          =   405
         Left            =   5520
         TabIndex        =   48
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblSKASF2 
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
         Height          =   405
         Left            =   4440
         TabIndex        =   50
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label21 
         Caption         =   "Strike 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   49
         Top             =   840
         Width           =   960
      End
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
         Height          =   405
         Left            =   4440
         TabIndex        =   22
         Top             =   1320
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
         Height          =   405
         Left            =   2280
         TabIndex        =   26
         Top             =   1320
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
         Height          =   405
         Left            =   3360
         TabIndex        =   32
         Top             =   1320
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
         Height          =   405
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "HR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   27
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
         Height          =   405
         Left            =   1320
         TabIndex        =   25
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "ASF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   24
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ASF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4440
         TabIndex        =   23
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
         Height          =   405
         Left            =   5520
         TabIndex        =   21
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "HR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   20
         Top             =   480
         Width           =   840
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Barrel Calculations.frx":0CF2
      Height          =   1575
      Left            =   8520
      TabIndex        =   16
      ToolTipText     =   "[DEPT CODE]"
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
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
   Begin VB.Frame Frame3 
      Caption         =   " Media SA "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   33
      Top             =   6960
      Width           =   7575
      Begin VB.Frame Frame1 
         Caption         =   " Gears "
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
         Left            =   240
         TabIndex        =   34
         Top             =   480
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   360
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Bindings        =   "090 Barrel Calculations.frx":0D06
         Height          =   855
         Left            =   240
         TabIndex        =   39
         ToolTipText     =   "[SHOT]"
         Top             =   2280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1508
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
         Left            =   5880
         TabIndex        =   47
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label Label17 
         Caption         =   "Shot : "
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
         Left            =   5040
         TabIndex        =   46
         Top             =   3480
         Width           =   780
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
         Left            =   240
         TabIndex        =   45
         Top             =   3480
         Width           =   1500
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
         Left            =   1920
         TabIndex        =   44
         Top             =   3480
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Select Case,Series,Dept_ID and Number of Gears"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   43
         Top             =   240
         Width           =   5880
      End
      Begin VB.Label Label12 
         Caption         =   "[1] Look in Dept Code Table 2 selected table for SHOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3840
         TabIndex        =   42
         Top             =   720
         Width           =   3585
      End
      Begin VB.Label Label13 
         Caption         =   "Shot Per Series Case Table 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   41
         Top             =   1800
         Width           =   3120
      End
      Begin VB.Label Label14 
         Caption         =   "[2] Look in Table 3 for Shot SA FROM [SHOT]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3840
         TabIndex        =   40
         Top             =   1440
         Width           =   3360
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Copper "
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
      Left            =   4440
      TabIndex        =   77
      Top             =   11280
      Width           =   3255
   End
   Begin VB.Label lblG1 
      Caption         =   "Gear 1 Qty:"
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
      TabIndex        =   52
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Dept Code Table 2 FROM [DEPT CODE]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6120
      TabIndex        =   30
      Top             =   240
      Width           =   2400
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Sq. In. * Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   29
      Top             =   6600
      Width           =   1920
   End
   Begin VB.Label Label9 
      Caption         =   "Sq. In. per Case Size FROM [PCS PER SIDE]"
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
      Left            =   8760
      TabIndex        =   28
      Top             =   6360
      Width           =   4440
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
      TabIndex        =   15
      Top             =   5520
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
      TabIndex        =   14
      Top             =   6120
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
      Left            =   2280
      TabIndex        =   13
      Top             =   6120
      Width           =   1605
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
      TabIndex        =   12
      Top             =   11580
      Width           =   1605
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
      TabIndex        =   11
      Top             =   11580
      Width           =   1620
   End
End
Attribute VB_Name = "frmBarrelCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAM_Click()

If (OptionMSA.Value = True) Then
    TYPE_CU = "MSA"
Else
    TYPE_CU = "PYRO"
End If

TOTAL_QTY = Val(txtQTY.Text)

If (TOTAL_QTY = 0) Then
    Exit Sub
End If

GEAR_1_QTY = Val(txtG1.Text)

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
If (Option7.Value = True) Then
    SERIES_ID = "100"
End If
If (Option2.Value = True) Then
    SERIES_ID = "200"
End If
If (Option9.Value = True) Then
    SERIES_ID = "900"
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

GearMax

If (GEAR_MAX_QTY < GEAR_1_QTY) Then

    MsgBox "Gear Quantity Greater than Max " & GEAR_MAX_QTY, vbInformation, "Plating"
    Exit Sub

End If

BarrelCalculation

'=========================================================
' DISPLAY
'=========================================================
lblPartSA.Caption = Format(SF * TOTAL_QTY / 144, "0.000")
lblMediaSA.Caption = Format(Shot_Qty / 144, "0.000")
lblSumSA.Caption = Format((SF * TOTAL_QTY + Shot_Qty) / 144, "0.000")

lblShot.Caption = SHOT_ID

lblSKASF1.Caption = SKTASF1
lblSKMIN1.Caption = SKTMIN1

lblASF1.Caption = ASF1
lblMIN1.Caption = MIN1

lblASF2.Caption = ASF2
lblMIN2.Caption = MIN2

lblSKASF2.Caption = ASF3
lblSKMIN2.Caption = MIN3
    
lblStrikeAmp1.Caption = Format(SKTASF1 * SA, "0.0")
lblStrikeAmpMin1.Caption = Format(SKTMIN1 * ASF3 * SA, "0.0")
  
lblBaseAmp.Caption = Format(ASF1 * SA, "0.0")
lblBaseAmpMin.Caption = Format(MIN1 * ASF1 * SA, "0.0")

lblFinishAmp.Caption = Format(ASF2 * SA, "0.0")
lblFinishAmpMin.Caption = Format(MIN2 * ASF2 * SA, "0.0")
 
lblStrikeAmp2.Caption = Format(ASF3 * SA, "0.0")
lblStrikeAmpMin2.Caption = Format(MIN3 * ASF3 * SA, "0.0")
 
End Sub

Private Sub cmdCal_Click()

If (OptionMSA.Value = True) Then
    TYPE_CU = "MSA"
Else
    TYPE_CU = "PYRO"
End If

TOTAL_QTY = Val(txtQTY.Text)

If (TOTAL_QTY = 0) Then
    Exit Sub
End If

GEAR_1_QTY = Val(txtG1.Text)

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

If (Option600SFL.Value = True) Then
    SERIES_ID = "600"
End If
If (Option800AB.Value = True) Then
    SERIES_ID = "810"
End If

If (Option1.Value = True) Then
    SERIES_ID = "100"
End If
If (Option7.Value = True) Then
    SERIES_ID = "100"
End If

If (Option2.Value = True) Then
    SERIES_ID = "200"
End If
If (Option9.Value = True) Then
    SERIES_ID = "900"
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

BarrelCalculation
  
lblShot.Caption = SHOT_ID

lblPartSA.Caption = Format(SF * TOTAL_QTY / 144, "0.000")
 
lblMediaSA.Caption = Format(Shot_Qty / 144, "0.000")
 
lblSumSA.Caption = Format((SF * TOTAL_QTY + Shot_Qty) / 144, "0.000")
 
End Sub


Private Sub Form_Load()

Caption = "Barrel Plating Calculations     " & ATC_DWG & "    " & ATC_VERSION
 
Data1.DatabaseName = DB_PLATING_TABLES
Data2.DatabaseName = DB_PLATING_TABLES
Data3.DatabaseName = DB_PLATING_TABLES
Data4.DatabaseName = DB_PLATING_TABLES

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

'090 Plating chg 12/03/2014 Extended Voltages

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SERIES CASE] " & _
       "FROM [SERIES CASE] " & _
       "WHERE mid([SERIES CASE],1,3) IN ('100','200','900','10E','600','800') ORDER BY mid([SERIES CASE],1,3)"
                                   
Data4.RecordSource = sSQL
Data4.Refresh

sSQLF = "    |^Valid Series"
MSFlexGrid4.FormatString = sSQLF


sSQL = "SELECT [DEPT_ID],[DESCRIPTION],[TANK DWG] " & _
       "FROM [DEPT CODE] WHERE [ACTIVE] = 1 AND [TANK]= 'Y' ORDER BY [DEPT_ID]"
                                   
Data1.RecordSource = sSQL
Data1.Refresh

sSQLF = "    |^Dept ID|Plating  Process         |Table                           "
MSFlexGrid1.FormatString = sSQLF
MSFlexGrid1.Height = 6000

sSQL = "SELECT [CASE],format([PCS PER SIDE MAX],'###,###'),[SHOT],format([SF],'0.000') " & _
       "FROM [PCS PER SIDE] " & _
       "WHERE [TYPE]= 'BARREL' ORDER BY [CASE]"
                                   
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    |^Case|>Pcs/Side |^Shot  |^Sq In. "
MSFlexGrid2.FormatString = sSQLF
MSFlexGrid2.Width = 4000

sSQL = "SELECT  [600 SFL],[100 AB],[100 CE],[200 AB],[200 CE] " & _
       "FROM [SHOT] WHERE [SHOT_ID] = 1"
       
Data3.RecordSource = sSQL
Data3.Refresh
 
sSQLF = "    |^600 SFL |^100 AB |^100 CE |^200 AB |^200 CE "

MSFlexGrid3.FormatString = sSQLF
MSFlexGrid3.Width = 5600

MSFlexGrid1_Click

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
DEPT_ID = Val(MSFlexGrid1.Text)
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1

End Sub

Private Sub optCaseE_Click()
lblG1.Visible = True
txtG1.Visible = True
End Sub

Private Sub Option1_Click()
optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True
End Sub
Private Sub Option7_Click()
optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True
End Sub

Private Sub Option2_Click()
optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = False
optCaseC.Enabled = False
optCaseA.Value = True
End Sub

Private Sub Option9_Click()
optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = False
optCaseC.Enabled = True
optCaseC.Value = True
End Sub
