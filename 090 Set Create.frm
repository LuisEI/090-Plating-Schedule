VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmSetCreate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "090  Create Schedule/Grouping"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Set Create.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerExitProgram 
      Interval        =   60000
      Left            =   12000
      Top             =   1560
   End
   Begin VB.CommandButton cmdOLEAN 
      BackColor       =   &H00FFC0FF&
      Caption         =   "BARREL OLEAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   9720
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   6960
      Width           =   1300
   End
   Begin VB.Frame fraBase 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Base"
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
      Left            =   6000
      TabIndex        =   89
      Top             =   6480
      Width           =   3735
      Begin VB.TextBox txtBaseID 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "MACHINE_B_ID"
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
         Height          =   420
         Left            =   360
         TabIndex        =   90
         ToolTipText     =   "Units Produced"
         Top             =   1320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox lblBaseAmpMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "BASE AMP MIN"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   12
         Text            =   "0000.0"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox lblBaseAmp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "BASE AMP"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   11
         Text            =   "00"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox txtBase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "EQ BASE"
         DataSource      =   "Data4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Bindings        =   "090 Set Create.frx":0CCA
         Height          =   975
         Left            =   120
         TabIndex        =   91
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblAMP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   2520
         TabIndex        =   96
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   7
         Left            =   1440
         TabIndex        =   95
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "M#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblMIN1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   93
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblASF1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   92
         Top             =   1200
         Width           =   600
      End
   End
   Begin VB.Frame fraFinish 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Finish "
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
      Left            =   11400
      TabIndex        =   82
      Top             =   6480
      Width           =   3615
      Begin VB.TextBox txtFinishID 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "MACHINE_F_ID"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   83
         ToolTipText     =   "Units Produced"
         Top             =   1320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox lblFinishAmpMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "FINISH AMP MIN"
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
         Height          =   360
         Left            =   2400
         TabIndex        =   17
         Text            =   "0000.0"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox lblFinishAmp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "FINISH AMP"
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
         Height          =   360
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox txtFinish 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "EQ FINISH"
         DataSource      =   "Data4"
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
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Bindings        =   "090 Set Create.frx":0CDE
         Height          =   975
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1720
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblAMP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   2400
         TabIndex        =   88
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   14
         Left            =   1320
         TabIndex        =   87
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "M#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   240
         TabIndex        =   86
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblASF2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblMIN2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblSQFT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblInfo 
         Caption         =   "SF:"
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
         Left            =   2160
         TabIndex        =   85
         Top             =   1200
         Width           =   315
      End
   End
   Begin VB.Frame fraStrike2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Strike "
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
      Left            =   9840
      TabIndex        =   77
      Top             =   6480
      Width           =   1455
      Begin VB.TextBox lblStrikeAmp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "SK2 AMP"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblStrikeAmpMin2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "SK2 MIN"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   78
         ToolTipText     =   "Units Produced"
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   17
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblSKMIN2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   80
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblSKASF2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   1680
         Width           =   600
      End
   End
   Begin VB.Frame fraStrike1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Strike "
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
      Left            =   4440
      TabIndex        =   70
      Top             =   6480
      Width           =   1455
      Begin VB.TextBox lblStrikeAmpMin1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "SK1 MIN"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   72
         ToolTipText     =   "Units Produced"
         Top             =   1200
         Width           =   825
      End
      Begin VB.TextBox lblStrikeAmp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "SK1 AMP"
         DataSource      =   "Data4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   71
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   240
         TabIndex        =   76
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblSKASF1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblSKMIN1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   73
         Top             =   1680
         Width           =   600
      End
   End
   Begin VB.TextBox txtRunQty 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      DataField       =   "RUN QTY"
      DataSource      =   "Data4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   69
      Text            =   "RUN QTY"
      ToolTipText     =   "Units Produced"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      DataField       =   "SPEED"
      DataSource      =   "Data4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   68
      Text            =   "SPEED"
      ToolTipText     =   "Units Produced"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtShot 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      DataField       =   "SHOT_ID"
      DataSource      =   "Data4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   67
      Text            =   "SHOT_ID"
      ToolTipText     =   "Units Produced"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdBARREL 
      BackColor       =   &H00FFC0FF&
      Caption         =   "BARREL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   9360
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdSBE 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SBE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   9000
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdCorrection 
      Caption         =   "Correction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   62
      Top             =   6960
      Width           =   1300
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Case Size "
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   53
      Top             =   2760
      Width           =   1815
      Begin VB.OptionButton optCaseR 
         BackColor       =   &H00C0FFFF&
         Caption         =   "180R "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   122
         Top             =   1920
         Width           =   1395
      End
      Begin VB.OptionButton OptionCaseL 
         BackColor       =   &H0080FF80&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1110
         TabIndex        =   114
         Top             =   3000
         Width           =   500
      End
      Begin VB.OptionButton optionCaseR 
         BackColor       =   &H0080FF80&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1110
         TabIndex        =   113
         Top             =   2640
         Width           =   500
      End
      Begin VB.OptionButton optionCaseB 
         BackColor       =   &H0080FF80&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   615
         TabIndex        =   112
         Top             =   2640
         Width           =   500
      End
      Begin VB.OptionButton optionCaseA 
         BackColor       =   &H0080FF80&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   111
         Top             =   2640
         Width           =   500
      End
      Begin VB.OptionButton OptionCaseF 
         BackColor       =   &H0080FF80&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   615
         TabIndex        =   106
         Top             =   3000
         Width           =   500
      End
      Begin VB.OptionButton OptionCaseS 
         BackColor       =   &H0080FF80&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   105
         Top             =   3000
         Width           =   500
      End
      Begin VB.OptionButton optCaseE 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   57
         Top             =   1560
         Width           =   800
      End
      Begin VB.OptionButton optCaseC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   800
      End
      Begin VB.OptionButton optCaseB 
         BackColor       =   &H00C0FFFF&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton optCaseA 
         BackColor       =   &H00C0FFFF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   240
         Index           =   22
         Left            =   360
         TabIndex        =   116
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "JAX"
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
         Index           =   21
         Left            =   120
         TabIndex        =   115
         Top             =   2280
         Width           =   1500
      End
   End
   Begin VB.TextBox txtSeries 
      Alignment       =   2  'Center
      DataField       =   "SERIES_ID"
      DataSource      =   "Data4"
      Height          =   300
      Left            =   3360
      TabIndex        =   52
      Text            =   "100"
      ToolTipText     =   "100/200 A/B/C/E"
      Top             =   6600
      Width           =   795
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Create Set "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   46
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton cmdDeleteSet 
         Caption         =   "Delete"
         Height          =   250
         Left            =   3120
         TabIndex        =   65
         Top             =   240
         Width           =   1000
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Series ID"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   2040
         TabIndex        =   58
         Top             =   480
         Width           =   2055
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFF80&
            Caption         =   "900 C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Top             =   1455
            Width           =   1700
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFF80&
            Caption         =   "700"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   1800
            Width           =   1700
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFF80&
            Caption         =   "[3] Olean"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   118
            Top             =   2580
            Width           =   1700
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFF80&
            Caption         =   "800 C/E  MSA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   117
            Top             =   2205
            Width           =   1700
         End
         Begin VB.OptionButton Option600SFL 
            BackColor       =   &H0080FF80&
            Caption         =   "600 S/F/L"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   108
            Top             =   3720
            Width           =   1700
         End
         Begin VB.OptionButton Option800AB 
            BackColor       =   &H0080FF80&
            Caption         =   "800 A/B/R"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   107
            Top             =   3360
            Width           =   1700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "100/710/800"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF80&
            Caption         =   "200 A/B"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   1700
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            Caption         =   "Barrel"
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
            Index           =   20
            Left            =   120
            TabIndex        =   121
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "JAX"
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
            TabIndex        =   110
            Top             =   3000
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "New"
         Height          =   250
         Left            =   2040
         TabIndex        =   49
         Top             =   240
         Width           =   1000
      End
      Begin VB.OptionButton optT 
         BackColor       =   &H00FFFF00&
         Caption         =   "Barrel"
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
         TabIndex        =   48
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optS 
         BackColor       =   &H00C0FFFF&
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
         Left            =   360
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox txtTYPE_ID 
      Alignment       =   2  'Center
      DataField       =   "TYPE_ID"
      DataSource      =   "Data4"
      Height          =   300
      Left            =   2040
      TabIndex        =   45
      Text            =   "1"
      Top             =   6600
      Width           =   1155
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6960
      Width           =   1300
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh6"
      Height          =   300
      Left            =   8880
      TabIndex        =   42
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 [GROUPING]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   9120
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SCHEDULE SETS"
      Top             =   720
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [GROUPING]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   11160
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Frame fraWS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GP_ID : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   31
      Top             =   7320
      Width           =   4215
      Begin VB.CommandButton cmdValidatePG 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Validate Plating"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Validate Plating Group"
         Top             =   4080
         Width           =   1500
      End
      Begin VB.CommandButton cmdValidateFormat 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Validate ATC Part"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton cmdValidateATCPart 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Validate W.O."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdPrintGrouping 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Print Plating Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3240
         Width           =   1500
      End
      Begin VB.TextBox txtLETTER_ID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LETTER_ID"
         DataSource      =   "Data3"
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
         MaxLength       =   1
         TabIndex        =   8
         ToolTipText     =   "Units Produced"
         Top             =   3600
         Width           =   345
      End
      Begin VB.CommandButton cmdDeleteWO 
         Caption         =   "Delete WO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   61
         Top             =   2775
         Width           =   1500
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add WO/Lot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   44
         Top             =   2400
         Width           =   1500
      End
      Begin VB.CommandButton cmdRefresh3 
         Caption         =   "Update Record"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   4080
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "P4 BASE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
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
         ToolTipText     =   "Units Produced"
         Top             =   3525
         Width           =   825
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "P2 BASE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Units Produced"
         Top             =   2775
         Width           =   825
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "P3 BASE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Units Produced"
         Top             =   3150
         Width           =   825
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "P1 BASE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         ToolTipText     =   "Units Produced"
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txtSQ 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "QTY"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         ToolTipText     =   "Units Produced"
         Top             =   1845
         Width           =   825
      End
      Begin VB.TextBox txtLot 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LOT NUM"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1470
         Width           =   2280
      End
      Begin VB.TextBox txtATCPart 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txtWorkOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         MaxLength       =   13
         TabIndex        =   0
         Top             =   240
         Width           =   2280
      End
      Begin VB.Label LabelBOM 
         BackColor       =   &H8000000B&
         Caption         =   "BOM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2520
         TabIndex        =   109
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Run Letter for Grouping must be present"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   360
         TabIndex        =   103
         Top             =   4440
         Width           =   3195
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run Letter:"
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
         Left            =   2640
         TabIndex        =   66
         Top             =   3600
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2NG/Head 4:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   3525
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2G/Head 3:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1NG/Head 2:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1G/Head 1:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Qty:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   35
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "W.O./Lot No :"
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
         Index           =   9
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATC Part :"
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
         Index           =   8
         Left            =   240
         TabIndex        =   33
         Top             =   1095
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lot Number :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   1470
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
      Height          =   300
      Left            =   8640
      TabIndex        =   29
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [GROUPING]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Frame fraDateSelect 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   300
         Left            =   3000
         TabIndex        =   40
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   300
         Left            =   1920
         TabIndex        =   28
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
         Height          =   300
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
         Height          =   300
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
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
         Format          =   121765889
         CurrentDate     =   38117
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "333"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dept_ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   50
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.TextBox txtSet 
      Alignment       =   2  'Center
      DataField       =   "SET NUMBER"
      DataSource      =   "Data4"
      Height          =   300
      Left            =   1200
      TabIndex        =   22
      Text            =   "1"
      Top             =   6600
      Width           =   795
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   3660
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Set Create.frx":0CF2
      Height          =   6255
      Left            =   4440
      TabIndex        =   21
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11033
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Set Create.frx":0D06
      Height          =   2655
      Left            =   4560
      TabIndex        =   30
      Top             =   9480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4683
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "090 Set Create.frx":0D1A
      Height          =   735
      Left            =   4560
      TabIndex        =   41
      Top             =   8760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TimerExitProgram_Timer"
      Height          =   300
      Left            =   12600
      TabIndex        =   104
      Top             =   1560
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set_ID:"
      Height          =   300
      Left            =   240
      TabIndex        =   23
      Top             =   6600
      Width           =   735
   End
End
Attribute VB_Name = "frmSetCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount = 0) Then
                                
      MsgBox "No Set in Place", vbCritical, "ATC Plating"
      FR_Table.Close
      FR_Database.Close
      Exit Sub
End If

sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
FR_Table.AddNew
FR_Table.Fields("[SET_ID]") = SET_ID
FR_Table.Fields("[WORK ORDER]") = "WORK ORDER"
FR_Table.Fields("[LOT NUM]") = "LOT NUM"
FR_Table.Fields("[ATC PART]") = "ATC PT"
FR_Table.Update
FR_Table.Close
FR_Database.Close

cmdRefresh2_Click
MSFlexGrid2_Click

txtWorkOrder.SetFocus

End Sub

Private Sub cmdBARREL_Click()

GearMax

Select Case NUMBER_HEADS
Case 1, 2, 3, 4
    If (GEAR_MAX_QTY < GEAR_1_QTY) Then
    
        MsgBox "Gear 1 Quantity Greater than Max " & GEAR_MAX_QTY, vbInformation, "Plating"
        Exit Sub
    
    End If
End Select

Select Case NUMBER_HEADS
Case 2, 3, 4
    If (GEAR_MAX_QTY < GEAR_2_QTY) Then
    
        MsgBox "Gear 2 Quantity Greater than Max " & GEAR_MAX_QTY, vbInformation, "Plating"
        Exit Sub
    
    End If
End Select

Select Case NUMBER_HEADS
Case 3, 4
    If (GEAR_MAX_QTY < GEAR_3_QTY) Then
    
        MsgBox "Gear 3 Quantity Greater than Max " & GEAR_MAX_QTY, vbInformation, "Plating"
        Exit Sub
    
    End If
End Select

Select Case NUMBER_HEADS
Case 4
    If (GEAR_MAX_QTY < GEAR_4_QTY) Then
    
        MsgBox "Gear 4 Quantity Greater than Max " & GEAR_MAX_QTY, vbInformation, "Plating"
        Exit Sub
    
    End If
End Select

BarrelCalculation

'=========================================================
' DISPLAY
'=========================================================

txtShot.Text = SHOT_ID
txtSpeed.Text = SPEED_ID


Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    FR_Table.Edit
    FR_Table.Fields("[SPEED 2]") = SPEED_ID
    FR_Table.Fields("[SPEED 3]") = SPEED_ID
    FR_Table.Fields("[SPEED 4]") = SPEED_ID
    FR_Table.Fields("[SPEED 5]") = SPEED_ID
    FR_Table.Fields("[SPEED 6]") = SPEED_ID
    FR_Table.Fields("[SPEED 7]") = SPEED_ID
    FR_Table.Fields("[SPEED 8]") = SPEED_ID
    FR_Table.Update
End If
FR_Table.Close

lblSKASF1.Visible = True
lblSKMIN1.Visible = True

lblSKASF2.Visible = True
lblSKMIN2.Visible = True

lblSKASF1.Caption = SKTASF1
lblSKMIN1.Caption = SKTMIN1

lblASF1.Caption = ASF1
lblMIN1.Caption = MIN1

lblASF2.Caption = ASF2
lblMIN2.Caption = MIN2

lblSKASF2.Caption = ASF3
lblSKMIN2.Caption = MIN3
    
lblSQFT.Caption = Format(SA, "0.0")

lblStrikeAmp1.Text = Format(SKTASF1 * SA, "0.0")
lblStrikeAmpMin1.Text = Format(SKTMIN1 * SKTASF1 * SA, "0.0")
  
lblBaseAmp.Text = Format(ASF1 * SA, "0.0")
lblBaseAmpMin.Text = Format(MIN1 * ASF1 * SA, "0.0")

lblFinishAmp.Text = Format(ASF2 * SA, "0.0")
lblFinishAmpMin.Text = Format(MIN2 * ASF2 * SA, "0.0")
 
lblStrikeAmp2.Text = Format(ASF3 * SA, "0.0")
lblStrikeAmpMin2.Text = Format(MIN3 * ASF3 * SA, "0.0")

End Sub

Private Sub cmdCalculate_Click()

DATE_ID = DTPicker3.Value

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
    
sSQL = "SELECT count([WORK ORDER]),format(sum([QTY]),'###,####')," & _
              "sum([P1 BASE]) AS [SQL QTY1]," & _
              "sum([P2 BASE]) AS [SQL QTY2]," & _
              "sum([P3 BASE]) AS [SQL QTY3]," & _
              "sum([P4 BASE]) AS [SQL QTY4] " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "GROUP BY [SET_ID]"
    
Set FR_Table = FR_Database.OpenRecordset(sSQL)

GEAR_1_QTY = 0
GEAR_2_QTY = 0
GEAR_3_QTY = 0
GEAR_4_QTY = 0
TOTAL_QTY = 0
NUMBER_HEADS = 0

If (FR_Table.RecordCount = 0) Then
    Exit Sub
Else
    
    Dim I As Integer
    Dim sBuff As String
    
    For I = 1 To 4
        sBuff = "[SQL QTY" & I & "]"
        If (FR_Table.Fields(sBuff) <> 0) Then
            NUMBER_HEADS = NUMBER_HEADS + 1
        End If
    Next I
    
    GEAR_1_QTY = FR_Table.Fields("[SQL QTY1]")
    GEAR_2_QTY = FR_Table.Fields("[SQL QTY2]")
    GEAR_3_QTY = FR_Table.Fields("[SQL QTY3]")
    GEAR_4_QTY = FR_Table.Fields("[SQL QTY4]")
            
    TOTAL_QTY = FR_Table.Fields("[SQL QTY1]") + FR_Table.Fields("[SQL QTY2]") + FR_Table.Fields("[SQL QTY3]") + FR_Table.Fields("[SQL QTY4]")
    
End If

If (TOTAL_QTY = 0) Then
    Exit Sub
End If

txtRunQty.Text = TOTAL_QTY

Data4.UpdateRecord

Set FR_Table = FR_Database.OpenRecordset(sSQL)

sSQL = "SELECT * FROM [SCHEDULE SETS] " & _
       "WHERE [SET_ID] =" & SET_ID & " AND " & _
             "[DATE_ID]=#" & DATE_ID & "# AND " & _
             "[DEPT_ID]=" & DEPT_ID
             
Set FR_Table = FR_Database.OpenRecordset(sSQL)

CASE_SIZE_ID = Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)
SERIES_ID = Val(Mid(FR_Table.Fields("[SERIES_ID]"), 1, 3))
  
'chg 05/19/2014
Dim EQ_BASE As Integer

EQ_BASE = FR_Table.Fields("[EQ BASE]")


Set FR_Table = FR_Database.OpenRecordset(sSQL)

sSQL = "SELECT * FROM [MACHINE] WHERE [NUMBER]=" & EQ_BASE
             
Set FR_Table = FR_Database.OpenRecordset(sSQL)


Select Case FR_Table.Fields("[SERIES]")
Case "MSA", "PYRO"
        TYPE_CU = FR_Table.Fields("[SERIES]")
Case Else
        TYPE_CU = "NA"
End Select

'Case 18, 73
'        TYPE_CU = "MSA"
'Case 17, 75
'        TYPE_CU = "PYRO"
'Case Else
'        TYPE_CU = "NA"
'End Select
  
Select Case TYPE_ID
Case "BARREL"
        
        Select Case SERIES_ID
        Case "300"
                    cmdOLEAN_Click
        Case Else
                    cmdBARREL_Click
        End Select

Case "SBE"
        cmdSBE_Click
End Select
  
Data4.UpdateRecord
  
cmdRefresh_Click
  
End Sub

Private Sub cmdCorrection_Click()

If (optT.Value = True) Then
    txtTYPE_ID.Text = "BARREL"
     TYPE_ID = "BARREL"
    If (Option7.Value = True) Then
            SERIES_ID = 100
    End If
Else
    txtTYPE_ID.Text = "SBE"
    TYPE_ID = "SBE"
    If (Option7.Value = True) Then
            SERIES_ID = 700
    End If
End If

If (Option1.Value = True) Then
    SERIES_ID = 100
End If
If (Option2.Value = True) Then
    SERIES_ID = 200
End If
If (Option9.Value = True) Then
    SERIES_ID = 900
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

If (Option3.Value = True) Then
    SERIES_ID = 300
    CASE_SIZE_ID = "N"
End If

If (Option800AB.Value = True) Then
    SERIES_ID = 810
End If
If (Option600SFL.Value = True) Then
    SERIES_ID = 600
End If

If (optionCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
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


txtSeries.Text = SERIES_ID & CASE_SIZE_ID

Select Case DEPT_ID
Case 285, 287, 525, 529
            Option3.Value = True
            txtSeries.Text = "300N"
            optT.Value = True
            txtTYPE_ID.Text = "BARREL"
            TYPE_ID = "BARREL"
Case 286, 288, 526, 528
            Option3.Value = True
            txtSeries.Text = "300N"
            optT.Value = True
            txtTYPE_ID.Text = "BARREL"
            TYPE_ID = "BARREL"
End Select

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]= " & SET_ID
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    FR_Table.Edit
    FR_Table.Fields("[SERIES_ID]") = SERIES_ID & CASE_SIZE_ID
    FR_Table.Fields("[TYPE_ID]") = TYPE_ID
    FR_Table.Update
End If
FR_Table.Close

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh6_Click

End Sub

Private Sub cmdCreate_Click()

' NY ATC PRODUCT
If (Option1.Value = True) Then
    SERIES_ID = 100
End If
If (Option2.Value = True) Then
    SERIES_ID = 200
End If
If (Option8.Value = True) Then
    SERIES_ID = 800
End If
If (Option9.Value = True) Then
    SERIES_ID = 900
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

If (optCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
End If

'JAX PRODUCT

If (Option800AB.Value = True) Then
    SERIES_ID = 810
End If
If (Option600SFL.Value = True) Then
    SERIES_ID = 600
End If

If (optionCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optionCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optionCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
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

'OLEAN PRODUCT
If (Option3.Value = True) Then
    SERIES_ID = 300
    CASE_SIZE_ID = "N"
End If


'Case 600, 810

'Select Case SERIES_ID
'Case 600
 '   If (optS.Value = False) Then
  '      MsgBox "600 and 800 Product Must be SBE", vbCritical, "ATC Plating System"
   '     Exit Sub
   ' End If

'End Select


If (Option3.Value = True) Then
    If (optT.Value = False) Then
        MsgBox "Olean Product must be Barrel", vbCritical, "ATC Plating System"
        Exit Sub
    End If
End If

DATE_ID = DTPicker3.Value

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [SET NUMBER] " & _
       "FROM [SCHEDULE SETS] " & _
       "WHERE [DATE_ID]=#" & DATE_ID & "# ORDER BY [SET NUMBER] DESC"
          
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim SET_COUNTER As Integer
Dim TOTAL_COUNT As Long
Dim END_COUNT As Long

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            TOTAL_COUNT = TOTAL_COUNT + 1
            END_COUNT = FR_Table.Fields("[SET NUMBER]")
            FR_Table.MoveNext
        Loop
End If

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim MAX_SET_NUM As Integer

If (FR_Table.RecordCount = 0) Then
        SET_NUMBER = 1
Else
        
        Do Until FR_Table.EOF
        
        '    SET_COUNTER = SET_COUNTER + 1
        '    If (SET_COUNTER <> FR_Table.Fields("[SET NUMBER]")) Then
        '        Exit Do
        '    End If
            
            If MAX_SET_NUM < FR_Table.Fields("[SET NUMBER]") Then
                MAX_SET_NUM = FR_Table.Fields("[SET NUMBER]")
            End If
            
            FR_Table.MoveNext
        Loop


        'If (SET_COUNTER <> END_COUNT) Then
        '    SET_NUMBER = SET_COUNTER
        'Else
            SET_NUMBER = MAX_SET_NUM + 1
        'End If

End If

sSQL = "SELECT [SET_ID],[TYPE_ID],[SERIES_ID],[DEPT_ID],[DATE_ID],[SET NUMBER] " & _
        "FROM [SCHEDULE SETS] WHERE [DATE_ID]=#" & DATE_ID & "# " & _
        "ORDER BY [SET_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

FR_Table.AddNew

If (optT.Value = True) Then
    FR_Table.Fields("[TYPE_ID]") = "BARREL"
    If (Option7.Value = True) Then
        SERIES_ID = 700
    End If
End If

If (optS.Value = True) Then
    FR_Table.Fields("[TYPE_ID]") = "SBE"
    If (Option7.Value = True) Then
        SERIES_ID = 700
    End If
End If

FR_Table.Fields("[SERIES_ID]") = SERIES_ID & CASE_SIZE_ID
FR_Table.Fields("[DEPT_ID]") = DEPT_ID
FR_Table.Fields("[DATE_ID]") = DTPicker3.Value
FR_Table.Fields("[SET NUMBER]") = SET_NUMBER

FR_Table.Update
FR_Table.Close
FR_Database.Close

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

MSFlexGrid2_Click

End Sub

Private Sub cmdDeleteSet_Click()
Dim iAns As Integer

iAns = MsgBox("Delete Set", vbQuestion + vbYesNo, "ATC Plating")

If (iAns = vbYes) Then

    Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
    
    Dim sSQL As String
    
    sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
        
    If (FR_Table.RecordCount <> 0) Then
            FR_Table.Delete
    End If
            
    sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
            
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
    If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            FR_Table.Delete
            FR_Table.MoveNext
        Loop
    End If
                   
    FR_Table.Close
    FR_Database.Close

cmdRefresh_Click

End If

End Sub

Private Sub cmdDeleteWO_Click()

Dim iAns As Integer

iAns = MsgBox("Delete Work Order", vbQuestion + vbYesNo, "ATC Plating")

If (iAns = vbYes) Then

    Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
    
    Dim sSQL As String
    
    sSQL = "SELECT * FROM [GROUPING] WHERE [GP_ID]=" & GP_ID
            
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
    FR_Table.Delete
    
    FR_Table.Close
    FR_Database.Close

cmdRefresh2_Click

End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()

DTPicker3.Value = DateAdd("D", 1, DTPicker3.Value)
DATE_ID = DTPicker3.Value

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

End Sub

Private Sub cmdOLEAN_Click()


SKTASF1 = 0
SKTMIN1 = 0
ASF1 = 0
MIN1 = 0
ASF2 = 0
MIN2 = 0
ASF3 = 0
MIN3 = 0

SKTASF2 = 0
SKTMIN2 = 0

'================================================================================
'   [1]  SURFACE AREA PART
'================================================================================
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * " & _
       "FROM [TBL PLATING OLEAN] " & _
       "WHERE [ID] = 1"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    ASF1 = FR_Table.Fields("[BASE AMP]")
    MIN1 = FR_Table.Fields("[BASE MIN]")
    ASF2 = FR_Table.Fields("[FINISH AMP]")
    MIN2 = FR_Table.Fields("[FINISH MIN]")
End If

FR_Table.Close


'=========================================================
' DISPLAY
'=========================================================
SHOT_ID = 0
SPEED_ID = 0
SA = 1
txtShot.Text = SHOT_ID
txtSpeed.Text = SPEED_ID

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    FR_Table.Edit
    FR_Table.Fields("[SPEED 2]") = SPEED_ID
    FR_Table.Fields("[SPEED 3]") = SPEED_ID
    FR_Table.Fields("[SPEED 4]") = SPEED_ID
    FR_Table.Fields("[SPEED 5]") = SPEED_ID
    FR_Table.Fields("[SPEED 6]") = SPEED_ID
    FR_Table.Fields("[SPEED 7]") = SPEED_ID
    FR_Table.Fields("[SPEED 8]") = SPEED_ID
    FR_Table.Update
End If
FR_Table.Close

lblSKASF1.Visible = True
lblSKMIN1.Visible = True

lblSKASF2.Visible = True
lblSKMIN2.Visible = True

lblSKASF1.Caption = SKTASF1
lblSKMIN1.Caption = SKTMIN1

lblASF1.Caption = ASF1
lblMIN1.Caption = MIN1

lblASF2.Caption = ASF2
lblMIN2.Caption = MIN2

lblSKASF2.Caption = ASF3
lblSKMIN2.Caption = MIN3
    
lblSQFT.Caption = Format(SA, "0.0")

lblStrikeAmp1.Text = Format(SKTASF1 * SA, "0.0")
lblStrikeAmpMin1.Text = Format(SKTMIN1 * SKTASF1 * SA, "0.0")
  
lblBaseAmp.Text = Format(ASF1 * SA * NUMBER_HEADS, "0.0")
lblBaseAmpMin.Text = Format(MIN1 * ASF1 * SA * NUMBER_HEADS, "0.0")

lblFinishAmp.Text = Format(ASF2 * SA * NUMBER_HEADS, "0.0")
lblFinishAmpMin.Text = Format(MIN2 * ASF2 * SA * NUMBER_HEADS, "0.0")
 
lblStrikeAmp2.Text = Format(ASF3 * SA, "0.0")
lblStrikeAmpMin2.Text = Format(MIN3 * ASF3 * SA, "0.0")


End Sub

Private Sub cmdPrevious_Click()

DTPicker3.Value = DateAdd("D", -1, DTPicker3.Value)
DATE_ID = DTPicker3.Value

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

End Sub

Private Sub cmdPrintGrouping_Click()

LETTER_ID = UCase(txtLETTER_ID.Text)
Get_DV  'UPDATE [DV] FROM [ATC PART] TABLE [GROUPING]
PrintGrouping

End Sub

Private Sub cmdRefresh_Click()

DATE_ID = DTPicker3.Value

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],[DATE_ID],[TYPE_ID],[SERIES_ID],[RUN QTY]," & _
              "[EQ BASE],[BASE AMP],[BASE AMP MIN]," & _
              "[EQ FINISH],[FINISH AMP],[FINISH AMP MIN]  " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]=#" & DATE_ID & "# AND [DEPT_ID]=" & DEPT_ID & " " & _
        "ORDER BY [SET_ID] DESC"
                                   
Data1.RecordSource = sSQL
Data1.Refresh


sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series||Base|Amp      |Amp Min  |Finish|Amp    |Amp Min   "

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefresh2_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [GP_ID],[WORK ORDER],mid([ATC PART],1,8),format([QTY],'###,####')," & _
              "format([P1 BASE],'###,####'),format([P2 BASE],'###,####')," & _
              "format([P3 BASE],'###,####'),format([P4 BASE],'###,####'),[LETTER_ID] " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "ORDER BY [GP_ID] DESC"
                                                                                                          
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    ||^Work Order/Lot    |^ATC Part    |>QTY        |>1G/H 1     |>1NG/H 2    |>2G/H 3     |>2NG/H 4  |^Run  "

MSFlexGrid2.FormatString = sSQLF

sSQL = "SELECT count([WORK ORDER]),format(sum([P1 BASE])+sum([P2 BASE])+sum([P3 BASE])+sum([P4 BASE]),'###,####')," & _
              "format(sum([P1 BASE]),'###,####'),format(sum([P2 BASE]),'###,####')," & _
              "format(sum([P3 BASE]),'###,####'),format(sum([P4 BASE]),'###,####') " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "GROUP BY [SET_ID]"
       
Data5.RecordSource = sSQL
Data5.Refresh

sSQLF = "    |^COUNT                   |>QTY        |>1G/H 1        |>1NG/H 2        |>2G/H 3        |>2NG/H 4     "

MSFlexGrid3.FormatString = sSQLF

End Sub

Private Sub cmdRefresh3_Click()

Data3.UpdateRecord
cmdRefresh2_Click

ValidateSeries

Dim X As Long
X = Val(Text3.Text) + Val(Text5.Text) + Val(Text4.Text) + Val(Text6.Text)

If (X = 0) Then
    MsgBox "No Pieces in Gear/Head", vbCritical, "Plating System"
End If

End Sub

Private Sub cmdRefresh6_Click()

Dim sSQL As String
Dim sSQLF As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    BASE_ID = FR_Table.Fields("[BASE_ID]")
    FINISH_ID = FR_Table.Fields("[FINISH_ID]")
End If
FR_Table.Close
FR_Database.Close

Select Case Mid(FINISH_ID, 1, 2)
Case "Lw"
        FINISH_ID = "Lw"
Case "Hg"
        FINISH_ID = "Hg"
End Select

PROCESS_ID = "BASE"

Select Case SERIES_ID
Case "800"
            sSQL = "SELECT [MACHINE_ID],[NUMBER],[NAME],[SERIES] " & _
                   "FROM [MACHINE] " & _
                   "WHERE   [TYPE]='" & TYPE_ID & "' AND " & _
                        "[PROCESS]='" & PROCESS_ID & "' AND " & _
                          "[BF_ID]='" & BASE_ID & "' AND " & _
                    "[LOCATION_ID]='" & LOCATION_ID & "' AND " & _
                         "[SERIES]='MSA' AND" & _
                        "[ACTIVE] = 1"

Case Else
            sSQL = "SELECT [MACHINE_ID],[NUMBER],[NAME],[SERIES] " & _
                   "FROM [MACHINE] " & _
                   "WHERE   [TYPE]='" & TYPE_ID & "' AND " & _
                        "[PROCESS]='" & PROCESS_ID & "' AND " & _
                          "[BF_ID]='" & BASE_ID & "' AND " & _
                    "[LOCATION_ID]='" & LOCATION_ID & "' AND " & _
                        "[ACTIVE] = 1"
End Select

Data6.RecordSource = sSQL
Data6.Refresh

sSQLF = "    ||^M#   |^Name   |<            "

MSFlexGrid6.FormatString = sSQLF

PROCESS_ID = "FINISH"

sSQL = "SELECT [MACHINE_ID],[NUMBER],[NAME] " & _
       "FROM [MACHINE] " & _
       "WHERE    [TYPE]='" & TYPE_ID & "' AND " & _
             "[PROCESS]='" & PROCESS_ID & "' AND " & _
               "[BF_ID]='" & FINISH_ID & "'AND " & _
         "[LOCATION_ID]='" & LOCATION_ID & "' AND " & _
             "[ACTIVE] = 1"
                                                                                                          
Data7.RecordSource = sSQL
Data7.Refresh

sSQLF = "    ||^M#   |^Name   "

MSFlexGrid7.FormatString = sSQLF

End Sub

Private Sub cmdReport_Click()
  
Select Case SERIES_CASE_ID
Case "300N"
           ExcelReportOlean
Case Else
           ExcelReport
End Select

End Sub

Private Sub cmdReset_Click()

DTPicker3.Value = Date
DATE_ID = DTPicker3.Value

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

End Sub

Private Sub cmdSBE_Click()

SBE_Calculation
'=========================================================
'
'=========================================================
    
'lblMediaSA.Caption = MEDIA_SA
'lblPartSA.Caption = Format(SA, "0.000")
'lblSumSA.Caption = Format(SA, "0.0")

lblASF1.Caption = ASF1
lblMIN1.Caption = MIN1
  
lblASF2.Caption = ASF2
lblMIN2.Caption = MIN2

lblSKASF1.Visible = False
lblSKMIN1.Visible = False

lblSKASF2.Visible = False
lblSKMIN2.Visible = False

txtShot.Text = SHOT_ID

lblSQFT.Caption = Format(SA, "0.0")

lblBaseAmp.Text = Format(ASF1 * SA, "0.0")
lblBaseAmpMin.Text = Format(MIN1 * ASF1 * SA / 60, "0.0")

lblFinishAmp.Text = Format(ASF2 * SA, "0.0")
lblFinishAmpMin.Text = Format(MIN2 * ASF2 * SA / 60, "0.0")
 
End Sub

Private Sub cmdValidateATCPart_Click()

If (Len(txtWorkOrder.Text) = 10) Then
    MsgBox "Not a Work Order", vbInformation, "Plating"
    Exit Sub
End If

'===========================================================================
'   LOOK UP ATC PART IN W.O. SCHEDULE
'===========================================================================

txtWorkOrder.Text = UCase(txtWorkOrder.Text)
ATC_PART_ID = "NA"

Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)

sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & txtWorkOrder.Text & "'"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
  
If (FR_Table.RecordCount <> 0) Then

    If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
            ATC_PART_ID = FR_Table.Fields("[ATC PART]")
            ValidATCPartCode (Trim(ATC_PART_ID))
    End If

End If
 
End Sub

Private Sub cmdValidateFormat_Click()

'If (Len(txtWorkOrder.Text) = 10) Then
'    MsgBox "Not a Work Order", vbInformation, "Plating"
'    Exit Sub
'End If

Select Case Mid(txtATCPart.Text, 1, 4)
Case "600F", "600S", "600L", "800A", "800B"


Case Else

        ValidATCPartCode (Trim(txtATCPart.Text))
        
        Select Case DEPT_ID
        Case 540, 539, 541, 549, 544, 551, 546, 552
                
                Select Case SERIES_ID
                Case "800"
                        If (Mid(txtATCPart.Text, 3) <> "800") Then
                            MsgBox "ATC Part Code not 800 Series", vbCritical, "ATC Plating System"
                        End If
                End Select
        
        Case Else
                
        End Select
End Select

End Sub

Private Sub cmdValidatePG_Click()

LETTER_ID = Trim(txtLETTER_ID.Text)
If Len(LETTER_ID & "X") <> 1 Then
    Screen.MousePointer = vbHourglass
    Get_DV
    frmMSGBOX.Show
    Screen.MousePointer = vbDefault
End If

'
'IF SBE USE ONLY 100/700 SERIES
'

'090 Plating chg 12/03/2014 Extended Voltages

Select Case TYPE_ID
Case "SBE"
            Dim sBuff As String

            Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
            Dim sSQL As String
            sSQL = "SELECT * " & _
                   "FROM [GROUPING] " & _
                   "WHERE [GP_ID]=" & GP_ID & " AND " & _
                     "MID([ATC PART],1,3) NOT IN ('100','710','180','10E','70E','71E','600','800')"
                 
            Set FR_Table = FR_Database.OpenRecordset(sSQL)
            If (FR_Table.RecordCount <> 0) Then
                   sBuff = "Work Order " & FR_Table.Fields("[WORK ORDER]") & vbNewLine
                   sBuff = sBuff & "Part Number " & FR_Table.Fields("[ATC PART]") & vbNewLine
                   sBuff = sBuff & "can not be used in SBE   " & vbNewLine
                   sBuff = sBuff & "SBE only 100,700,600,800 Series"
            End If
            FR_Table.Close
            FR_Database.Close
            
            MsgBox sBuff, vbCritical + vbInformation, "ATC Plating Schedule"
                        
End Select

End Sub

Private Sub Form_Load()

Caption = "Create Schedule/Grouping     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION
Data5.DatabaseName = DB_PLATING_TERMINATION
Data6.DatabaseName = DB_PLATING_TERMINATION
Data7.DatabaseName = DB_PLATING_TERMINATION

If (ENABLE5_SP_TEST = 1) Then
    cmdValidateATCPart.Visible = True
    cmdValidateFormat.Visible = True
End If

MSFlexGrid1.Width = 10800
MSFlexGrid2.Width = 9400
MSFlexGrid3.Width = 9000
MSFlexGrid3.Height = 700

lblCode.Caption = DEPT_ID

LabelBOM.Caption = ""

Select Case DEPT_ID
Case 285, 525, 287, 529
            Option3.Enabled = True
End Select
Select Case DEPT_ID
Case 286, 526, 288, 528
            Option3.Enabled = True
End Select

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Dim sSQL As String

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    lblDesc.Caption = FR_Table.Fields("[DESCRIPTION]")
    Select Case FR_Table.Fields("[SBE]")
    Case "Y"
                optS.Visible = True
    Case "N"
                optS.Visible = False
    Case Else
    End Select
    
    Select Case FR_Table.Fields("[STRIKE1]")
    Case "Y"
                fraStrike1.Visible = True
    Case "N"
                fraStrike1.Visible = False
    Case Else
    End Select
    
    Select Case FR_Table.Fields("[STRIKE2]")
    Case "Y"
                fraStrike2.Visible = True
    Case "N"
                fraStrike2.Visible = False
    Case Else
    End Select
    
End If
FR_Table.Close
FR_Database.Close

DTPicker3.Value = Date
DATE_ID = DTPicker3.Value

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

cmdRefresh6_Click   'LIST MACHINE

Option1_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
SET_ID = Val(MSFlexGrid1.Text)
  
LabelBOM.Caption = ""
  
MSFlexGrid1.Col = 6
SERIES_ID = Val(MSFlexGrid1.Text)
SERIES_CASE_ID = MSFlexGrid1.Text
  
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

Dim sSQL As String
sSQL = "SELECT * FROM [SCHEDULE SETS] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND " & _
             "[DATE_ID]=#" & DATE_ID & "# AND " & _
             "[DEPT_ID]=" & DEPT_ID
                                
Data4.RecordSource = sSQL
Data4.Refresh

TYPE_ID = txtTYPE_ID.Text

Select Case TYPE_ID
Case "SBE"
            lblAMP(0).Caption = "Amp Hr"
            lblAMP(1).Caption = "Amp Hr"
Case "BARREL"
            lblAMP(0).Caption = "Amp Min"
            lblAMP(1).Caption = "Amp Min"
End Select

cmdRefresh6_Click

cmdRefresh2_Click
MSFlexGrid2_Click

lblSKASF1.Caption = ""
lblSKMIN1.Caption = ""

lblASF1.Caption = ""
lblMIN1.Caption = ""

lblASF2.Caption = ""
lblMIN2.Caption = ""

lblSKASF2.Caption = ""
lblSKMIN2.Caption = ""

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
GP_ID = Val(MSFlexGrid2.Text)
  
fraWS.Caption = "GP_ID : " & GP_ID
  
MSFlexGrid2.Col = 3
ATC_PART_ID = MSFlexGrid2.Text
  
Select Case Mid(ATC_PART_ID, 4, 1)
Case "A", "B", "R", "F", "S", "L"
  
        ValidPart (ATC_PART_ID)
        
End Select
  
  
Dim sSQL As String
sSQL = "SELECT * FROM [GROUPING] WHERE [GP_ID]=" & GP_ID
                                
Data3.RecordSource = sSQL
Data3.Refresh
  
MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

End Sub


Private Sub MSFlexGrid6_Click()

MSFlexGrid6.Col = 1
txtBaseID.Text = Val(MSFlexGrid6.Text)
  
MSFlexGrid6.Col = 2
txtBase.Text = Val(MSFlexGrid6.Text)
     
MSFlexGrid6.Col = 0
MSFlexGrid6.ColSel = MSFlexGrid6.Cols - 1

End Sub

Private Sub MSFlexGrid7_Click()

MSFlexGrid7.Col = 1
txtFinishID.Text = Val(MSFlexGrid7.Text)
  
MSFlexGrid7.Col = 2
txtFinish.Text = Val(MSFlexGrid7.Text)
     
MSFlexGrid7.Col = 0
MSFlexGrid7.ColSel = MSFlexGrid7.Cols - 1

End Sub

Private Sub Option1_Click()

optT.Visible = True
optS.Visible = True

If optS.Value = True Then
    Option2.Enabled = False
    Option9.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
Else
    Option2.Enabled = True
    Option9.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
End If

optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

End Sub

Private Sub Option2_Click()
optT.Visible = True
If optS.Value = True Then
    Option2.Enabled = False
    Option9.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
Else
    Option2.Enabled = True
    Option9.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
End If

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = False
optCaseC.Enabled = False
optCaseA.Value = True

optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False
End Sub

Private Sub Option3_Click()
optT.Visible = True
 
End Sub

Private Sub Option600SFL_Click()

optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = False
optCaseC.Enabled = False
 
optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
 
OptionCaseS.Value = True
OptionCaseS.Enabled = True
OptionCaseL.Enabled = True
OptionCaseF.Enabled = True

Select Case DEPT_ID
Case 539, 549, 551
        optT.Value = True
        optT.Visible = True
        optS.Value = False
Case Else
        optS.Value = True
        optS.Visible = True
        optT.Value = False
End Select

End Sub

Private Sub Option7_Click()
optT.Visible = True
If optS.Value = True Then
    Option2.Enabled = False
    Option9.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
Else
    Option2.Enabled = True
    Option9.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
End If

optCaseA.Enabled = True
optCaseB.Enabled = True
optCaseE.Enabled = True
optCaseC.Enabled = True

optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

End Sub

Private Sub Option8_Click()
optT.Visible = True
If optS.Value = True Then
    Option2.Enabled = False
    Option9.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
Else
    Option2.Enabled = True
    Option9.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
End If

optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = True
optCaseC.Enabled = True

optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

End Sub

Private Sub Option800AB_Click()

optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = False
optCaseC.Enabled = False
 
optionCaseA.Enabled = True
optionCaseB.Enabled = True
optionCaseR.Enabled = True
optionCaseA.Value = True

OptionCaseS.Enabled = False
OptionCaseL.Enabled = False
OptionCaseF.Enabled = False

Select Case DEPT_ID
Case 539, 549, 551
        optT.Value = True
        optT.Visible = True
        optS.Value = False
Case Else
        optS.Value = True
        optS.Visible = True
        optT.Value = False
End Select

End Sub

Private Sub Option9_Click()
optT.Visible = True
If optS.Value = True Then
    Option2.Enabled = False
    Option9.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
Else
    Option2.Enabled = True
    Option9.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
End If

optCaseA.Enabled = False
optCaseB.Enabled = False
optCaseE.Enabled = False
optCaseC.Enabled = True
optCaseC.Value = True

optionCaseA.Enabled = False
optionCaseB.Enabled = False
optionCaseR.Enabled = False
OptionCaseS.Enabled = False
OptionCaseF.Enabled = False
OptionCaseL.Enabled = False

End Sub

Private Sub optS_Click()

Option2.Enabled = False
Option9.Enabled = False
Option7.Enabled = False
Option8.Enabled = False

cmdRefresh6_Click
End Sub

Private Sub optT_Click()

Option2.Enabled = True
Option9.Enabled = True
Option7.Enabled = True
Option8.Enabled = True

cmdRefresh6_Click
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10)
End Sub

Private Sub Text1_LostFocus()

Text1.Text = UCase(Text1.Text)

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

Private Sub TimerExitProgram_Timer()

If Format(Time, "hh AM/PM") = "01 AM" Then

    Select Case DataBase_MODE
    Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            DataBase_MODE = DATABASE_MODE_REM_NY
            LOCATION_ID = "NY"
    End Select

    ConfigComputer_DB (2)
    End
End If

Strangelove

End Sub

Private Sub txtATCPart_GotFocus()
txtATCPart.SelStart = 0
txtATCPart.SelLength = Len(txtATCPart)
End Sub

Private Sub txtATCPart_LostFocus()

txtATCPart.Text = UCase(txtATCPart.Text)

Select Case SERIES_CASE_ID
Case "300N"

Case Else
        Select Case Len(txtWorkOrder.Text)
        Case 10
                'LOT NUMBER
        Case 12
                ValidATCPartCodeTest (Trim(txtATCPart.Text))
        End Select
End Select

End Sub

Private Sub txtLETTER_ID_GotFocus()
txtLETTER_ID.SelStart = 0
txtLETTER_ID.SelLength = Len(txtLETTER_ID)
End Sub

Private Sub txtLETTER_ID_LostFocus()
txtLETTER_ID.Text = UCase(txtLETTER_ID.Text)
End Sub

Private Sub txtLot_GotFocus()
txtLot.SelStart = 0
txtLot.SelLength = Len(txtLot)
End Sub


Private Sub txtLot_LostFocus()

txtLot.Text = UCase(txtLot.Text)

End Sub

Private Sub txtSQ_GotFocus()
txtSQ.SelStart = 0
txtSQ.SelLength = Len(txtSQ)
End Sub

Private Sub txtWorkOrder_GotFocus()
txtWorkOrder.SelStart = 0
txtWorkOrder.SelLength = Len(txtWorkOrder)
End Sub

Private Sub txtWorkOrder_LostFocus()

ALERT_MESSAGE = "ok"

txtWorkOrder.Text = Replace(txtWorkOrder.Text, ".", "")

txtWorkOrder.Text = Mid(UCase(txtWorkOrder.Text), 1, 12)

LabelBOM.Caption = ""

Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)

If (Val(txtSQ.Text) = 0) Then

    'If (Len(txtWorkOrder.Text) = 10) Then
     '      txtATCPart.Text = LotDecode(txtWorkOrder.Text)
      '     txtLot.Text = txtWorkOrder.Text
    'Else
    
            Dim sSQL As String
            
            sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & txtWorkOrder.Text & "'"
                          
            Set FR_Table = FR_Database.OpenRecordset(sSQL)
             
            If (FR_Table.RecordCount <> 0) Then
            
                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                        txtATCPart.Text = FR_Table.Fields("[ATC PART]")
                End If
                If IsNull(FR_Table.Fields("[LOT NUM]")) = False Then
                        txtLot.Text = FR_Table.Fields("[LOT NUM]")
                End If
                If IsNull(FR_Table.Fields("[START QTY]")) = False Then
                        txtSQ.Text = FR_Table.Fields("[START QTY]")
                Else
                        txtSQ.Text = 0
                End If
                
                Select Case Mid(txtATCPart.Text, 1, 4)
                Case "600F", "600S", "600L", "800A", "800B"
                
                Case Else
                        Select Case DEPT_ID
                        Case 553, 554, 558
                                If IsNull(FR_Table.Fields("[BOM]")) = False Then
                                    LabelBOM.Caption = FR_Table.Fields("[BOM]")
                                    If Mid(FR_Table.Fields("[BOM]"), 9, 1) = "C" Then
                                        ALERT_MESSAGE = "REVISAR" & vbNewLine & "ESTAS PIEZAS" & vbNewLine & "NO TIENEN PLATING"
                                    End If
                                End If
                        Case Else
                                If IsNull(FR_Table.Fields("[BOM]")) = False Then
                                    LabelBOM.Caption = FR_Table.Fields("[BOM]")
                                    If Mid(FR_Table.Fields("[BOM]"), 9, 1) <> "C" Then
                                        ALERT_MESSAGE = "REVISAR" & vbNewLine & "ESTAS PIEZAS" & vbNewLine & "YA TIENEN PLATING"
                                    End If
                                End If
                        End Select
                        
                        Select Case DEPT_ID
                        Case 286, 288
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) = "C" Then
                                        ALERT_MESSAGE = "NO NECESITA" & vbNewLine & "PLATINADO"
                                    Else
                                    'OK
                                    End If
                                End If
                        End Select
                
                        Select Case DEPT_ID
                        Case 286
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) = "T" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 288
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) = "W" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 523, 538
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) <> "Y" Or Mid(FR_Table.Fields("[ATC PART]"), 9, 2) = "YN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        
                        Case 528
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) <> "T" Or Mid(FR_Table.Fields("[ATC PART]"), 9, 2) = "TN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 533, 526
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 1) <> "W" Or Mid(FR_Table.Fields("[ATC PART]"), 9, 2) = "WN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 534
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If (Mid(FR_Table.Fields("[ATC PART]"), 9, 1) <> "P" And Mid(FR_Table.Fields("[ATC PART]"), 9, 1) <> "W" And Mid(FR_Table.Fields("[ATC PART]"), 9, 2) <> "P ") Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 551
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 2) <> "TN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 539
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 2) <> "WN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        Case 552
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 2) <> "PN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                                
                        Case 549
                                If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
                                   
                                    If Mid(FR_Table.Fields("[ATC PART]"), 9, 2) <> "YN" Then
                                        ALERT_MESSAGE = "PLATINADO" & vbNewLine & "EQUIVOCADO"
                                    End If
                                End If
                        
                        End Select
                        
                
                
                End Select
            Else
                If (Len(txtWorkOrder.Text) = 10) Then
                    Select Case Mid(txtWorkOrder.Text, 1, 1)
                    Case "A" To "Z"
                            txtATCPart.Text = LotDecode(txtWorkOrder.Text)
                            txtLot.Text = txtWorkOrder.Text
                    End Select
                End If
            End If
            
            FR_Table.Close
            FR_Database.Close

End If

Select Case ALERT_MESSAGE
Case "ok"

Case Else
            frmAlert.Show vbModal
End Select

Data4.UpdateRecord

End Sub
