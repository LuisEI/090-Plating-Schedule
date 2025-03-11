VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmSetReview 
   BackColor       =   &H00FFFFFF&
   Caption         =   "090 Review Schedule/Grouping"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Set Review.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOLEAN 
      BackColor       =   &H00FFC0FF&
      Caption         =   "BARREL OLEAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   11520
      Visible         =   0   'False
      Width           =   1800
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   7200
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrintTXT 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Print Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   3360
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.TextBox txtRunQty 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      DataField       =   "RUN QTY"
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
      Left            =   11640
      TabIndex        =   87
      Text            =   "RUN QTY"
      ToolTipText     =   "Units Produced"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
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
      Left            =   11640
      TabIndex        =   86
      Text            =   "SPEED"
      ToolTipText     =   "Units Produced"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtShot 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      DataField       =   "SHOT_ID"
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
      Left            =   11640
      TabIndex        =   85
      Text            =   "SHOT_ID"
      ToolTipText     =   "Units Produced"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdDEPT 
      Caption         =   "Dept_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      TabIndex        =   76
      Top             =   2880
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "[1] New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton cmdBARREL 
      BackColor       =   &H00FFC0FF&
      Caption         =   "BARREL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   11160
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdSBE 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SBE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   10800
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Frame fraStrike1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Strike "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4560
      TabIndex        =   67
      Top             =   6960
      Width           =   1455
      Begin VB.TextBox lblStrikeAmp1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "SK1 AMP"
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
         Left            =   240
         TabIndex        =   69
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblStrikeAmpMin1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "SK1 MIN"
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
         Left            =   240
         TabIndex        =   68
         ToolTipText     =   "Units Produced"
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblSKMIN1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   720
         TabIndex        =   95
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblSKASF1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   94
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   71
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   70
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.Frame fraStrike2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Strike "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9960
      TabIndex        =   62
      Top             =   6960
      Width           =   1455
      Begin VB.TextBox lblStrikeAmpMin2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "SK2 MIN"
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
         Left            =   240
         TabIndex        =   64
         ToolTipText     =   "Units Produced"
         Top             =   1200
         Width           =   825
      End
      Begin VB.TextBox lblStrikeAmp2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "SK2 AMP"
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
         Left            =   240
         TabIndex        =   63
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblSKASF2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   97
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblSKMIN2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   720
         TabIndex        =   96
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   66
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   65
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Case Size "
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   52
      Top             =   2040
      Width           =   1815
      Begin VB.OptionButton optCaseR 
         BackColor       =   &H00C0FFFF&
         Caption         =   "180R "
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
         TabIndex        =   122
         Top             =   2220
         Width           =   1395
      End
      Begin VB.OptionButton OptionCaseS 
         BackColor       =   &H0080FF80&
         Caption         =   "S"
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
         TabIndex        =   118
         Top             =   3240
         Width           =   500
      End
      Begin VB.OptionButton OptionCaseF 
         BackColor       =   &H0080FF80&
         Caption         =   "F"
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
         Left            =   615
         TabIndex        =   117
         Top             =   3240
         Width           =   500
      End
      Begin VB.OptionButton optionCaseA 
         BackColor       =   &H0080FF80&
         Caption         =   "A"
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
         TabIndex        =   116
         Top             =   2880
         Width           =   500
      End
      Begin VB.OptionButton optionCaseB 
         BackColor       =   &H0080FF80&
         Caption         =   "B"
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
         Left            =   615
         TabIndex        =   115
         Top             =   2880
         Width           =   500
      End
      Begin VB.OptionButton optionCaseR 
         BackColor       =   &H0080FF80&
         Caption         =   "R"
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
         Left            =   1110
         TabIndex        =   114
         Top             =   2880
         Width           =   500
      End
      Begin VB.OptionButton OptionCaseL 
         BackColor       =   &H0080FF80&
         Caption         =   "L"
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
         Left            =   1110
         TabIndex        =   113
         Top             =   3240
         Width           =   500
      End
      Begin VB.OptionButton optCaseE 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E"
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
         Left            =   120
         TabIndex        =   56
         Top             =   1845
         Width           =   495
      End
      Begin VB.OptionButton optCaseC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C"
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
         Left            =   120
         TabIndex        =   55
         Top             =   1470
         Width           =   495
      End
      Begin VB.OptionButton optCaseB 
         BackColor       =   &H00C0FFFF&
         Caption         =   "B"
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
         Left            =   120
         TabIndex        =   54
         Top             =   1095
         Width           =   1200
      End
      Begin VB.OptionButton optCaseA 
         BackColor       =   &H00C0FFFF&
         Caption         =   "A"
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
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Value           =   -1  'True
         Width           =   495
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
         Index           =   24
         Left            =   240
         TabIndex        =   121
         Top             =   360
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
         Height          =   360
         Index           =   22
         Left            =   120
         TabIndex        =   119
         Top             =   2520
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Plating Set "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   49
      Top             =   1320
      Width           =   4215
      Begin VB.CommandButton cmdDeleteSet 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2880
         TabIndex        =   108
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdCorrection 
         Caption         =   "Correction"
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
         TabIndex        =   81
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txtSet 
         Alignment       =   2  'Center
         DataField       =   "SET NUMBER"
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
         Left            =   1080
         TabIndex        =   79
         Text            =   "1"
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox txtTYPE_ID 
         Alignment       =   2  'Center
         DataField       =   "TYPE_ID"
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
         Left            =   1920
         TabIndex        =   78
         Text            =   "1"
         Top             =   5040
         Width           =   1155
      End
      Begin VB.TextBox txtSeries 
         Alignment       =   2  'Center
         DataField       =   "SERIES_ID"
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
         Left            =   3360
         TabIndex        =   77
         Text            =   "100"
         ToolTipText     =   "100/200 A/B/C/E"
         Top             =   5040
         Width           =   675
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Series ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   2040
         TabIndex        =   57
         Top             =   480
         Width           =   2055
         Begin VB.OptionButton Option800AB 
            BackColor       =   &H0080FF80&
            Caption         =   "800 A/B/R"
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
            TabIndex        =   111
            Top             =   3120
            Width           =   1700
         End
         Begin VB.OptionButton Option600SFL 
            BackColor       =   &H0080FF80&
            Caption         =   "600 S/F/L"
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
            TabIndex        =   110
            Top             =   3480
            Width           =   1700
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFF80&
            Caption         =   "800 C/E  MSA"
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
            TabIndex        =   106
            Top             =   2040
            Width           =   1700
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFF80&
            Caption         =   "[3] Olean"
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
            Left            =   120
            TabIndex        =   101
            Top             =   2400
            Width           =   1700
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFF80&
            Caption         =   "700"
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
            TabIndex        =   82
            Top             =   1680
            Width           =   1700
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFF80&
            Caption         =   "900 C"
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
            TabIndex        =   60
            Top             =   1320
            Width           =   1700
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "100/710/800"
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
            TabIndex        =   59
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF80&
            Caption         =   "200 A/B"
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
            TabIndex        =   58
            Top             =   960
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
            Height          =   360
            Index           =   23
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   1700
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
            Index           =   21
            Left            =   120
            TabIndex        =   112
            Top             =   2760
            Width           =   1695
         End
      End
      Begin VB.OptionButton optT 
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
         Height          =   375
         Left            =   240
         TabIndex        =   51
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
         Height          =   375
         Left            =   1200
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Set_ID:"
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
         TabIndex        =   80
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "333"
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
         Left            =   360
         TabIndex        =   75
         Top             =   4560
         Width           =   855
      End
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7200
      Width           =   1500
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 FROM [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   4680
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh6"
      Height          =   300
      Left            =   11280
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [MACHINE] "
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 FROM [GROUPING] WHERE [SET_ID]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   10200
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [GROUPING] WHERE [GP_ID]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   9720
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.Frame fraWS 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   7560
      Width           =   4215
      Begin VB.CommandButton cmdValidatePG 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Validate Plating"
         Height          =   300
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Validate Plating Group"
         Top             =   3960
         Width           =   1500
      End
      Begin VB.CommandButton cmdPrintGrouping 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Print Plating Group"
         Height          =   300
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   3120
         Width           =   1500
      End
      Begin VB.TextBox txtLETTER_ID 
         Alignment       =   2  'Center
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
         TabIndex        =   83
         ToolTipText     =   "Units Produced"
         Top             =   3480
         Width           =   345
      End
      Begin VB.CommandButton cmdDeleteWO 
         Caption         =   "Delete WO"
         Height          =   300
         Left            =   2520
         TabIndex        =   61
         Top             =   2775
         Width           =   1500
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add WO/Lot"
         Height          =   300
         Left            =   2520
         TabIndex        =   30
         Top             =   2400
         Width           =   1500
      End
      Begin VB.CommandButton cmdRefresh3 
         Caption         =   "UpdateRecord"
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   3960
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFC0&
         DataField       =   "P4 BASE"
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
         Left            =   1560
         TabIndex        =   20
         ToolTipText     =   "Units Produced"
         Top             =   3525
         Width           =   825
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFC0&
         DataField       =   "P2 BASE"
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
         Left            =   1560
         TabIndex        =   19
         ToolTipText     =   "Units Produced"
         Top             =   2775
         Width           =   825
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
         DataField       =   "P3 BASE"
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
         Left            =   1560
         TabIndex        =   18
         ToolTipText     =   "Units Produced"
         Top             =   3150
         Width           =   825
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         DataField       =   "P1 BASE"
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
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "Units Produced"
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txtSQ 
         BackColor       =   &H00FFFFC0&
         DataField       =   "QTY"
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
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Units Produced"
         Top             =   1845
         Width           =   825
      End
      Begin VB.TextBox txtLot 
         BackColor       =   &H00FFFFC0&
         DataField       =   "LOT NUM"
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1470
         Width           =   2280
      End
      Begin VB.TextBox txtATCPart 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
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
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txtWorkOrder 
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
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
         Left            =   1560
         MaxLength       =   13
         TabIndex        =   9
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label LabelBOM 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   2520
         TabIndex        =   109
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Run Letter for Grouping must be present"
         Height          =   300
         Index           =   20
         Left            =   480
         TabIndex        =   107
         Top             =   4320
         Width           =   3195
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FROM [GROUPING] "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   1560
         TabIndex        =   103
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run Letter:"
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
         Index           =   11
         Left            =   2520
         TabIndex        =   84
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2NG/Head 4:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   3525
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2G/Head 3:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1NG/Head 2:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1G/Head 1:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Qty:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   16
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "W.O./Lot No :"
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
         Index           =   9
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ATC Part :"
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
         Index           =   8
         Left            =   240
         TabIndex        =   14
         Top             =   1095
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lot Number :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1470
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
      Height          =   300
      Left            =   8400
      TabIndex        =   6
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [GROUPING] WHERE [SET_ID]"
      Connect         =   "Access"
      DatabaseName    =   "\\NY-ENG\SPC Network\Data Base\TERMINATION And PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GROUPING"
      Top             =   9360
      Visible         =   0   'False
      Width           =   5340
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
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdRefresh 
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
         Left            =   3000
         TabIndex        =   25
         Top             =   720
         Width           =   1000
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
         Left            =   1920
         TabIndex        =   5
         Top             =   735
         Width           =   1000
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
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1000
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         Format          =   124387329
         CurrentDate     =   38117
      End
      Begin VB.Label lblInfo 
         Caption         =   "[SCHEDULE SETS]"
         Height          =   300
         Index           =   13
         Left            =   120
         TabIndex        =   104
         Top             =   840
         Width           =   1665
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   4020
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Set Review.frx":0CCA
      Height          =   6735
      Left            =   4440
      TabIndex        =   0
      ToolTipText     =   "FROM [SCHEDULE SETS]"
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11880
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollBars      =   2
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Set Review.frx":0CDE
      Height          =   2295
      Left            =   4680
      TabIndex        =   7
      ToolTipText     =   "FROM [GROUPING] "
      Top             =   9840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollBars      =   2
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
      Bindings        =   "090 Set Review.frx":0CF2
      Height          =   615
      Left            =   4680
      TabIndex        =   26
      ToolTipText     =   "FROM [GROUPING] "
      Top             =   9120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
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
   Begin VB.Frame fraFinish 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Finish "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   11520
      TabIndex        =   31
      Top             =   6960
      Width           =   3615
      Begin VB.TextBox txtFinish 
         Alignment       =   2  'Center
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
         TabIndex        =   44
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblFinishAmp 
         Alignment       =   2  'Center
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
         TabIndex        =   43
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblFinishAmpMin 
         Alignment       =   2  'Center
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
         TabIndex        =   42
         Text            =   "0000.0"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox txtFinishID 
         BackColor       =   &H0000FFFF&
         DataField       =   "MACHINE_F_ID"
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
         TabIndex        =   41
         ToolTipText     =   "Units Produced"
         Top             =   1320
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Bindings        =   "090 Set Review.frx":0D06
         Height          =   975
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "FROM [MACHINE]"
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1720
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
         TabIndex        =   99
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label lblSQFT 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   98
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblMIN2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2880
         TabIndex        =   93
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblASF2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2160
         TabIndex        =   92
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "M#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   48
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   47
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblAMP 
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   46
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame fraBase 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Base"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6120
      TabIndex        =   32
      Top             =   6960
      Width           =   3735
      Begin VB.TextBox txtBase 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "EQ BASE"
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
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblBaseAmp 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "BASE AMP"
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
         TabIndex        =   35
         Text            =   "00"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox lblBaseAmpMin 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         DataField       =   "BASE AMP MIN"
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
         Left            =   2520
         TabIndex        =   34
         Text            =   "0000.0"
         ToolTipText     =   "Units Produced"
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox txtBaseID 
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
         TabIndex        =   33
         ToolTipText     =   "Units Produced"
         Top             =   1320
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Bindings        =   "090 Set Review.frx":0D1A
         Height          =   975
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "FROM [MACHINE]"
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
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
      Begin VB.Label lblASF1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3000
         TabIndex        =   91
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblMIN1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3000
         TabIndex        =   90
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "M#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   40
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   39
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblAMP 
         Caption         =   "Amp Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmSetReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
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

txtShot.Text = SHOT_ID
txtSpeed.Text = SPEED_ID


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

SKTASF1 = 0
SKTMIN1 = 0
ASF1 = 0
MIN1 = 0
ASF2 = 0
MIN2 = 0
ASF3 = 0
MIN3 = 0

DATE_ID = DTPicker3.Value

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
    
sSQL = "SELECT count([WORK ORDER])," & _
         "format(sum([QTY]),'###,####')," & _
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
       "WHERE [SET_ID]=" & SET_ID & " AND " & _
             "[DATE_ID]=#" & DATE_ID & "# AND " & _
             "[DEPT_ID]=" & DEPT_ID
             
Set FR_Table = FR_Database.OpenRecordset(sSQL)

CASE_SIZE_ID = Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)
SERIES_ID = Val(Mid(FR_Table.Fields("[SERIES_ID]"), 1, 3))
  
  
'chg 12/13/2011
Select Case FR_Table.Fields("[EQ BASE]")
Case 18, 73
        TYPE_CU = "MSA"
Case 17, 75
        TYPE_CU = "PYRO"
Case Else
        TYPE_CU = "NA"
End Select
  
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

If (Option3.Value = True) Then
    SERIES_ID = 300
    CASE_SIZE_ID = "N"
End If

txtSeries.Text = SERIES_ID & CASE_SIZE_ID

cmdRefresh6_Click

End Sub

Private Sub cmdCreate_Click()


DATE_ID = DTPicker3.Value

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT count([SET_ID]) AS [SQL COUNT] " & _
        "FROM [SCHEDULE SETS] WHERE [DATE_ID]=#" & DATE_ID & "#" & _
        "GROUP BY [DATE_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
        SET_NUMBER = 1
Else
        SET_NUMBER = FR_Table.Fields("[SQL COUNT]") + 1
End If

sSQL = "SELECT [SET_ID],[TYPE_ID],[SERIES_ID],[DEPT_ID],[DATE_ID],[SET NUMBER] " & _
        "FROM [SCHEDULE SETS] WHERE [DATE_ID]=#" & DATE_ID & "# " & _
        "ORDER BY [SET_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

FR_Table.AddNew

If (optT.Value = True) Then
    FR_Table.Fields("[TYPE_ID]") = "BARREL"
Else
    FR_Table.Fields("[TYPE_ID]") = "SBE"
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

If (Option800AB.Value = True) Then
    SERIES_ID = 810
End If
If (Option600SFL.Value = True) Then
    SERIES_ID = 600
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

Private Sub cmdDept_Click()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Dim sSQL As String

sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
   ' lblDesc.Caption = FR_Table.Fields("[DESCRIPTION]")
'    lblFinish.Caption = FR_Table.Fields("[FINISH_ID]")
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

sSQL = "SELECT * FROM [TBL PLATING OLEAN] WHERE [ID] = 1"
       
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
  
  '*NUMBER_HEADS
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
Get_DV
PrintGrouping

End Sub

Private Sub cmdPrintTXT_Click()

Set_Calculation
             
Dim sFilename As String
sFilename = "C:\ATC\test.txt"

Dim iFilenum As Integer
iFilenum = FreeFile

Open sFilename For Output Shared As #iFilenum
    
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim STRIKE1 As String
Dim STRIKE2 As String

Dim BASE_ID As String
Dim FINISH_ID As String
Dim OVERPLATE As String

If (FR_Table.RecordCount <> 0) Then
    STRIKE1 = FR_Table.Fields("[STRIKE1]")
    STRIKE2 = FR_Table.Fields("[STRIKE2]")
    BASE_ID = FR_Table.Fields("[BASE_ID]")
    FINISH_ID = FR_Table.Fields("[FINISH_ID]")
    OVERPLATE = FR_Table.Fields("[OVERPLATE]")
End If
            
Screen.MousePointer = vbHourglass
                
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT *  FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
                                                                                                                                              
    Print #iFilenum, "DATE : "; FR_Table.Fields("[DATE_ID]")

    Print #iFilenum, "Set Number : "; FR_Table.Fields("[SET NUMBER]")

    Print #iFilenum, "Case : "; Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)     'CASE
    
    Print #iFilenum, "Dept : "; FR_Table.Fields("[DEPT_ID]")                   'DEPT CODE
    
End If

'==============================================================================
'       WORK ORDER PER GEAR
'==============================================================================
If (GEAR_1_QTY <> 0) Then

        sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " AND [P1 BASE]<>0 "
               
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        If (FR_Table.RecordCount <> 0) Then
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                Print #iFilenum, Str(Mid(FR_Table.Fields("[WORK ORDER]"), 1, 6)):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3) & "." & Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 11, 2)
            Else
                Print #iFilenum, FR_Table.Fields("[WORK ORDER]")
            End If
        Else
        
        End If
Else
        Print #iFilenum, ""
End If

If (GEAR_2_QTY <> 0) Then

        sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " AND [P2 BASE]<>0 "
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        
        If (FR_Table.RecordCount <> 0) Then
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                Print #iFilenum, Str(Mid(FR_Table.Fields("[WORK ORDER]"), 1, 6)):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3) & "." & Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 11, 2)
            Else
                Print #iFilenum, FR_Table.Fields("[WORK ORDER]")
            End If
        Else
        
        End If
Else
        Print #iFilenum, ""
End If

If (GEAR_3_QTY <> 0) Then

        sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " AND [P3 BASE]<>0 "
               
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        If (FR_Table.RecordCount <> 0) Then
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                Print #iFilenum, Str(Mid(FR_Table.Fields("[WORK ORDER]"), 1, 6)):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3) & "." & Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 11, 2)
            Else
                Print #iFilenum, FR_Table.Fields("[WORK ORDER]")
            End If
        Else
        
        End If
Else
        Print #iFilenum, ""
End If

If (GEAR_4_QTY <> 0) Then

        sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " AND [P4 BASE]<>0 "
               
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        If (FR_Table.RecordCount <> 0) Then
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                Print #iFilenum, Str(Mid(FR_Table.Fields("[WORK ORDER]"), 1, 6)):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3) & "." & Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1):
                Print #iFilenum, Mid(FR_Table.Fields("[WORK ORDER]"), 11, 2)
            Else
                Print #iFilenum, FR_Table.Fields("[WORK ORDER]")
            End If
        Else
        
        End If
Else
        Print #iFilenum, ""
End If
                 
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT *  FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    
   Select Case FR_Table.Fields("[TYPE_ID]")
   Case "BARREL"
   
                Print #iFilenum, "Run Qty :"; TOTAL_QTY
                Print #iFilenum, "Media "; SHOT_ID
                Print #iFilenum, ""; Surface; Area; ";SA"
                
                If (STRIKE1 = "Y") Then
                    Print #iFilenum, "Strike 1"
                    Print #iFilenum, "Table AMP ", SKTASF1
                    Print #iFilenum, "Table MIN ", SKTMIN1
                    Print #iFilenum, "Amps ", Format(FR_Table.Fields("[SK1 AMP]"), "0.0")
                    Print #iFilenum, "Amp Min ", Format(FR_Table.Fields("[SK1 MIN]"), "0.0")
                Else
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                End If
                                
                'BASE PROCESS
                If (OVERPLATE = "N") Then
                    Print #iFilenum, "BASE "; BASE_ID
                    Print #iFilenum, FR_Table.Fields("[EQ BASE]")
                    
                    Print #iFilenum, "Table AMP ", SKTASF1
                    Print #iFilenum, "Table MIN ", SKTMIN1
                    Print #iFilenum, "Amps ", Format(FR_Table.Fields("[BASE AMP]"), "0.0")
                    Print #iFilenum, "Amp Min ", Format(FR_Table.Fields("[BASE AMP MIN]"), "0.0")
                                                                            
                Else
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                End If
                
                If (STRIKE2 = "Y") Then
                    Print #iFilenum, "Strike 2"
                    Print #iFilenum, "Table AMP", ASF3
                    Print #iFilenum, "Table MIN", MIN3
                    Print #iFilenum, "Amps ", Format(FR_Table.Fields("[SK2 AMP]"), "0.0")
                    Print #iFilenum, "Amp Min ", Format(FR_Table.Fields("[SK2 MIN]"), "0.0")
                Else
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                    Print #iFilenum, ""
                End If
                
                'FINISH PROCESS
                Print #iFilenum, "FINISH "; FINISH_ID
                Print #iFilenum, FR_Table.Fields("[EQ FINISH]")
                Print #iFilenum, "Table AMP", SKTASF1
                Print #iFilenum, "Table MIN", SKTMIN1
                Print #iFilenum, "Amp", Format(FR_Table.Fields("[FINISH AMP]"), "0.0")
                Print #iFilenum, "Amp Min", Format(FR_Table.Fields("[FINISH AMP MIN]"), "0.0")
                                                
    End Select


End If
        

FR_Table.Close
FR_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Close iFilenum

MsgBox "Excel Update Complete", vbInformation, "ATC Plating"
                                         
End Sub

Private Sub cmdRefresh_Click()

DATE_ID = DTPicker3.Value

Dim sSQL As String
sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],[DATE_ID],[TYPE_ID],[SERIES_ID],[RUN QTY]," & _
              "[EQ BASE],[BASE AMP],[BASE AMP MIN]," & _
              "[EQ FINISH],[FINISH AMP],[FINISH AMP MIN]  " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]=#" & DATE_ID & "# " & _
        "ORDER BY [SET_ID] DESC"
                                   
Data1.RecordSource = sSQL
Data1.Refresh

Dim sSQLF As String
sSQLF = "    |Set_ID|^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series||Base|Amp      |Amp Min  |Finish|Amp    |Amp Min   "

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

sSQLF = "    ||^Work Order/Lot    |^ATC Part    |>QTY        |>1G/H 1        |>1NG/H 2        |>2G/H 3        |>2NG/H 4      |^Run  "
MSFlexGrid2.FormatString = sSQLF
       
sSQL = "SELECT count([WORK ORDER]),format(sum([P1 BASE])+sum([P2 BASE])+sum([P3 BASE])+sum([P4 BASE]),'###,####')," & _
              "format(sum([P1 BASE]),'###,####'),format(sum([P2 BASE]),'###,####')," & _
              "format(sum([P3 BASE]),'###,####'),format(sum([P4 BASE]),'###,####') " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "GROUP BY [SET_ID]"
       
Data5.RecordSource = sSQL
Data5.Refresh

sSQLF = "    |^COUNT                   |>QTY        |>1G/H 1        |>1NG/H 2        |>2G/H 3        |>2NG/H 4      "

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

sSQLF = "    ||^M#   |^Name   |<               "

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

'lblASF1.Caption = ASF1
'lblMIN1.Caption = MIN1
  
'lblASF2.Caption = ASF2
'lblMIN2.Caption = MIN2

'lblStrike1.Visible = False
'lblStrike2.Visible = False

'lblSKASF1.Visible = False
'lblSKMIN1.Visible = False

'lblSKASF2.Visible = False
'lblSKMIN2.Visible = False

'lblSQFT.Caption = Format(SA, "0.0")

txtShot.Text = SHOT_ID

lblBaseAmp.Text = Format(ASF1 * SA, "0.0")
lblBaseAmpMin.Text = Format(MIN1 * ASF1 * SA / 60, "0.0")

lblFinishAmp.Text = Format(ASF2 * SA, "0.0")
lblFinishAmpMin.Text = Format(MIN2 * ASF2 * SA / 60, "0.0")
 
End Sub


Private Sub cmdValidatePG_Click()

LETTER_ID = Trim(txtLETTER_ID.Text)
If Len(LETTER_ID & "X") <> 1 Then
    Screen.MousePointer = vbHourglass
    Get_DV      '======================= UPDATE [DV] FROM [ATC PART] TABLE [GROUPING]
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
                   "FROM [GROUPING] WHERE [GP_ID]=" & GP_ID & " AND " & _
                                     "MID([ATC PART],1,3) NOT IN ('100','710','180','10E','70E','71E')"
                 
            Set FR_Table = FR_Database.OpenRecordset(sSQL)
            If (FR_Table.RecordCount <> 0) Then
             
                   sBuff = "Work Order " & FR_Table.Fields("[WORK ORDER]") & vbNewLine
                   sBuff = sBuff & "Part Number " & FR_Table.Fields("[ATC PART]") & vbNewLine
                   sBuff = sBuff & "can not be used in SBE   " & vbNewLine
                   sBuff = sBuff & "SBE only 100 and 700 Series"
            End If
            FR_Table.Close
            FR_Database.Close
            
            MsgBox sBuff, vbCritical + vbInformation, "ATC Plating Schedule"
                        
End Select

End Sub

Private Sub Form_Load()

Caption = "Review Schedule/Grouping     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION
Data5.DatabaseName = DB_PLATING_TERMINATION
Data6.DatabaseName = DB_PLATING_TERMINATION
Data7.DatabaseName = DB_PLATING_TERMINATION

'MSFlexGrid1.Top = Me.Height - 800
'MSFlexGrid1.Left = Me.Height - 800
MSFlexGrid1.Width = 10800
'MSFlexGrid1.Height = Me.Height - 80

'MSFlexGrid2.Top = Me.Height - 800
'MSFlexGrid2.Left = Me.Height - 800
MSFlexGrid2.Width = 9400
'MSFlexGrid2.Height = 1600

'MSFlexGrid2.Top = Me.Height - 800
'MSFlexGrid2.Left = Me.Height - 800
MSFlexGrid3.Width = 9000
MSFlexGrid3.Height = 1000

DTPicker3.Value = Date
DATE_ID = DTPicker3.Value

LabelBOM.Caption = ""

cmdRefresh_Click
MSFlexGrid1_Click

cmdRefresh2_Click

cmdRefresh6_Click   'LIST MACHINE

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
SET_ID = Val(MSFlexGrid1.Text)
  
LabelBOM.Caption = ""
  
MSFlexGrid1.Col = 2
DEPT_ID = Val(MSFlexGrid1.Text)
  
lblCode.Caption = DEPT_ID
  
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

cmdDept_Click

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
  
  
MSFlexGrid2.Col = 3
ATC_PART_ID = MSFlexGrid2.Text
  
Select Case Mid(ATC_PART_ID, 4, 1)
Case "A", "B", "R", "F", "S", "L"
  
        ValidPart (ATC_PART_ID)
        
End Select

'GP_ID = 68054
  
fraWS.Caption = "GP_ID : " & GP_ID
  
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

Private Sub txtATCPart_GotFocus()
txtATCPart.SelStart = 0
txtATCPart.SelLength = Len(txtATCPart)
End Sub

Private Sub txtATCPart_LostFocus()
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
LabelBOM.Caption = ""

txtWorkOrder.Text = Replace(txtWorkOrder.Text, ".", "")

txtWorkOrder.Text = Mid(UCase(txtWorkOrder.Text), 1, 12)

Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)

If (Val(txtSQ.Text) = 0) Then

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
