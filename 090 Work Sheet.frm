VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmWorkSheet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "090 OEE Plating Worksheet"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Work Sheet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh6"
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
      Left            =   7200
      TabIndex        =   67
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [MACHINE] "
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 FROM [MACHINE] "
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9000
      Visible         =   0   'False
      Width           =   3840
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
      Height          =   1935
      Left            =   6720
      TabIndex        =   62
      Top             =   6720
      Width           =   2175
      Begin VB.TextBox txtFinish 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "EQ FINISH"
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
         Height          =   420
         Left            =   720
         TabIndex        =   64
         ToolTipText     =   "EQ FINISH"
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox txtFinishID 
         BackColor       =   &H00C0FFFF&
         DataField       =   "MACHINE_F_ID"
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
         Height          =   420
         Left            =   1680
         TabIndex        =   63
         ToolTipText     =   "MACHINE_F_ID"
         Top             =   360
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Bindings        =   "090 Work Sheet.frx":0CCA
         Height          =   975
         Left            =   240
         TabIndex        =   65
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
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
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   555
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
      Height          =   1935
      Left            =   4440
      TabIndex        =   57
      Top             =   6720
      Width           =   2175
      Begin VB.TextBox txtBase 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "EQ BASE"
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
         Height          =   420
         Left            =   720
         TabIndex        =   59
         ToolTipText     =   "EQ BASE"
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox txtBaseID 
         BackColor       =   &H00C0FFFF&
         DataField       =   "MACHINE_B_ID"
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
         Height          =   420
         Left            =   1680
         TabIndex        =   58
         ToolTipText     =   "MACHINE_B_ID"
         Top             =   360
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Bindings        =   "090 Work Sheet.frx":0CDE
         Height          =   975
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
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
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame fraSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SET_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9120
      TabIndex        =   45
      Top             =   6720
      Width           =   5895
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 5"
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
         Left            =   1800
         TabIndex        =   79
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 6"
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
         Left            =   2760
         TabIndex        =   78
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 7"
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
         TabIndex        =   77
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 8"
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
         Left            =   4680
         TabIndex        =   76
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "FN HEAD 1"
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
         Left            =   1800
         TabIndex        =   75
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "FN HEAD 2"
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
         Left            =   2760
         TabIndex        =   74
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "FN HEAD 3"
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
         TabIndex        =   73
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "FN HEAD 4"
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
         Left            =   4680
         TabIndex        =   72
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "HEAD 4"
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
         Left            =   4680
         TabIndex        =   56
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "HEAD 3"
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
         TabIndex        =   55
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "HEAD 2"
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
         Left            =   2760
         TabIndex        =   54
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "HEAD 1"
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
         Left            =   1800
         TabIndex        =   52
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 4"
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
         Left            =   4680
         TabIndex        =   51
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 3"
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
         TabIndex        =   50
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED 2"
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
         Left            =   2760
         TabIndex        =   48
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         DataField       =   "SPEED"
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
         Left            =   1800
         TabIndex        =   47
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Finish Speed:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   81
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Finish Serial #"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   80
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "2NG/H4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   4680
         TabIndex        =   71
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "2G/H3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3720
         TabIndex        =   70
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "1NG/H2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2760
         TabIndex        =   69
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Serial #"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   53
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "1G/H1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1800
         TabIndex        =   49
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Speed:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   46
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [SCHEDULE SETS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1305
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1995
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1650
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2685
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   6600
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
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   3495
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   85
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   84
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   83
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmdRefresh1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   82
         Top             =   840
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   122159105
         CurrentDate     =   38117
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit to Main"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   9840
      Width           =   2160
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [WORK SHEET PT] "
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   4200
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Work Sheet.frx":0CF2
      Height          =   4455
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7858
      _Version        =   393216
      BackColorSel    =   16744703
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
   Begin VB.Frame fraWS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WS_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   4215
      Begin VB.CommandButton cmDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
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
         Left            =   2760
         TabIndex        =   68
         Top             =   840
         Width           =   1320
      End
      Begin VB.CommandButton cmdSub 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sub from Stop"
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
         TabIndex        =   2
         Top             =   2280
         Width           =   1440
      End
      Begin VB.CommandButton cmdStopTime 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Stop Time"
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
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   1440
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add to Start"
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
         Left            =   360
         TabIndex        =   1
         Top             =   2280
         Width           =   1440
      End
      Begin VB.TextBox txtTotalTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TOTAL TIME"
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
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Total Time in Minutes"
         Top             =   3120
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
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
         CustomFormat    =   "h:mm AM/PM"
         Format          =   122159106
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "STOP TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
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
         CustomFormat    =   "h:mm AM/PM"
         Format          =   122159106
         CurrentDate     =   38117
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run:"
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
         Index           =   10
         Left            =   2040
         TabIndex        =   44
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "DATE_ID"
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
         Height          =   360
         Index           =   6
         Left            =   2760
         TabIndex        =   42
         ToolTipText     =   "DATE_ID"
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create :"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   43
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "DATE_ID"
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
         Index           =   5
         Left            =   2760
         TabIndex        =   40
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XXX"
         DataField       =   "CODE_ID"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   39
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "WS_ID"
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
         Index           =   2
         Left            =   1200
         TabIndex        =   38
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "SET_ID"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   37
         ToolTipText     =   "SET_ID"
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Time (m):"
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
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FROM [WORK SHEET PT]"
         Height          =   300
         Index           =   12
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Code ID :"
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
         TabIndex        =   28
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set No:"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set ID :"
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
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Set Selection "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   9495
      Begin VB.CommandButton cmdRefresh2 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   36
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrevious2 
         Caption         =   "Day  <<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "Day  >>"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   33
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdReset2 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   122159105
         CurrentDate     =   38117
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Bindings        =   "090 Work Sheet.frx":0D06
         Height          =   1695
         Left            =   2520
         TabIndex        =   41
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2990
         _Version        =   393216
         BackColorSel    =   16776960
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         SelectionMode   =   1
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
      Begin VB.Label lblInfo 
         Caption         =   "FROM [SCHEDULE SETS]"
         Height          =   300
         Index           =   13
         Left            =   240
         TabIndex        =   87
         Top             =   1680
         Width           =   2235
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "FROM [SCHEDULE SETS]"
      Height          =   300
      Index           =   11
      Left            =   9240
      TabIndex        =   86
      Top             =   9600
      Width           =   2235
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   360
      TabIndex        =   21
      Top             =   1305
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   1650
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   1995
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   2340
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   5
      Left            =   360
      TabIndex        =   17
      Top             =   2685
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblEQ 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Work Station"
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
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   3345
   End
   Begin VB.Label txtOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPERATOR"
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
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label txtShift 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shift"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

DTPicker2.Value = DateAdd("n", txtTotalTime.Text, DTPicker1.Value)

Data4.UpdateRecord

cmdRefresh1_Click

End Sub

Private Sub cmdCode_Click(Index As Integer)

fraWS.Enabled = True

DATE_ID = DTPicker3.Value

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
              
sSQL = "SELECT * FROM [WORK SHEET PT] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [CODE_ID]=" & Val(lblCode(Index).Caption)
                           
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount = 0) Then
        FR_Table.AddNew
        WS_ID = FR_Table.Fields("[WS_ID]")
        FR_Table.Fields("[SET_ID]") = SET_ID
        FR_Table.Fields("[DATE_ID]") = DATE_ID
        FR_Table.Fields("[OP_ID]") = OP_ID
        FR_Table.Fields("[CODE_ID]") = Val(lblCode(Index).Caption)
        
        FR_Table.Fields("[START TIME]") = Format(Time, "hh:mm am/pm")
        
        Select Case Val(lblCode(Index).Caption)
        Case 400, 600   'check code
                FR_Table.Fields("[TOTAL TIME]") = 10
                FR_Table.Fields("[STOP TIME]") = Format(DateAdd("n", 10, Time), "hh:mm am/pm")
        Case 300, 500
                FR_Table.Fields("[TOTAL TIME]") = 0
        End Select
                
        FR_Table.Update
Else
        Dim sBuff As String
        sBuff = "[DATE_ID] " & FR_Table.Fields("[DATE_ID]") & vbNewLine
        sBuff = sBuff & "[SET_ID] " & FR_Table.Fields("[SET_ID]") & vbNewLine
        sBuff = sBuff & "[OP_ID] " & FR_Table.Fields("[OP_ID]") & vbNewLine
        sBuff = sBuff & "[CODE_ID] " & FR_Table.Fields("[CODE_ID]") & vbNewLine
        
        MsgBox "Already Started " & sBuff, vbInformation, "ATC Plating"

End If

FR_Table.Close
FR_Database.Close

cmdRefresh1_Click

End Sub

Private Sub cmDelete_Click()
Dim iAns As Integer

iAns = MsgBox("Delete Work Sheet Item", vbQuestion + vbYesNo, "ATC Plating")

If (iAns = vbYes) Then

    Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
    
    Dim sSQL As String
                
    sSQL = "SELECT * FROM [WORK SHEET PT] WHERE [WS_ID]=" & WS_ID
    
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
    If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            FR_Table.Delete
            FR_Table.MoveNext
        Loop
    End If
                   
    FR_Table.Close
    FR_Database.Close

   cmdRefresh1_Click

End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
 
Private Sub cmdNext_Click()
DTPicker3.Value = DateAdd("D", 1, DTPicker3.Value)
cmdRefresh1_Click
MSFlexGrid1_Click
End Sub

Private Sub cmdNext2_Click()

DTPicker4.Value = DateAdd("D", 1, DTPicker4.Value)
cmdRefresh2_Click

End Sub

Private Sub cmdPrevious_Click()

DTPicker3.Value = DateAdd("D", -1, DTPicker3.Value)

cmdRefresh1_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdPrevious2_Click()

DTPicker4.Value = DateAdd("D", -1, DTPicker4.Value)
cmdRefresh2_Click

cmdRefresh1_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh1_Click()

DATE_ID = DTPicker3.Value

Dim sSQL As String
Dim sSQLF As String

Select Case 0
Case 0
            sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                          "[WORK SHEET PT].[SET_ID]," & _
                          "[SCHEDULE SETS].[DATE_ID]," & _
                          "[SCHEDULE SETS].[SET NUMBER]," & _
                          "[SCHEDULE SETS].[TYPE_ID]," & _
                          "[WORK SHEET PT].[OP_ID]," & _
                          "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST]," & _
                          "[WORK SHEET PT].[DATE_ID]," & _
                          "[WORK SHEET PT].[CODE_ID]," & _
                          "format([START TIME],'h:mm AM/PM')," & _
                          "format([STOP TIME],'h:mm AM/PM')," & _
                          "[WORK SHEET PT].[TOTAL TIME] " & _
                  "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
                  "WHERE [WORK SHEET PT].[SET_ID]  = [SCHEDULE SETS].[SET_ID] AND " & _
                        "[WORK SHEET PT].[OP_ID]   = [BARCODE].[OP_ID] AND " & _
                        "[SCHEDULE SETS].[DEPT_ID] =" & DEPT_ID & " AND " & _
                        "[WORK SHEET PT].[DATE_ID] =#" & DATE_ID & "# " & _
                  "ORDER BY [WORK SHEET PT].[WS_ID] DESC"
Case 1
                sSQL = "SELECT [WS_ID]," & _
                              "[WORK SHEET PT].[SET_ID]," & _
                              "[SCHEDULE SETS].[DATE_ID]," & _
                              "[SCHEDULE SETS].[SET NUMBER]," & _
                              "[SCHEDULE SETS].[TYPE_ID]," & _
                              "[WORK SHEET PT].[DATE_ID]," & _
                              "[WORK SHEET PT].[CODE_ID]," & _
                              "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                              "[WORK SHEET PT].[TOTAL TIME] " & _
                      "FROM [WORK SHEET PT],[SCHEDULE SETS] " & _
                      "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                            "[OP_ID]=" & OP_ID & " AND " & _
                            "[SCHEDULE SETS].[DEPT_ID] =" & DEPT_ID & " AND " & _
                            "[WORK SHEET PT].[DATE_ID] =#" & DATE_ID & "# " & _
                      "ORDER BY [WORK SHEET PT].[CODE_ID] ASC"
End Select
 
sSQLF = "   ||^Set ID  |^Create Date |^Set No.  |^Type         |^|<Operator               |^Actual  Date |^Code    |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  "
 
 
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh2_Click()

Dim DATE_ID As String
DATE_ID = DTPicker4.Value

Dim sSQL As String
Dim sSQLF As String

Select Case 1
Case 0
        MSFlexGrid2.Width = 10800
        
        sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],format([DATE_ID],'mm/dd/yy'),[TYPE_ID],[SERIES_ID],[RUN QTY]," & _
                      "[EQ BASE],[BASE AMP],[BASE AMP MIN]," & _
                      "[EQ FINISH],[FINISH AMP],[FINISH AMP MIN]  " & _
                "FROM [SCHEDULE SETS] " & _
                "WHERE [DATE_ID]=#" & DATE_ID & "# AND [DEPT_ID]=" & DEPT_ID & " " & _
                "ORDER BY [SET_ID]"
        sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series||Base|Amp      |Amp Min  |Finish|Amp    |Amp Min   "
        
Case 1
        MSFlexGrid2.Width = 6000
        
        sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],format([DATE_ID],'mm/dd/yy'),[TYPE_ID],[SERIES_ID] " & _
                "FROM [SCHEDULE SETS] " & _
                "WHERE [DATE_ID]=#" & DATE_ID & "# AND [DEPT_ID]=" & DEPT_ID & " " & _
                "ORDER BY [SET_ID]"
                       
        sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series "
       
End Select

Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF

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

' [TYPE]    SBE/BARREL
' [PROCESS] BASE/FINISH
' [BF_ID]   Nickel,Copper Low/Hi

sSQL = "SELECT [MACHINE_ID],[NUMBER],[NAME] " & _
       "FROM [MACHINE] " & _
       "WHERE [TYPE]='" & TYPE_ID & "' AND " & _
             "[PROCESS]='" & PROCESS_ID & "' AND " & _
             "[LOCATION_ID]='" & LOCATION_ID & "' AND " & _
             "[BF_ID]='" & BASE_ID & "' AND [ACTIVE] = 1"
                                                                                                          
Data6.RecordSource = sSQL
Data6.Refresh

sSQLF = "    ||^M#   |^Name   "

MSFlexGrid6.FormatString = sSQLF

PROCESS_ID = "FINISH"

sSQL = "SELECT [MACHINE_ID],[NUMBER],[NAME] " & _
       "FROM [MACHINE] " & _
       "WHERE [TYPE]='" & TYPE_ID & "' AND " & _
             "[PROCESS]='" & PROCESS_ID & "' AND " & _
             "[LOCATION_ID]='" & LOCATION_ID & "' AND " & _
             "[BF_ID]='" & FINISH_ID & "'AND [ACTIVE] = 1"
                                                                                                          
Data7.RecordSource = sSQL
Data7.Refresh

sSQLF = "    ||^M#   |^Name   "

MSFlexGrid7.FormatString = sSQLF

End Sub

Private Sub cmdReset_Click()

WS_ID = -1
DTPicker3.Value = Date

cmdRefresh1_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdReset2_Click()

Dim DATE_ID As String
DTPicker4.Value = Date
DATE_ID = DTPicker4.Value

Dim sSQL As String
Dim sSQLF As String

Select Case 1
Case 0
MSFlexGrid2.Width = 10800

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],format([DATE_ID],'mm/dd/yy'),[TYPE_ID],[SERIES_ID],[RUN QTY]," & _
              "[EQ BASE],[BASE AMP],[BASE AMP MIN]," & _
              "[EQ FINISH],[FINISH AMP],[FINISH AMP MIN]  " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]=#" & DATE_ID & "# AND [DEPT_ID]=" & DEPT_ID & " " & _
        "ORDER BY [SET_ID]"
sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series||Base|Amp      |Amp Min  |Finish|Amp    |Amp Min   "
        
Case 1
MSFlexGrid2.Width = 6000

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],format([DATE_ID],'mm/dd/yy'),[TYPE_ID],[SERIES_ID] " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]=#" & DATE_ID & "# AND [DEPT_ID]=" & DEPT_ID & " " & _
        "ORDER BY [SET_ID]"
        
        
sSQLF = "    ||^DEPT_ID|^SET #|^DATE_ID      |^SBE/Barrel   |^Series "
       
End Select

Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF
End Sub

Private Sub cmdStopTime_Click()

DTPicker2.Value = Format(Time, "hh:mm am/pm")

Dim stime As String
If (DTPicker1.Value > DTPicker2.Value) Then
    stime = DateDiff("n", DTPicker1.Value, DTPicker2.Value) + 1440
Else
    stime = DateDiff("n", DTPicker1.Value, DTPicker2.Value)
End If

txtTotalTime.Text = stime

Data4.UpdateRecord

Dim sSQL As String

FLAG_WO_ID = FLAG_ID(DEPT_ID)
   
If (FLAG_WO_ID <> 0) Then
    
    Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)
        
    Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)
    sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
    Set TO_Table = TO_Database.OpenRecordset(sSQL)
    
    If (TO_Table.RecordCount <> 0) Then
        Do Until TO_Table.EOF
            WO_ID = TO_Table.Fields("[WORK ORDER]")
            If (Len(WO_ID) = 12) Then
                sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & WO_ID & "'"
                Set FR_Table = FR_Database.OpenRecordset(sSQL)
                If (FR_Table.RecordCount <> 0) Then
                    FR_Table.Edit
                    FR_Table.Fields("[FLAG_NY]") = FLAG_WO_ID
                    FR_Table.Update
                Else
                    If (Len(WO_ID) = 12) Then
                        If (IsNumeric(Mid(WO_ID, 1, 6)) = False Or IsNumeric(Mid(WO_ID, 7, 6)) = False) Then
                        Else
                            FR_Table.AddNew
                            FR_Table.Fields("[DATE_ID]") = DATE_ID
                            FR_Table.Fields("[WORK ORDER]") = TO_Table.Fields("[WORK ORDER]")
                            FR_Table.Fields("[ATC PART]") = TO_Table.Fields("[ATC PART]")
                            FR_Table.Fields("[LOT NUM]") = TO_Table.Fields("[LOT NUM]")
                            FR_Table.Fields("[START QTY]") = TO_Table.Fields("[QTY]")
                            FR_Table.Fields("[FLAG_NY]") = FLAG_WO_ID
                            FR_Table.Fields("[LT1]") = 1
                            FR_Table.Update
                        End If
                    End If
                End If
             End If
             TO_Table.MoveNext
         Loop
     End If
     FR_Database.Close
     TO_Database.Close

End If

cmdRefresh1_Click

End Sub

Private Sub cmdSub_Click()

DTPicker1.Value = DateAdd("n", -Val(txtTotalTime.Text), DTPicker2.Value)

Data4.UpdateRecord

cmdRefresh1_Click
End Sub

Private Sub Form_Activate()

DTPicker1.Value = Format(Time, "hh:mm am/pm")
DTPicker2.Value = Format(Time, "hh:mm am/pm")

cmdRefresh1_Click
MSFlexGrid1_Click

cmdRefresh2_Click

End Sub

Private Sub Form_Load()

Caption = "OEE Plating Worksheet     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION

Data6.DatabaseName = DB_PLATING_TERMINATION
Data7.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11280
'MSFlexGrid1.Height = 3000
MSFlexGrid1.ForeColorSel = vbBlack

  'SFlexGrid2.Height = 2000
MSFlexGrid2.ForeColorSel = vbBlack

Dim sSQL As String
Dim I As Integer

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT * FROM [TBL CODES] "
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Do Until FR_Table.EOF
    cmdCode(I).Visible = True
    lblCode(I).Visible = True
    lblCode(I).Caption = FR_Table.Fields("[CODE_ID]")
    cmdCode(I).Caption = FR_Table.Fields("[DESCRIPTION]")
    I = I + 1
    FR_Table.MoveNext
Loop

WS_ID = -1

'
'   DISPLAY OPERATOR AND MACHINE INFORMATION
'
sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    txtOperator.Caption = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
    
    Select Case FR_Table.Fields("[SHIFT_ID]")
    Case "D"
            txtShift.Caption = "Day"
    Case "E"
            txtShift.Caption = "Evening"
    End Select
End If

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    lblEQ.Caption = FR_Table.Fields("[DESCRIPTION]")
Else
    lblEQ.Caption = ""
End If
FR_Table.Close
FR_Database.Close

Dim SHIFT_TIME As String
SHIFT_TIME = "3 AM"

Dim START_TIME As Date
START_TIME = Format(Time, "h AM/PM")

If (START_TIME < SHIFT_TIME) Then
    'Change Date -1
    DATE_ID = DateAdd("d", -1, Date$)
Else
    DATE_ID = Date$
End If

DTPicker3.Value = DATE_ID
DTPicker4.Value = DATE_ID
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
WS_ID = Val(MSFlexGrid1.Text)

fraWS.Caption = "WS_ID : " & WS_ID

MSFlexGrid1.Col = 2
SET_ID = Val(MSFlexGrid1.Text)

fraSet.Caption = "SET_ID : " & SET_ID

MSFlexGrid1.Col = 5
TYPE_ID = MSFlexGrid1.Text

'MSFlexGrid1.Col = 6
'OP_ID = Val(MSFlexGrid1.Text)

MSFlexGrid1.Col = 9
CODE_ID = Val(MSFlexGrid1.Text)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET PT] WHERE [WS_ID]=" & WS_ID
Data4.RecordSource = sSQL
Data4.Refresh
        
sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
Data3.RecordSource = sSQL
Data3.Refresh
        
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

cmdRefresh6_Click   'LIST MACHINE

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
SET_ID = Val(MSFlexGrid2.Text)

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1

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

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
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
Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10)
End Sub
Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11)
End Sub
Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12)
End Sub
Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16)
End Sub
Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15)
End Sub
Private Sub Text14_GotFocus()
Text14.SelStart = 0
Text14.SelLength = Len(Text14)
End Sub
Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13)
End Sub
Private Sub txtTotalTime_GotFocus()
txtTotalTime.SelStart = 0
txtTotalTime.SelLength = Len(txtTotalTime)
End Sub
