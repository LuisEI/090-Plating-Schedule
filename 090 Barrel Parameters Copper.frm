VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBarrelParametersCU 
   Caption         =   "090 Barrel Plating Copper Parameter Tables"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Barrel Parameters Copper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "[3] ATC 115-121 CU Lw / Sn 100,600 Series"
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
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Dept 540,544,546"
      Top             =   360
      Width           =   4800
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] ATC 115-121 CU Lw 200/900 Series"
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
      Left            =   7920
      TabIndex        =   4
      ToolTipText     =   "Dept 540,544,546"
      Top             =   780
      Width           =   4800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] ATC 115-121 CU (Strike) Hg 100 Series"
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
      Left            =   7920
      TabIndex        =   5
      ToolTipText     =   "Dept 541"
      Top             =   1200
      Width           =   4800
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 FROM [121 CU 2 LW]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\TERMINATION And PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "121 CU 2 LW"
      Top             =   4920
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [121 CU 2 HG]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\TERMINATION And PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "121 CU 2 HG"
      Top             =   5400
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 FROM [121 CU 1]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\TERMINATION And PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "121 CU 1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Frame fraCopperLW 
      Caption         =   " Copper / Lw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5160
      TabIndex        =   64
      Top             =   7560
      Width           =   4815
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
         DataField       =   "TYPE_CU"
         DataSource      =   "Data7"
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
         Left            =   3480
         TabIndex        =   30
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text60 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ASF SK"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   24
         Top             =   1560
         Width           =   1080
      End
      Begin VB.TextBox Text57 
         BackColor       =   &H00C0E0FF&
         DataField       =   "MIN SK"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   1920
         Width           =   1080
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "UpdateRecord"
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3600
         Width           =   1800
      End
      Begin VB.TextBox Text42 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MEDIA SPEED"
         DataSource      =   "Data7"
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
         Left            =   3480
         TabIndex        =   32
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text41 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CASE SERIES"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BARREL"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text39 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MAX"
         DataSource      =   "Data7"
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
         Left            =   3480
         TabIndex        =   31
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text38 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MIN"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN LP"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   3600
         Width           =   1080
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF LW"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   27
         Top             =   2880
         Width           =   1080
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN LW"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   3240
         Width           =   1080
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF CU"
         DataSource      =   "Data7"
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
         Left            =   1440
         TabIndex        =   26
         Top             =   2280
         Width           =   1080
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN CU"
         DataSource      =   "Data7"
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
         Left            =   3480
         TabIndex        =   33
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "FROM [121 CU 2 LW]"
         Height          =   300
         Left            =   2880
         TabIndex        =   92
         Top             =   3240
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF SK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   60
         Left            =   240
         TabIndex        =   88
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN SK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   57
         Left            =   240
         TabIndex        =   87
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Speed :"
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
         Index           =   42
         Left            =   2640
         TabIndex        =   79
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Series :"
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
         Index           =   41
         Left            =   240
         TabIndex        =   78
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Barrel :"
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
         Index           =   40
         Left            =   240
         TabIndex        =   77
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Max :"
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
         Index           =   39
         Left            =   2640
         TabIndex        =   76
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Min :"
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
         Index           =   38
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN PN :"
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
         Index           =   30
         Left            =   240
         TabIndex        =   69
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF LW:"
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
         Index           =   29
         Left            =   240
         TabIndex        =   68
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN WN/TN:"
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
         Index           =   28
         Left            =   240
         TabIndex        =   67
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF CU:"
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
         Index           =   27
         Left            =   240
         TabIndex        =   66
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN CU :"
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
         Left            =   2640
         TabIndex        =   65
         Top             =   2280
         Width           =   1065
      End
   End
   Begin VB.Frame fraCopperHG 
      Caption         =   " Copper / Hg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   10200
      TabIndex        =   57
      Top             =   7560
      Width           =   4815
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         DataField       =   "TYPE_CU"
         DataSource      =   "Data6"
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
         Left            =   3480
         TabIndex        =   45
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text62 
         BackColor       =   &H00C0E0FF&
         DataField       =   "MIN SK"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   39
         Top             =   1920
         Width           =   1080
      End
      Begin VB.TextBox Text61 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ASF SK"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   38
         Top             =   1560
         Width           =   1080
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "UpdateRecord"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3840
         Width           =   1800
      End
      Begin VB.TextBox Text47 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MIN"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text46 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MAX"
         DataSource      =   "Data6"
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
         Left            =   3480
         TabIndex        =   46
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text45 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BARREL"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   37
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text44 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CASE SERIES"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   35
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text43 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MEDIA SPEED"
         DataSource      =   "Data6"
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
         Left            =   3480
         TabIndex        =   47
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN CU"
         DataSource      =   "Data6"
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
         Left            =   3480
         TabIndex        =   48
         Top             =   2280
         Width           =   1080
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF CU"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   40
         Top             =   2280
         Width           =   1080
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00C0FFFF&
         DataField       =   "MIN LW"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   42
         Top             =   3120
         Width           =   1080
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00C0FFFF&
         DataField       =   "ASF LW"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   41
         Top             =   2760
         Width           =   1080
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF HG"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   43
         Top             =   3480
         Width           =   1080
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN HG"
         DataSource      =   "Data6"
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
         Left            =   1440
         TabIndex        =   44
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "FROM [121 CU 2 HG]"
         Height          =   300
         Left            =   2760
         TabIndex        =   93
         Top             =   3480
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN SK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   62
         Left            =   240
         TabIndex        =   90
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF SK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   61
         Left            =   240
         TabIndex        =   89
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Min :"
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
         Index           =   47
         Left            =   240
         TabIndex        =   84
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Max :"
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
         Index           =   46
         Left            =   2640
         TabIndex        =   83
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Barrel :"
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
         Index           =   45
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Series :"
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
         Index           =   44
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Speed :"
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
         Index           =   43
         Left            =   2640
         TabIndex        =   80
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN CU :"
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
         Index           =   36
         Left            =   2640
         TabIndex        =   63
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF CU:"
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
         Index           =   35
         Left            =   240
         TabIndex        =   62
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN SK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   240
         TabIndex        =   61
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF SK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   240
         TabIndex        =   60
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF HG:"
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
         Index           =   16
         Left            =   240
         TabIndex        =   59
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN HG :"
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
         Left            =   240
         TabIndex        =   58
         Top             =   3840
         Width           =   1065
      End
   End
   Begin VB.OptionButton Option7 
      Caption         =   "[7] ATC 115-170 CU (Strike) Hg 100/800 Series"
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
      TabIndex        =   2
      ToolTipText     =   "Dept 541"
      Top             =   1200
      Width           =   4800
   End
   Begin VB.OptionButton Option6 
      Caption         =   "[6] ATC 115-170 CU Lw 200/900 Series"
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
      TabIndex        =   1
      ToolTipText     =   "Dept 540,544,546"
      Top             =   780
      Width           =   4800
   End
   Begin VB.Frame fraCopperTin 
      Caption         =   " Copper / Tin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   51
      Top             =   7560
      Width           =   4815
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         DataField       =   "TYPE_CU"
         DataSource      =   "Data5"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text56 
         BackColor       =   &H00C0E0FF&
         DataField       =   "MIN SK"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   1080
      End
      Begin VB.TextBox Text55 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ASF SK"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1080
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "UpdateRecord"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3600
         Width           =   1800
      End
      Begin VB.TextBox Text37 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MIN"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text34 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MAX"
         DataSource      =   "Data5"
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
         Left            =   3480
         TabIndex        =   17
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox Text33 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BARREL"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CASE SERIES"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text31 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MEDIA SPEED"
         DataSource      =   "Data5"
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
         Left            =   3480
         TabIndex        =   18
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text26 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN TIN"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   3240
         Width           =   1080
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF TIN"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   2880
         Width           =   1080
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN TIN PN"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   3600
         Width           =   1080
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF CU"
         DataSource      =   "Data5"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   2280
         Width           =   1080
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN CU"
         DataSource      =   "Data5"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "FROM [121 CU 1]"
         Height          =   300
         Left            =   2760
         TabIndex        =   91
         Top             =   3240
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN SK :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   56
         Left            =   240
         TabIndex        =   86
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF SK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   55
         Left            =   240
         TabIndex        =   85
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Min :"
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
         Index           =   37
         Left            =   240
         TabIndex        =   74
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Qty Max :"
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
         Index           =   34
         Left            =   2640
         TabIndex        =   73
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Barrel :"
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
         Index           =   33
         Left            =   240
         TabIndex        =   72
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Series :"
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
         Index           =   32
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "Speed :"
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
         Index           =   31
         Left            =   2640
         TabIndex        =   70
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN WN/TN:"
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
         Index           =   26
         Left            =   240
         TabIndex        =   56
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF TIN:"
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
         Index           =   25
         Left            =   240
         TabIndex        =   55
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN PN :"
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
         Index           =   24
         Left            =   240
         TabIndex        =   54
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF CU:"
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
         Left            =   240
         TabIndex        =   53
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN CU :"
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
         Left            =   2640
         TabIndex        =   52
         Top             =   2280
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   1320
      TabIndex        =   50
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5] ATC 115-170 CU Lw / Sn 100/800 Series"
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
      TabIndex        =   0
      ToolTipText     =   "Dept 540,544,546"
      Top             =   360
      Value           =   -1  'True
      Width           =   4800
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 MSFlexGrid1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Barrel Parameters Copper.frx":0CCA
      Height          =   1215
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
      _Version        =   393216
      AllowBigSelection=   0   'False
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
   Begin VB.Frame Frame2 
      Caption         =   " PYRO "
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
      Left            =   7680
      TabIndex        =   95
      Top             =   120
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   " SA "
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
      TabIndex        =   94
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label LabelDBM 
      AutoSize        =   -1  'True
      Caption         =   "DBM"
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
      TabIndex        =   104
      Top             =   12000
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "[TYPE_CU]='MSA'"
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
      Left            =   5400
      TabIndex        =   103
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "[TYPE_CU]='PYRO'"
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
      Left            =   13200
      TabIndex        =   102
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 2 HG]"
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
      Left            =   13200
      TabIndex        =   101
      Top             =   1320
      Width           =   1890
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 2 LW]"
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
      Left            =   13200
      TabIndex        =   100
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 1] "
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
      Left            =   13200
      TabIndex        =   99
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 2 HG]"
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
      Left            =   5400
      TabIndex        =   98
      Top             =   1320
      Width           =   1890
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 2 LW]"
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
      Left            =   5400
      TabIndex        =   97
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FROM [121 CU 1]"
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
      Left            =   5400
      TabIndex        =   96
      Top             =   600
      Width           =   1545
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   4650
      Left            =   13680
      Picture         =   "090 Barrel Parameters Copper.frx":0CDE
      Top             =   1920
      Width           =   1965
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4395
      Left            =   11640
      Picture         =   "090 Barrel Parameters Copper.frx":1D820
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4155
      Left            =   8640
      Picture         =   "090 Barrel Parameters Copper.frx":3734E
      Top             =   1920
      Visible         =   0   'False
      Width           =   2280
   End
End
Attribute VB_Name = "frmBarrelParametersCU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String
 
If (Option5.Value = True) Then

    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                  "[BARREL],[MEDIA SPEED]," & _
                  "[ASF CU],[MIN CU],[ASF TIN],[MIN TIN],[MIN TIN PN] " & _
           "FROM [121 CU 1] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','700A','700B','100C','100E','200A','200B','900C','800C','800E') AND " & _
                 "[TYPE_CU]='MSA' ORDER BY [CASE SIZE],[CASE SERIES],[DV MAX]"
                                       
     sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel   |^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^TIN ASF|^TIN Min|^TIN Min PN"

End If

If (Option6.Value = True) Then

    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF CU],[MIN CU],[ASF LW],[MIN LW],[MIN LP]" & _
           "FROM [121 CU 2 LW] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','100C','100E','700A','700B','200A','200B','900C') AND " & _
                 "[TYPE_CU]='MSA' ORDER BY [CASE SIZE],[DV MAX]"

                                           
    sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel|^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^LW ASF|^WN&TN Min|^PN Min"
    
End If

If (Option7.Value = True) Then
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF CU],[MIN CU],[ASF LW],[MIN LW],[ASF HG],[MIN HG] " & _
           "FROM [121 CU 2 HG] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','700A','700B','100C','100E','200A','200B','900C','800C','800E') AND" & _
                 "[TYPE_CU]='MSA' ORDER BY [CASE SIZE],[CASE SERIES],[DV MAX]"
                                           
    sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel|^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^LW ASF|^LW Min|^HG ASF|^HG Min"

End If

If (Option3.Value = True) Then

    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                  "[BARREL],[MEDIA SPEED]," & _
                  "[ASF CU],[MIN CU],[ASF TIN],[MIN TIN],[MIN TIN PN] " & _
           "FROM [121 CU 1] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','700A','700B','100C','100E','200A','200B','900C','600F','600S','600L','800A','800B') AND " & _
                 "[TYPE_CU]='PYRO' ORDER BY [CASE SIZE],[DV MAX]"
                                       
     sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel   |^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^TIN ASF|^TIN Min|^TIN Min PN"

End If

If (Option2.Value = True) Then

    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF CU],[MIN CU],[ASF LW],[MIN LW],[MIN LP]" & _
           "FROM [121 CU 2 LW] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','100C','100E','200A','200B','900C','700A','700B') AND " & _
                 "[TYPE_CU]='PYRO' ORDER BY [CASE SIZE],[DV MAX]"

                                           
    sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel|^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^LW ASF|^WN&TN Min|^PN Min"
    
End If


If (Option1.Value = True) Then
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF CU],[MIN CU],[ASF LW],[MIN LW],[ASF HG],[MIN HG] " & _
           "FROM [121 CU 2 HG] " & _
           "WHERE [CASE SERIES] IN ('100A','100B','700A','700B','100C','100E','200A','200B','900C') AND " & _
                 "[TYPE_CU]='PYRO' ORDER BY [CASE SIZE],[DV MAX]"
                                           
    sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel|^Speed"
    sSQLF = sSQLF & "|^CU ASF|^CU Min|^LW ASF|^LW Min|^HG ASF|^HG Min"

End If

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub Command1_Click()
Data5.UpdateRecord
cmdRefresh_Click
End Sub

Private Sub Command2_Click()
Data7.UpdateRecord
cmdRefresh_Click
End Sub

Private Sub Command3_Click()
Data6.UpdateRecord
cmdRefresh_Click
End Sub

Private Sub Form_Load()

Caption = "Barrel Plating Copper Parameter Tables     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TABLES
Data5.DatabaseName = DB_PLATING_TABLES
Data6.DatabaseName = DB_PLATING_TABLES
Data7.DatabaseName = DB_PLATING_TABLES

LabelDBM.Caption = DB_PLATING_TABLES

cmdRefresh_Click
MSFlexGrid1_Click

MSFlexGrid1.Left = 0
MSFlexGrid1.Width = 10700
MSFlexGrid1.Height = 5400

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
TABLE_ID = Val(MSFlexGrid1.Text)

Dim sSQL As String
 
If (Option5.Value = True Or Option3.Value = True) Then
    fraCopperTin.Enabled = True
    sSQL = "SELECT * FROM [121 CU 1] WHERE [ID]=" & TABLE_ID
Else
    fraCopperTin.Enabled = False
    sSQL = "SELECT * FROM [121 CU 1] WHERE [ID]=" & -1
End If
Data5.RecordSource = sSQL
Data5.Refresh
   
If (Option6.Value = True Or Option2.Value = True) Then
    fraCopperLW.Enabled = True
    sSQL = "SELECT * FROM [121 CU 2 LW] WHERE [ID]=" & TABLE_ID
Else
    fraCopperLW.Enabled = False
    sSQL = "SELECT * FROM [121 CU 2 LW] WHERE [ID]=" & -1
End If
Data7.RecordSource = sSQL
Data7.Refresh

If (Option7.Value = True Or Option1.Value = True) Then
    fraCopperHG.Enabled = True
    sSQL = "SELECT * FROM [121 CU 2 HG] WHERE [ID]=" & TABLE_ID
Else
    fraCopperHG.Enabled = False
    sSQL = "SELECT * FROM [121 CU 2 HG] WHERE [ID]=" & -1
End If

Data6.RecordSource = sSQL
Data6.Refresh

MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1

If (ENABLE1_ATC_TABLES = 1) Then
    fraCopperTin.Enabled = True
    fraCopperLW.Enabled = True
    fraCopperHG.Enabled = True
Else
    fraCopperTin.Enabled = False
    fraCopperLW.Enabled = False
    fraCopperHG.Enabled = False
End If

End Sub

Private Sub Option5_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option7_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option6_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option1_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option2_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option4_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option3_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option8_Click()
cmdRefresh_Click
MSFlexGrid1_Click
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
