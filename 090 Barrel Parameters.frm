VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBarrelParametersNickel 
   Caption         =   "090 Barrel Plating Parameter Tables Nickel"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Barrel Parameters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "[1] ATC 115-115 Nickel/Tin"
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
      TabIndex        =   0
      ToolTipText     =   "Dept 529"
      Top             =   120
      Value           =   -1  'True
      Width           =   3400
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Barrel Parameters.frx":0CCA
      Height          =   2775
      Left            =   10080
      TabIndex        =   33
      ToolTipText     =   "[PCS PER SIDE]"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4895
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
   Begin VB.Data Data8 
      Caption         =   "Data8 FROM [REPLATE BARREL]"
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
      RecordSource    =   "121 CU 2 LW"
      Top             =   5760
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.Frame fraReplate 
      Caption         =   " Replate Barrel "
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
      Left            =   6480
      TabIndex        =   54
      Top             =   6240
      Width           =   2775
      Begin VB.TextBox Text59 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   31
         Top             =   2640
         Width           =   1080
      End
      Begin VB.TextBox Text58 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   30
         Top             =   2280
         Width           =   1080
      End
      Begin VB.TextBox Text54 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MEDIA SPEED"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   1860
         Width           =   1080
      End
      Begin VB.TextBox Text53 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CASE SERIES"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text52 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BARREL"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   1485
         Width           =   1080
      End
      Begin VB.TextBox Text51 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MAX"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   27
         Top             =   1110
         Width           =   1080
      End
      Begin VB.TextBox Text50 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MIN"
         DataSource      =   "Data8"
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
         Left            =   1440
         TabIndex        =   26
         Top             =   735
         Width           =   1080
      End
      Begin VB.CommandButton Command4 
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "FROM [REPLATE BARREL]"
         Height          =   300
         Left            =   240
         TabIndex        =   63
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN  :"
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
         Index           =   59
         Left            =   240
         TabIndex        =   61
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF :"
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
         Index           =   58
         Left            =   240
         TabIndex        =   60
         Top             =   2280
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
         Height          =   375
         Index           =   54
         Left            =   240
         TabIndex        =   59
         Top             =   1860
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
         Height          =   375
         Index           =   53
         Left            =   240
         TabIndex        =   58
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
         Height          =   375
         Index           =   52
         Left            =   240
         TabIndex        =   57
         Top             =   1485
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
         Height          =   375
         Index           =   51
         Left            =   240
         TabIndex        =   56
         Top             =   1110
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
         Height          =   375
         Index           =   50
         Left            =   240
         TabIndex        =   55
         Top             =   735
         Width           =   945
      End
   End
   Begin VB.OptionButton Option8 
      Caption         =   "[8] ATC Replate Barrel"
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
      TabIndex        =   4
      ToolTipText     =   "Dept 541"
      Top             =   1380
      Width           =   3525
   End
   Begin VB.Frame fraNickel 
      Caption         =   " Nickel "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   240
      TabIndex        =   35
      Top             =   6240
      Width           =   6015
      Begin VB.TextBox Text49 
         BackColor       =   &H00C0FFFF&
         DataField       =   "ASF AU SK"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   4080
         Width           =   1080
      End
      Begin VB.TextBox Text48 
         BackColor       =   &H00C0FFFF&
         DataField       =   "MIN AU SK"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   4455
         Width           =   1080
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MIN"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   855
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DV MAX"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   1230
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BARREL"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1605
         Width           =   1080
      End
      Begin VB.CommandButton cmdRefresh3 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5280
         Width           =   1800
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CASE SERIES"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   1080
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MEDIA SPEED"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   1980
         Width           =   1080
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN HG"
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
         Left            =   4560
         TabIndex        =   14
         Top             =   1560
         Width           =   1080
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF HG"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF TIN"
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
         Left            =   4560
         TabIndex        =   18
         Top             =   3240
         Width           =   1080
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN TIN"
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
         Left            =   4560
         TabIndex        =   19
         Top             =   3615
         Width           =   1080
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN LW P"
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
         Left            =   4560
         TabIndex        =   17
         Top             =   2760
         Width           =   1080
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN LW W"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   2400
         Width           =   1080
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF LW"
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
         Left            =   4560
         TabIndex        =   15
         Top             =   2040
         Width           =   1080
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN AU"
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
         Left            =   4560
         TabIndex        =   23
         Top             =   5205
         Width           =   1080
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF AU"
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
         Left            =   4560
         TabIndex        =   22
         Top             =   4830
         Width           =   1080
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ASF NI"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MIN NI"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   735
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "FROM [092 NICKEL HG]"
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
         Left            =   240
         TabIndex        =   68
         Top             =   4920
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "FROM [092 NICKEL LW] "
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
         Left            =   240
         TabIndex        =   67
         Top             =   4440
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "FROM [088 NICKEL GOLD]"
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
         Left            =   240
         TabIndex        =   66
         Top             =   3960
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "FROM [115 NICKEL TIN]"
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
         Left            =   240
         TabIndex        =   62
         Top             =   3480
         Width           =   2265
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
         Height          =   375
         Index           =   49
         Left            =   3120
         TabIndex        =   53
         Top             =   4080
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
         Height          =   375
         Index           =   48
         Left            =   3120
         TabIndex        =   52
         Top             =   4455
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
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   51
         Top             =   855
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   50
         Top             =   1230
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
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   49
         Top             =   1605
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
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   48
         Top             =   480
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
         Height          =   375
         Index           =   19
         Left            =   360
         TabIndex        =   47
         Top             =   1980
         Width           =   945
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
         Height          =   375
         Index           =   7
         Left            =   3120
         TabIndex        =   46
         Top             =   1575
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
         Height          =   375
         Index           =   8
         Left            =   3120
         TabIndex        =   45
         Top             =   1200
         Width           =   1065
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
         Height          =   375
         Index           =   14
         Left            =   3120
         TabIndex        =   44
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN TIN :"
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
         Left            =   3120
         TabIndex        =   43
         Top             =   3615
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN LW P:"
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
         Left            =   3120
         TabIndex        =   42
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN LW W :"
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
         Left            =   3120
         TabIndex        =   41
         Top             =   2400
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
         Height          =   375
         Index           =   10
         Left            =   3120
         TabIndex        =   40
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN AU :"
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
         Left            =   3120
         TabIndex        =   39
         Top             =   5205
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF AU:"
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
         Left            =   3120
         TabIndex        =   38
         Top             =   4830
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "ASF NI:"
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
         Left            =   3120
         TabIndex        =   37
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "MIN NI :"
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
         Left            =   3120
         TabIndex        =   36
         Top             =   735
         Width           =   945
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [092 NICKEL HG]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "088 NICKEL GOLD"
      Top             =   5760
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   1080
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] ATC 115-092 Nickel/HG"
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
      TabIndex        =   3
      ToolTipText     =   "Dept 524,537"
      Top             =   1380
      Width           =   3400
   End
   Begin VB.Data Data3 
      Caption         =   "Data3  FROM [SHOT]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [PCS PER SIDE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3] ATC 115-092 Nickel/LW"
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
      TabIndex        =   2
      ToolTipText     =   "Dept 530,525,285,532"
      Top             =   960
      Width           =   3400
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] ATC 115-088 Nickel (Strike) Gold "
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
      TabIndex        =   1
      ToolTipText     =   "Dept 535"
      Top             =   540
      Width           =   4125
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 MSFlexGrid1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Barrel Parameters.frx":0CDE
      Height          =   1215
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "090 Barrel Parameters.frx":0CF2
      Height          =   855
      Left            =   10440
      TabIndex        =   64
      ToolTipText     =   "[SHOT]"
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.Label lblInfo 
      Caption         =   "Gear Shot Quantities FROM [SHOT]"
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
      Index           =   18
      Left            =   10440
      TabIndex        =   65
      Top             =   240
      Width           =   4185
   End
End
Attribute VB_Name = "frmBarrelParametersNickel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRefresh_Click()
Dim sSQL As String
Dim sSQLF As String

If (Option1.Value = True) Then
                                   
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF NI],[MIN NI],[ASF TIN],[MIN TIN] " & _
           "FROM [115 NICKEL TIN] " & _
           "WHERE [CASE SERIES] IN ('700A','700B','100A','100B','100C','100E','200A','200B','900C') ORDER BY [CASE SIZE],[DV MAX]"
                                
    sSQLF = "       ||^Series    |>Min     |>Max       "
    sSQLF = sSQLF & "|^Barrel   |^Speed    "
    sSQLF = sSQLF & "|^NI ASF|^NI MIN |^TIN ASF|^TIN MIN"
End If

If (Option2.Value = True) Then

    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF NI],[MIN NI],[ASF AU SK],[MIN AU SK],[ASF AU],[MIN AU] " & _
           "FROM [088 NICKEL GOLD] " & _
           "WHERE [CASE SERIES] IN ('700A','700B','100A','100B','100C','100E','200A','200B','900C') ORDER BY [CASE SIZE],[DV MAX]"
    
     sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel   |^Speed    "
    sSQLF = sSQLF & "|^NI ASF|^NI Min|^SK ASF|^SK Min|^AU ASF|^AU Min"

End If

If (Option3.Value = True) Then
    
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                  "[BARREL],[MEDIA SPEED]," & _
                  "[ASF NI],[MIN NI],[ASF LW],[MIN LW W],[MIN LW P] " & _
           "FROM [092 NICKEL LW] " & _
           "WHERE [CASE SERIES] IN ('700A','700B','100A','100B','100C','100E','200A','200B','900C') ORDER BY [CASE SIZE],[DV MAX]"
     
     sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel   |^Speed"
    sSQLF = sSQLF & "|^NI ASF|^NI Min|^LW ASF|^LW Min W|^LW Min P"
    
End If
If (Option4.Value = True) Then
    
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF NI],[MIN NI],[ASF HG],[MIN HG] " & _
           "FROM [092 NICKEL HG] " & _
           "WHERE [CASE SERIES] IN ('700A','700B','100A','100B','100C','100E','200A','200B','900C') ORDER BY [CASE SIZE],[DV MAX]"
                                           
     sSQLF = "       ||^Series    |>Min     |>Max     "
    sSQLF = sSQLF & "|^Barrel |^Speed "
    sSQLF = sSQLF & "|^NI ASF|^NI MIN |^HG ASF|^HG MIN"
    
End If
 
If (Option8.Value = True) Then
                                   
    sSQL = "SELECT [ID],[CASE SERIES],[DV MIN],[DV MAX]," & _
                    "[BARREL],[MEDIA SPEED]," & _
                    "[ASF],[MIN] " & _
           "FROM [REPLATE BARREL] " & _
           "WHERE [CASE SERIES] IN ('700A','700B','100A','100B','100C','100E','200A','200B','900C') ORDER BY [CASE SIZE],[DV MAX]"
                                
    sSQLF = "       ||^Series    |>Min     |>Max          "
    sSQLF = sSQLF & "|^Barrel   |^Speed    "
    sSQLF = sSQLF & "|^ASF    |^MIN    "
End If
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefresh3_Click()
Data4.UpdateRecord
cmdRefresh_Click
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

Private Sub Command4_Click()
Data8.UpdateRecord
cmdRefresh_Click
End Sub

Private Sub Form_Load()

Caption = "Barrel Plating Parameter Tables [Nickel]    " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TABLES
Data2.DatabaseName = DB_PLATING_TABLES
Data3.DatabaseName = DB_PLATING_TABLES
Data4.DatabaseName = DB_PLATING_TABLES

Data8.DatabaseName = DB_PLATING_TABLES
 
 
If (ENABLE1_ATC_TABLES = 1) Then
    fraNickel.Enabled = True
    fraReplate.Enabled = True
Else
    fraNickel.Enabled = False
    fraReplate.Enabled = False
End If
 
cmdRefresh_Click

MSFlexGrid1_Click

MSFlexGrid1.Left = 0
'MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11000
MSFlexGrid1.Height = 4300

Dim sSQL As String
sSQL = "SELECT [CASE],format([PCS PER SIDE MAX],'###,###'),[SHOT],format([SF],'0.000') " & _
       "FROM [PCS PER SIDE] WHERE [TYPE]='BARREL' ORDER BY [CASE]"
                                   
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    |^Case   |>Pcs/Sd    |^Shot   |>Sq In.    "
MSFlexGrid2.FormatString = sSQLF
MSFlexGrid2.Width = 4300

sSQL = "SELECT [100 AB],[100 CE],[200 AB],[200 CE] " & _
       "FROM [SHOT] WHERE [SHOT_ID] = 1"
                                   
Data3.RecordSource = sSQL
Data3.Refresh
 
Dim sSQLF3 As String
sSQLF3 = "    |^100 AB |^100 CE |^200 AB |^200 CE "

MSFlexGrid3.FormatString = sSQLF3
MSFlexGrid3.Width = 4300

End Sub


Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
TABLE_ID = Val(MSFlexGrid1.Text)

Dim sSQL As String

If (Option8.Value = True) Then
    fraReplate.Enabled = True
    sSQL = "SELECT * FROM [REPLATE BARREL] WHERE [ID]=" & TABLE_ID
Else
    fraReplate.Enabled = False
    sSQL = "SELECT * FROM [REPLATE BARREL] WHERE [ID]=" & -1
End If
Data8.RecordSource = sSQL
Data8.Refresh

sSQL = "SELECT * FROM [088 NICKEL GOLD] WHERE [ID]=" & -1

If (Option1.Value = True) Then
    sSQL = "SELECT * FROM [115 NICKEL TIN] WHERE [ID]=" & TABLE_ID
End If

If (Option2.Value = True) Then
    sSQL = "SELECT * FROM [088 NICKEL GOLD] WHERE [ID]=" & TABLE_ID
End If

If (Option3.Value = True) Then
    sSQL = "SELECT * FROM [092 NICKEL LW] WHERE [ID]=" & TABLE_ID
End If

If (Option4.Value = True) Then
      sSQL = "SELECT *  FROM [092 NICKEL HG] WHERE [ID]=" & TABLE_ID
End If

Data4.RecordSource = sSQL
Data4.Refresh

MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1

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

Private Sub optNickelTin_Click()

End Sub

Private Sub Option3_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Option8_Click()
cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12)
End Sub

Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13)
End Sub

Private Sub Text19_GotFocus()
Text19.SelStart = 0
Text19.SelLength = Len(Text19)
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4)
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

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5)
End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6)
End Sub
