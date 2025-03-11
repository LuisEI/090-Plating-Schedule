VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkSheetR 
   Caption         =   "090 OEE Plating Worksheet Review"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Work Sheet Review.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   7080
      TabIndex        =   73
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [WORK SHEET PT]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Frame fraWS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   53
      Top             =   6840
      Width           =   3975
      Begin VB.TextBox txtTotalTime 
         BackColor       =   &H00C0FFC0&
         DataField       =   "TOTAL TIME"
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
         Left            =   1560
         TabIndex        =   57
         Text            =   "0"
         ToolTipText     =   "Total Time in Minutes"
         Top             =   3120
         Width           =   585
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to Start"
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
         TabIndex        =   56
         Top             =   2280
         Width           =   1440
      End
      Begin VB.CommandButton cmdStopTime 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Stop Time"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3120
         Width           =   1440
      End
      Begin VB.CommandButton cmdSub 
         Caption         =   "Sub from Stop"
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
         Left            =   2400
         TabIndex        =   54
         Top             =   2280
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   360
         TabIndex        =   58
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
         Format          =   48889858
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "STOP TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   2400
         TabIndex        =   59
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
         Format          =   48889858
         CurrentDate     =   38117
      End
      Begin VB.Label lblDept_ID 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "777"
         DataField       =   "DEPT_ID"
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
         Left            =   2520
         TabIndex        =   72
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   61
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   960
         TabIndex        =   64
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   960
         TabIndex        =   66
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblInfo 
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
         Left            =   120
         TabIndex        =   71
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblInfo 
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
         Left            =   120
         TabIndex        =   70
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblInfo 
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
         Left            =   120
         TabIndex        =   69
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label lblInfo 
         Caption         =   "FROM [WORK SHEET PT]"
         Height          =   300
         Index           =   12
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label lblInfo 
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
         TabIndex        =   67
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "SET NUMBER"
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
         Index           =   2
         Left            =   960
         TabIndex        =   65
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   2520
         TabIndex        =   63
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label lblInfo 
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
         Left            =   1680
         TabIndex        =   62
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblInfo 
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
         Left            =   1800
         TabIndex        =   60
         Top             =   1800
         Width           =   675
      End
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5]Complete Finish (500,600)"
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
      TabIndex        =   52
      Top             =   3840
      Width           =   3105
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] Complete Base (300,400)"
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
      TabIndex        =   49
      Top             =   3480
      Width           =   3225
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9600
      TabIndex        =   24
      Top             =   6840
      Width           =   5535
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1560
         TabIndex        =   40
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   39
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3480
         TabIndex        =   38
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4440
         TabIndex        =   37
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1560
         TabIndex        =   36
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   35
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3480
         TabIndex        =   34
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4440
         TabIndex        =   33
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4440
         TabIndex        =   32
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3480
         TabIndex        =   31
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   30
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1560
         TabIndex        =   29
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4440
         TabIndex        =   28
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3480
         TabIndex        =   27
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2520
         TabIndex        =   26
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1560
         TabIndex        =   25
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label lblSet 
         Caption         =   "Base Speed:"
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
         TabIndex        =   48
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "1G/H1"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   47
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Caption         =   "Base Serial #"
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
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "1NG/H2"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   45
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "2G/H3"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   44
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         Caption         =   "2NG/H4"
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
         Left            =   4440
         TabIndex        =   43
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblSet 
         Caption         =   "Finish Serial #"
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
         Index           =   5
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label lblSet 
         Caption         =   "Finish Speed:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   1515
      End
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3] All Incomplete Time"
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
      TabIndex        =   23
      Top             =   2400
      Width           =   3105
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] Active Finish (500,600)"
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
      TabIndex        =   22
      Top             =   3120
      Width           =   3105
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] Active Base (300,400)"
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
      TabIndex        =   21
      Top             =   2760
      Value           =   -1  'True
      Width           =   3105
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh6"
      Height          =   300
      Left            =   7320
      TabIndex        =   17
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 FROM [MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Frame fraFinish 
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
      Height          =   1935
      Left            =   7200
      TabIndex        =   12
      Top             =   6840
      Width           =   2295
      Begin VB.TextBox txtFinish 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "EQ FINISH"
         DataSource      =   "Data3"
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
         Height          =   420
         Left            =   720
         TabIndex        =   14
         ToolTipText     =   "Units Produced"
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
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Units Produced"
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Bindings        =   "090 Work Sheet Review.frx":0CCA
         Height          =   975
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
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
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame fraBase 
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
      Height          =   1935
      Left            =   4800
      TabIndex        =   7
      Top             =   6840
      Width           =   2295
      Begin VB.TextBox txtBase 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "EQ BASE"
         DataSource      =   "Data3"
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
         Height          =   420
         Left            =   720
         TabIndex        =   9
         ToolTipText     =   "Units Produced"
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
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "Units Produced"
         Top             =   1560
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Bindings        =   "090 Work Sheet Review.frx":0CDE
         Height          =   975
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
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
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   555
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
      RecordSource    =   " "
      Top             =   10080
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.Frame fraDateSelect 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3255
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
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Value           =   -1  'True
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
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdRefresh1 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   240
         TabIndex        =   1
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
         Format          =   48889857
         CurrentDate     =   38117
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
         TabIndex        =   2
         Top             =   1200
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
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
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
         Format          =   48889857
         CurrentDate     =   38117
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Work Sheet Review.frx":0CF2
      Height          =   6615
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11668
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
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
   Begin VB.Label lblInfo 
      Caption         =   "FROM [SCHEDULE SETS]"
      Height          =   300
      Index           =   11
      Left            =   9600
      TabIndex        =   74
      Top             =   9600
      Width           =   2595
   End
   Begin VB.Label lblSet 
      Caption         =   "Finish Setup,Check [500,600]"
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
      Left            =   240
      TabIndex        =   51
      Top             =   5160
      Width           =   2955
   End
   Begin VB.Label lblSet 
      Caption         =   "Base Setup,Check [300,400] "
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
      Left            =   240
      TabIndex        =   50
      Top             =   4800
      Width           =   3075
   End
End
Attribute VB_Name = "frmWorkSheetR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

DTPicker2.Value = DateAdd("n", txtTotalTime.Text, DTPicker1.Value)

Data4.UpdateRecord

cmdRefresh1_Click

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
 
Private Sub cmdNext_Click()

If (optWeek.Value = True) Then
    DTPicker3.Value = DateAdd("WW", 1, DTPicker3.Value)
    DTPicker4.Value = DateAdd("WW", 1, DTPicker4.Value)
End If

If (optDay.Value = True) Then
    DTPicker3.Value = DateAdd("D", 1, DTPicker3.Value)
    DTPicker4.Value = DateAdd("D", 1, DTPicker4.Value)
End If

cmdRefresh_Click
MSFlexGrid1_Click
End Sub

Private Sub cmdPrevious_Click()

If (optWeek.Value = True) Then
    DTPicker3.Value = DateAdd("WW", -1, DTPicker3.Value)
    DTPicker4.Value = DateAdd("WW", -1, DTPicker4.Value)
End If

If (optDay.Value = True) Then
    DTPicker3.Value = DateAdd("D", -1, DTPicker3.Value)
    DTPicker4.Value = DateAdd("D", -1, DTPicker4.Value)
End If

cmdRefresh_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh_Click()

DATE_START_ID = Format(DTPicker3.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker4.Value, "mm/dd/yyyy")

Dim SQL As Integer
Dim sSQL As String
Dim sSQLF As String

If (Option1.Value = True) Then
    SQL = 0
End If
If (Option2.Value = True) Then
    SQL = 0
End If
If (Option3.Value = True) Then
    SQL = 0
End If
If (Option4.Value = True) Then
    SQL = 1
End If
If (Option5.Value = True) Then
    SQL = 1
End If

Select Case SQL
Case 0

    sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                  "[SCHEDULE SETS].[DEPT_ID] AS [SQL DEPT_ID]," & _
                  "[WORK SHEET PT].[SET_ID]," & _
                  "[SCHEDULE SETS].[DATE_ID]," & _
                  "[SCHEDULE SETS].[SET NUMBER] AS [SQL SET NUMBER]," & _
                  "[SCHEDULE SETS].[TYPE_ID] AS [SQL TYPE_ID]," & _
                  "[WORK SHEET PT].[OP_ID]," & _
                  "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST] AS [SQL OPERATOR]," & _
                  "[WORK SHEET PT].[DATE_ID]," & _
                  "[WORK SHEET PT].[CODE_ID]," & _
                  "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                  "[WORK SHEET PT].[TOTAL TIME] " & _
          "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
          "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
                "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# "

    If (Option3.Value = True) Then
        sSQL = sSQL & "AND [WORK SHEET PT].[CODE_ID] IN (300,400,500,600) AND  [WORK SHEET PT].[TOTAL TIME]= 0 "
    End If
    
    If (Option1.Value = True) Then
        sSQL = sSQL & "AND [WORK SHEET PT].[CODE_ID] IN (300,400)"
    End If
    If (Option2.Value = True) Then
        sSQL = sSQL & "AND [WORK SHEET PT].[CODE_ID] IN (500,600)"
    End If
    
    sSQL = sSQL & " ORDER BY [SCHEDULE SETS].[DEPT_ID] ASC," & _
                            "[SCHEDULE SETS].[SET NUMBER] ASC," & _
                            "[WORK SHEET PT].[CODE_ID] ASC"
                            
sSQLF = "   ||^Dept_ID|^Set ID|^Create Date|^Set No.|^Type         |^|<Operator                |^Actual  Date |^Code    |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  "
                            

Case 1

    sSQL = "SELECT first([WORK SHEET PT].[WS_ID])," & _
                  "first([SCHEDULE SETS].[DEPT_ID])," & _
                  "first([WORK SHEET PT].[SET_ID])," & _
                  "first([SCHEDULE SETS].[DATE_ID])," & _
                  "first([SCHEDULE SETS].[SET NUMBER])," & _
                  "first([SCHEDULE SETS].[TYPE_ID])," & _
                  "first([WORK SHEET PT].[DATE_ID])," & _
                  "min([WORK SHEET PT].[CODE_ID]) + max([WORK SHEET PT].[CODE_ID])," & _
                  "sum([WORK SHEET PT].[TOTAL TIME]) " & _
          "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
          "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
                "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# "

        If (Option4.Value = True) Then
            'COMPLETE BASE
            sSQL = sSQL & "AND [WORK SHEET PT].[CODE_ID] IN (300,400) "
            'sSQL = sSQL & "AND ([WORK SHEET PT].[CODE_ID] IN (400) AND [WORK SHEET PT].[TOTAL TIME]<> 0) "
            
            sSQL = sSQL & "GROUP BY [SCHEDULE SETS].[DEPT_ID]," & _
                              "[SCHEDULE SETS].[SET NUMBER] " & _
                          "HAVING   min([WORK SHEET PT].[CODE_ID]) + max([WORK SHEET PT].[CODE_ID]) = 700 "
        End If
        If (Option5.Value = True) Then
            'COMPLETE FINISH
            sSQL = sSQL & "AND [WORK SHEET PT].[CODE_ID] IN (500,600) "
            
            'sSQL = sSQL & "AND ([WORK SHEET PT].[CODE_ID] IN (500) AND [WORK SHEET PT].[TOTAL TIME]<> 0) "
            'sSQL = sSQL & "AND ([WORK SHEET PT].[CODE_ID] IN (600) AND [WORK SHEET PT].[TOTAL TIME]<> 0) "
            
            sSQL = sSQL & "GROUP BY [SCHEDULE SETS].[DEPT_ID]," & _
                              "[SCHEDULE SETS].[SET NUMBER] " & _
                          "HAVING   min([WORK SHEET PT].[CODE_ID]) + max([WORK SHEET PT].[CODE_ID]) = 1100 "
        End If

sSQLF = "   ||^Dept_ID|^Set ID|^Create Date|^Set No.|^Type         |^Actual  Date |^Code    |^Total Time"


End Select
 
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF
End Sub

Private Sub cmdRefresh1_Click()

cmdRefresh_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh2_Click()

Dim DATE_ID As String
'DATE_ID = DTPicker4.Value

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
'MSFlexGrid2.Width = 6000

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
sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID

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

If (optWeek.Value = True) Then
    DTPicker3.Value = Format(DateAdd("d", -DTPicker3.DayOfWeek + 2, DTPicker3.Value), "mm/dd/yyyy")
    DTPicker4.Value = Format(DateAdd("d", 6, DTPicker3.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
   DTPicker3.Value = DTPicker3.Value
   DTPicker4.Value = DTPicker3.Value
End If

cmdRefresh_Click
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

Dim sTime As String
If (DTPicker1.Value > DTPicker2.Value) Then
    sTime = DateDiff("n", DTPicker1.Value, DTPicker2.Value) + 1440
Else
    sTime = DateDiff("n", DTPicker1.Value, DTPicker2.Value)
End If

txtTotalTime.Text = sTime

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

Private Sub Form_Load()

Caption = "OEE Plating Worksheet Review     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION
Data6.DatabaseName = DB_PLATING_TERMINATION
Data7.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11500
MSFlexGrid1.ForeColorSel = vbBlack

WS_ID = -1

DTPicker3.Value = Date

If (optWeek.Value = True) Then
    DTPicker3.Value = Format(DateAdd("d", -DTPicker3.DayOfWeek + 2, DTPicker3.Value), "mm/dd/yyyy")
    DTPicker4.Value = Format(DateAdd("d", 6, DTPicker3.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
   DTPicker3.Value = DTPicker3.Value
   DTPicker4.Value = DTPicker3.Value
End If
  
cmdRefresh1_Click
MSFlexGrid1_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub MSFlexGrid1_Click()

fraWS.Enabled = True
If (Option4.Value = True) Then
    fraWS.Enabled = False
    Exit Sub
End If
If (Option5.Value = True) Then
    fraWS.Enabled = False
    Exit Sub
End If

MSFlexGrid1.Col = 1
WS_ID = Val(MSFlexGrid1.Text)
fraWS.Caption = "WS_ID : " & WS_ID

MSFlexGrid1.Col = 2
DEPT_ID = Val(MSFlexGrid1.Text)

MSFlexGrid1.Col = 3
SET_ID = Val(MSFlexGrid1.Text)

MSFlexGrid1.Col = 6
TYPE_ID = MSFlexGrid1.Text

MSFlexGrid1.Col = 7
OP_ID = Val(MSFlexGrid1.Text)

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

Private Sub optDay_Click()

DTPicker3.Value = Format(DateAdd("d", -DTPicker3.DayOfWeek + 2, DTPicker3.Value), "mm/dd/yyyy")
DTPicker4.Value = Format(DateAdd("d", 6, DTPicker3.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Day  <<"
cmdNext.Caption = "Day  >>"

End Sub

Private Sub Option1_Click()
cmdRefresh1_Click
End Sub

Private Sub Option2_Click()
cmdRefresh1_Click
End Sub

Private Sub Option3_Click()
cmdRefresh1_Click
End Sub

Private Sub Option4_Click()
cmdRefresh1_Click
End Sub

Private Sub Option5_Click()
cmdRefresh1_Click
End Sub

Private Sub optWeek_Click()

DTPicker3.Value = Format(DateAdd("d", -DTPicker3.DayOfWeek + 2, DTPicker3.Value), "mm/dd/yyyy")
DTPicker4.Value = Format(DateAdd("d", 6, DTPicker3.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Week  <<"
cmdNext.Caption = "Week  >>"

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
