VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSummary 
   Caption         =   "090  Summary Review Plating"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Summary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandSQL 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SQL"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   11520
      Width           =   1575
   End
   Begin VB.OptionButton Option16 
      Caption         =   "[16] Set Count per EQ Base"
      Height          =   300
      Left            =   360
      TabIndex        =   84
      Top             =   5160
      Width           =   2625
   End
   Begin VB.OptionButton Option15 
      Caption         =   "[15]  SBE/Barrel Total"
      Height          =   300
      Left            =   360
      TabIndex        =   82
      Top             =   5760
      Value           =   -1  'True
      Width           =   2625
   End
   Begin VB.CommandButton cmdLift 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lift Term"
      Height          =   300
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   4800
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton cmdPlatingLoad 
      BackColor       =   &H000080FF&
      Caption         =   "Elect Load Plating ID"
      Height          =   300
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6240
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H000080FF&
      Caption         =   "Test"
      Height          =   300
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5520
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdTests 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PG Test"
      Height          =   300
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   11040
      Width           =   1125
   End
   Begin VB.CommandButton cmdWOSearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "WO Search"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   11040
      Width           =   1605
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Show"
      Height          =   300
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   10320
      Width           =   1000
   End
   Begin VB.TextBox txtLETTER_ID 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   2760
      TabIndex        =   68
      Text            =   " "
      ToolTipText     =   "LETTER_ID"
      Top             =   10680
      Width           =   600
   End
   Begin VB.TextBox txtSET_ID 
      Height          =   300
      Left            =   1560
      TabIndex        =   67
      Text            =   " "
      ToolTipText     =   "SET_ID"
      Top             =   10680
      Width           =   1200
   End
   Begin VB.CommandButton cmdMix 
      BackColor       =   &H000080FF&
      Caption         =   "Mix Test"
      Height          =   300
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdCT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Table"
      Height          =   300
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   10320
      Width           =   1000
   End
   Begin VB.Frame fraTable 
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   3840
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 9"
         DataSource      =   "Data2"
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
         Index           =   9
         Left            =   3000
         TabIndex        =   53
         Text            =   " "
         Top             =   4620
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 8"
         DataSource      =   "Data2"
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
         Index           =   8
         Left            =   3000
         TabIndex        =   52
         Text            =   " "
         Top             =   4200
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 7"
         DataSource      =   "Data2"
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
         Index           =   7
         Left            =   3000
         TabIndex        =   51
         Text            =   " "
         Top             =   3780
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 6"
         DataSource      =   "Data2"
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
         Index           =   6
         Left            =   3000
         TabIndex        =   50
         Text            =   " "
         Top             =   3360
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 5"
         DataSource      =   "Data2"
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
         Index           =   5
         Left            =   3000
         TabIndex        =   49
         Text            =   " "
         Top             =   2940
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 4"
         DataSource      =   "Data2"
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
         Index           =   4
         Left            =   3000
         TabIndex        =   48
         Text            =   " "
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 3"
         DataSource      =   "Data2"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   47
         Text            =   " "
         Top             =   2100
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 2"
         DataSource      =   "Data2"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   46
         Text            =   " "
         Top             =   1680
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 1"
         DataSource      =   "Data2"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   45
         Text            =   " "
         Top             =   1260
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 0"
         DataSource      =   "Data2"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   44
         Text            =   " "
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 1"
         DataSource      =   "Data2"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   43
         Text            =   " "
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 2"
         DataSource      =   "Data2"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   42
         Text            =   " "
         Top             =   1920
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 3"
         DataSource      =   "Data2"
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
         Index           =   3
         Left            =   1560
         TabIndex        =   41
         Text            =   " "
         Top             =   2340
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 4"
         DataSource      =   "Data2"
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
         Index           =   4
         Left            =   1560
         TabIndex        =   40
         Text            =   " "
         Top             =   2760
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 5"
         DataSource      =   "Data2"
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
         Index           =   5
         Left            =   1560
         TabIndex        =   39
         Text            =   " "
         Top             =   3180
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 6"
         DataSource      =   "Data2"
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
         Index           =   6
         Left            =   1560
         TabIndex        =   38
         Text            =   " "
         Top             =   3600
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 7"
         DataSource      =   "Data2"
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
         Index           =   7
         Left            =   1560
         TabIndex        =   37
         Text            =   " "
         Top             =   4020
         Width           =   1200
      End
      Begin VB.TextBox txtTol 
         DataField       =   "Tol Code 8"
         DataSource      =   "Data2"
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
         Index           =   8
         Left            =   1560
         TabIndex        =   36
         Text            =   " "
         Top             =   4440
         Width           =   1200
      End
      Begin VB.TextBox txtTolLimit 
         DataField       =   "Bin Limit 0"
         DataSource      =   "Data2"
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
         Index           =   0
         Left            =   3000
         TabIndex        =   35
         Text            =   " "
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 1"
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
         Index           =   16
         Left            =   720
         TabIndex        =   64
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 2"
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
         Index           =   17
         Left            =   720
         TabIndex        =   63
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 3"
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
         Left            =   720
         TabIndex        =   62
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 4"
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
         Index           =   19
         Left            =   720
         TabIndex        =   61
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 5"
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
         Index           =   20
         Left            =   720
         TabIndex        =   60
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 6"
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
         Index           =   21
         Left            =   720
         TabIndex        =   59
         Top             =   3180
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 7"
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
         Index           =   22
         Left            =   720
         TabIndex        =   58
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 8"
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
         Index           =   23
         Left            =   720
         TabIndex        =   57
         Top             =   4020
         Width           =   795
      End
      Begin VB.Label lblTol 
         Caption         =   "Bin 9"
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
         Index           =   24
         Left            =   720
         TabIndex        =   56
         Top             =   4440
         Width           =   795
      End
      Begin VB.Label lblTol 
         Alignment       =   2  'Center
         Caption         =   "Tolerance Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   1560
         TabIndex        =   55
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblTol 
         Alignment       =   2  'Center
         Caption         =   "Bin Limit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   54
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DV"
      Height          =   300
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Convert DV"
      Top             =   9960
      Width           =   525
   End
   Begin VB.CommandButton cmdPrintGrouping 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Report"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10320
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Excel Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   3615
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lift Term JR"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1080
         Width           =   1600
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lift Term NY"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1080
         Width           =   1600
      End
      Begin VB.CommandButton cmdExcelNew 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Tank/Case Per Day"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
         Width           =   1600
      End
      Begin VB.TextBox txtDefect 
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   1080
         TabIndex        =   78
         Text            =   " "
         ToolTipText     =   "LETTER_ID"
         Top             =   360
         Width           =   480
      End
      Begin VB.OptionButton optCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   960
         TabIndex        =   77
         ToolTipText     =   "FROM [TBL CALCULATION],[DEPT CODE]"
         Top             =   720
         Width           =   915
      End
      Begin VB.OptionButton optDefect 
         Caption         =   "Defect"
         Height          =   300
         Left            =   120
         TabIndex        =   76
         ToolTipText     =   "FROM [TBL CALCULATION],[DEPT CODE]"
         Top             =   360
         Width           =   825
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   300
         Left            =   120
         TabIndex        =   75
         ToolTipText     =   "FROM [TBL CALCULATION],[DEPT CODE]"
         Top             =   720
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Work Orders"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   360
         Width           =   1600
      End
   End
   Begin VB.OptionButton Option14 
      Caption         =   "[14] Plate Grouping"
      Height          =   300
      Left            =   360
      TabIndex        =   28
      Top             =   9960
      Width           =   2265
   End
   Begin VB.OptionButton Option13 
      Caption         =   "[13] W.O. Errors"
      Height          =   300
      Left            =   360
      TabIndex        =   27
      Top             =   9000
      Width           =   2625
   End
   Begin VB.OptionButton Option12 
      Caption         =   "[12] Department by Series"
      Height          =   300
      Left            =   360
      TabIndex        =   26
      Top             =   6420
      Width           =   2625
   End
   Begin VB.OptionButton Option11 
      Caption         =   "[11] Code,Series,SBE"
      Height          =   300
      Left            =   360
      TabIndex        =   25
      Top             =   8700
      Width           =   2625
   End
   Begin VB.OptionButton Option10 
      Caption         =   "[10] Code,Series,Barrel"
      Height          =   300
      Left            =   360
      TabIndex        =   24
      Top             =   8400
      Width           =   2625
   End
   Begin VB.CommandButton cmdCalculationEQFinish 
      Caption         =   "[2] Calculation EQ Finish"
      Height          =   300
      Left            =   3960
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.OptionButton Option9 
      Caption         =   "[9] Base EQ"
      Height          =   300
      Left            =   360
      TabIndex        =   22
      Top             =   9300
      Width           =   2625
   End
   Begin VB.CommandButton cmdCalculationEQ 
      Caption         =   "[2] Calculation EQ Base"
      Height          =   300
      Left            =   3960
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.OptionButton Option8 
      Caption         =   "[8] WO Count"
      Height          =   300
      Left            =   360
      TabIndex        =   20
      Top             =   5460
      Width           =   2625
   End
   Begin VB.OptionButton Option7 
      Caption         =   "[7] OEE Codes"
      Height          =   300
      Left            =   360
      TabIndex        =   19
      Top             =   6120
      Width           =   2625
   End
   Begin VB.OptionButton Option6 
      Caption         =   "[6] Set Count"
      Height          =   300
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   2625
   End
   Begin VB.CommandButton cmdCalculation 
      Caption         =   "[1] Calculation for Dept"
      Height          =   300
      Left            =   3960
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5] Case Size"
      Height          =   300
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] "
      Top             =   4500
      Width           =   2625
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] Series"
      Height          =   300
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] "
      Top             =   4200
      Width           =   2625
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3] Finish EQ"
      Height          =   300
      Left            =   360
      TabIndex        =   12
      Top             =   9600
      Width           =   2625
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] EQ Base && Finish"
      Height          =   300
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "FROM [TBL CALCULATION EQ],[MACHINE]"
      Top             =   3900
      Width           =   2625
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
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton OptionYear 
         Caption         =   "Year"
         Height          =   300
         Left            =   1800
         TabIndex        =   86
         Top             =   1440
         Width           =   825
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
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1065
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Range (From - To)"
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
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   2505
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
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Week  >>"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Week  <<"
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
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
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
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   1065
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
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   945
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   600
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
         Format          =   16515073
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1800
         TabIndex        =   10
         Top             =   600
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
         Format          =   16515073
         CurrentDate     =   38117
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Date or Opt [1-13] Change click Refresh Command Button"
         Height          =   465
         Left            =   360
         TabIndex        =   83
         Top             =   120
         Width           =   2835
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] Department"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "FROM [TBL CALCULATION],[DEPT CODE]"
      Top             =   3600
      Width           =   2625
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   4980
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Summary.frx":0CCA
      Height          =   1095
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "[TBL CALCULATION]"
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393216
      ScrollBars      =   2
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
   Begin VB.Label lblTol 
      Caption         =   "Plating_ID :"
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   73
      Top             =   10680
      Width           =   1035
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCalculation_Click()

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [TBL CALCULATION]"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            FR_Table.Edit
            FR_Table.Fields("[SET COUNT]") = 0
            FR_Table.Fields("[WO COUNT]") = 0
            FR_Table.Fields("[SUM QTY]") = 0
            FR_Table.Fields("[PROCESS TIME]") = 0
            FR_Table.Fields("[OEE TIME]") = 0
            FR_Table.Fields("[PERFORMANCE]") = 0
            FR_Table.Update
            FR_Table.MoveNext
        Loop
End If


Select Case LOCATION_ID
Case "NY"

            sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID]) AS [SQL DEPT_ID]," & _
                          "first([DEPT CODE].[DESCRIPTION])," & _
                          "count([SCHEDULE SETS].[SET_ID])  AS [SQL SET COUNT]," & _
                          "format(sum([RUN QTY]),'###,##0') AS [SQL SUM QTY]" & _
                    "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                    "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                          "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                          "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                          "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                    "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
Case "JR"

        sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID]) AS [SQL DEPT_ID]," & _
                          "first([DEPT CODE].[DESCRIPTION])," & _
                      "count([SCHEDULE SETS].[SET_ID])  AS [SQL SET COUNT]," & _
                      "format(sum([RUN QTY]),'###,##0') AS [SQL SUM QTY]" & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
End Select




Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            
            sSQL = "SELECT * FROM [TBL CALCULATION] WHERE [DEPT_ID]=" & FR_Table.Fields("[SQL DEPT_ID]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[SET COUNT]") = FR_Table.Fields("[SQL SET COUNT]")
                TO_Table.Fields("[SUM QTY]") = FR_Table.Fields("[SQL SUM QTY]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        
        Loop
End If

Dim I As Integer

'================================
' OEE TOTAL TIME
'================================

sSQL = "SELECT  first([SCHEDULE SETS].[DEPT_ID])    AS [SQL DEPT_ID]," & _
               "first([DEPT CODE].[DESCRIPTION])," & _
               "sum  ([WORK SHEET PT].[TOTAL TIME]/60) AS [SQL OEE TIME] " & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT]" & _
        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION] WHERE [DEPT_ID]=" & FR_Table.Fields("[SQL DEPT_ID]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[OEE TIME]") = FR_Table.Fields("[SQL OEE TIME]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        
        Loop
End If

'================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    SBE HR
'================================================
        
sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID]) AS [SQL DEPT_ID]," & _
              "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) + sum([FINISH AMP MIN])/sum([FINISH AMP]+0.0001) AS [SQL TIME]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[SCHEDULE SETS].[TYPE_ID] = 'SBE'" & _
        "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION] WHERE [DEPT_ID]=" & FR_Table.Fields("[SQL DEPT_ID]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[PROCESS TIME]") = FR_Table.Fields("[SQL TIME]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        
        Loop
End If
                
'==================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    BARREL
'==================================================
        
sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID]) AS [SQL DEPT_ID]," & _
              "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) + sum([FINISH AMP MIN])/sum([FINISH AMP]+0.0001) AS [SQL TIME]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[SCHEDULE SETS].[TYPE_ID] = 'BARREL'" & _
        "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION] WHERE [DEPT_ID]=" & FR_Table.Fields("[SQL DEPT_ID]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[PROCESS TIME]") = TO_Table.Fields("[PROCESS TIME]") + (FR_Table.Fields("[SQL TIME]") / 60)
                TO_Table.Update
            End If
            FR_Table.MoveNext
        
        Loop
End If

'================================================
'   WO COUNT
'================================================
        
sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID]) AS [SQL DEPT_ID]," & _
              "count([GROUPING].[SET_ID])       AS [SQL WO COUNT]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "GROUP BY [SCHEDULE SETS].[DEPT_ID]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION] WHERE [DEPT_ID]=" & FR_Table.Fields("[SQL DEPT_ID]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[WO COUNT]") = FR_Table.Fields("[SQL WO COUNT]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        
        Loop
End If

End Sub

Private Sub cmdCalculationEQ_Click()

'[2] Calculation EQ Base

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT * FROM [TBL CALCULATION EQ]"
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            FR_Table.Edit
            FR_Table.Fields("[SET COUNT]") = 0
            FR_Table.Fields("[WO COUNT]") = 0
            FR_Table.Fields("[SUM QTY]") = 0
            FR_Table.Fields("[PROCESS TIME]") = 0
            FR_Table.Fields("[OEE TIME]") = 0
            FR_Table.Fields("[PERFORMANCE]") = 0
            FR_Table.Update
            FR_Table.MoveNext
        Loop
End If

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE]) AS [SQL EQ BASE]," & _
                      "count([SCHEDULE SETS].[SET_ID])  AS [SQL SET COUNT]," & _
                      "sum  ([SCHEDULE SETS].[RUN QTY]) AS [SQL SUM QTY]" & _
               "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
               "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                     "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                     "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                     "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                     "[DEPT CODE].[DEPT_ID] NOT IN (555,556,557,553,554,558) " & _
               "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "JR"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE]) AS [SQL EQ BASE]," & _
                      "count([SCHEDULE SETS].[SET_ID])  AS [SQL SET COUNT]," & _
                      "sum  ([SCHEDULE SETS].[RUN QTY]) AS [SQL SUM QTY]" & _
               "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
               "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                     "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                     "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                     "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                     "[DEPT CODE].[DEPT_JR_ID] NOT IN (555,556,557,553,554,558) " & _
               "GROUP BY [SCHEDULE SETS].[EQ BASE]"
End Select
                
Set FR_Table = FR_Database.OpenRecordset(sSQL)
'
'       [SQL SET COUNT],[SQL SUM QTY]
'
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ BASE]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[SET COUNT]") = FR_Table.Fields("[SQL SET COUNT]")
                TO_Table.Fields("[SUM QTY]") = FR_Table.Fields("[SQL SUM QTY]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================
' OEE TOTAL TIME SQL OEE TIME]
'================================
Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])    AS [SQL EQ BASE]," & _
                      "sum  ([WORK SHEET PT].[TOTAL TIME]/60) AS [SQL OEE TIME] " & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT]" & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[WORK SHEET PT].[CODE_ID] IN (300,400)" & _
                "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "JR"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])    AS [SQL EQ BASE]," & _
                      "sum  ([WORK SHEET PT].[TOTAL TIME]/60) AS [SQL OEE TIME] " & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT]" & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[WORK SHEET PT].[CODE_ID] IN (300,400)" & _
                "GROUP BY [SCHEDULE SETS].[EQ BASE]"
End Select
                                                                
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ BASE]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
            TO_Table.Edit
            TO_Table.Fields("[OEE TIME]") = FR_Table.Fields("[SQL OEE TIME]")
            TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    SBE HR
'================================================
Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])           AS [SQL EQ BASE]," & _
                      "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) AS [SQL TIME]" & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[SCHEDULE SETS].[TYPE_ID] = 'SBE'" & _
                "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "JR"
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])           AS [SQL EQ BASE]," & _
                      "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) AS [SQL TIME]" & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[SCHEDULE SETS].[TYPE_ID] = 'SBE'" & _
                "GROUP BY [SCHEDULE SETS].[EQ BASE]"
End Select
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ BASE]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
            TO_Table.Edit
            TO_Table.Fields("[PROCESS TIME]") = FR_Table.Fields("[SQL TIME]")
            TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If
                
'==================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    BARREL
'==================================================
Select Case LOCATION_ID
Case "NY"
            sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])           AS [SQL EQ BASE]," & _
                          "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) AS [SQL TIME]" & _
                    "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                    "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                          "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                          "[SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                          "[SCHEDULE SETS].[TYPE_ID] = 'BARREL'" & _
                    "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "JR"
            sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])           AS [SQL EQ BASE]," & _
                          "sum([BASE AMP MIN])/sum([BASE AMP]+0.0001) AS [SQL TIME]" & _
                    "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
                    "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                          "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                          "[SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                          "[SCHEDULE SETS].[TYPE_ID] = 'BARREL'" & _
                    "GROUP BY [SCHEDULE SETS].[EQ BASE]"
End Select
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ BASE]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[PROCESS TIME]") = TO_Table.Fields("[PROCESS TIME]") + (FR_Table.Fields("[SQL TIME]") / 60)
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================================
'   WO COUNT
'================================================
Select Case LOCATION_ID
Case "NY"
    sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])    AS [SQL EQ BASE]," & _
                  "count([GROUPING].[SET_ID])          AS [SQL WO COUNT]" & _
            "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
            "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "JR"
    sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])    AS [SQL EQ BASE]," & _
                  "count([GROUPING].[SET_ID])          AS [SQL WO COUNT]" & _
            "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
            "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [SCHEDULE SETS].[EQ BASE]"
End Select
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ BASE]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[WO COUNT]") = FR_Table.Fields("[SQL WO COUNT]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'MsgBox "Complete", vbInformation, "ATC Plating"

End Sub

Private Sub cmdCalculationEQFinish_Click()

'[2] Calculation EQ Finish

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim DEPT_LOCATION As String
Select Case LOCATION_ID
Case "NY"
            DEPT_LOCATION = "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND "
Case "JR"
            DEPT_LOCATION = "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND "
End Select

sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH]) AS [SQL EQ FINISH]," & _
              "count([SCHEDULE SETS].[SET_ID])    AS [SQL SET COUNT]," & _
              "sum  ([SCHEDULE SETS].[RUN QTY])   AS [SQL SUM QTY]" & _
       "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
       DEPT_LOCATION & _
             "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
             "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
             "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
             "[DEPT CODE].[DEPT_ID] NOT IN (555,556,557,553,554,558) " & _
       "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
                
Set FR_Table = FR_Database.OpenRecordset(sSQL)
'
'       [SQL SET COUNT],[SQL SUM QTY]
'
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ FINISH]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[SET COUNT]") = FR_Table.Fields("[SQL SET COUNT]")
                TO_Table.Fields("[SUM QTY]") = FR_Table.Fields("[SQL SUM QTY]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================
' OEE TOTAL TIME SQL OEE TIME]
'================================
sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])    AS [SQL EQ FINISH]," & _
              "sum  ([WORK SHEET PT].[TOTAL TIME]/60)   AS [SQL OEE TIME] " & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT]" & _
        DEPT_LOCATION & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[WORK SHEET PT].[CODE_ID] IN (500,600)" & _
        "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
                                                                
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ FINISH]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[OEE TIME]") = FR_Table.Fields("[SQL OEE TIME]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    SBE HR
'================================================

sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])             AS [SQL EQ FINISH]," & _
              "sum([FINISH AMP MIN])/sum([FINISH AMP]+0.0001) AS [SQL TIME]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
        DEPT_LOCATION & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[SCHEDULE SETS].[TYPE_ID] = 'SBE'" & _
        "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ FINISH]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
            TO_Table.Edit
            TO_Table.Fields("[PROCESS TIME]") = FR_Table.Fields("[SQL TIME]")
            TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If
                
'==================================================
'   DEPARTMENT PROCESS HOURS PER TYPE    BARREL
'==================================================

sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])               AS [SQL EQ FINISH]," & _
              "sum([FINISH AMP MIN])/sum([FINISH AMP]+0.0001) AS [SQL TIME]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT] " & _
        DEPT_LOCATION & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[SCHEDULE SETS].[TYPE_ID] = 'BARREL'" & _
        "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ FINISH]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[PROCESS TIME]") = TO_Table.Fields("[PROCESS TIME]") + (FR_Table.Fields("[SQL TIME]") / 60)
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'================================================
'   WO COUNT
'================================================

sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])    AS [SQL EQ FINISH]," & _
              "count([GROUPING].[SET_ID])            AS [SQL WO COUNT]" & _
        "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
        DEPT_LOCATION & _
              "[SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
              "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
              "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        Do Until FR_Table.EOF
            sSQL = "SELECT * FROM [TBL CALCULATION EQ] WHERE [NUMBER]=" & FR_Table.Fields("[SQL EQ FINISH]")
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                TO_Table.Edit
                TO_Table.Fields("[WO COUNT]") = FR_Table.Fields("[SQL WO COUNT]")
                TO_Table.Update
            End If
            FR_Table.MoveNext
        Loop
End If

'MsgBox "Complete", vbInformation, "ATC Plating"

End Sub

Private Sub cmdConvert_Click()

Screen.MousePointer = vbHourglass
Get_DV
Screen.MousePointer = vbDefault
 
MsgBox "Complete", vbInformation, "ATC Plating"

End Sub

Private Sub cmdCT_Click()

SET_ID = Trim(txtSET_ID.Text)
LETTER_ID = Trim(txtLETTER_ID.Text)

If (fraTable.Visible = True) Then
    fraTable.Visible = False
Else
    fraTable.Visible = True
End If

Select Case Calculate_Plating_Table
Case 1
        fraTable.Caption = SET_ID & LETTER_ID & " Problem Missing Lot Number"
Case 2
        fraTable.Caption = SET_ID & LETTER_ID & " No Sort Req"
                                                'DV and TOL same one DV
Case 3
        fraTable.Caption = SET_ID & LETTER_ID & " Lot Sort Required"
                                                'DV same and TOL Diff One DV Lot Sort
Case 4
        fraTable.Caption = SET_ID & LETTER_ID & " Problem 11 Too Many Bins"
                                                'Second Sort Required Not Able to Generate Sort
Case 5
        fraTable.Caption = SET_ID & LETTER_ID & " OK 2nd Lot Sort Req"
                                                ' Work Order conatins Lot Number
Case 6
        fraTable.Caption = SET_ID & LETTER_ID & " OK 2nd Sort Req"
                                                ' Grouping of equal DV diff Tolerance
Case 7
        fraTable.Caption = SET_ID & LETTER_ID & " Problem Sort Order"
                                                'Not Able To Generate Sort Table
Case 0
        fraTable.Caption = SET_ID & LETTER_ID & " OK"
End Select

Dim I As Integer
For I = 0 To 9
    txtTolLimit(I).Text = ""
    If (gdBinLimit(I + 1) <> 0) Then
        Select Case gdBinLimit(I + 1)
        Case 0 To 10
                txtTolLimit(I).Text = Format(gdBinLimit(I + 1), "#,###,##0.00")
        Case 10 To 20
                txtTolLimit(I).Text = Format(gdBinLimit(I + 1), "#,###,##0.0")
        Case Else
                txtTolLimit(I).Text = Format(gdBinLimit(I + 1), "#,###,##0")
        End Select
    End If
    
Next I
For I = 0 To 8
    txtTol(I).Text = gsBinTol(I + 1)
Next I

End Sub

Private Sub cmdExcel_Click()

DATE_START_ID = DTPicker1.Value
DATE_END_ID = DTPicker2.Value
  
Dim objExcel As Object
Set objExcel = CreateObject("EXCEL.SHEET")

objExcel.Application.Visible = True
            
Screen.MousePointer = vbHourglass

Dim sSQL As String
                                  
Set TO_Database = OpenDatabase(DB_OEE_VISUAL)

'Set FR2_Database = OpenDatabase(SERVER_DB_NY & "OEE SPM JR MASTER.mdb")

Set FR2_Database = OpenDatabase(DB_OEE_VISUAL)

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [SCHEDULE SETS].[SET NUMBER] AS [SQL SET NUMBER]," & _
              "[SCHEDULE SETS].[DEPT_ID]    AS [SQL DEPT_ID]," & _
              "[SCHEDULE SETS].[DATE_ID]    AS [SQL SCHED DATE_ID]," & _
              "[WORK SHEET PT].[DATE_ID]    AS [SQL DATE_ID]," & _
              "[SCHEDULE SETS].[TYPE_ID]    AS [SQL TYPE_ID]," & _
              "[SCHEDULE SETS].[SERIES_ID]  AS [SQL SERIES_ID]," & _
              "[SCHEDULE SETS].[EQ BASE]    AS [SQL EQ BASE]," & _
              "[SCHEDULE SETS].[EQ FINISH]  AS [SQL EQ FINISH]," & _
                   "[GROUPING].[WORK ORDER] AS [SQL WORK ORDER]," & _
                   "[GROUPING].[LOT NUM]    AS [SQL LOT NUMBER]," & _
                   "[GROUPING].[ATC PART]   AS [SQL ATC PART]," & _
                   "[GROUPING].[QTY]        AS [SQL QTY]" & _
       "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
       "WHERE [SCHEDULE SETS].[SET_ID]  = [GROUPING].[SET_ID] AND " & _
             "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
             "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
             "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# "

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim iRow As Integer, iCol As Integer
iRow = 1

objExcel.Application.Cells(iRow, 1).Value = "[SET NUMBER]"
objExcel.Application.Cells(iRow, 2).Value = "[DEPT_ID]"
objExcel.Application.Cells(iRow, 3).Value = "Schedule Date"
objExcel.Application.Cells(iRow, 4).Value = "[DATE PLATE]"
objExcel.Application.Cells(iRow, 5).Value = "[TYPE_ID]"
objExcel.Application.Cells(iRow, 6).Value = "[SERIES_ID]"
                                                                                        
objExcel.Application.Cells(iRow, 7).Value = "[WORK ORDER]"
objExcel.Application.Cells(iRow, 8).Value = "[LOT NUMBER]"
objExcel.Application.Cells(iRow, 9).Value = "[ATC PART]"
objExcel.Application.Cells(iRow, 10).Value = "[SQL QTY]"


objExcel.Application.Cells(iRow, 11).Value = "Close Date"
objExcel.Application.Cells(iRow, 12).Value = "Dept"
objExcel.Application.Cells(iRow, 13).Value = "Defect"
objExcel.Application.Cells(iRow, 14).Value = "Closed Qty"

objExcel.Application.Cells(iRow, 15).Value = "[SQL EQ BASE]"
objExcel.Application.Cells(iRow, 16).Value = "[SQL EQ FINISH]"

If (optDefect.Value = True) Then
   Dim DEFECT_ID As Long
   DEFECT_ID = Val(txtDefect.Text)
End If

If (FR_Table.RecordCount <> 0) Then
   Do Until FR_Table.EOF
        iRow = iRow + 1
                
        If (optDefect.Value = True) Then
        
                sSQL = "SELECT [WORK SHEET].[DATE_ID] AS [SQL DATE_ID]," & _
                              "[WORK SHEET].[CODE_ID] AS [SQL CODE_ID] " & _
                        "FROM [WORK SHEET],[DEFECT LIST],[DEFECTS]" & _
                        "WHERE [WORK SHEET].[WS_ID]      = [DEFECTS].[WS_ID] AND " & _
                             "[DEFECT LIST].[DEFECT_ID]  = [DEFECTS].[DEFECT_ID] AND " & _
                              "[WORK SHEET].[WORK ORDER] ='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                                 "[DEFECTS].[DEFECT_ID]  =" & DEFECT_ID
                
                Set TO_Table = TO_Database.OpenRecordset(sSQL)
                                
                objExcel.Application.Cells(iRow, 1).Value = FR_Table.Fields("[SQL SET NUMBER]")
                objExcel.Application.Cells(iRow, 2).Value = FR_Table.Fields("[SQL DEPT_ID]")
                objExcel.Application.Cells(iRow, 3).Value = Format(FR_Table.Fields("[SQL SCHED DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 4).Value = Format(FR_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 5).Value = FR_Table.Fields("[SQL TYPE_ID]")
                objExcel.Application.Cells(iRow, 6).Value = FR_Table.Fields("[SQL SERIES_ID]")
                                                                                                        
                objExcel.Application.Cells(iRow, 7).Value = FR_Table.Fields("[SQL WORK ORDER]")
                objExcel.Application.Cells(iRow, 8).Value = FR_Table.Fields("[SQL LOT NUMBER]")
                objExcel.Application.Cells(iRow, 9).Value = FR_Table.Fields("[SQL ATC PART]")
                objExcel.Application.Cells(iRow, 10).Value = FR_Table.Fields("[SQL QTY]")
                If (TO_Table.RecordCount <> 0) Then
                        objExcel.Application.Cells(iRow, 11).Value = Format(TO_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                        objExcel.Application.Cells(iRow, 12).Value = TO_Table.Fields("[SQL CODE_ID]")
                End If
                objExcel.Application.Cells(iRow, 15).Value = FR_Table.Fields("[SQL EQ BASE]")
                objExcel.Application.Cells(iRow, 16).Value = FR_Table.Fields("[SQL EQ FINISH]")
 
        End If
                
        If (optCancel.Value = True) Then
        
                Dim SQL_CODE_ID(11) As Integer
                Dim FIND_DEPT_ID As Integer
        
                SQL_CODE_ID(1) = 988    'CANCEL
                SQL_CODE_ID(2) = 985    'CLOSE
                
                SQL_CODE_ID(3) = 705    'JR
                SQL_CODE_ID(4) = 706    'JR
                SQL_CODE_ID(5) = 729    'JR
                SQL_CODE_ID(6) = 735    'JR JUAREZ DB
                
                SQL_CODE_ID(7) = 670    'NY
                SQL_CODE_ID(8) = 692    'NY
                SQL_CODE_ID(9) = 693    'NY
                
                SQL_CODE_ID(10) = 296    'LOT
                SQL_CODE_ID(11) = 297    'LOT
                
                Dim I As Integer
                For I = 1 To 11
                
                    sSQL = "SELECT [DEFECT LIST].[DESCRIPTION] AS [SQL Description]," & _
                                   "[WORK SHEET].[DATE_ID]     AS [SQL DATE_ID]," & _
                                   "[WORK SHEET].[QUANTITY]    AS [SQL INSPECTED]," & _
                                   "[WORK SHEET].[CODE_ID]     AS [SQL CODE_ID] " & _
                        "FROM [WORK SHEET],[DEFECT LIST],[DEFECTS]" & _
                        "WHERE [WORK SHEET].[WS_ID]      = [DEFECTS].[WS_ID] AND " & _
                             "[DEFECT LIST].[DEFECT_ID]  = [DEFECTS].[DEFECT_ID] AND " & _
                              "[WORK SHEET].[WORK ORDER] ='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                              "[WORK SHEET].[CODE_ID]    = " & SQL_CODE_ID(I)
                
                    Select Case SQL_CODE_ID(I)
                    Case 988
                            Set TO_Table = TO_Database.OpenRecordset(sSQL)
                            If (TO_Table.RecordCount <> 0) Then
                                    objExcel.Application.Cells(iRow, 11).Value = Format(TO_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                                    objExcel.Application.Cells(iRow, 12).Value = TO_Table.Fields("[SQL CODE_ID]")
                                    objExcel.Application.Cells(iRow, 13).Value = TO_Table.Fields("[SQL Description]")
                                    objExcel.Application.Cells(iRow, 14).Value = TO_Table.Fields("[SQL INSPECTED]")
                                    FIND_DEPT_ID = 1
                            End If
                    Case 705, 706, 729
                            Set TO_Table = TO_Database.OpenRecordset(sSQL)
                            If (TO_Table.RecordCount <> 0) Then
                                    objExcel.Application.Cells(iRow, 11).Value = Format(TO_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                                    objExcel.Application.Cells(iRow, 12).Value = TO_Table.Fields("[SQL CODE_ID]")
                                    objExcel.Application.Cells(iRow, 13).Value = TO_Table.Fields("[SQL Description]")
                                    objExcel.Application.Cells(iRow, 14).Value = TO_Table.Fields("[SQL INSPECTED]")
                                    objExcel.Application.Cells(iRow, 17).Value = "JR"
                                    Exit For
                            End If
                    Case 735
                            Set FR2_Table = FR2_Database.OpenRecordset(sSQL)
                            If (FR2_Table.RecordCount <> 0) Then
                                    objExcel.Application.Cells(iRow, 11).Value = Format(FR2_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                                    objExcel.Application.Cells(iRow, 12).Value = FR2_Table.Fields("[SQL CODE_ID]")
                                    objExcel.Application.Cells(iRow, 13).Value = FR2_Table.Fields("[SQL Description]")
                                    objExcel.Application.Cells(iRow, 14).Value = FR2_Table.Fields("[SQL INSPECTED]")
                                    objExcel.Application.Cells(iRow, 17).Value = "JR"
                                    Exit For
                            End If
                    Case Else
                            Set TO_Table = TO_Database.OpenRecordset(sSQL)
                            If (TO_Table.RecordCount <> 0) Then
                                    objExcel.Application.Cells(iRow, 11).Value = Format(TO_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                                    objExcel.Application.Cells(iRow, 12).Value = TO_Table.Fields("[SQL CODE_ID]")
                                    objExcel.Application.Cells(iRow, 13).Value = TO_Table.Fields("[SQL Description]")
                                    objExcel.Application.Cells(iRow, 14).Value = TO_Table.Fields("[SQL INSPECTED]")
                                    If (FIND_DEPT_ID = 1) Then
                                        objExcel.Application.Cells(iRow, 17).Value = "NY"
                                    End If
                                    Exit For
                            End If
                    End Select
                
                Next I
                
                objExcel.Application.Cells(iRow, 1).Value = FR_Table.Fields("[SQL SET NUMBER]")
                objExcel.Application.Cells(iRow, 2).Value = FR_Table.Fields("[SQL DEPT_ID]")
                objExcel.Application.Cells(iRow, 3).Value = Format(FR_Table.Fields("[SQL SCHED DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 4).Value = Format(FR_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 5).Value = FR_Table.Fields("[SQL TYPE_ID]")
                objExcel.Application.Cells(iRow, 6).Value = FR_Table.Fields("[SQL SERIES_ID]")
                                                                                                        
                objExcel.Application.Cells(iRow, 7).Value = FR_Table.Fields("[SQL WORK ORDER]")
                objExcel.Application.Cells(iRow, 8).Value = FR_Table.Fields("[SQL LOT NUMBER]")
                objExcel.Application.Cells(iRow, 9).Value = FR_Table.Fields("[SQL ATC PART]")
                objExcel.Application.Cells(iRow, 10).Value = FR_Table.Fields("[SQL QTY]")
                objExcel.Application.Cells(iRow, 15).Value = FR_Table.Fields("[SQL EQ BASE]")
                objExcel.Application.Cells(iRow, 16).Value = FR_Table.Fields("[SQL EQ FINISH]")
                
        End If
        If (optAll.Value = True) Then
                objExcel.Application.Cells(iRow, 1).Value = FR_Table.Fields("[SQL SET NUMBER]")
                objExcel.Application.Cells(iRow, 2).Value = FR_Table.Fields("[SQL DEPT_ID]")
                objExcel.Application.Cells(iRow, 3).Value = Format(FR_Table.Fields("[SQL SCHED DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 4).Value = Format(FR_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                objExcel.Application.Cells(iRow, 5).Value = FR_Table.Fields("[SQL TYPE_ID]")
                objExcel.Application.Cells(iRow, 6).Value = FR_Table.Fields("[SQL SERIES_ID]")
                                                                                                        
                objExcel.Application.Cells(iRow, 7).Value = FR_Table.Fields("[SQL WORK ORDER]")
                objExcel.Application.Cells(iRow, 8).Value = FR_Table.Fields("[SQL LOT NUMBER]")
                objExcel.Application.Cells(iRow, 9).Value = FR_Table.Fields("[SQL ATC PART]")
                objExcel.Application.Cells(iRow, 10).Value = FR_Table.Fields("[SQL QTY]")
                objExcel.Application.Cells(iRow, 15).Value = FR_Table.Fields("[SQL EQ BASE]")
                objExcel.Application.Cells(iRow, 16).Value = FR_Table.Fields("[SQL EQ FINISH]")
        End If
        FR_Table.MoveNext
   Loop
End If
       
FR2_Database.Close
FR_Database.Close
TO_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Dim sFile As String
sFile = "c:\ATC\PLATING WO DOWNLOAD.xls"

objExcel.SaveAs sFile
objExcel.Application.Quit
Set objExcel = Nothing

MsgBox "Excel Update Complete", vbInformation, "ATC Plating"

End Sub

Private Sub cmdExcelNew_Click()

DATE_START_ID = DTPicker1.Value
DATE_END_ID = DTPicker2.Value
  
Dim objExcel As Object
Set objExcel = CreateObject("EXCEL.SHEET")

objExcel.Application.Visible = True
            
Screen.MousePointer = vbHourglass

Dim sSQL As String
Dim sSQL2 As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL2 = "SELECT [MACHINE].[NUMBER]                            AS [SQL MACHINE_ID]," & _
               "[MACHINE].[TYPE]                              AS [SQL MACH TYPE]," & _
               "[MACHINE].[DESCRIPTION]                       AS [SQL MACH DESC]," & _
               "mid([PROCESS],1,1)& mid(lcase([PROCESS]),2,5) AS [SQL PROCESS]  " & _
       "FROM [MACHINE]  " & _
       "WHERE [DEPT_ID]='PT' AND [ACTIVE] = 1 AND" & _
             "[LOCATION_ID]='" & LOCATION_ID & "'" & _
       "ORDER BY [TYPE],[NUMBER],mid([PROCESS],1,1)"
         
Set FR_Table = FR_Database.OpenRecordset(sSQL2)

Dim iDayOfWeek As Integer

Dim iRow As Integer, iCol As Integer
iRow = 3

objExcel.Application.Cells(iRow, 1).Value = "Tank"
objExcel.Application.Cells(iRow, 2).Value = "Type"
objExcel.Application.Cells(iRow, 3).Value = "Description"

Dim I As Integer

For I = 0 To 5
        objExcel.Application.Cells(iRow, 4 + 2 * I).Value = Format(DateAdd("d", I, DATE_START_ID), "ddd")
        objExcel.Application.Cells(iRow, 5 + 2 * I).Value = Format(DateAdd("d", I, DATE_START_ID), "MM/DD/YYYY")
Next I

iRow = 5
If (FR_Table.RecordCount <> 0) Then
   Do Until FR_Table.EOF
                
        iRow = iRow + 1
        objExcel.Application.Cells(iRow, 1).Value = FR_Table.Fields("[SQL MACHINE_ID]")
        objExcel.Application.Cells(iRow, 2).Value = FR_Table.Fields("[SQL MACH TYPE]")
        objExcel.Application.Cells(iRow, 3).Value = FR_Table.Fields("[SQL MACH DESC]")
                
        Select Case FR_Table.Fields("[SQL PROCESS]")
        
        Case "Base"
            '9 BASE EQ
            sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])      AS [SQL MACHINE_ID]," & _
                                "first([MACHINE].[TYPE])         AS [SQL MACH TYPE]," & _
                                "first([MACHINE].[DESCRIPTION])  AS [SQL MACH DESC]," & _
                             "first([WORK SHEET PT].[DATE_ID])   AS [SQL DATE_ID]," & _
                          "count([SCHEDULE SETS].[SET_ID])       AS [SQL SET COUNT]," & _
                            "sum([SCHEDULE SETS].[RUN QTY])      AS [SQL SUM QTY]" & _
                    "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT],[MACHINE] " & _
                    "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                          "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                          "[SCHEDULE SETS].[EQ BASE] = [MACHINE].[NUMBER] AND " & _
                          "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                          "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                          "[SCHEDULE SETS].[EQ BASE]=" & FR_Table.Fields("[SQL MACHINE_ID]") & " AND " & _
                          "[DEPT CODE].[DEPT_ID] NOT IN (555,556,557,553,554,558)" & _
                    "GROUP BY [SCHEDULE SETS].[EQ BASE],[WORK SHEET PT].[DATE_ID]"
        
        Case "Finish"
           '3 Finish EQ
           sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])    AS [SQL MACHINE_ID]," & _
                               "first([MACHINE].[TYPE])         AS [SQL MACH TYPE]," & _
                               "first([MACHINE].[DESCRIPTION])  AS [SQL MACH DESC]," & _
                            "first([WORK SHEET PT].[DATE_ID])   AS [SQL DATE_ID]," & _
                         "count([SCHEDULE SETS].[SET_ID])       AS [SQL SET COUNT]," & _
                           "sum([SCHEDULE SETS].[RUN QTY])      AS [SQL SUM QTY]" & _
                   "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT],[MACHINE] " & _
                   "WHERE [SCHEDULE SETS].[DEPT_ID]= [DEPT CODE].[DEPT_ID] AND " & _
                         "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                         "[SCHEDULE SETS].[EQ FINISH] = [MACHINE].[NUMBER] AND " & _
                         "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                         "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                         "[SCHEDULE SETS].[EQ FINISH]=" & FR_Table.Fields("[SQL MACHINE_ID]") & " " & _
                   "GROUP BY [SCHEDULE SETS].[EQ FINISH],[WORK SHEET PT].[DATE_ID]"
        
        End Select
        
        Set TO_Table = TO_Database.OpenRecordset(sSQL)
        If (TO_Table.RecordCount <> 0) Then
            Do Until TO_Table.EOF
                ' 1 Sunday to 7 Saturday
                iDayOfWeek = Format(TO_Table.Fields("[SQL DATE_ID]"), "W")
                objExcel.Application.Cells(iRow, iDayOfWeek * 2).Value = Format(TO_Table.Fields("[SQL SUM QTY]"), "###,##0")
                objExcel.Application.Cells(iRow, 1 + iDayOfWeek * 2).Value = TO_Table.Fields("[SQL SET COUNT]")
                
                TO_Table.MoveNext
            Loop
        End If
        FR_Table.MoveNext
   Loop
                                                                                              
End If
                                                                                                 
iRow = iRow + 2

Dim CASE_SIZE(5) As String
CASE_SIZE(1) = "A"
CASE_SIZE(2) = "B"
CASE_SIZE(3) = "R"
CASE_SIZE(4) = "C"
CASE_SIZE(5) = "E"

Dim CASE_DATE_QUANTITY(5, 7) As Long
Dim CASE_DATE_SET_COUNT(5, 7) As Integer

Dim iRowStart As Integer
iRowStart = iRow
                                   
sSQL = "SELECT   [WORK SHEET PT].[DATE_ID]    AS [SQL DATE_ID]," & _
                "[SCHEDULE SETS].[SET_ID]     AS [SQL SET ID]," & _
                "[SCHEDULE SETS].[RUN QTY]    AS [SQL QTY]" & _
        "FROM [SCHEDULE SETS],[WORK SHEET PT] " & _
        "WHERE [SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID]  AND " & _
              "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# "
                   
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
       
       For I = 1 To 5
            sSQL = "SELECT count([GROUPING].[SET_ID]) " & _
                   "FROM [GROUPING] " & _
                   "WHERE mid([GROUPING].[ATC PART],4,1)='" & CASE_SIZE(I) & "' AND " & _
                             "[GROUPING].[SET_ID]=" & FR_Table.Fields("[SQL SET ID]") & " " & _
                    "GROUP BY mid([GROUPING].[ATC PART],4,1),[GROUPING].[SET_ID]"
                              
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
            If (TO_Table.RecordCount <> 0) Then
                Do Until TO_Table.EOF
                     ' 1 Sunday to 7 Saturday
                     iDayOfWeek = Format(FR_Table.Fields("[SQL DATE_ID]"), "W")
                     CASE_DATE_QUANTITY(I, iDayOfWeek) = CASE_DATE_QUANTITY(I, iDayOfWeek) + FR_Table.Fields("[SQL QTY]")
                     CASE_DATE_SET_COUNT(I, iDayOfWeek) = CASE_DATE_SET_COUNT(I, iDayOfWeek) + 1
                    
                     TO_Table.MoveNext
                Loop
            End If
       
       Next I
       
       FR_Table.MoveNext
    Loop
End If
                                                                                                     
iRow = iRowStart
Dim j As Integer

For j = 1 To 5
         objExcel.Application.Cells(iRow + j, 2).Value = CASE_SIZE(j)
Next j
' 1 Sunday to 7 Saturday
For I = 2 To 7
    For j = 1 To 5
         objExcel.Application.Cells(iRow + j, I * 2).Value = Format(CASE_DATE_QUANTITY(j, I), "###,##0")
         objExcel.Application.Cells(iRow + j, 1 + I * 2).Value = CASE_DATE_SET_COUNT(j, I)
    Next j
Next I
'===========================================================
'       Equipment Downtime CODE 101 UNPLANNED DOWNTIME
'===========================================================
iRow = 25

Dim sDate(7) As String
sDate(1) = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
sDate(2) = DateAdd("D", 1, sDate(1))
sDate(3) = DateAdd("D", 2, sDate(1))
sDate(4) = DateAdd("D", 3, sDate(1))
sDate(5) = DateAdd("D", 4, sDate(1))
sDate(6) = DateAdd("D", 5, sDate(1))
sDate(7) = DateAdd("D", 6, sDate(1))
                                                             
                                   
                        ' (iRow,iCol)
'==========================================================================================
'   1.0 Header  Horizontal
'==========================================================================================
For I = 1 To 6
    objExcel.Application.Cells(iRow + 2, 2 * I + 2).Value = Format(sDate(I), "dddd")  ' Day of Week
    objExcel.Application.Cells(iRow + 2, 2 * I + 3).Value = sDate(I)                  ' Date
    objExcel.Application.Cells(iRow + 3, 2 * I + 2).Value = "Quantity"
    objExcel.Application.Cells(iRow + 3, 2 * I + 3).Value = "Count"
Next I

Set FR_Table = FR_Database.OpenRecordset(sSQL2)

If (FR_Table.RecordCount <> 0) Then
        
    objExcel.Application.Cells(iRow + 4, 2).Value = "Down Time"
        
    iRow = iRow + 5
    
    Do Until FR_Table.EOF
    
            'Per Machine
            sSQL = "SELECT  first([MACHINE].[NUMBER])           AS [SQL MACHINE_ID]," & _
                           "first([MACHINE].[DESCRIPTION])      AS [SQL DESCRIPTION]," & _
                        "first([WORK SHEET].[DATE_ID])          AS [SQL DATE_ID]," & _
                          "sum([WORK SHEET].[TOTAL TIME])       AS [SQL SUM TIME]," & _
                        "count([WORK SHEET].[TOTAL TIME])       AS [SQL COUNT] " & _
               "FROM [WORK SHEET],[MACHINE]" & _
               "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[NUMBER] AND " & _
                      "[WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[WORK SHEET].[TOTAL TIME]<>0 AND " & _
                      "[WORK SHEET].[MACHINE_ID]= " & FR_Table.Fields("[SQL MACHINE_ID]") & " AND " & _
                      "[WORK SHEET].[CODE_ID] IN (101) " & _
               "GROUP BY [WORK SHEET].[MACHINE_ID],[WORK SHEET].[DATE_ID] "
            
        iCount = iCount + 1
                                                                
        objExcel.Application.Cells(iRow, 1).Value = FR_Table.Fields("[SQL MACHINE_ID]")
        objExcel.Application.Cells(iRow, 2).Value = FR_Table.Fields("[SQL MACH TYPE]")
        objExcel.Application.Cells(iRow, 3).Value = FR_Table.Fields("[SQL MACH DESC]")
                                
        Set TO_Table = TO_Database.OpenRecordset(sSQL)
        Do Until TO_Table.EOF
            ' 1 Sunday to 7 Saturday
            iDayOfWeek = Format(TO_Table.Fields("[SQL DATE_ID]"), "W")
            objExcel.Application.Cells(iRow, iDayOfWeek * 2).Value = Format(TO_Table.Fields("[SQL SUM TIME]"), "##,###,##0")
            objExcel.Application.Cells(iRow, iDayOfWeek * 2 + 1).Value = Format(TO_Table.Fields("[SQL COUNT]"), "##,###,##0")
            TO_Table.MoveNext
        Loop
                
        FR_Table.MoveNext
        iRow = iRow + 1
    Loop
End If
           
TO_Database.Close
FR_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Dim sFile As String
sFile = "c:\ATC\PLATING SUM.xls"

objExcel.SaveAs sFile
objExcel.Application.Quit
Set objExcel = Nothing

MsgBox "Excel Update Complete", vbInformation, "ATC Plating"
 
End Sub

Private Sub cmdLift_Click()

DATE_START_ID = DTPicker1.Value
DATE_END_ID = DTPicker2.Value
            
Dim ENCAPSULATED As String
            
Screen.MousePointer = vbHourglass

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim SET_NUM As Integer

Dim sSQL As String
Dim sSQL2 As String

sSQL = "SELECT * FROM [TBL LIFT TERM]"

Set TO_Table = TO_Database.OpenRecordset(sSQL)
If (TO_Table.RecordCount <> 0) Then
   Do Until TO_Table.EOF
            TO_Table.Delete
            TO_Table.MoveNext
   Loop
End If
Set TO_Table = TO_Database.OpenRecordset(sSQL)
                                                                                                    
Set FR2_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set FR4_Database = OpenDatabase(DB_PLATING_TERMINATION)

Select Case LOC_SOURCE_ID
Case "NY"
        Set FR_Database = OpenDatabase(DB_OEE_VISUAL)
        Set FR3_Database = OpenDatabase(DB_OEE_VISUAL)
Case "JR"
        Set FR_Database = OpenDatabase(SERVER_DB_NY & "OEE SPM JR MASTER.mdb")
        Set FR3_Database = OpenDatabase(SERVER_DB_NY & "OEE SPM JR MASTER.mdb")
End Select

sSQL = "SELECT [WORK SHEET].[DATE_ID]    AS [SQL DATE_ID]," & _
              "[WORK SHEET].[WORK ORDER] AS [SQL WORK ORDER]," & _
              "[WORK SHEET].[LOT NUM]    AS [SQL LOT NUM]," & _
              "[WORK SHEET].[ATC PART]   AS [SQL ATC PART]," & _
              "[WORK SHEET].[QUANTITY]   AS [SQL QUANTITY]," & _
              "[WORK SHEET].[CODE_ID]    AS [SQL CODE_ID] " & _
        "FROM [WORK SHEET] " & _
        "WHERE [WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
              "[WORK SHEET].[CODE_ID] = 513 AND " & _
          "MID([WORK SHEET].[ATC PART],1,4) IN ('100A','100B','100E','710A','200A','200B','700A','700B','800C','800E','830C','830E') AND " & _
          "MID([WORK SHEET].[ATC PART],10,1) IN ('N','1','3','5','7','9') "
                                                                                                                             
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
   Do Until FR_Table.EOF
        TO_Table.AddNew
                
        'If FR_Table.Fields("[SQL WORK ORDER]") = "781384091001" Then
        '    Beep
        'End If
        
        TO_Table.Fields("[WORK ORDER]") = FR_Table.Fields("[SQL WORK ORDER]")
         
        sSQL = "SELECT [SCHEDULE SETS].[SET NUMBER] AS [SQL SET NUMBER]," & _
                      "[SCHEDULE SETS].[DEPT_ID]    AS [SQL DEPT_ID]," & _
                      "[SCHEDULE SETS].[DATE_ID]    AS [SQL SCHED DATE_ID]," & _
                      "[WORK SHEET PT].[DATE_ID]    AS [SQL DATE_ID]," & _
                      "[SCHEDULE SETS].[TYPE_ID]    AS [SQL TYPE_ID]," & _
                      "[SCHEDULE SETS].[SERIES_ID]  AS [SQL SERIES_ID]," & _
                      "[SCHEDULE SETS].[EQ BASE]    AS [SQL EQ BASE]," & _
                      "[SCHEDULE SETS].[HEAD 1]     AS [SQL HEAD 1]," & _
                      "[SCHEDULE SETS].[HEAD 2]     AS [SQL HEAD 2]," & _
                      "[SCHEDULE SETS].[HEAD 3]     AS [SQL HEAD 3]," & _
                      "[SCHEDULE SETS].[HEAD 4]     AS [SQL HEAD 4]," & _
                           "[GROUPING].[P1 BASE]    AS [SQL P1 BASE]," & _
                           "[GROUPING].[P2 BASE]    AS [SQL P2 BASE]," & _
                           "[GROUPING].[P3 BASE]    AS [SQL P3 BASE]," & _
                           "[GROUPING].[P4 BASE]    AS [SQL P4 BASE]," & _
                           "[GROUPING].[WORK ORDER] AS [SQL WORK ORDER]," & _
                           "[GROUPING].[LOT NUM]    AS [SQL LOT NUMBER]," & _
                           "[GROUPING].[ATC PART]   AS [SQL ATC PART]," & _
                           "[GROUPING].[QTY]        AS [SQL QTY]" & _
               "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
               "WHERE [SCHEDULE SETS].[SET_ID]     = [GROUPING].[SET_ID] AND " & _
                     "[SCHEDULE SETS].[SET_ID]     = [WORK SHEET PT].[SET_ID] AND " & _
                          "[GROUPING].[WORK ORDER] ='" & FR_Table.Fields("[SQL WORK ORDER]") & "'"
        
        Set FR2_Table = FR2_Database.OpenRecordset(sSQL)
        If (FR2_Table.RecordCount <> 0) Then
            TO_Table.Fields("[DATE_ID]") = Format(FR2_Table.Fields("[SQL SCHED DATE_ID]"), "MM/DD/YYYY")
            TO_Table.Fields("[SET NUMBER]") = FR2_Table.Fields("[SQL SET NUMBER]")
        End If
        TO_Table.Update
        FR_Table.MoveNext
   Loop
End If
              
Dim objExcel As Object
Set objExcel = CreateObject("EXCEL.SHEET")

objExcel.Application.Visible = True
              
Dim iRow As Integer, iCol As Integer
Dim PBASE(4) As Long

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [TBL LIFT TERM] ORDER BY [DATE_ID],[SET NUMBER]"

Set TO_Table = TO_Database.OpenRecordset(sSQL)

iRow = 1

Dim ROW16 As String
 
If (TO_Table.RecordCount <> 0) Then

   Do Until TO_Table.EOF
        'If TO_Table.Fields("[WORK ORDER]") = "781384091001" Then
         '   Beep
        'End If
        sSQL = "SELECT [WORK SHEET].[DATE_ID]      AS [SQL DATE_ID]," & _
                      "[WORK SHEET].[WORK ORDER]   AS [SQL WORK ORDER]," & _
                      "[WORK SHEET].[LOT NUM]      AS [SQL LOT NUM]," & _
                      "[WORK SHEET].[ATC PART]     AS [SQL ATC PART]," & _
                      "[WORK SHEET].[QUANTITY]     AS [SQL QUANTITY]," & _
                         "[DEFECTS].[DEFECT_ID]    AS [SQL DEFECT ID]," & _
                     "[DEFECT LIST].[DESCRIPTION]  AS [SQL DESCRIPTION]," & _
                      "[WORK SHEET].[CODE_ID]      AS [SQL CODE_ID] " & _
              "FROM [WORK SHEET],[DEFECT LIST],[DEFECTS]" & _
              "WHERE [WORK SHEET].[WS_ID]      = [DEFECTS].[WS_ID] AND " & _
                   "[DEFECT LIST].[DEFECT_ID]  = [DEFECTS].[DEFECT_ID] AND " & _
                    "[WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[WORK SHEET].[CODE_ID] = 513 AND " & _
                    "[WORK SHEET].[WORK ORDER]='" & TO_Table.Fields("[WORK ORDER]") & "'"
                                                                
        Set FR_Table = FR_Database.OpenRecordset(sSQL)

        If (iRow = 1) Then
        '************************** HEADER **********************************************
                objExcel.Application.Cells(iRow, 1).Value = "Dept"
                objExcel.Application.Cells(iRow, 2).Value = "Plating Date"
                objExcel.Application.Cells(iRow, 3).Value = "Schedule Date"
                objExcel.Application.Cells(iRow, 4).Value = "WORK ORDER"
                objExcel.Application.Cells(iRow, 5).Value = "LOT NUMBER"
                objExcel.Application.Cells(iRow, 6).Value = "ATC PART"
                
                objExcel.Application.Cells(iRow, 7).Value = "Base Tank"
                
                objExcel.Application.Cells(iRow, 8).Value = "Set"
                objExcel.Application.Cells(iRow, 9).Value = "Barrel #1"
                objExcel.Application.Cells(iRow, 10).Value = "Barrel #2"
                
                objExcel.Application.Cells(iRow, 11).Value = "1G"
                objExcel.Application.Cells(iRow, 12).Value = "1NG"
                objExcel.Application.Cells(iRow, 13).Value = "2G"
                objExcel.Application.Cells(iRow, 14).Value = "2NG"
                
                'objExcel.Application.Cells(iRow, 15).Value = "Sum Qty"
                objExcel.Application.Cells(iRow, 16).Value = "S 1G"
                objExcel.Application.Cells(iRow, 17).Value = "S 1NG"
                objExcel.Application.Cells(iRow, 18).Value = "S 2G"
                objExcel.Application.Cells(iRow, 19).Value = "S 2NG"
                
                objExcel.Application.Cells(iRow, 20).Value = "Insp. Date"
                objExcel.Application.Cells(iRow, 21).Value = "513 Insp"
                objExcel.Application.Cells(iRow, 22).Value = "670 Yield"
                objExcel.Application.Cells(iRow, 23).Value = "988 Canc"
        End If
        iRow = iRow + 1
    
    If (FR_Table.RecordCount <> 0) Then
            ROW16 = ""
            ' THERE IS A DEFECT DETECTED ********************************************************
            Select Case FR_Table.Fields("[SQL DEFECT ID]")
            Case 113
                    objExcel.Application.Cells(iRow, 21).Value = "Pass"
            Case 127
                    objExcel.Application.Cells(iRow, 21).Value = "Inspect"
            Case 170
                    objExcel.Application.Cells(iRow, 21).Value = "Fail"
            Case Else
                    objExcel.Application.Cells(iRow, 21).Value = "D" & FR_Table.Fields("[SQL DEFECT ID]")
            End Select
            
            ROW16 = objExcel.Application.Cells(iRow, 21).Value
            
            '090 Plating chg 03/08/2013 NO EXCEPTION
            Select Case FR_Table.Fields("[SQL DEFECT ID]")
            Case 127, 113, 170
            
                       sSQL = "SELECT [WORK SHEET].[DATE_ID]    AS [SQL DATE_ID]," & _
                                     "[WORK SHEET].[WORK ORDER] AS [SQL WORK ORDER] " & _
                             "FROM [WORK SHEET] " & _
                             "WHERE [WORK SHEET].[WORK ORDER]='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                                   "[WORK SHEET].[CODE_ID] = 988"
                                   
                       Set FR3_Table = FR3_Database.OpenRecordset(sSQL)
                       If (FR3_Table.RecordCount <> 0) Then
                               objExcel.Application.Cells(iRow, 23).Value = "YES"
                       Else
                               objExcel.Application.Cells(iRow, 23).Value = "NO"
                       End If
                    '
                    'LOCATION NY CODE 670   JR CODE 735,703
                    '
                      Select Case LOC_SOURCE_ID
                      Case "NY"
                                sSQL = "SELECT sum([WORK SHEET].[QUANTITY])    AS [SQL QUANTITY]," & _
                                            "first([WORK SHEET].[START QTY])   AS [SQL START QTY]  " & _
                                     "FROM [WORK SHEET] " & _
                                     "WHERE [WORK SHEET].[WORK ORDER]='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                                           "[WORK SHEET].[CODE_ID] IN (670,705,706) " & _
                                     "GROUP BY [WORK SHEET].[WORK ORDER]"
                                     
                                  'IF NOT FOUND THEN LOOK JR
                                                                          
                       Case "JR"
                                sSQL = "SELECT sum([WORK SHEET].[QUANTITY])    AS [SQL QUANTITY]," & _
                                            "first([WORK SHEET].[START QTY])   AS [SQL START QTY]  " & _
                                     "FROM [WORK SHEET] " & _
                                     "WHERE [WORK SHEET].[WORK ORDER]='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                                           "[WORK SHEET].[CODE_ID] IN (735,703) " & _
                                     "GROUP BY [WORK SHEET].[WORK ORDER]"
                       End Select
                       
                       Set FR3_Table = FR3_Database.OpenRecordset(sSQL)
                       If (FR3_Table.RecordCount <> 0) Then
                                objExcel.Application.Cells(iRow, 22).Value = Format(FR3_Table.Fields("[SQL QUANTITY]") / FR3_Table.Fields("[SQL START QTY]"), "0%")
                       End If
                
            End Select
            '090 Plating chg 03/08/2013
            Select Case FR_Table.Fields("[SQL DEFECT ID]")
            Case 113, 127, 170
                    ' ADD COMMENT LINE
                     sSQL = "SELECT  [TBL PROCESS COMMENT].[COMMENT] AS [SQL COMMENT] " & _
                            "FROM [WORK SHEET],[TBL PROCESS COMMENT] " & _
                            "WHERE [WORK SHEET].[WS_ID]     = [TBL PROCESS COMMENT].[WS_ID] AND " & _
                                  "[WORK SHEET].[WORK ORDER]='" & FR_Table.Fields("[SQL WORK ORDER]") & "' AND " & _
                                  "[WORK SHEET].[CODE_ID] = 290 "
                              
                    Set FR3_Table = FR3_Database.OpenRecordset(sSQL)
                    If (FR3_Table.RecordCount <> 0) Then
                        objExcel.Application.Cells(iRow, 24).Value = FR3_Table.Fields("[SQL COMMENT]")
                    End If
                    
                    objExcel.Application.Cells(iRow, 25).Value = FR_Table.Fields("[SQL DESCRIPTION]")
                    
            End Select
            
            ' THERE IS A DEFECT DETECTED END ********************************************************
        
        Else
        sSQL = "SELECT [WORK SHEET].[DATE_ID]    AS [SQL DATE_ID]," & _
                      "[WORK SHEET].[WORK ORDER] AS [SQL WORK ORDER]," & _
                      "[WORK SHEET].[LOT NUM]    AS [SQL LOT NUM]," & _
                      "[WORK SHEET].[ATC PART]   AS [SQL ATC PART]," & _
                      "[WORK SHEET].[QUANTITY]   AS [SQL QUANTITY] " & _
              "FROM [WORK SHEET] " & _
              "WHERE [WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[WORK SHEET].[CODE_ID] = 513 AND " & _
                    "[WORK SHEET].[WORK ORDER]='" & TO_Table.Fields("[WORK ORDER]") & "'"
                                                                                                                                
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
                
        End If
        
        objExcel.Application.Cells(iRow, 4).Value = FR_Table.Fields("[SQL WORK ORDER]")
        objExcel.Application.Cells(iRow, 5).Value = FR_Table.Fields("[SQL LOT NUM]")
        objExcel.Application.Cells(iRow, 6).Value = FR_Table.Fields("[SQL ATC PART]")
                
        If (Mid(FR_Table.Fields("[SQL ATC PART]"), 9, 1) = "E") Then
                If (objExcel.Application.Cells(iRow, 21).Value = "Inspect") Then
                    objExcel.Application.Cells(iRow, 21).Value = "Encap"
                End If
        End If
        
        objExcel.Application.Cells(iRow, 20).Value = Format(FR_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                
        sSQL = "SELECT [SCHEDULE SETS].[SET NUMBER] AS [SQL SET NUMBER]," & _
                      "[SCHEDULE SETS].[SET_ID]     AS [SQL SET_ID]," & _
                      "[SCHEDULE SETS].[DEPT_ID]    AS [SQL DEPT_ID]," & _
                      "[SCHEDULE SETS].[DATE_ID]    AS [SQL SCHED DATE_ID]," & _
                      "[WORK SHEET PT].[DATE_ID]    AS [SQL DATE_ID]," & _
                      "[SCHEDULE SETS].[TYPE_ID]    AS [SQL TYPE_ID]," & _
                      "[SCHEDULE SETS].[SERIES_ID]  AS [SQL SERIES_ID]," & _
                      "[SCHEDULE SETS].[EQ BASE]    AS [SQL EQ BASE]," & _
                      "[SCHEDULE SETS].[HEAD 1]     AS [SQL HEAD 1]," & _
                      "[SCHEDULE SETS].[HEAD 2]     AS [SQL HEAD 2]," & _
                      "[SCHEDULE SETS].[HEAD 3]     AS [SQL HEAD 3]," & _
                      "[SCHEDULE SETS].[HEAD 4]     AS [SQL HEAD 4]," & _
                           "[GROUPING].[P1 BASE]    AS [SQL P1 BASE]," & _
                           "[GROUPING].[P2 BASE]    AS [SQL P2 BASE]," & _
                           "[GROUPING].[P3 BASE]    AS [SQL P3 BASE]," & _
                           "[GROUPING].[P4 BASE]    AS [SQL P4 BASE]," & _
                           "[GROUPING].[WORK ORDER] AS [SQL WORK ORDER]," & _
                           "[GROUPING].[LOT NUM]    AS [SQL LOT NUMBER]," & _
                           "[GROUPING].[ATC PART]   AS [SQL ATC PART]," & _
                           "[GROUPING].[P1 BASE]+[GROUPING].[P2 BASE]+[GROUPING].[P3 BASE]+[GROUPING].[P4 BASE] AS [SQL QTY]" & _
               "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
               "WHERE [SCHEDULE SETS].[SET_ID]     = [GROUPING].[SET_ID] AND " & _
                     "[SCHEDULE SETS].[SET_ID]     = [WORK SHEET PT].[SET_ID] AND " & _
                          "[GROUPING].[WORK ORDER] ='" & FR_Table.Fields("[SQL WORK ORDER]") & "'"
        
        Set FR2_Table = FR2_Database.OpenRecordset(sSQL)
 
        If (TO_Table.RecordCount <> 0) Then
            
            If (FR2_Table.RecordCount <> 0) Then
                SET_NUM = FR2_Table.Fields("[SQL SET NUMBER]")
            End If
            Do Until FR2_Table.EOF
                    objExcel.Application.Cells(iRow, 4).Value = FR_Table.Fields("[SQL WORK ORDER]")
                    objExcel.Application.Cells(iRow, 5).Value = FR_Table.Fields("[SQL LOT NUM]")
                    objExcel.Application.Cells(iRow, 6).Value = FR_Table.Fields("[SQL ATC PART]")
                    
                   '090 Plating chg 03/08/2013
                   
                    objExcel.Application.Cells(iRow, 7).Value = FR2_Table.Fields("[SQL EQ BASE]")
            
                    objExcel.Application.Cells(iRow, 20).Value = Format(FR_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
            
                    objExcel.Application.Cells(iRow, 1).Value = FR2_Table.Fields("[SQL DEPT_ID]")
                    objExcel.Application.Cells(iRow, 2).Value = Format(FR2_Table.Fields("[SQL DATE_ID]"), "MM/DD/YYYY")
                    objExcel.Application.Cells(iRow, 3).Value = Format(FR2_Table.Fields("[SQL SCHED DATE_ID]"), "MM/DD/YYYY")
                    
                 '   objExcel.Application.Cells(iRow, 7).Value = FR2_Table.Fields("[SQL QTY]")
                    
                    objExcel.Application.Cells(iRow, 8).Value = FR2_Table.Fields("[SQL SET NUMBER]")
                                      
                    ' QUANTITY
                    PBASE(1) = Val(FR2_Table.Fields("[SQL P1 BASE]"))
                    PBASE(2) = Val(FR2_Table.Fields("[SQL P2 BASE]"))
                    PBASE(3) = Val(FR2_Table.Fields("[SQL P3 BASE]"))
                    PBASE(4) = Val(FR2_Table.Fields("[SQL P4 BASE]"))
                    If PBASE(1) <> 0 Or PBASE(2) <> 0 Then
                            If PBASE(1) <> 0 Then
                                objExcel.Application.Cells(iRow, 9).Value = FR2_Table.Fields("[SQL HEAD 1]")
                            Else
                                objExcel.Application.Cells(iRow, 9).Value = FR2_Table.Fields("[SQL HEAD 2]")
                            End If
                    End If
                    If PBASE(3) <> 0 Or PBASE(4) <> 0 Then
                            If PBASE(3) <> 0 Then
                                objExcel.Application.Cells(iRow, 10).Value = FR2_Table.Fields("[SQL HEAD 3]")
                            Else
                                objExcel.Application.Cells(iRow, 10).Value = FR2_Table.Fields("[SQL HEAD 4]")
                            End If
                    End If
                    objExcel.Application.Cells(iRow, 11).Value = FR2_Table.Fields("[SQL P1 BASE]")
                    objExcel.Application.Cells(iRow, 12).Value = FR2_Table.Fields("[SQL P2 BASE]")
                    objExcel.Application.Cells(iRow, 13).Value = FR2_Table.Fields("[SQL P3 BASE]")
                    objExcel.Application.Cells(iRow, 14).Value = FR2_Table.Fields("[SQL P4 BASE]")
                    
                                       ' SUM GROUPINGS
                    sSQL = "SELECT sum([P1 BASE])+sum([P2 BASE])+sum([P3 BASE])+sum([P4 BASE]) AS [SQL SUM]," & _
                                 "sum([P1 BASE]) AS [SQL SG1],sum([P2 BASE]) AS [SQL SG2]," & _
                                 "sum([P3 BASE]) AS [SQL SG3],sum([P4 BASE]) AS [SQL SG4] " & _
                          "FROM [GROUPING] WHERE [SET_ID]=" & FR2_Table.Fields("[SQL SET_ID]") & " " & _
                          "GROUP BY [SET_ID]"
                    Set FR4_Table = FR4_Database.OpenRecordset(sSQL)
                    
                  '  objExcel.Application.Cells(iRow, 15).Value = FR4_Table.Fields("[SQL SUM]")
                    objExcel.Application.Cells(iRow, 16).Value = FR4_Table.Fields("[SQL SG1]")
                    objExcel.Application.Cells(iRow, 17).Value = FR4_Table.Fields("[SQL SG2]")
                    objExcel.Application.Cells(iRow, 18).Value = FR4_Table.Fields("[SQL SG3]")
                    objExcel.Application.Cells(iRow, 19).Value = FR4_Table.Fields("[SQL SG4]")
                                        
                    FR2_Table.MoveNext
                    If (FR2_Table.EOF = False) Then
                        If (SET_NUM <> FR2_Table.Fields("[SQL SET NUMBER]")) Then
                            iRow = iRow + 1
                            objExcel.Application.Cells(iRow, 21).Value = ROW16
                        End If
                        SET_NUM = FR2_Table.Fields("[SQL SET NUMBER]")
                    End If
            Loop
        End If
                        
        TO_Table.MoveNext
   Loop
End If
                          
FR_Database.Close
FR2_Database.Close
FR3_Database.Close
FR4_Database.Close
TO_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Dim sFile As String

Select Case LOC_SOURCE_ID
Case "NY"
        sFile = "C:\ATC\LIFT TERM NY.xls"
Case "JR"
        sFile = "C:\ATC\LIFT TERM JR.xls"
End Select

objExcel.SaveAs sFile
objExcel.Application.Quit
Set objExcel = Nothing

MsgBox "Excel Download Complete " & sFile, vbInformation, "ATC Plating"

End Sub


Private Sub cmdMix_Click()

' COMPARE [GROUP BY 7 COUNT] TO [GROUP BY 8 COUNT]

Dim COUNT7 As Integer
Dim COUNT8 As Integer

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
                
'[LOT NUM],[ATC PART]
                
sSQL = "SELECT first([ATC PART]) AS [SQL COUNT] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "GROUP BY  mid([ATC PART],5,3) "
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT7 = COUNT7 + 1
        FR_Table.MoveNext
    Loop
End If

sSQL = "SELECT first([ATC PART]) AS [SQL COUNT] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "GROUP BY  mid([ATC PART],5,4)  "
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT8 = COUNT8 + 1
        FR_Table.MoveNext
    Loop
End If

FR_Table.Close
FR_Database.Close

If (COUNT7 = COUNT8) Then
    MsgBox "OK"
Else
    MsgBox "NG"
End If

End Sub

Private Sub cmdNext_Click()

If (optWeek.Value = True) Then
    DTPicker1.Value = DateAdd("WW", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("WW", 1, DTPicker2.Value)
End If

If (optDay.Value = True) Then
    DTPicker1.Value = DateAdd("D", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("D", 1, DTPicker2.Value)
End If

If (optMonth.Value = True) Then
    DTPicker1.Value = DateAdd("M", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("M", 1, DTPicker2.Value)
End If

If (OptionYear.Value = True) Then
    DTPicker1.Value = DateAdd("YYYY", 1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("YYYY", 1, DTPicker2.Value)
End If

'cmdRefresh_Click

End Sub

Private Sub cmdPlatingLoad_Click()
  
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT [SET_ID]    AS [SQL SET_ID]," & _
              "[LETTER_ID] AS [SQL LETTER_ID]" & _
       "FROM [GROUPING] " & _
       "WHERE [WORK ORDER]='" & WO_ID & "' AND [LETTER_ID]<> ' ' "

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    SET_ID = FR_Table.Fields("[SQL SET_ID]")
    LETTER_ID = FR_Table.Fields("[SQL LETTER_ID]")
    MsgBox "Found " & SET_ID & LETTER_ID

    sSQL = "SELECT sum([QTY]) AS [SQL QTY] " & _
           "FROM [GROUPING] " & _
           "WHERE [SET_ID]=" & SET_ID & " AND " & _
              "[LETTER_ID]='" & LETTER_ID & "' " & _
           "GROUP BY [SET_ID],[LETTER_ID] "
    
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    
    MsgBox "Qty " & FR_Table.Fields("[SQL QTY]")
Else
     MsgBox "Not Found"
End If
FR_Table.Close
FR_Database.Close
     
        
End Sub

Private Sub cmdPrevious_Click()

If (optWeek.Value = True) Then
    DTPicker1.Value = DateAdd("WW", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("WW", -1, DTPicker2.Value)
End If

If (optDay.Value = True) Then
    DTPicker1.Value = DateAdd("D", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("D", -1, DTPicker2.Value)
End If

If (optMonth.Value = True) Then
    DTPicker1.Value = DateAdd("M", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("M", -1, DTPicker2.Value)
End If

If (OptionYear.Value = True) Then
    DTPicker1.Value = DateAdd("YYYY", -1, DTPicker1.Value)
    DTPicker2.Value = DateAdd("YYYY", -1, DTPicker2.Value)
End If
'cmdRefresh_Click

End Sub

Private Sub cmdPrintGrouping_Click()

Get_DV
PrintGrouping

End Sub

Private Sub cmdRefresh_Click()

Screen.MousePointer = vbHourglass

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

If (Option16.Value = True) Then
    '16 SET COUNT PER EQ
    sSQL = "SELECT      first([SCHEDULE SETS].[MACHINE_B_ID])       AS [SQL MACHINE_ID]," & _
                       "first([MACHINE].[NUMBER])                   AS [SQL NUMBER]," & _
                       "first([MACHINE].[NAME])                     AS [SQL MNAME]," & _
                       "count([SCHEDULE SETS].[SET_ID])             AS [SQL SET COUNT]," & _
                  "format(sum([SCHEDULE SETS].[RUN QTY]),'###,##0') AS [SQL SUM QTY]" & _
            "FROM [SCHEDULE SETS],[WORK SHEET PT],[MACHINE] " & _
            "WHERE [SCHEDULE SETS].[SET_ID]        = [WORK SHEET PT].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[MACHINE_B_ID]  = [MACHINE].[MACHINE_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [SCHEDULE SETS].[MACHINE_B_ID]"
     
    sSQLF = "    |^|^Number   |^Name            |^Set   Count|>Sum Qty       "
                
End If

If (Option1.Value = True) Then

        ' NEW CALCULATED TABLE Department
        
        cmdCalculation_Click
        
        Select Case LOCATION_ID
        Case "NY"
                sSQL = "SELECT [TBL CALCULATION].[DEPT_ID]," & _
                                    "[DEPT CODE].[DESCRIPTION]," & _
                              "[TBL CALCULATION].[SET COUNT]," & _
                              "[TBL CALCULATION].[WO COUNT]," & _
                              "format([TBL CALCULATION].[SUM QTY],'###,##0')," & _
                              "format([TBL CALCULATION].[PROCESS TIME],'###,##0.0')," & _
                              "format([TBL CALCULATION].[OEE TIME],'###,##0.0')" & _
                        "FROM [TBL CALCULATION],[DEPT CODE] " & _
                        "WHERE [TBL CALCULATION].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                              "[DEPT CODE].[ACTIVE] = 1 "
        Case "JR"
                sSQL = "SELECT [TBL CALCULATION].[DEPT_ID]," & _
                                    "[DEPT CODE].[DESCRIPTION]," & _
                              "[TBL CALCULATION].[SET COUNT]," & _
                              "[TBL CALCULATION].[WO COUNT]," & _
                              "format([TBL CALCULATION].[SUM QTY],'###,##0')," & _
                              "format([TBL CALCULATION].[PROCESS TIME],'###,##0.0')," & _
                              "format([TBL CALCULATION].[OEE TIME],'###,##0.0')" & _
                        "FROM [TBL CALCULATION],[DEPT CODE] " & _
                        "WHERE [TBL CALCULATION].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                              "[DEPT CODE].[ACTIVE] = 1 "
                              
        End Select

        sSQLF = "    |^DEPT    |^                                  |>Count Set|>WO Count|>Sum Qty       |>Process (Hr)    |>OEE Time (Hr) "
         
End If

If (Option2.Value = True) Then

        cmdCalculationEQ_Click
        
        cmdCalculationEQFinish_Click
  
        ' BASE AND FINISH EQ
               sSQL = "SELECT [TBL CALCULATION EQ].[NUMBER]," & _
                                        "[MACHINE].[PROCESS]," & _
                                        "[MACHINE].[TYPE]," & _
                             "[TBL CALCULATION EQ].[SET COUNT]," & _
                             "[TBL CALCULATION EQ].[WO COUNT]," & _
                      "format([TBL CALCULATION EQ].[SUM QTY],'###,##0')," & _
                      "format([TBL CALCULATION EQ].[PROCESS TIME],'###,##0.0')," & _
                      "format([TBL CALCULATION EQ].[OEE TIME],'###,##0.0')" & _
                "FROM [TBL CALCULATION EQ],[MACHINE] " & _
                "WHERE [TBL CALCULATION EQ].[NUMBER] = [MACHINE].[NUMBER] AND [MACHINE].[DEPT_ID]='PT' " & _
                "ORDER BY [PROCESS],[TYPE]"

        sSQLF = "    |^EQ# |^Process    |^Type            |>Count Set|>WO Count|>Sum Qty       |>Process (Hr)         |>OEE Time (Hr)"
        
End If

If (Option4.Value = True) Then

        '4 Series
        sSQL = "SELECT first(mid([GROUPING].[ATC PART],1,4))," & _
                      "count([SCHEDULE SETS].[SET_ID])," & _
                      "format(sum([GROUPING].[QTY]),'###,##0') " & _
                "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
                "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                "GROUP BY mid([GROUPING].[ATC PART],1,4) "
         
        sSQLF = "    |^SERIES|^Count Set|>Sum Qty       "

End If

If (Option5.Value = True) Then

        Select Case LOCATION_ID
        Case "NY"
                '5 CASE SIZE
                sSQL = "SELECT first(mid([GROUPING].[ATC PART],4,1))    AS [SQL CASE SIZE]," & _
                              "first([WORK SHEET PT].[DATE_ID])         AS [SQL DATE]," & _
                              "count([SCHEDULE SETS].[SET_ID])          AS [SQL SET COUNT]," & _
                              "format(sum([GROUPING].[QTY]),'###,##0')  AS [SQL QTY] " & _
                        "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
                        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                              "[SCHEDULE SETS].[SET_ID]  = [GROUPING].[SET_ID] AND " & _
                              "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                              "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                        "GROUP BY mid([GROUPING].[ATC PART],4,1) "
         Case "JR"
                '5 CASE SIZE
                sSQL = "SELECT first(mid([GROUPING].[ATC PART],4,1))    AS [SQL CASE SIZE]," & _
                              "first([WORK SHEET PT].[DATE_ID])         AS [SQL DATE]," & _
                              "count([SCHEDULE SETS].[SET_ID])          AS [SQL SET COUNT]," & _
                              "format(sum([GROUPING].[QTY]),'###,##0')  AS [SQL QTY] " & _
                        "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING],[WORK SHEET PT] " & _
                        "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                              "[SCHEDULE SETS].[SET_ID]  = [GROUPING].[SET_ID] AND " & _
                              "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                              "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                              "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                        "GROUP BY mid([GROUPING].[ATC PART],4,1)  "
        
        
        End Select
        
        sSQLF = "    |^Case Size|^Date                    |^Count Set|>Sum Qty       "

End If

If (Option12.Value = True) Then
        Select Case LOCATION_ID
        Case "NY"
        sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID])," & _
                          "first([DEPT CODE].[DESCRIPTION])," & _
                      "first([SCHEDULE SETS].[SERIES_ID])," & _
                       "first(mid([GROUPING].[ATC PART],5,3))," & _
                           "count([GROUPING].[QTY])," & _
                      "format(sum([GROUPING].[QTY]),'###,##0') " & _
              "FROM [SCHEDULE SETS],[GROUPING],[DEPT CODE],[WORK SHEET PT] " & _
              "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                    "[SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                    "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                    "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[TYPE_ID] IN ('BARREL') AND [LETTER_ID] BETWEEN 'A' AND 'Z' " & _
              "GROUP BY [SCHEDULE SETS].[DEPT_ID],[SCHEDULE SETS].[SERIES_ID],mid([GROUPING].[ATC PART],5,3) "
        Case "JR"
        sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID])," & _
                          "first([DEPT CODE].[DESCRIPTION])," & _
                      "first([SCHEDULE SETS].[SERIES_ID])," & _
                       "first(mid([GROUPING].[ATC PART],5,3))," & _
                           "count([GROUPING].[QTY])," & _
                      "format(sum([GROUPING].[QTY]),'###,##0') " & _
              "FROM [SCHEDULE SETS],[GROUPING],[DEPT CODE],[WORK SHEET PT] " & _
              "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                    "[SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_JR_ID] AND " & _
                    "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                    "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[TYPE_ID] IN ('BARREL') AND [LETTER_ID] BETWEEN 'A' AND 'Z' " & _
              "GROUP BY [SCHEDULE SETS].[DEPT_ID],[SCHEDULE SETS].[SERIES_ID],mid([GROUPING].[ATC PART],5,3) "
        End Select
        sSQLF = "    |^Code_ID|^Code Description    |^SERIES|^Value    |^W.O. Count|>Sum Qty        "
         
End If

If (Option13.Value = True) Then
         '13 ERRORS
        sSQL = "SELECT [SCHEDULE SETS].[DATE_ID],[DEPT_ID],[WORK ORDER]," & _
                      "mid([GROUPING].[ATC PART],4,1)," & _
                      "[SCHEDULE SETS].[SET NUMBER]," & _
                      "format([GROUPING].[QTY],'###,##0') " & _
                "FROM [SCHEDULE SETS],[GROUPING]" & _
                "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "mid([GROUPING].[ATC PART],4,1) NOT IN ('A','B','C','E','R')"
         
        sSQLF = "   |^Schedule Set Date|^Dept_ID  |^Work Order/Lot Number |^SERIES|^Count #|>Sum Qty       "

End If

If (Option10.Value = True) Then

        '10 Code,Series,Barrel
        
        sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID])," & _
                    "first([SCHEDULE SETS].[SERIES_ID])," & _
                    "count([GROUPING].[QTY])," & _
                    "format(sum([GROUPING].[QTY]),'###,##0')," & _
                    "format(sum([BASE AMP MIN])/(sum([BASE AMP]) + 0.1),'###,##0')     AS [SQL BT]," & _
                    "format(sum([FINISH AMP MIN])/(sum([FINISH AMP]) + 0.1),'###,##0') AS [SQL FT], " & _
                    "format(val([SQL BT])+val([SQL FT]),'###,##0') " & _
              "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
              "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                    "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                    "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
              "GROUP BY [SCHEDULE SETS].[DEPT_ID],[SCHEDULE SETS].[SERIES_ID] "
        
        sSQLF = "    |^Code_ID|^SERIES|^W.O. Count|>Sum Qty       |BT    |FT      |TT  (m) "
End If
If (Option11.Value = True) Then

        '11 Code,Series,SBE
        
        sSQL = "SELECT first([SCHEDULE SETS].[DEPT_ID])," & _
                    "first([SCHEDULE SETS].[SERIES_ID])," & _
                    "count([GROUPING].[QTY])," & _
                    "format(sum([GROUPING].[QTY]),'###,##0')," & _
                    "format(sum([BASE AMP MIN])/(sum([BASE AMP]) + 0.1),'###,##0')     AS [SQL BT]," & _
                    "format(sum([FINISH AMP MIN])/(sum([FINISH AMP]) + 0.1),'###,##0') AS [SQL FT], " & _
                    "format(val([SQL BT])+val([SQL FT]),'###,##0') " & _
              "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
              "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                    "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                    "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[TYPE_ID] IN ('SBE') " & _
              "GROUP BY [SCHEDULE SETS].[DEPT_ID],[SCHEDULE SETS].[SERIES_ID] "
        
        sSQLF = "    |^Code_ID|^SERIES|^W.O. Count|>Sum Qty       |BT    |FT      |TT (hr)  "
End If


If (Option6.Value = True) Then
    '6 SET COUNT
    sSQL = "SELECT      first([WORK SHEET PT].[DATE_ID])            AS [SQL DATE_ID]," & _
                       "count([SCHEDULE SETS].[SET_ID])             AS [SQL SET COUNT]," & _
                  "format(sum([SCHEDULE SETS].[RUN QTY]),'###,##0') AS [SQL SUM QTY]" & _
            "FROM [SCHEDULE SETS],[WORK SHEET PT] " & _
            "WHERE [SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [WORK SHEET PT].[DATE_ID]"
     
    sSQLF = "    |^Date ID         |^Set   Count|>Sum Qty       "
                
End If



If (Option7.Value = True) Then
     ' 7 OEE Codes
      sSQL = "SELECT first([WORK SHEET PT].[DATE_ID])," & _
                    "first([WORK SHEET PT].[CODE_ID])," & _
                    "count([WORK SHEET PT].[CODE_ID]) " & _
            "FROM [WORK SHEET PT],[SCHEDULE SETS] " & _
            "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [WORK SHEET PT].[DATE_ID],[WORK SHEET PT].[CODE_ID]          "
    
    sSQLF = "    |^Date ID         |^CODE_ID|>Count           "

End If
If (Option8.Value = True) Then
    '8 WORK ORDER COUNT
    sSQL = "SELECT first([WORK SHEET PT].[DATE_ID])       AS [SQL DATE_ID]," & _
                    "count([GROUPING].[SET_ID])           AS [SQL WO COUNT]," & _
               "format(sum([GROUPING].[QTY]),'###,##0')   AS [SQL SUM QTY]" & _
            "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT]" & _
            "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [WORK SHEET PT].[DATE_ID]"
     
    sSQLF = "    |^Date ID         |^WO   Count|>Sum Qty       "
        
End If
If (Option15.Value = True) Then
    '15 SBE/BARREL TOTAL

    sSQL = "SELECT      first([SCHEDULE SETS].[TYPE_ID])            AS [SQL DATE_ID]," & _
                       "count([SCHEDULE SETS].[SET_ID])             AS [SQL SET COUNT]," & _
                  "format(sum([SCHEDULE SETS].[RUN QTY]),'###,##0') AS [SQL SUM QTY]" & _
            "FROM [SCHEDULE SETS],[WORK SHEET PT] " & _
            "WHERE [SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
            "GROUP BY [WORK SHEET PT].[CODE_ID],[SCHEDULE SETS].[TYPE_ID]"
     
    sSQLF = "    |^TYPE ID         |^Set   Count|>Sum Qty       "
                                              
End If

If (Option9.Value = True) Then
        
       Select Case LOCATION_ID
       Case "NY", "JR"
        '9 BASE EQ
        sSQL = "SELECT first([SCHEDULE SETS].[EQ BASE])," & _
                      "first([MACHINE].[DESCRIPTION])," & _
                      "count([SCHEDULE SETS].[SET_ID])," & _
                     "format(sum([RUN QTY]),'###,##0') " & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT],[MACHINE] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID]  = [WORK SHEET PT].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[EQ BASE] = [MACHINE].[NUMBER] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[DEPT CODE].[DEPT_ID] NOT IN (555,556,557,553,554,558)" & _
                "GROUP BY [SCHEDULE SETS].[EQ BASE]"
            
        End Select
        
        sSQLF = "    |^Tank     |<Description             |^Count Set|>Sum Qty       "
End If
If (Option3.Value = True) Then
      '3 Finish EQ
      Select Case LOCATION_ID
      Case "NY"
      
        sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])," & _
                      "first([MACHINE].[DESCRIPTION])," & _
                      "count([SCHEDULE SETS].[SET_ID])," & _
                     "format(sum([RUN QTY]),'###,##0') " & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT],[MACHINE] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID]= [DEPT CODE].[DEPT_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[EQ FINISH] = [MACHINE].[NUMBER] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
      Case "JR"
      
        sSQL = "SELECT first([SCHEDULE SETS].[EQ FINISH])," & _
                      "first([MACHINE].[DESCRIPTION])," & _
                      "count([SCHEDULE SETS].[SET_ID])," & _
                     "format(sum([RUN QTY]),'###,##0') " & _
                "FROM [SCHEDULE SETS],[DEPT CODE],[WORK SHEET PT],[MACHINE] " & _
                "WHERE [SCHEDULE SETS].[DEPT_ID]= [DEPT CODE].[DEPT_JR_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[EQ FINISH] = [MACHINE].[NUMBER] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
      
      End Select
      
        sSQLF = "    |^Tank     |<Description             |^Count Set|>Sum Qty       "
End If
 

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

Screen.MousePointer = vbDefault

End Sub



Private Sub cmdReset_Click()
DTPicker1.Value = Date

If (optWeek.Value = True) Then
    DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
    DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
   DTPicker1.Value = DTPicker1.Value
   DTPicker2.Value = DTPicker1.Value
End If

cmdRefresh_Click

End Sub


Private Sub cmdTest_Click()

Dim iFilenum As Integer
iFilenum = FreeFile

Dim sFilename As String
sFilename = "C:\ATC\PLATING TEST COUNT.TXT"

Open sFilename For Append Shared As #iFilenum

Dim TOTAL(10) As Integer
Dim COUNT As Integer
Dim ROW As Integer

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)
'[LOT NUM],[ATC PART]
                
sSQL = "SELECT first([SET_ID]) as [SQL SET_ID]," & _
              "first([LETTER_ID]) AS [SQL LETTER_ID]  " & _
       "FROM [GROUPING] WHERE [LETTER_ID]<> ' '" & _
       "GROUP BY [SET_ID]&[LETTER_ID]"
       
Set TO_Table = TO_Database.OpenRecordset(sSQL)

TO_Table.MoveLast
COUNT = TO_Table.RecordCount
TO_Table.MoveFirst
COUNT = 0
Do Until TO_Table.EOF
                
        SET_ID = TO_Table.Fields("[SQL SET_ID]")
        LETTER_ID = TO_Table.Fields("[SQL LETTER_ID]")
                  
        Select Case Calculate_Plating_Table
        Case 0 To 10
        TOTAL(Calculate_Plating_Table) = TOTAL(Calculate_Plating_Table) + 1
        Case Else
                    COUNT = COUNT + 1
                    'Print #iFilenum, SET_ID, LETTER_ID, Calculate_Plating_Table
        End Select
       ROW = ROW + 1
       cmdTest.Caption = ROW
       DoEvents
       TO_Table.MoveNext
Loop

TO_Table.Close
TO_Database.Close
Dim I As Integer
For I = 0 To 10
    Print #iFilenum, I, TOTAL(I)
Next I

Close iFilenum
MsgBox "Complete " & COUNT

End Sub

Private Sub cmdTests_Click()

SET_ID = Trim(txtSET_ID.Text)
LETTER_ID = Trim(txtLETTER_ID.Text)

frmMSGBOX.Show
 
End Sub

Private Sub cmdWOSearch_Click()

Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)
     
Dim sSQL As String

sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & WO_ID & "'"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    If (FR_Table.Fields("[ATC PART]") <> vbNull) Then
            ATC_PART_ID = FR_Table.Fields("[ATC PART]")
    End If
    If (FR_Table.Fields("[LOT NUM]") <> vbNull) Then
            LOT_ID = FR_Table.Fields("[LOT NUM]")
    End If
    
    Select Case ATC_PART_ID
    Case "ATC PT"
            MsgBox "Not Found"
            Exit Sub
    Case Else
    
    End Select
    Select Case LOT_ID
    Case "LOT NUM"
            MsgBox "Not Found"
            Exit Sub
    Case Else
    
    End Select
    MsgBox "Found" & vbNewLine & "LOT_ID : " & LOT_ID & vbNewLine & "ATC_PART_ID : " & ATC_PART_ID
Else
    MsgBox "Not Found"
End If

End Sub



Private Sub Command1_Click()
LOC_SOURCE_ID = "NY"
cmdLift_Click
End Sub

Private Sub Command2_Click()
LOC_SOURCE_ID = "JR"
cmdLift_Click
End Sub

Private Sub Command3_Click()

SET_ID = Trim(txtSET_ID.Text)
LETTER_ID = Trim(txtLETTER_ID.Text)

Screen.MousePointer = vbHourglass
Dim sSQL As String
Dim sSQLF As String

Select Case 0
Case 0
sSQL = "SELECT [GP_ID],[SET_ID],[WORK ORDER],[LOT NUM],[ATC PART]," & _
              "format([QTY],'###,####'),format([DV],'###,###0.0')," & _
              "[LETTER_ID] " & _
       "FROM [GROUPING] " & _
       "WHERE [LETTER_ID]='" & LETTER_ID & "' AND [SET_ID]=" & SET_ID & " " & _
       "ORDER BY [DV] ASC"
                                                                                                          
sSQLF = "    ||^SET_ID|^Work Order/Lot    |^Lot Number        |<ATC Part                         |>QTY      |>Design Value    |^Run  "
        
End Select

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF
Screen.MousePointer = vbDefault

End Sub

Private Sub CommandSQL_Click()
Screen.MousePointer = vbHourglass

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

'"[MACHINE].[TYPE]," & _

'[TYPE_ID]

        '4 Series
        sSQL = "SELECT first([SCHEDULE SETS].[TYPE_ID])," & _
                      "first(mid([GROUPING].[ATC PART],1,4))," & _
                      "count([SCHEDULE SETS].[SET_ID])," & _
                      "format(sum([GROUPING].[QTY]),'###,##0') " & _
                "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT] " & _
                "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                      "[SCHEDULE SETS].[SET_ID] = [WORK SHEET PT].[SET_ID] AND " & _
                      "[WORK SHEET PT].[CODE_ID] = 500 AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
                "GROUP BY [SCHEDULE SETS].[TYPE_ID],mid([GROUPING].[ATC PART],1,4) "
         
        sSQLF = "    |^BARREL/SBE  |^SERIES|^Count Set|>Sum Qty       "


Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

Screen.MousePointer = vbDefault


End Sub

Private Sub Form_Load()

Caption = "Summary Review Plating   " & ATC_DWG & "    " & ATC_VERSION

'DB_PLATING_TERMINATION = SERVER_DB_NY & "Plating JR 09-20-2017.MDB"

Data1.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11500
MSFlexGrid1.Height = Me.Height - 800

DTPicker1.Value = Date

If (optWeek.Value = True) Then
    DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
    DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")
End If

If (optDay.Value = True) Then
   DTPicker1.Value = DTPicker1.Value
   DTPicker2.Value = DTPicker1.Value
End If

cmdRefresh_Click

End Sub


Private Sub MSFlexGrid1_Click()

'[DEPT_ID]

Dim sSQL As String
Dim sSQLF As String
If (Option4.Value = True) Then
         
    sSQL = "SELECT mid([GROUPING].[ATC PART],1,4)," & _
                  "[GROUPING].[ATC PART]," & _
                  "[SCHEDULE SETS].[DEPT_ID]," & _
                  "[SCHEDULE SETS].[DATE_ID]," & _
                  "[SCHEDULE SETS].[SET NUMBER]," & _
                  "format([GROUPING].[QTY],'###,##0') " & _
            "FROM [SCHEDULE SETS],[GROUPING] " & _
            "WHERE [SCHEDULE SETS].[SET_ID] = [GROUPING].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                  "mid([GROUPING].[ATC PART],1,3) NOT IN ('100','700','710','200','900','10E','70E','71E')"
     
    sSQLF = "    |^SERIES|ATC Part         |^DEPT_ID |^DATE_ID            |^SET_ID|>Qty        "
    
    Data1.RecordSource = sSQL
    Data1.Refresh
    
    MSFlexGrid1.FormatString = sSQLF

End If

If (Option5.Value = True) Then
         
    sSQL = "SELECT mid([GROUPING].[ATC PART],4,1)," & _
                  "[GROUPING].[ATC PART]," & _
                  "[SCHEDULE SETS].[DEPT_ID]," & _
                  "[SCHEDULE SETS].[DATE_ID]," & _
                  "[SCHEDULE SETS].[SET NUMBER] " & _
            "FROM [SCHEDULE SETS],[DEPT CODE],[GROUPING] " & _
            "WHERE [SCHEDULE SETS].[DEPT_ID] = [DEPT CODE].[DEPT_ID] AND " & _
                  "[SCHEDULE SETS].[SET_ID]  = [GROUPING].[SET_ID] AND " & _
                  "[SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                  "mid([GROUPING].[ATC PART],4,1) NOT IN ('A','B','C','E','R')"
             
    sSQLF = "    |^SERIES|ATC Part         |^DEPT_ID |^DATE_ID            |^SET_ID"
            
    Data1.RecordSource = sSQL
    Data1.Refresh
    
    MSFlexGrid1.FormatString = sSQLF

End If


If (Option14.Value = True) Then

    MSFlexGrid1.Col = 2
    SET_ID = Val(MSFlexGrid1.Text)
    txtSET_ID.Text = SET_ID

    MSFlexGrid1.Col = 3
    WO_ID = MSFlexGrid1.Text

    MSFlexGrid1.Col = 8
    LETTER_ID = UCase(MSFlexGrid1.Text)
    txtLETTER_ID.Text = LETTER_ID
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

End If

End Sub

Private Sub optDay_Click()

DTPicker1.Value = Format(Date, "mm/dd/yyyy")
DTPicker2.Value = Format(Date, "mm/dd/yyyy")

cmdPrevious.Caption = "Day  <<"
cmdNext.Caption = "Day  >>"

End Sub

 


Private Sub Option14_Click()

Screen.MousePointer = vbHourglass
Dim sSQL As String
Dim sSQLF As String

Select Case 0
Case 0
        sSQL = "SELECT [GP_ID],[SET_ID],[WORK ORDER],[LOT NUM],[ATC PART]," & _
                      "format([QTY],'###,####'),format([DV],'###,###0.0')," & _
                      "[LETTER_ID] " & _
               "FROM [GROUPING] WHERE [LETTER_ID]<> ' ' " & _
               "ORDER BY [SET_ID] DESC,[LETTER_ID] ASC,[DV] ASC"
                                                                                                                  
        sSQLF = "    ||^SET_ID|^Work Order/Lot    |^Lot Number        |<ATC Part                         |>QTY      |>Design Value    |^Run  "

Case 1
        sSQL = "SELECT COUNT([GP_ID]),FIRST([SET_ID]&[LETTER_ID]) " & _
               "FROM [GROUPING] WHERE   [LETTER_ID]<> ' ' " & _
               "GROUP BY [SET_ID]&[LETTER_ID] ORDER BY COUNT([GP_ID]) DESC"
        
        sSQLF = "    |            |"

Case 2
        sSQL = "SELECT first(mid([ATC PART],8,1)),count(mid([ATC PART],8,1)) " & _
               "FROM [GROUPING] WHERE   [LETTER_ID]<> ' ' AND " & _
                        "mid([ATC PART],1,3) IN ('100','200','700','710','800','830','900','10E','70E','71E')" & _
               "GROUP BY mid([ATC PART],8,1) ORDER BY first(mid([ATC PART],8,1)) DESC"
        
        sSQL = "SELECT MID([ATC PART],1,8)  " & _
               "FROM [GROUPING] WHERE   [LETTER_ID]<> ' ' AND " & _
                        "mid([ATC PART],1,3) IN ('100','200','700','710','800','830','900','10E','70E','71E')" & _
               "ORDER BY mid([ATC PART],8,1) DESC"
        
        sSQLF = "    |                         |Count            "
End Select

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF
Screen.MousePointer = vbDefault

End Sub


Private Sub OptionYear_Click()
cmdPrevious.Caption = "Year  <<"
cmdNext.Caption = "Year  >>"

DTPicker1.Value = Format(Date, "1/1/yyyy")
DTPicker2.Value = Format(DateAdd("yyyy", 1, DTPicker1.Value), "mm/dd/yyyy")

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

Private Sub txtWorkOrder_GotFocus()
txtWorkOrder.SelStart = 0
txtWorkOrder.SelLength = Len(txtWorkOrder)
End Sub

Private Sub txtLETTER_ID_GotFocus()
txtLETTER_ID.SelStart = 0
txtLETTER_ID.SelLength = Len(txtLETTER_ID)
End Sub

Private Sub txtLETTER_ID_LostFocus()
txtLETTER_ID.Text = UCase(txtLETTER_ID.Text)
End Sub

Private Sub txtSET_ID_GotFocus()
txtSET_ID.SelStart = 0
txtSET_ID.SelLength = Len(txtSET_ID)
End Sub
