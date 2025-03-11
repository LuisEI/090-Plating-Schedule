VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSBEParameters 
   Caption         =   "090 SBE Plating Parameter Tables DWG NO. 115-144"
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19140
   Icon            =   "090 SBE Parameters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   19140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   11640
      TabIndex        =   25
      Top             =   7320
      Width           =   3135
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         DataField       =   "CERAMIC"
         DataSource      =   "Data3"
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
         TabIndex        =   60
         Top             =   3840
         Width           =   800
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         DataField       =   "LOT CODE"
         DataSource      =   "Data3"
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
         TabIndex        =   59
         Top             =   3840
         Width           =   800
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         DataField       =   "Min4"
         DataSource      =   "Data3"
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
         TabIndex        =   47
         Top             =   3360
         Width           =   800
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         DataField       =   "ASF4"
         DataSource      =   "Data3"
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
         TabIndex        =   46
         Top             =   3360
         Width           =   800
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         DataField       =   "Min3"
         DataSource      =   "Data3"
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
         TabIndex        =   45
         Top             =   2940
         Width           =   800
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         DataField       =   "ASF3"
         DataSource      =   "Data3"
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
         TabIndex        =   44
         Top             =   2940
         Width           =   800
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         DataField       =   "ASF3"
         DataSource      =   "Data3"
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
         TabIndex        =   35
         Top             =   2100
         Width           =   800
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         DataField       =   "Min3"
         DataSource      =   "Data3"
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
         TabIndex        =   34
         Top             =   2100
         Width           =   800
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         DataField       =   "ASF4"
         DataSource      =   "Data3"
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
         TabIndex        =   33
         Top             =   2520
         Width           =   800
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         DataField       =   "Min4"
         DataSource      =   "Data3"
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
         TabIndex        =   32
         Top             =   2520
         Width           =   800
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "UpdateRecord"
         Height          =   360
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         DataField       =   "Case"
         DataSource      =   "Data3"
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
         Left            =   960
         TabIndex        =   30
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         DataField       =   "Min2"
         DataSource      =   "Data3"
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
         TabIndex        =   29
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         DataField       =   "ASF2"
         DataSource      =   "Data3"
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
         TabIndex        =   28
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         DataField       =   "Min1"
         DataSource      =   "Data3"
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
         TabIndex        =   27
         Top             =   1260
         Width           =   800
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         DataField       =   "ASF1"
         DataSource      =   "Data3"
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
         TabIndex        =   26
         Top             =   1260
         Width           =   800
      End
      Begin VB.Label Label11 
         Caption         =   "MAX"
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
         Left            =   2160
         TabIndex        =   56
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label10 
         Caption         =   "MIN"
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
         Left            =   1320
         TabIndex        =   55
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "CASE"
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
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label8 
         Caption         =   "RP ASF"
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
         Left            =   240
         TabIndex        =   53
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "LW ASF"
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
         Left            =   240
         TabIndex        =   52
         Top             =   2100
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "TN ASF"
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
         Left            =   240
         TabIndex        =   51
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "NI ASF"
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
         Left            =   240
         TabIndex        =   50
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "DV"
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
         Left            =   240
         TabIndex        =   49
         Top             =   2940
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "QTY"
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
         Left            =   240
         TabIndex        =   48
         Top             =   3360
         Width           =   780
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [TBL SBE]"
      Connect         =   "Access"
      DatabaseName    =   "\\Ny-eng\spc network\Data Base\PLATING JR.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL SBE"
      Top             =   3600
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Frame FrameSBE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   5655
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "QTY"
         DataSource      =   "Data1"
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
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "CHIP VOL"
         DataSource      =   "Data1"
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
         TabIndex        =   6
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "MEDIA VOL"
         DataSource      =   "Data1"
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
         TabIndex        =   7
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "MEDIA SF"
         DataSource      =   "Data1"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         DataField       =   "CASE SIZE"
         DataSource      =   "Data1"
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
         TabIndex        =   4
         Top             =   240
         Width           =   405
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "UpdateRecord"
         Height          =   360
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
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
      Height          =   1575
      Left            =   6120
      TabIndex        =   22
      Top             =   120
      Width           =   8415
      Begin VB.OptionButton Option800AB 
         Caption         =   "800 A/B/R    JAX"
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
         Left            =   4200
         TabIndex        =   40
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton Option600SFL 
         Caption         =   "600 S/F/L     JAX"
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
         TabIndex        =   39
         Top             =   960
         Width           =   2775
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
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
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
         TabIndex        =   11
         Top             =   360
         Width           =   2055
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
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   1335
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
         Left            =   5880
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdTBL_SBE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FROM [TBL SBE] "
      Height          =   300
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [TBL SBE 144]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\PLATING.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL SPE 144"
      Top             =   5520
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.CommandButton cmdSPE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "<< FROM [TBL SBE 144]"
      Height          =   300
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 FROM [TBL SBE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [TBL SBE 144]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   4380
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
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton optCaseR 
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
         Left            =   4680
         TabIndex        =   57
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton OptionJAXR 
         Caption         =   "R (800R)"
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
         TabIndex        =   43
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton OptionJAXB 
         Caption         =   "B (800B)"
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
         Left            =   1800
         TabIndex        =   42
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton OptionJAXA 
         Caption         =   "A (800A)"
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
         TabIndex        =   41
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton OptionCaseF 
         Caption         =   "F (600F)"
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
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OptionCaseL 
         Caption         =   "L (600L)"
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
         Left            =   1800
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OptionCaseS 
         Caption         =   "S (600S)"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   960
         Width           =   1335
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   855
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
         Left            =   1560
         TabIndex        =   1
         Top             =   360
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
         Left            =   2520
         TabIndex        =   2
         Top             =   360
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
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [PCS PER SIDE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   4020
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 SBE Parameters.frx":0CCA
      Height          =   3255
      Left            =   6120
      TabIndex        =   16
      ToolTipText     =   "[PCS PER SIDE]"
      Top             =   8520
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5741
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
      Bindings        =   "090 SBE Parameters.frx":0CDE
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "[TBL SBE 144]"
      Top             =   3480
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
      Bindings        =   "090 SBE Parameters.frx":0CF2
      Height          =   1335
      Left            =   6120
      TabIndex        =   15
      ToolTipText     =   "[TBL SBE]"
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      _Version        =   393216
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "ATC Plating Tables.MDB"
      Height          =   195
      Left            =   9840
      TabIndex        =   58
      Top             =   12240
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "<< FROM [TBL SBE 144]"
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
      Left            =   6840
      TabIndex        =   23
      Top             =   12000
      Width           =   2340
   End
   Begin VB.Label Label15 
      Caption         =   "Lookup FROM [TBL SBE] * Sum SA"
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
      Left            =   6240
      TabIndex        =   19
      Top             =   1680
      Width           =   4500
   End
   Begin VB.Label Label7 
      Caption         =   "FROM [PCS PER SIDE]"
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
      Left            =   6120
      TabIndex        =   18
      Top             =   8160
      Width           =   3300
   End
End
Attribute VB_Name = "frmSBEParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()

Dim Qty As Long
Qty = Val(txtQTY.Text)

If (Qty = 0) Then
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

If (Option1.Value = True) Then
    SERIES_ID = 100
End If
If (Option2.Value = True) Then
    SERIES_ID = 200
End If

'================================================================================
'   [1]  SURFACE AREA PART
'================================================================================
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String

sSQL = "SELECT [CASE],[PCS PER SIDE MAX],[SHOT],[SF] " & _
       "FROM [PCS PER SIDE] " & _
       "WHERE [CASE] ='" & CASE_SIZE_ID & "'"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim SF As Double
SF = FR_Table.Fields("[SF]")

lblPartSA.Caption = Format(FR_Table.Fields("[SF]") * Qty, "0.000")

Dim PART_SA As Double

PART_SA = (FR_Table.Fields("[SF]") * Qty) / 144

'================================================================================
'   [2]  MEDIA SURFACE AREA
'================================================================================

sSQL = "SELECT [QTY],[MEDIA SF] " & _
       "FROM [TBL SBE 144] " & _
       "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & Qty & " ORDER BY [QTY] DESC"
Set FR_Table = FR_Database.OpenRecordset(sSQL)

lblMediaSA.Caption = FR_Table.Fields("[MEDIA SF]")

Dim Media_SA As Double
Media_SA = FR_Table.Fields("[MEDIA SF]")

Dim SA_Sum As Double

SA_Sum = PART_SA + Media_SA

lblSumSA.Caption = Format(SA_Sum, "0.000")

'================================================================================
'   [3]  AMPS/MIN
'================================================================================
       
sSQL = "SELECT * FROM [TBL SBE] WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] =" & SERIES_ID
                                       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

lblNickelAmps.Caption = Format(SA_Sum * FR_Table.Fields("[ASF1]"), "0.0")
lblTinAmps.Caption = Format(SA_Sum * FR_Table.Fields("[ASF2]"), "0.0")
lblSolderAmps.Caption = Format(SA_Sum * FR_Table.Fields("[ASF3]"), "0.0")

lblNickelMin.Caption = Format(SA_Sum * FR_Table.Fields("[MIN1]") * FR_Table.Fields("[ASF1]") / 60, "0.0")
lblTinMin.Caption = Format(SA_Sum * FR_Table.Fields("[MIN2]") * FR_Table.Fields("[ASF2]") / 60, "0.0")
lblSolderMin.Caption = Format(SA_Sum * FR_Table.Fields("[MIN3]") * FR_Table.Fields("[ASF3]") / 60, "0.0")

lblRPAmps.Caption = Format(SA_Sum * FR_Table.Fields("[ASF4]"), "0.0")
lblRPMin.Caption = Format(SA_Sum * FR_Table.Fields("[MIN4]") * FR_Table.Fields("[ASF4]") / 60, "0.0")

'
'   REPLATING BARREL CODE_ID 555
'
' FR_Table.Fields("[MEDIA SF]")
Select Case SERIES_ID
Case 100  '100,710,700
            Media_SA = 7.125
            '[MEDIA VOL] = 150
Case 200  '200,900
            Media_SA = 9.5
             '[MEDIA VOL] = 200
End Select

SA_Sum = PART_SA + Media_SA

lbl555Amps.Caption = Format(SA_Sum * FR_Table.Fields("[ASF4]"), "0.0")
lbl555Min.Caption = Format(SA_Sum * FR_Table.Fields("[MIN4]") * FR_Table.Fields("[ASF4]") / 60, "0.0")

FR_Table.Close
FR_Database.Close

End Sub

Private Sub cmdSPE_Click()

TABLE_ID = 0

If (optCaseA.Value = True) Then
    CASE_SIZE_ID = "A"
End If
If (optCaseB.Value = True) Then
    CASE_SIZE_ID = "B"
End If
If (optCaseR.Value = True) Then
    CASE_SIZE_ID = "R"
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

If (OptionJAXA.Value = True) Then
    CASE_SIZE_ID = "A"
    TABLE_ID = 1
End If
If (OptionJAXB.Value = True) Then
    CASE_SIZE_ID = "B"
    TABLE_ID = 1
End If
If (OptionJAXR.Value = True) Then
    CASE_SIZE_ID = "R"
    TABLE_ID = 1
End If

Dim sSQL As String

Select Case TABLE_ID
Case 0
        sSQL = "SELECT [ID],format([QTY],'#,###,###')," & _
                               "[CHIP VOL]," & _
                               "[MEDIA VOL]," & _
                        "format([MEDIA SF],'0.000')," & _
                               "[FLOW] " & _
               "FROM [TBL SBE 144] " & _
               "WHERE [CASE SIZE]='" & CASE_SIZE_ID & "' " & _
               "ORDER BY [QTY] DESC"
       
Case 1
        sSQL = "SELECT [ID],format([QTY],'#,###,###')," & _
                                "[CHIP VOL]," & _
                                "[MEDIA VOL]," & _
                         "format([MEDIA SF],'0.000')," & _
                                "[FLOW] " & _
                "FROM [TBL SBE ABR JAX] WHERE [CASE SIZE]='" & CASE_SIZE_ID & "' ORDER BY [QTY] DESC"

End Select
                                   
Data4.RecordSource = sSQL
Data4.Refresh

Dim sSQLF4 As String
sSQLF4 = "    ||>Quantity    |>Chip Vol|^Media Vol|^Media SF|^Flow  "
MSFlexGrid4.FormatString = sSQLF4

MSFlexGrid4.Height = 8600
MSFlexGrid4.Width = 5800

MSFlexGrid5_Click
MSFlexGrid2_Click

End Sub

Private Sub cmdTBL_SBE_Click()
 

Dim sSQL As String
sSQL = "SELECT [ID],[CASE],[SERIES]," & _
              "format([ASF1],'0.00'),[MIN1]," & _
              "format([ASF2],'0.00'),[MIN2]," & _
              "format([ASF3],'0.00'),[MIN3], " & _
              "format([ASF4],'0.00'),[MIN4]," & _
              " [MODE],[DV MIN],[DV MAX]," & _
              "format([QTY MIN],'#,###,##0'),format([QTY MAX],'#,###,##0') " & _
        "FROM [TBL SBE] " & _
        "WHERE [CASE] IN ('A','B','C','E','S','F','L','R') AND "
                                   
                                   
If (Option1.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 100"
End If
If (Option2.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 200"
End If
If (Option7.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 700"
End If
If (Option9.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 200"
End If

If (Option600SFL.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 600"
End If
If (Option800AB.Value = True) Then
    sSQL = sSQL & "[SERIES_TYPE] = 810"
End If

Data5.RecordSource = sSQL
Data5.Refresh
 
Dim sSQLF As String
sSQLF = "   ||^Case||^NI ASF|^NI Min|^TN ASF|^TN Min|^LW ASF|^LW Min|^RP ASF|^RP Min||>DV MIN|>DV  MAX|>QTY MIN  |>QTY MAX  "
 
MSFlexGrid5.FormatString = sSQLF
MSFlexGrid5.Height = 4000
MSFlexGrid5.Width = 11500

End Sub

Private Sub cmdUpdate_Click()

Data1.UpdateRecord
cmdSPE_Click
MSFlexGrid4_Click

End Sub

Private Sub Command1_Click()

Data3.UpdateRecord
cmdTBL_SBE_Click

End Sub

Private Sub Form_Load()

Caption = "SBE Plating Parameter Tables DWG NO. 115-144    " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TABLES
Data2.DatabaseName = DB_PLATING_TABLES
Data3.DatabaseName = DB_PLATING_TABLES
Data4.DatabaseName = DB_PLATING_TABLES
Data5.DatabaseName = DB_PLATING_TABLES
 
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

cmdTBL_SBE_Click

sSQL = "SELECT [CASE],format([PCS PER SIDE MAX],'###,###'),format([SF],'0.0000')" & _
       "FROM [PCS PER SIDE] WHERE [TYPE]='SPE' ORDER BY[CASE]"
                                   
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "      |^Case|>Pcs/Sd      |^Sq In.    "
MSFlexGrid2.FormatString = sSQLF
'MSFlexGrid2.Height = 2000
MSFlexGrid2.Width = 3600

cmdSPE_Click
MSFlexGrid4_Click

If (ENABLE1_ATC_TABLES = 1) Then
    FrameSBE.Enabled = True
Else
    FrameSBE.Enabled = False
End If


End Sub

Private Sub MSFlexGrid2_Click()

Select Case CASE_SIZE_ID
Case "A"
              MSFlexGrid2.ROW = 1
Case "B"
              MSFlexGrid2.ROW = 2
Case "C"
              MSFlexGrid2.ROW = 3
Case "E"
              MSFlexGrid2.ROW = 4
Case "F"
              MSFlexGrid2.ROW = 5
Case "L"
              MSFlexGrid2.ROW = 6
Case "R"
              MSFlexGrid2.ROW = 7
Case "S"
              MSFlexGrid2.ROW = 8
End Select

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

End Sub

Private Sub MSFlexGrid4_Click()

MSFlexGrid4.Col = 1
TABLE_ID = Val(MSFlexGrid4.Text)

Dim TABLE_A_ID As Long

TABLE_A_ID = 0

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

If (OptionCaseS.Value = True) Then
    CASE_SIZE_ID = "S"
End If
If (OptionCaseF.Value = True) Then
    CASE_SIZE_ID = "F"
End If
If (OptionCaseL.Value = True) Then
    CASE_SIZE_ID = "L"
End If

If (OptionJAXA.Value = True) Then
    CASE_SIZE_ID = "A"
    TABLE_A_ID = 1
End If
If (OptionJAXB.Value = True) Then
    CASE_SIZE_ID = "B"
    TABLE_A_ID = 1
End If
If (OptionJAXR.Value = True) Then
    CASE_SIZE_ID = "R"
    TABLE_A_ID = 1
End If


Dim sSQL As String

Select Case TABLE_A_ID
Case 0
        sSQL = "SELECT * FROM [TBL SBE 144] WHERE [ID]=" & TABLE_ID
Case 1
        sSQL = "SELECT * FROM [TBL SBE ABR JAX] WHERE [ID]=" & TABLE_ID
End Select
                                                                
Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid4.Col = 0
MSFlexGrid4.ColSel = MSFlexGrid4.Cols - 1 '10

End Sub

Private Sub MSFlexGrid5_Click()

MSFlexGrid5.Col = 1
TABLE_ID = Val(MSFlexGrid5.Text)

Dim sSQL As String
sSQL = "SELECT * FROM [TBL SBE] WHERE [ID]=" & TABLE_ID
                                
Data3.RecordSource = sSQL
Data3.Refresh

MSFlexGrid5.Col = 0
MSFlexGrid5.ColSel = MSFlexGrid5.Cols - 1 '10

End Sub

Private Sub optCaseA_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub optCaseB_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub optCaseC_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub optCaseE_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub optCaseR_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub Option1_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option2_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option600SFL_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option7_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option800AB_Click()
cmdTBL_SBE_Click
End Sub

Private Sub Option9_Click()
cmdTBL_SBE_Click
End Sub

Private Sub OptionCaseF_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub OptionCaseL_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub OptionCaseS_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub OptionJAXA_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub OptionJAXB_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub OptionJAXR_Click()
cmdSPE_Click
MSFlexGrid4_Click
End Sub

Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11)
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12)
End Sub

Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13)
End Sub

Private Sub Text14_GotFocus()
Text14.SelStart = 0
Text14.SelLength = Len(Text14)
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
