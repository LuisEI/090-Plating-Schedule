VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "090 Plating Main Schedule"
   ClientHeight    =   12195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15885
   Icon            =   "090 MAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12195
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerExitProgram 
      Interval        =   60000
      Left            =   5760
      Top             =   2880
   End
   Begin VB.CommandButton cmdRemoteNY 
      Caption         =   "NY"
      Height          =   300
      Left            =   11640
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdRemoteJR 
      Caption         =   "JR"
      Height          =   300
      Left            =   12480
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdUpdateWO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update Data Bases"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8340
      Width           =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   11160
      Top             =   2040
   End
   Begin VB.CommandButton cmdDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "EQ Down Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Equipment Down Time Screen"
      Top             =   5040
      Width           =   2880
   End
   Begin VB.CommandButton cmdRT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Real Time (Amp Hr)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7920
      Width           =   2880
   End
   Begin VB.CommandButton cmdSQL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Search SQL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6060
      Width           =   2880
   End
   Begin VB.CommandButton cmdDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dept Codes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Plating  Department Codes"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.CommandButton cmdReviewSets 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Review Sets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3540
      Width           =   2880
   End
   Begin VB.CommandButton cmdEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Plating Equipment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "ATC Plating Equipment"
      Top             =   6780
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Frame fraTables 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ATC Tables "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   11640
      TabIndex        =   16
      Top             =   9000
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdTankCU 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Barrel Tables Copper"
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
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1980
         Width           =   3000
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SBE Calculation"
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
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   900
         Width           =   3000
      End
      Begin VB.CommandButton cmdBC 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Barrel Calculation"
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
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2400
         Width           =   3000
      End
      Begin VB.CommandButton cmdSBE 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SBE (Spouted Bed Electrode)"
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
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   3000
      End
      Begin VB.CommandButton cmdTank 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Barrel Tables Nickel"
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
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1560
         Width           =   3000
      End
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   10920
      PasswordChar    =   "*"
      TabIndex        =   13
      Text            =   "XXXX"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "WS Review"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Work Sheet Review"
      Top             =   4620
      Width           =   2880
   End
   Begin VB.Data Data3 
      Caption         =   "Data3  FROM [MACHINE]"
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
      Top             =   7200
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.CommandButton cmdRefresh3 
      Caption         =   "Refresh3"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   7800
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CommandButton cmdWS 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Work Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   2880
   End
   Begin VB.CommandButton cmdSummary 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Summaries"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   2880
   End
   Begin VB.CommandButton cmdSetReview 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Create Sets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2880
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CommandButton cmdRefresh1 
      Caption         =   "Refresh1"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [DEPT CODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   3540
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 MAIN.frx":0CCA
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "[BARCODE]"
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 MAIN.frx":0CDE
      Height          =   975
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "[DEPT CODE]"
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "090 MAIN.frx":0CF2
      Height          =   975
      Left            =   4920
      TabIndex        =   10
      ToolTipText     =   "[MACHINE]"
      Top             =   6120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1720
      _Version        =   393216
      ScrollTrack     =   -1  'True
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
   Begin VB.Label LabelFormResolution 
      Alignment       =   2  'Center
      Caption         =   "Form Res"
      Height          =   255
      Left            =   10200
      TabIndex        =   34
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TimerExitProgram_Timer"
      Height          =   300
      Left            =   6240
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mode"
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
      Left            =   14520
      TabIndex        =   29
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JR"
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
      Left            =   13560
      TabIndex        =   28
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SBE"
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
      Left            =   1680
      TabIndex        =   15
      Top             =   11520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barrel"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3180
      Left            =   240
      Picture         =   "090 MAIN.frx":0D06
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   4500
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3180
      Left            =   240
      Picture         =   "090 MAIN.frx":233FB
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   4500
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   11640
      Picture         =   "090 MAIN.frx":488FA
      Top             =   0
      Width           =   4170
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code_ID"
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
      Left            =   13080
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   11880
      TabIndex        =   6
      Top             =   2520
      Width           =   3240
   End
   Begin VB.Label lblCode_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "555"
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
      Left            =   11880
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu mnuOP 
      Caption         =   "Operators"
   End
   Begin VB.Menu mnuConfigPlating 
      Caption         =   "Configuration Plating"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub cmdBC_Click()
frmBarrelCalculation.Show vbModal
End Sub

Private Sub cmdDept_Click()
frmDept.Show vbModal
End Sub

Private Sub cmdDown_Click()

If (OP_ID = 0) Then
    MsgBox "No Operator selected ", vbInformation, "EDB System"
    Exit Sub
End If

frmWorkSheetDT.Show
frmMain.Hide

End Sub

Private Sub cmdEQ_Click()

frmEquipment.Show vbModal

End Sub

Private Sub cmdRefresh1_Click()

Dim sSQL As String
sSQL = "SELECT [OP_ID], [FIRST] & ' ' & [LAST],[SHIFT_ID] " & _
       "FROM [BARCODE] " & _
       "WHERE [ACTIVE] = 1 AND [DEPT_ID]='PT' AND [LOCATION_ID]='" & LOCATION_ID & "' " & _
       "ORDER BY [LAST],[FIRST]"
                                   
Data1.RecordSource = sSQL
Data1.Refresh

Dim sSQLF As String
sSQLF = "    ||<Operator                                |^Shift"

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefresh2_Click()

Dim sSQL As String
Dim sSQLF As String

    
Select Case LOCATION_ID
Case "JR"
        sSQL = "SELECT [DEPT_JR_ID],[DESCRIPTION],[SBE],[TANK] " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "

        sSQL = sSQL & "[LOC_JR]='" & LOCATION_ID & "' ORDER BY [DEPT_ID]"
Case "NY"
        sSQL = "SELECT [DEPT_ID],[DESCRIPTION],[SBE],[TANK] " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "
        sSQL = sSQL & "[LOC_NY]='" & LOCATION_ID & "' ORDER BY [DEPT_ID]"
End Select

sSQLF = "    |^Code ID|<Base / Finish              |^SBE  |^TANK"

Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF
 
End Sub

Private Sub cmdRefresh3_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [NUMBER],[TYPE]," & _
              "[DESCRIPTION]," & _
              "mid([PROCESS],1,1)& mid(lcase([PROCESS]),2,5) AS [SQL PROCESS]," & _
              "[LOCATION_ID]  " & _
       "FROM [MACHINE]  " & _
       "WHERE [DEPT_ID]='PT' AND [ACTIVE] = 1 AND "
         
sSQL = sSQL & "[LOCATION_ID]='" & LOCATION_ID & "'"

sSQL = sSQL & "ORDER BY [TYPE] DESC,mid([PROCESS],1,1),[NUMBER]"

sSQLF = "    |^M ID|^TYPE       |Description               |^Process |^L_ID"

Data3.RecordSource = sSQL
Data3.Refresh

MSFlexGrid3.FormatString = sSQLF
 
End Sub


Private Sub cmdRemoteJR_Click()

Screen.MousePointer = vbHourglass

DataBase_MODE = 3
LOCATION_ID = "JR"

cmdRemoteNY.FontBold = False
cmdRemoteJR.FontBold = True

DataBase_Address
Form_Load
Form_Activate
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRemoteNY_Click()

Screen.MousePointer = vbHourglass

DataBase_MODE = 0
LOCATION_ID = "NY"

cmdRemoteNY.FontBold = True
cmdRemoteJR.FontBold = False

DataBase_Address
Form_Load
Form_Activate
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdReviewSets_Click()
frmSetReview.Show
End Sub

Private Sub cmdRT_Click()
frmReal_Time.Show
End Sub

Private Sub cmdSBE_Click()
 
frmSBEParameters.Show         'SBE PLATTING PARAMETERS

End Sub

Private Sub cmdSetReview_Click()

If (DEPT_ID = 0) Then
    MsgBox "Select Dept Code", vbInformation, "ATC"
    Exit Sub
End If
If (OP_ID = 0) Then
    MsgBox "Select Operator", vbInformation, "ATC"
    Exit Sub
End If

frmMain.Hide

frmSetCreate.Show

End Sub

Private Sub cmdSQL_Click()
frmSQL.Show vbModal
End Sub

Private Sub cmdSummary_Click()

frmSummary.Show vbModal

End Sub

Private Sub cmdTank_Click()
frmBarrelParametersNickel.Show        'BARREL PLATTING PARAMETERS
End Sub

Private Sub cmdTankCU_Click()
frmBarrelParametersCU.Show        'BARREL PLATTING PARAMETERS
End Sub

Private Sub cmdUpdateWO_Click()

Select Case LOCATION_ID
Case "NY"

Case "JR"
        Screen.MousePointer = vbHourglass

        Update_WO_Schedule
        Screen.MousePointer = vbDefault
        MsgBox "Update Complete", vbInformation, "EDB System"
        
Case Else

End Select

End Sub

Private Sub cmdWS_Click()

If (OP_ID = 0 Or DEPT_ID = 0) Then
    MsgBox "No Operator or Equipment selected ", vbInformation, "EDB System"
    Exit Sub
End If

frmWorkSheet.Show

frmMain.Hide

End Sub

Private Sub Command1_Click()

frmMain.Hide
frmWorkSheetR.Show

End Sub

Private Sub Command3_Click()
frmSBECalculation.Show vbModal
End Sub

Private Sub Form_Activate()

Select Case DB_SOURCE_ID
Case 0
        cmdRemoteNY.Visible = False
        cmdRemoteJR.Visible = False
Case 1
        cmdRemoteNY.Visible = True
        cmdRemoteJR.Visible = True
End Select

Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
        lblMode.Caption = "REM NY"
Case DATABASE_MODE_REM_JR
        lblMode.Caption = "REM JR"
Case DATABASE_MODE_LCL
        lblMode.Caption = "LCL"
Case DATABASE_MODE_FIL
        lblMode.Caption = "FILE"
End Select

'If (ENABLE1_ATC_TABLES = 1) Then
    fraTables.Visible = True
'Else
 '   fraTables.Visible = False
'End If

If (ENABLE2_DEPT_CODES = 1) Then
    cmdDEPT.Visible = True
Else
    cmdDEPT.Visible = False
End If

If (ENABLE4_EQ = 1) Then
    cmdEQ.Visible = True
Else
    cmdEQ.Visible = False
End If

End Sub

Private Sub Form_Load()


Caption = "Plating Main Schedule" & Space(8) & ATC_DWG & Space(8) & ATC_VERSION
Caption = Caption & Space(8) & "IP " & IP_ADDRESS & Space(8) & strComputerName

lblLocation = LOCATION_ID

LabelFormResolution.Caption = frmMain.Width / Screen.TwipsPerPixelX & " by " & frmMain.Height / Screen.TwipsPerPixelY

MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 4400
MSFlexGrid1.Height = 5000

MSFlexGrid2.Top = 0
MSFlexGrid2.Width = 5400
MSFlexGrid2.Height = 6000

MSFlexGrid3.Width = 5400
MSFlexGrid3.Height = 6200

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data3.DatabaseName = DB_PLATING_TERMINATION

cmdRefresh1_Click
MSFlexGrid1_Click

cmdRefresh2_Click
MSFlexGrid2_Click

cmdRefresh3_Click

mnuOP.Visible = False
mnuConfigPlating.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim iAns As Integer
iAns = MsgBox("Exit Program", vbYesNo, "ATC Plating System")
If (iAns = vbYes) Then
    Cancel = 0
    
    Select Case DataBase_MODE
    Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            DataBase_MODE = DATABASE_MODE_REM_NY
            LOCATION_ID = "NY"
    End Select
    
    ConfigComputer_DB (1)
    End
Else
    Cancel = 1
End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then
     If (mnuOP.Visible = False) Then
        mnuOP.Visible = True
        mnuConfigPlating.Visible = True
        cmdDEPT.Visible = True
     Else
        mnuOP.Visible = False
        mnuConfigPlating.Visible = False
        cmdDEPT.Visible = False
     End If
End If

End Sub

Private Sub mnuConfigPlating_Click()

frmConfiguration.Show vbModal

End Sub

Private Sub mnuOP_Click()

frmOperator.Show vbModal

End Sub
 
Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
OP_ID = Val(MSFlexGrid1.Text)
  
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Dim sSQL As String
sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    lblOperator.Caption = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
Else
    lblOperator.Caption = ""
End If
FR_Table.Close
FR_Database.Close
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
DEPT_ID = Val(MSFlexGrid2.Text)
  
lblCode_ID.Caption = DEPT_ID
   
MSFlexGrid2.Col = 2
lblDesc.Caption = MSFlexGrid2.Text
   
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
     
MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1

End Sub

Private Sub Timer1_Timer()
Configuration (FWRITE)
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

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPasswordv)
End Sub
