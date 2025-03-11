VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkSheetDT 
   Caption         =   "090 Plating Worksheet Down Time"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "090 Work Sheet DT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2 [MACHINE]"
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
      Top             =   6840
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] Summary Planned and Unplanned DT all Departments"
      Height          =   375
      Left            =   4440
      TabIndex        =   32
      Top             =   9165
      Width           =   4695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3] Selected EQ History"
      Height          =   375
      Left            =   4440
      TabIndex        =   31
      Top             =   8790
      Width           =   5055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] Down Time History [TOTAL TIME] <>  0"
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      ToolTipText     =   "Time <> 0"
      Top             =   8415
      Width           =   5055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] Down Machines [TOTAL TIME] = 0"
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      ToolTipText     =   "Time = 0"
      Top             =   8040
      Value           =   -1  'True
      Width           =   4215
   End
   Begin VB.CommandButton cmdRefreshDisplay 
      BackColor       =   &H00FFC0FF&
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
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10080
      Width           =   2000
   End
   Begin VB.Frame fraDepartment 
      Caption         =   " Department "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   3975
      Begin VB.CommandButton cmdRefresh2 
         Caption         =   "Refresh2"
         Height          =   300
         Left            =   840
         TabIndex        =   25
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Bindings        =   "090 Work Sheet DT.frx":0CCA
         Height          =   1095
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "FROM [MACHINE]"
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         _Version        =   393216
         BackColorSel    =   16711680
         AllowBigSelection=   0   'False
         SelectionMode   =   1
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
      Begin VB.Label lblEQ 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   3345
      End
   End
   Begin VB.Frame fraWS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   3975
      Begin VB.CommandButton cmdDelete 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   1440
      End
      Begin VB.CommandButton cmdRefreshNew 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1680
         Width           =   1440
      End
      Begin VB.TextBox txtATCPart 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   18
         Text            =   "txt20"
         Top             =   1200
         Width           =   2520
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
         Left            =   2280
         TabIndex        =   3
         Top             =   2160
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   1440
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
         TabIndex        =   2
         Top             =   2160
         Width           =   1440
      End
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
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Total Time in Minutes"
         Top             =   3240
         Width           =   465
      End
      Begin VB.TextBox txtWorkOrder 
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
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
         Left            =   1200
         MaxLength       =   17
         TabIndex        =   1
         Text            =   "txt17"
         Top             =   720
         Width           =   2520
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   360
         TabIndex        =   14
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
         Format          =   123994114
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "STOP TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
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
         Format          =   123994114
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "DATE_ID"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   1680
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
         Format          =   123994113
         CurrentDate     =   38117
      End
      Begin VB.Label lblInfo 
         Caption         =   "FROM [WORK SHEET]"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label lblInfo 
         Caption         =   "Comment:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1185
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
         TabIndex        =   17
         Top             =   3240
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         Caption         =   "Problem :"
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
         TabIndex        =   16
         Top             =   720
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Code 21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   21
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Code 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   20
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   2400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [WORK SHEET],[MACHINE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit to Main"
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
      Left            =   6960
      TabIndex        =   10
      Top             =   10080
      Width           =   2000
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [WORK SHEET]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WORK SHEET"
      Top             =   1080
      Visible         =   0   'False
      Width           =   5580
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Work Sheet DT.frx":0CDE
      Height          =   1095
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "FROM [WORK SHEET],[MACHINE]"
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
   Begin MSComCtl2.DTPicker DTPicker4 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   6240
      TabIndex        =   33
      Top             =   9600
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   123994113
      CurrentDate     =   38117
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   7800
      TabIndex        =   34
      Top             =   9600
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   123994113
      CurrentDate     =   38117
   End
   Begin VB.CheckBox Check1 
      Caption         =   "[1] Per Selected Department"
      Height          =   375
      Left            =   8880
      TabIndex        =   28
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Caption         =   "FROM [WORK SHEET],[MACHINE]"
      Height          =   300
      Index           =   3
      Left            =   12000
      TabIndex        =   37
      Top             =   8040
      Width           =   2745
   End
   Begin VB.Label lblInfo 
      Caption         =   "Per Date Range"
      Height          =   300
      Index           =   0
      Left            =   4920
      TabIndex        =   35
      Top             =   9600
      Width           =   1425
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   10920
      Picture         =   "090 Work Sheet DT.frx":0CF2
      Top             =   9720
      Width           =   4170
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
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
      Index           =   21
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C0"
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
      Index           =   20
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   600
   End
   Begin VB.Label txtOperator 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label txtShift 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmWorkSheetDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

DTPicker2.Value = DateAdd("n", txtTotalTime.Text, DTPicker1.Value)

Data4.UpdateRecord

cmdRefreshDisplay_Click

End Sub

Private Sub cmdCode_Click(Index As Integer)

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET PT] " & _
       "WHERE [DATE_ID]=#" & DATE_ID & "# AND [OP_ID]=" & OP_ID & " AND [MACHINE_ID]=" & MACHINE_ID
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
FR_Table.AddNew
WS_ID = FR_Table.Fields("[WS_ID]")

FR_Table.Fields("[OP_ID]") = OP_ID
FR_Table.Fields("[DATE_ID]") = DATE_ID
FR_Table.Fields("[MACHINE_ID]") = MACHINE_ID
FR_Table.Fields("[CODE_ID]") = Val(lblCode(Index).Caption)
 
'FR_Table.Fields("[WORK ORDER]") = "Problem"
'FR_Table.Fields("[ATC PART]") = "Comment"

FR_Table.Fields("[START TIME]") = Format(Time, "hh:mm am/pm")
FR_Table.Fields("[TOTAL TIME]") = 0

FR_Table.Update

FR_Table.Close
FR_Database.Close

txtWorkOrder.SetFocus

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdDelete_Click()

Dim iAns As Integer
iAns = MsgBox("Delete WS_ID " & WS_ID, vbYesNo, "ATC Termination System")
If (iAns = vbYes) Then

        Dim sSQL As String
        
        Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
        sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            FR_Table.Delete
        End If
         
        FR_Table.Close
        FR_Database.Close
        
        cmdRefreshDisplay_Click
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
 
Private Sub cmdNext_Click()

DTPicker3.Value = DateAdd("D", 1, DTPicker3.Value)
WS_ID = -1

cmdRefreshDisplay_Click
MSFlexGrid1_Click
 
End Sub

Private Sub cmdPrevious_Click()

DTPicker3.Value = DateAdd("D", -1, DTPicker3.Value)
WS_ID = -1
 
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh_Click()

WS_ID = -1
 
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh2_Click()

Dim sSQLF As String
Dim sSQL As String

sSQL = "SELECT [MACHINE_ID],[NUMBER],[TYPE],[DESCRIPTION]   " & _
       "FROM [MACHINE]  " & _
       "WHERE [DEPT_ID]='PT' AND [ACTIVE] = 1 " & _
       "ORDER BY [TYPE],[NUMBER],mid([PROCESS],1,1)"

sSQLF = "    ||^M#|^                  |Description                 "
 
Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF
MSFlexGrid2.Width = 3700
MSFlexGrid2.Height = 4000

End Sub

Private Sub cmdRefreshDisplay_Click()

DATE_START_ID = DTPicker4.Value
DATE_END_ID = DTPicker5.Value

Dim sSQL As String
 

 If (Option1.Value = True) Then

 sSQL = "SELECT [WS_ID],[CODE_ID]," & _
               "[MACHINE].[NUMBER],[MACHINE].[NAME]," & _
               "[WORK ORDER],[ATC PART],[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM')," & _
              "format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
       "FROM [WORK SHEET],[MACHINE] " & _
       "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]=0 AND " & _
             "[WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
             "[MACHINE].[DEPT_ID]='PT' " & _
       "ORDER BY [WS_ID] DESC"
       
       sSQL = "SELECT [WS_ID],[CODE_ID]," & _
               "[MACHINE].[NUMBER],[MACHINE].[NAME]," & _
               "[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM')," & _
              "format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
       "FROM [WORK SHEET PT],[MACHINE] " & _
       "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]=0 AND " & _
             "[WORK SHEET PT].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
             "[MACHINE].[DEPT_ID]='PT' " & _
       "ORDER BY [WS_ID] DESC"
       
End If

If (Option2.Value = True) Then

 sSQL = "SELECT [WS_ID],[CODE_ID]," & _
               "[MACHINE].[NUMBER],[MACHINE].[NAME]," & _
               "[WORK ORDER],[ATC PART],[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM')," & _
              "format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
       "FROM [WORK SHEET],[MACHINE] " & _
       "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]<>0 AND " & _
             "[WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
             "[MACHINE].[DEPT_ID]='PT' " & _
       "ORDER BY [WS_ID] DESC"
       
        sSQL = "SELECT [WS_ID],[CODE_ID]," & _
               "[MACHINE].[NUMBER],[MACHINE].[NAME]," & _
               "[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM')," & _
              "format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
       "FROM [WORK SHEET PT],[MACHINE] " & _
       "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]<>0 AND " & _
             "[WORK SHEET PT].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
             "[MACHINE].[DEPT_ID]='PT' " & _
       "ORDER BY [WS_ID] DESC"
       
End If

If (Option3.Value = True) Then

 sSQL = "SELECT [WS_ID],[CODE_ID],[MACHINE].[MACHINE],[MACHINE].[Description],[WORK ORDER],[ATC PART],[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
      "FROM [WORK SHEET],[MACHINE] " & _
      "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]<> 0 AND " & _
            "[WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
            "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " " & _
      "ORDER BY [WS_ID] DESC"
      
      sSQL = "SELECT [WS_ID],[CODE_ID],[MACHINE].[MACHINE],[MACHINE].[Description],[DATE_ID]," & _
              "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM'),[TOTAL TIME] " & _
      "FROM [WORK SHEET PT],[MACHINE] " & _
      "WHERE [CODE_ID] IN (100,101) AND [TOTAL TIME]<> 0 AND " & _
            "[WORK SHEET PT].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
            "[WORK SHEET PT].[MACHINE_ID] =" & MACHINE_ID & " " & _
      "ORDER BY [WS_ID] DESC"
      
End If
 
Dim sSQLF As String
sSQLF = "   ||^Code|^M#    |<Description      |<Problem                     |<Comment                      |^Date               |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  "

If (Option4.Value = True) Then
sSQL = "SELECT         first([MACHINE].[MACHINE])," & _
                      "first([MACHINE].[Description])," & _
                   "count([WORK SHEET PT].[WS_ID])," & _
                   "first([WORK SHEET PT].[CODE_ID])," & _
                   "first([WORK SHEET PT].[DATE_ID])," & _
                    "last([WORK SHEET PT].[DATE_ID])," & _
              "format(sum([WORK SHEET PT].[TOTAL TIME]),'##,###') " & _
       "FROM [WORK SHEET PT],[MACHINE] " & _
       "WHERE [WORK SHEET PT].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
             "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
             "[WORK SHEET PT].[CODE_ID] IN (100,100) AND [TOTAL TIME] <> 0 " & _
       "GROUP BY [WORK SHEET PT].[MACHINE_ID],[WORK SHEET PT].[CODE_ID]"
       
sSQLF = "   |^M#    |<Description                  |>Count|^Code|^Date               |^Date               "
sSQLF = sSQLF & "|Total Time"

End If

Data1.RecordSource = sSQL
Data1.Refresh
 
MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefreshNew_Click()

Data4.UpdateRecord
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdReset_Click()

DTPicker3.Value = Date
WS_ID = -1
 
cmdRefreshDisplay_Click
MSFlexGrid1_Click
 
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

cmdRefreshDisplay_Click

End Sub

Private Sub cmdSub_Click()

DTPicker1.Value = DateAdd("n", -Val(txtTotalTime.Text), DTPicker2.Value)

Data4.UpdateRecord

cmdRefreshDisplay_Click

End Sub

Private Sub Form_Load()

Caption = "Plating Worksheet Down Time     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION
Data4.DatabaseName = DB_PLATING_TERMINATION
 
MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11000
MSFlexGrid1.Height = 8000
MSFlexGrid1.ForeColorSel = vbBlack

lblCode(20).Caption = "100"
cmdCode(20).Caption = "Planned Downtime"
lblCode(21).Caption = "101"
cmdCode(21).Caption = "UnPlanned Downtime"

DTPicker3.Value = Date
DTPicker5.Value = Date
DTPicker4.Value = DateAdd("m", -6, Date)

WS_ID = -1

'
'   DISPLAY OPERATOR AND MACHINE INFORMATION
'
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String

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

DATE_ID = Date$

DTPicker1.Value = Format(Time, "hh:mm am/pm")
DTPicker2.Value = Format(Time, "hh:mm am/pm")

cmdRefreshDisplay_Click
MSFlexGrid1_Click

cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub
 
Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
WS_ID = Val(MSFlexGrid1.Text)

fraWS.Caption = "WS_ID: " & WS_ID
fraWS.Enabled = True

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET PT] WHERE [WS_ID]=" & WS_ID
Data4.RecordSource = sSQL
Data4.Refresh
        
If (WS_ID = -1) Then

End If

fraWS.Enabled = True

MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
MACHINE_ID = Val(MSFlexGrid2.Text)
        
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    lblEQ.Caption = FR_Table.Fields("[NUMBER]") & " " & FR_Table.Fields("[NAME]")
Else
    lblEQ.Caption = ""
End If
FR_Table.Close
FR_Database.Close
                
MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

End Sub

Private Sub Option1_Click()
cmdRefreshDisplay_Click
End Sub

Private Sub Option2_Click()
cmdRefreshDisplay_Click
End Sub

Private Sub txtATCPart_GotFocus()
txtATCPart.SelStart = 0
txtATCPart.SelLength = Len(txtATCPart)
End Sub

Private Sub txtOrderQty_GotFocus()
txtOrderQty.SelStart = 0
txtOrderQty.SelLength = Len(txtOrderQty)
End Sub

Private Sub txtSQ_GotFocus()
txtSQ.SelStart = 0
txtSQ.SelLength = Len(txtSQ)
End Sub

Private Sub txtTotalTime_GotFocus()
txtTotalTime.SelStart = 0
txtTotalTime.SelLength = Len(txtTotalTime)
End Sub

Private Sub txtWorkOrder_GotFocus()
txtWorkOrder.SelStart = 0
txtWorkOrder.SelLength = Len(txtWorkOrder)
End Sub
