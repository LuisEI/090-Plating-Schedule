VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReal_Time 
   Caption         =   "090 Plating Real Time View "
   ClientHeight    =   12495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "090 Real Time.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12495
   ScaleWidth      =   15885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandRefreshSA 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Refresh SA/SQ FT"
      Height          =   300
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   2400
   End
   Begin VB.CommandButton CommandSAF 
      BackColor       =   &H00FFC0C0&
      Caption         =   "FINISH Sum  SA/SQ FT"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2880
      Width           =   2400
   End
   Begin VB.CommandButton CommandSA 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BASE Sum  SA/SQ FT"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Width           =   2400
   End
   Begin VB.CheckBox Check2 
      Caption         =   "[2] All Equipment "
      Height          =   300
      Left            =   1080
      TabIndex        =   20
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdAmp 
      BackColor       =   &H00FFC0FF&
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
      Height          =   360
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Summed Codes"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   2400
   End
   Begin VB.CommandButton cmdSelected 
      BackColor       =   &H00FFC0FF&
      Caption         =   "History M_ID :"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   2400
   End
   Begin VB.CheckBox Check1 
      Caption         =   "[1] [TOTAL TIME]= 0 "
      Height          =   300
      Left            =   1080
      TabIndex        =   16
      Top             =   2880
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] CODE (500,600) FINISH EQUIPMENT"
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   3585
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] CODE (300,400) BASE EQUIPMENT"
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Value           =   -1  'True
      Width           =   3585
   End
   Begin VB.CommandButton cmdRefresh1 
      Caption         =   "Refresh1"
      Height          =   300
      Left            =   10680
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   6420
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
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3375
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
         TabIndex        =   8
         Top             =   1440
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
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
         Height          =   360
         Left            =   360
         TabIndex        =   3
         Top             =   720
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
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   945
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
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
   Begin VB.Data Data1 
      Caption         =   "Data1FROM [MACHINE]"
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
      Width           =   3540
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "090 Real Time.frx":0CCA
      Height          =   4455
      Left            =   9960
      TabIndex        =   0
      ToolTipText     =   "FROM [MACHINE]"
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7858
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
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "090 Real Time.frx":0CDE
      Height          =   1095
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE]"
      Top             =   4560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1931
      _Version        =   393216
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
      Caption         =   "SETUP [300,500], CHECK [400,600]"
      Height          =   300
      Index           =   2
      Left            =   480
      TabIndex        =   13
      Top             =   3840
      Width           =   2985
   End
   Begin VB.Label lblInfo 
      Caption         =   "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE]"
      Height          =   300
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   4425
   End
End
Attribute VB_Name = "frmReal_Time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAmp_Click()

Screen.MousePointer = vbHourglass

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

Select Case PROCESS_ID
Case "BASE"
                                 
      sSQL = "SELECT first([SCHEDULE SETS].[TYPE_ID])                    AS [SQL TYPE_ID]," & _
                    "first([SCHEDULE SETS].[EQ BASE])                    AS [SQL EQ_ID]," & _
                    "first([SCHEDULE SETS].[EQ BASE])                    AS [SQL DESC]," & _
                    "first([WORK SHEET PT].[DATE_ID])," & _
                     "last([WORK SHEET PT].[DATE_ID])," & _
               "format(sum([SCHEDULE SETS].[BASE AMP MIN]),'###,##0')" & _
            "FROM [WORK SHEET PT],[SCHEDULE SETS] " & _
            "WHERE [WORK SHEET PT].[SET_ID]  = [SCHEDULE SETS].[SET_ID] AND " & _
                  "[WORK SHEET PT].[CODE_ID] IN (300) AND " & _
                  "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                  "[SCHEDULE SETS].[EQ BASE] = " & MACHINE_ID & " " & _
             "GROUP BY [SCHEDULE SETS].[EQ BASE]"
Case "FINISH"
                         
        sSQL = "SELECT first([SCHEDULE SETS].[TYPE_ID])                    AS [SQL TYPE_ID]," & _
                      "first([SCHEDULE SETS].[EQ FINISH])                  AS [SQL EQ_ID]," & _
                      "first([SCHEDULE SETS].[EQ FINISH])                  AS [SQL DESC]," & _
                      "first([WORK SHEET PT].[DATE_ID])," & _
                       "last([WORK SHEET PT].[DATE_ID])," & _
                 "format(sum([SCHEDULE SETS].[FINISH AMP MIN]),'###,##0')" & _
              "FROM [WORK SHEET PT],[SCHEDULE SETS] " & _
              "WHERE [WORK SHEET PT].[SET_ID]    = [SCHEDULE SETS].[SET_ID] AND " & _
                    "[WORK SHEET PT].[CODE_ID] IN (500) AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                    "[SCHEDULE SETS].[EQ FINISH] = " & MACHINE_ID & " " & _
               "GROUP BY [SCHEDULE SETS].[EQ FINISH]"
End Select

sSQLF = "   |^Type         |^EQ    |^Description                            "
sSQLF = sSQLF & "|^Start Date           |^End Date          |>Amp Min       "
                                                    
Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRefresh1_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [NUMBER],[TYPE],[DESCRIPTION],mid([PROCESS],1,1)& mid(lcase([PROCESS]),2,5)  " & _
       "FROM [MACHINE] " & _
       "WHERE [DEPT_ID]='PT' AND " & _
             "[ACTIVE] = 1 AND " & _
             "[LOCATION_ID]='" & LOCATION_ID & "' " & _
       "ORDER BY mid([PROCESS],1,1),[TYPE]"

sSQLF = "    |^M ID|^                     |^Description                      |^               "

Data1.RecordSource = sSQL
Data1.Refresh
MSFlexGrid1.FormatString = sSQLF

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

cmdRefresh1_Click
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

cmdRefresh1_Click

End Sub

Private Sub cmdRefresh_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

' "[SCHEDULE SETS].[EQ BASE] =" & MACHINE_ID

If (Option1.Value = True) Then
        'BASE EQUIPMENT AND PROCESS
        sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                      "[SCHEDULE SETS].[DEPT_ID]                    AS [SQL DEPT_ID]," & _
                      "[WORK SHEET PT].[SET_ID]," & _
                      "[SCHEDULE SETS].[DATE_ID]," & _
                      "[SCHEDULE SETS].[SET NUMBER]                 AS [SQL SET NUMBER]," & _
                      "[SCHEDULE SETS].[TYPE_ID]                    AS [SQL TYPE_ID]," & _
                      "[SCHEDULE SETS].[EQ BASE]                    AS [SQL EQ_ID]," & _
                      "[WORK SHEET PT].[OP_ID]," & _
                      "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST] AS [SQL OPERATOR]," & _
                      "[WORK SHEET PT].[DATE_ID]," & _
                      "[WORK SHEET PT].[CODE_ID]," & _
                      "format([START TIME],'h:mm AM/PM')," & _
                      "format([STOP TIME],'h:mm AM/PM')," & _
                      "[WORK SHEET PT].[TOTAL TIME] " & _
              "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
              "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                    "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
                    "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND "
                   
        If (Check2.Value = vbChecked) Then
        sSQL = sSQL & "[SCHEDULE SETS].[EQ BASE] IN (11,14,15,17,18) AND "
        Else
        sSQL = sSQL & "[SCHEDULE SETS].[EQ BASE]= " & MACHINE_ID & " AND "
        End If
        
        sSQL = sSQL & "[WORK SHEET PT].[CODE_ID] IN (300,400)"
        If (Check1.Value = vbChecked) Then
        sSQL = sSQL & " AND [WORK SHEET PT].[TOTAL TIME]= 0 "
        End If
        sSQL = sSQL & " ORDER BY [SCHEDULE SETS].[EQ BASE] ASC," & _
                        "[SCHEDULE SETS].[SET NUMBER] ASC," & _
                        "[WORK SHEET PT].[CODE_ID] ASC"
End If
                        
If (Option2.Value = True) Then
        'FINISH EQUIPMENT AND PROCESS
        sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                        "[SCHEDULE SETS].[DEPT_ID]                            AS [SQL DEPT_ID]," & _
                        "[WORK SHEET PT].[SET_ID]," & _
                        "[SCHEDULE SETS].[DATE_ID]," & _
                        "[SCHEDULE SETS].[SET NUMBER]                         AS [SQL SET NUMBER]," & _
                        "[SCHEDULE SETS].[TYPE_ID]                            AS [SQL TYPE_ID]," & _
                        "[SCHEDULE SETS].[EQ FINISH]                          AS [SQL EQ_ID]," & _
                        "[WORK SHEET PT].[OP_ID]," & _
                        "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST] AS [SQL OPERATOR]," & _
                        "[WORK SHEET PT].[DATE_ID]," & _
                        "[WORK SHEET PT].[CODE_ID]," & _
                        "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                        "[WORK SHEET PT].[TOTAL TIME] " & _
                "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
                "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                      "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
                      "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND "
            
        If (Check2.Value = vbChecked) Then
        sSQL = sSQL & "[SCHEDULE SETS].[EQ FINISH] IN (21,22,23,26,27,28) AND "
        Else
        sSQL = sSQL & "[SCHEDULE SETS].[EQ FINSIH]= " & MACHINE_ID & " AND "
        End If
            
        sSQL = sSQL & "[WORK SHEET PT].[CODE_ID] IN (500,600)"
        
        If (Check1.Value = vbChecked) Then
        sSQL = sSQL & " AND [WORK SHEET PT].[TOTAL TIME]= 0 "
        End If
        sSQL = sSQL & " ORDER BY [SCHEDULE SETS].[EQ BASE] ASC," & _
                                "[SCHEDULE SETS].[SET NUMBER] ASC," & _
                                "[WORK SHEET PT].[CODE_ID] ASC"
End If

sSQLF = "   ||^Dept_ID|^Set ID|^Create Date|^Set No.|^Type         |^EQ    |^|<Operator                |^Actual  Date |^Code    |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  "
                                 
Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF
                                
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

Private Sub cmdSearch_Click()

WO_ID = txtWorkOrder.Text

Dim sSQL As String
Dim sSQLF As String
                                                                                                         
              '"[EQ BASE],[BASE AMP],[BASE AMP MIN]," & _
              '"[EQ FINISH],[FINISH AMP],[FINISH AMP MIN]  " & _
              '[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
              '[WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND "
              'DISTINCT

sSQL = "SELECT [SCHEDULE SETS].[DATE_ID]," & _
              "[GROUPING].[WORK ORDER]," & _
              "[GROUPING].[ATC PART]," & _
              "[GROUPING].[LOT NUM]," & _
              "[GROUPING].[QTY]," & _
              "[SCHEDULE SETS].[DEPT_ID]," & _
              "[SCHEDULE SETS].[SET NUMBER]," & _
              "[SCHEDULE SETS].[TYPE_ID]," & _
              "[SCHEDULE SETS].[EQ BASE]," & _
              "[SCHEDULE SETS].[EQ FINISH] " & _
       "FROM [SCHEDULE SETS],[GROUPING]" & _
       "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
             "[GROUPING].[WORK ORDER]='" & WO_ID & "'"
                                                                                                          
Data2.RecordSource = sSQL
Data2.Refresh

sSQLF = "    |^Create Date    |^Work Order           |^ATC Part                     |^Lot Number        |Quantity     "
sSQLF = sSQLF & "|^Dept|^Set#|^Type        |^Base EQ|^Finish EQ"
 
MSFlexGrid2.FormatString = sSQLF

    sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                  "[WORK SHEET PT].[SET_ID]," & _
                  "[SCHEDULE SETS].[DATE_ID]," & _
                  "[SCHEDULE SETS].[SET NUMBER]," & _
                  "[SCHEDULE SETS].[TYPE_ID]," & _
                  "[WORK SHEET PT].[OP_ID]," & _
                  "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST]," & _
                  "[WORK SHEET PT].[DATE_ID]," & _
                  "[WORK SHEET PT].[CODE_ID]," & _
                  "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                  "[WORK SHEET PT].[TOTAL TIME] " & _
          "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
          "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
                "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
                "[SCHEDULE SETS].[DEPT_ID]=" & DEPT_ID & " AND " & _
                "[WORK SHEET PT].[DATE_ID] =#" & DATE_ID & "# " & _
          "ORDER BY [WORK SHEET PT].[WS_ID] DESC"


sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
              "[WORK SHEET PT].[CODE_ID]," & _
              "[WORK SHEET PT].[DATE_ID]," & _
              "[BARCODE].[FIRST] & ' ' & [BARCODE].[LAST]," & _
              "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
              "[WORK SHEET PT].[TOTAL TIME] " & _
       "FROM [SCHEDULE SETS],[GROUPING],[WORK SHEET PT],[BARCODE] " & _
       "WHERE [GROUPING].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
             "[WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
             "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
             "[GROUPING].[WORK ORDER]='" & WO_ID & "'"
                                                            
Data3.RecordSource = sSQL
Data3.Refresh

sSQLF = "    ||^CODE|^Work Date       |Plating Operator              |^Start Time           |^Stop Time            |Total Time"
                              
MSFlexGrid3.FormatString = sSQLF

End Sub

Private Sub cmdSelected_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

Select Case PROCESS_ID
Case "BASE"

            sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                          "[SCHEDULE SETS].[DEPT_ID]                    AS [SQL DEPT_ID]," & _
                          "[WORK SHEET PT].[SET_ID]," & _
                          "[SCHEDULE SETS].[DATE_ID]," & _
                          "[SCHEDULE SETS].[SET NUMBER]                 AS [SQL SET NUMBER]," & _
                          "[SCHEDULE SETS].[TYPE_ID]                    AS [SQL TYPE_ID]," & _
                          "[SCHEDULE SETS].[EQ BASE]                    AS [SQL EQ_ID]," & _
                          "[WORK SHEET PT].[OP_ID]," & _
                            "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST] AS [SQL OPERATOR]," & _
                          "[WORK SHEET PT].[DATE_ID]," & _
                          "[WORK SHEET PT].[CODE_ID]," & _
                   "format([WORK SHEET PT].[START TIME],'h:mm AM/PM')," & _
                   "format([WORK SHEET PT].[STOP TIME],'h:mm AM/PM')," & _
                          "[WORK SHEET PT].[TOTAL TIME]," & _
                          "[SCHEDULE SETS].[BASE AMP]," & _
                   "format([SCHEDULE SETS].[BASE AMP MIN],'0')" & _
                  "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
                  "WHERE [WORK SHEET PT].[SET_ID]  = [SCHEDULE SETS].[SET_ID] AND " & _
                        "[WORK SHEET PT].[OP_ID]   = [BARCODE].[OP_ID] AND " & _
                        "[WORK SHEET PT].[CODE_ID] IN (300) AND " & _
                        "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND "
                        
            sSQL = sSQL & "[SCHEDULE SETS].[EQ BASE] = " & MACHINE_ID & " "
            
Case "FINISH"

            sSQL = "SELECT [WORK SHEET PT].[WS_ID]," & _
                          "[SCHEDULE SETS].[DEPT_ID]                    AS [SQL DEPT_ID]," & _
                          "[WORK SHEET PT].[SET_ID]," & _
                          "[SCHEDULE SETS].[DATE_ID]," & _
                          "[SCHEDULE SETS].[SET NUMBER]                 AS [SQL SET NUMBER]," & _
                          "[SCHEDULE SETS].[TYPE_ID]                    AS [SQL TYPE_ID]," & _
                          "[SCHEDULE SETS].[EQ FINISH]                  AS [SQL EQ_ID]," & _
                          "[WORK SHEET PT].[OP_ID]," & _
                            "mid([BARCODE].[FIRST],1,1) & '. ' & [BARCODE].[LAST] AS [SQL OPERATOR]," & _
                          "[WORK SHEET PT].[DATE_ID]," & _
                          "[WORK SHEET PT].[CODE_ID]," & _
                   "format([WORK SHEET PT].[START TIME],'h:mm AM/PM')," & _
                   "format([WORK SHEET PT].[STOP TIME],'h:mm AM/PM')," & _
                          "[WORK SHEET PT].[TOTAL TIME]," & _
                          "[SCHEDULE SETS].[FINISH AMP]," & _
                   "format([SCHEDULE SETS].[FINISH AMP MIN],'0')" & _
                  "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
                  "WHERE [WORK SHEET PT].[SET_ID]  = [SCHEDULE SETS].[SET_ID] AND " & _
                        "[WORK SHEET PT].[OP_ID]   = [BARCODE].[OP_ID] AND " & _
                        "[WORK SHEET PT].[CODE_ID] IN (500) AND " & _
                        "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND "
            
            sSQL = sSQL & "[SCHEDULE SETS].[EQ FINISH] = " & MACHINE_ID & " "
End Select
                            
sSQLF = "   ||^Dept_ID|^Set ID|^Create Date|^Set #|^Type         |^EQ    |^|<Operator                "
sSQLF = sSQLF & "|^Actual  Date |^Code|^Start            |^Stop         |Time|Amp   |Amp Min"
                                                    
Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub Command1_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String
                             
sSQL = "SELECT first([WORK SHEET PT].[DATE_ID])," & _
              "first([WORK SHEET PT].[CODE_ID])," & _
              "count([WORK SHEET PT].[CODE_ID]) " & _
      "FROM [WORK SHEET PT],[SCHEDULE SETS],[BARCODE] " & _
      "WHERE [WORK SHEET PT].[SET_ID] = [SCHEDULE SETS].[SET_ID] AND " & _
            "[WORK SHEET PT].[OP_ID]  = [BARCODE].[OP_ID] AND " & _
            "[WORK SHEET PT].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
      "GROUP BY [WORK SHEET PT].[DATE_ID],[WORK SHEET PT].[CODE_ID]"
 
sSQLF = "    |^DATE_ID       |^Code      |Count Code"
                              
Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub CommandRefreshSA_Click()
DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SET_ID],[DEPT_ID],[SET NUMBER],[DATE_ID],[TYPE_ID],[SERIES_ID]," & _
              "[EQ BASE]," & _
              "[EQ FINISH],[PART_SA],[SA] " & _
        "FROM [SCHEDULE SETS] " & _
        "WHERE [DATE_ID]BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# " & _
        "ORDER BY [DATE_ID] DESC"
                                   
Dim COUNT As Long

Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Table = TO_Database.OpenRecordset(sSQL)

If (TO_Table.RecordCount <> 0) Then
    Do Until TO_Table.EOF
    
        Select Case Mid(TO_Table.Fields("[SERIES_ID]"), 1, 3)
        Case "300"
        
        Case Else
                SET_ID = TO_Table.Fields("[SET_ID]")
                
                Set_Calculation
                
                TO_Table.Edit
                TO_Table.Fields("[PART_SA]") = Format(PART_SA, "0.0")          'PART SA
                TO_Table.Fields("[SA]") = Format(SA, "0.0")                    'SQ FT
                TO_Table.Update
                COUNT = COUNT + 1
                
                CommandRefreshSA.Caption = COUNT
                CommandRefreshSA.Refresh
                DoEvents
        End Select
        TO_Table.MoveNext
    Loop
End If
TO_Database.Close

MsgBox "Ok Count " & COUNT, vbInformation, "ATC"

End Sub

Private Sub CommandSA_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String
                                  
sSQL = "SELECT format(min([SCHEDULE SETS].[DATE_ID]),'mm/dd/yy')," & _
              "format(max([SCHEDULE SETS].[DATE_ID]),'mm/dd/yy')," & _
              "first([SCHEDULE SETS].[EQ BASE])," & _
              "count([SCHEDULE SETS].[EQ BASE])," & _
        "format(sum([SCHEDULE SETS].[PART_SA]),'####,##0.0')," & _
        "format(sum([SCHEDULE SETS].[SA]),'####,##0.0') " & _
       "FROM [SCHEDULE SETS],[MACHINE] " & _
       "WHERE [SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
             "[SCHEDULE SETS].[EQ BASE] = [MACHINE].[NUMBER] AND " & _
             "[MACHINE].[DEPT_ID]='PT' AND " & _
             "[MACHINE].[ACTIVE] = 1 " & _
       "GROUP BY [EQ BASE]"
 
sSQLF = "    |^Start           |^End             |^EQ       |>Count        |>Sum  SA       |>Sum  SQ FT    "
                              
Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub CommandSAF_Click()

DATE_START_ID = Format(DTPicker1.Value, "mm/dd/yyyy")
DATE_END_ID = Format(DTPicker2.Value, "mm/dd/yyyy")

Dim sSQL As String
Dim sSQLF As String
                                  
sSQL = "SELECT format(min([SCHEDULE SETS].[DATE_ID]),'mm/dd/yy')," & _
              "format(max([SCHEDULE SETS].[DATE_ID]),'mm/dd/yy')," & _
                   "first([SCHEDULE SETS].[EQ FINISH])," & _
                   "count([SCHEDULE SETS].[EQ FINISH])," & _
              "format(sum([SCHEDULE SETS].[PART_SA]),'####,##0.0')," & _
              "format(sum([SCHEDULE SETS].[SA]),'####,##0.0') " & _
       "FROM [SCHEDULE SETS],[MACHINE] " & _
       "WHERE [SCHEDULE SETS].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
             "[SCHEDULE SETS].[EQ FINISH] = [MACHINE].[NUMBER] AND " & _
             "([MACHINE].[DEPT_ID]='PT' AND " & _
              "[MACHINE].[PROCESS]='FINISH' AND " & _
              "[MACHINE].[ACTIVE] = 1) AND " & _
              "[MACHINE].[LOCATION_ID]='" & LOCATION_ID & "'  " & _
       "GROUP BY [EQ FINISH]"
 
sSQLF = "    |^Start           |^End             |^EQ       |>Count        |>Sum  SA       |>Sum  SQ FT    "
                              
Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub Form_Load()

Caption = "Plating Real Time View     " & ATC_DWG & "    " & ATC_VERSION

Data1.DatabaseName = DB_PLATING_TERMINATION
Data2.DatabaseName = DB_PLATING_TERMINATION

MSFlexGrid1.Top = 0
'MSFlexGrid1.Width = 11500
'MSFlexGrid1.Height = 4000

MSFlexGrid2.Left = 0
MSFlexGrid2.Width = 15000
MSFlexGrid2.Height = 6000

DTPicker1.Value = Date

optWeek_Click

cmdRefresh1_Click
MSFlexGrid1_Click   'select MACHINE AND PROCESS_ID

cmdRefresh_Click

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
MACHINE_ID = Val(MSFlexGrid1.Text)
  
cmdSelected.Caption = "History M_ID : " & MACHINE_ID
cmdAmp.Caption = "Amp Min >> " & MACHINE_ID

MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
sSQL = "SELECT * FROM [MACHINE] WHERE [NUMBER]=" & MACHINE_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
If (FR_Table.RecordCount <> 0) Then
    PROCESS_ID = FR_Table.Fields("[PROCESS]")
End If
FR_Table.Close
FR_Database.Close

End Sub

Private Sub optDay_Click()
DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")

cmdPrevious.Caption = "Day  <<"
cmdNext.Caption = "Day  >>"

End Sub

Private Sub Option1_Click()
cmdRefresh_Click
End Sub

Private Sub Option2_Click()
cmdRefresh_Click
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
