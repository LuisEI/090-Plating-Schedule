VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMsgBoxUpdate 
   Caption         =   "Message Box"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "090 MessageBox Update.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   5685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefreshM_ID 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Refresh M_ID"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   3075
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Copy PLATING JR.MDB to Server"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpdateTables 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ATC Electrical Test Tables"
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpdateSched 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update DEPT ,MACHINE"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpdateSchedule 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Master Schedule"
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   3075
   End
   Begin VB.CommandButton cmdUpServer 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Records to Server"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   3075
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   840
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
      Format          =   24576001
      CurrentDate     =   38117
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
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
      Format          =   24576001
      CurrentDate     =   38117
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      Caption         =   "Data Base Updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   3180
   End
End
Attribute VB_Name = "frmMsgBoxUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()

On Error GoTo Network_Copy_Error

Dim sSQL As String

sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='1234567'"

frmMain.Data1.DatabaseName = DB_MASTER_SCHEDULE
frmMain.Data1.RecordSource = sSQL
frmMain.Data1.Refresh

frmMain.Data2.DatabaseName = DB_MASTER_SCHEDULE
frmMain.Data2.RecordSource = sSQL
frmMain.Data2.Refresh

frmMain.Data3.DatabaseName = DB_MASTER_SCHEDULE
frmMain.Data3.RecordSource = sSQL
frmMain.Data3.Refresh

Dim SourceFile As String
Dim DestinationFile As String

Select Case LOCATION_ID
Case "JR"
        
        SourceFile = "C:\ATC\PLATING JR.MDB"
        DestinationFile = "\\Jzfs\ATC\Data Base\PLATING JR.MDB"
                                  
        Screen.MousePointer = vbHourglass
        
        FileCopy SourceFile, DestinationFile
        
        Screen.MousePointer = vbDefault
        
        MsgBox "Successful", vbInformation, "ATC DataBase System"

End Select

Dim sSQLF As String

frmMain.Data1.DatabaseName = DB_PLATING_TERMINATION

sSQL = "SELECT [OP_ID], [FIRST] & ' ' & [LAST] " & _
       "FROM [BARCODE] " & _
       "WHERE [ACTIVE] = 1 AND [DEPT_ID]='PT' AND [LOCATION_ID]='" & LOCATION_ID & "' " & _
       "ORDER BY [LAST],[FIRST]"
sSQLF = "    ||<Operator                                "

frmMain.Data1.RecordSource = sSQL
frmMain.Data1.Refresh
frmMain.MSFlexGrid1.FormatString = sSQLF

frmMain.Data2.DatabaseName = DB_PLATING_TERMINATION

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT [DEPT_ID],[DESCRIPTION],[SBE] " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "
               
        sSQL = sSQL & "[LOC_NY]='NY' ORDER BY [DEPT_ID]"
Case "JR"
        sSQL = "SELECT [DEPT_JR_ID],[DESCRIPTION],[SBE] " & _
               "FROM [DEPT CODE] " & _
               "WHERE [ACTIVE]=1 AND "
        sSQL = sSQL & "[LOC_JR]='JR' ORDER BY [DEPT_ID]"
End Select
sSQLF = "    |^Code ID|<Base / Finish              |^SBE  "
                                   
frmMain.Data2.RecordSource = sSQL
frmMain.Data2.Refresh
frmMain.MSFlexGrid2.FormatString = sSQLF

frmMain.Data3.DatabaseName = DB_PLATING_TERMINATION

sSQL = "SELECT [NUMBER],[TYPE],[DESCRIPTION],mid([PROCESS],1,1)& mid(lcase([PROCESS]),2,5),[LOCATION_ID]  " & _
       "FROM [MACHINE]  " & _
       "WHERE [DEPT_ID]='PT' AND [ACTIVE] = 1 AND "
        
Select Case LOCATION_ID
Case "NY"
       sSQL = sSQL & "[LOCATION_ID]='NY' ORDER BY [LOCATION_ID] DESC,[TYPE],[NUMBER],mid([PROCESS],1,1)"
Case "JR"
       sSQL = sSQL & "[LOCATION_ID]='JR' ORDER BY [LOCATION_ID] DESC,[TYPE],[NUMBER],mid([PROCESS],1,1)"
End Select
             
sSQLF = "    |^M ID|^TYPE       |Description               |^Process |^L_ID"

frmMain.Data3.RecordSource = sSQL
frmMain.Data3.Refresh
frmMain.MSFlexGrid3.FormatString = sSQLF

Exit Sub
Network_Copy_Error:

Screen.MousePointer = vbDefault

MsgBox "Unsuccessful", vbCritical, "ATC DataBase System"
  
Exit Sub

End Sub

Private Sub cmdRefreshM_ID_Click()

Screen.MousePointer = vbHourglass
Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT [MACHINE_ID] FROM [WORK SHEET] "
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
   Do Until FR_Table.EOF
   FR_Table.Edit
        FR_Table.Fields("[MACHINE_ID]") = MACHINE_ID
        FR_Table.Update
        FR_Table.MoveNext
   Loop
End If
FR_Table.Close
FR_Database.Close

Screen.MousePointer = vbDefault

MsgBox "Refresh Complete", vbInformation, "Data Base Operation"

End Sub

Private Sub cmdUpdateSched_Click()

On Error GoTo Network_Mode_Error

'============ STEP 1
Select Case 0
Case 0
        'Dim SourceFile As String
        'Dim DestinationFile As String
        
        'SourceFile = "\\Jzfs\ATC\Data Base\BARCODE JR.MDB"
        'DestinationFile = "C:\ATC\BARCODE JR.MDB"
         
        'Screen.MousePointer = vbHourglass
        
        'FileCopy SourceFile, DestinationFile
        
        'Screen.MousePointer = vbDefault
        
        Dim sSQL As String
        sSQL = "SELECT * FROM [BARCODE]"
        frmMain.Data2.RecordSource = sSQL
        frmMain.Data2.Refresh
        frmMain.Data3.RecordSource = sSQL
        frmMain.Data3.Refresh
End Select

MsgBox "Step 1 Succesful", vbInformation, "ATC DataBase System"

' Set 2 DELETE TABLE
Dim DBs As Database

Dim sFilenameTo As String
sFilenameTo = "C:\ATC\PLATING JR.MDB"

Dim sTableName As String
sTableName = "DROP TABLE [MACHINE]"
 
Set DBs = OpenDatabase(sFilenameTo)

DBs.Execute sTableName

sTableName = "DROP TABLE [DEPT CODE]"
 
Set DBs = OpenDatabase(sFilenameTo)

DBs.Execute sTableName

DBs.Close

MsgBox "Set 2 Delete Complete", vbInformation, "Data Base Operation"

Screen.MousePointer = vbHourglass
 
' Copy Table from One DataBase to Another DataBase


Set FR_Database = OpenDatabase("C:\ATC\PLATING SPEC TABLES.MDB")

sTableName = "[MACHINE]"
 
'TO DATABASE
                                                   
sSQL = "SELECT " & sTableName & ".* " & _
       "INTO " & sTableName & " IN '" & sFilenameTo & "' " & _
       "FROM " & sTableName & ""
                                          
FR_Database.Execute sSQL

sTableName = "[DEPT CODE]"
 
'TO DATABASE
                                                   
sSQL = "SELECT " & sTableName & ".* " & _
       "INTO " & sTableName & " IN '" & sFilenameTo & "' " & _
       "FROM " & sTableName & ""
                                          
FR_Database.Execute sSQL

FR_Database.Close

Screen.MousePointer = vbDefault

MsgBox "Add Table Complete [From - To Option]", vbInformation, "Data Base Operation"

Exit Sub
Network_Mode_Error:

Screen.MousePointer = vbDefault
MsgBox "Unsuccesful", vbCritical, "ATC DataBase System"
  
Exit Sub

End Sub

Private Sub cmdUpdateSchedule_Click()

On Error GoTo Network_Mode_Error

Dim SourceFile As String
Dim DestinationFile As String
 
SourceFile = "S:\ATC\Data Base\WO SCHED MASTER.MDB"
DestinationFile = "C:\ATC\WO SCHED MASTER.MDB"
        
Select Case LOCATION_ID
Case "JR"
        
        Screen.MousePointer = vbHourglass
        
        FileCopy SourceFile, DestinationFile
        
        Screen.MousePointer = vbDefault

End Select

MsgBox "Succesful", vbInformation, "ATC DataBase System"

Exit Sub
Network_Mode_Error:

Screen.MousePointer = vbDefault
MsgBox "Unsuccesful", vbCritical, "ATC DataBase System"
  
Exit Sub

End Sub

Private Sub cmdUpdateTables_Click()

On Error GoTo Network_Mode_Error_Tables

Dim SourceFile As String
Dim DestinationFile As String

SourceFile = "\\Jzfs\ATC\Data Base\ATC Electrical Test.MDB"
DestinationFile = "C:\ATC\SORT\ATC Electrical Test.MDB"

Screen.MousePointer = vbHourglass

FileCopy SourceFile, DestinationFile

Screen.MousePointer = vbDefault

MsgBox "Succesful", vbInformation, "ATC Data Base System"

Exit Sub
Network_Mode_Error_Tables:

Screen.MousePointer = vbDefault
MsgBox "Unsuccesful", vbCritical, "ATC Data Base System"
  
Exit Sub

End Sub

Private Sub cmdUpServer_Click()

DATE_START_ID = DTPicker1.Value
DATE_END_ID = DTPicker2.Value

Dim START_TIME As Double
START_TIME = Timer

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

'Set TO_Database = OpenDatabase("C:\ATC\Data Base\OEE WORK STATION JR.MDB")

Set TO_Database = OpenDatabase("\\Jzfs\ATC\Data Base\OEE WORK STATION JR.MDB")

Dim sSQL As String

sSQL = "SELECT * FROM [WORK SHEET]" & _
       "WHERE         ( ( LEN([PRODUCT]) IN (12) AND " & _
             "(ISNUMERIC(MID([PRODUCT],1,6))= TRUE AND " & _
              "ISNUMERIC(MID([PRODUCT],7,6))= TRUE) ) OR " & _
                      " (LEN([PRODUCT]) IN (10) AND " & _
             "(ISNUMERIC(MID([PRODUCT],1,1))= FALSE AND " & _
              "ISNUMERIC(MID([PRODUCT],2,2))= TRUE  AND " & _
              "ISNUMERIC(MID([PRODUCT],4,1))= FALSE AND " & _
              "ISNUMERIC(MID([PRODUCT],8,3))= FALSE))) AND " & _
            "( [DATE] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# )"
       
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
 
        Dim I As Long
        Dim iCount As Long
        
        Do Until FR_Table.EOF
         
                sSQL = "SELECT * FROM [WORK SHEET] " & _
                        "WHERE  [PRODUCT]    ='" & FR_Table.Fields("[PRODUCT]") & "'  AND " & _
                               "[MACHINE_ID] = " & FR_Table.Fields("[MACHINE_ID]") & " AND " & _
                               "[CODE]       = " & FR_Table.Fields("[CODE]")
                    
                Set TO_Table = TO_Database.OpenRecordset(sSQL)
         
                If (TO_Table.RecordCount = 0) Then
                    TO_Table.AddNew
                    I = I + 1
                Else
                    TO_Table.Edit
                End If
                
                iCount = iCount + 1
                
               ' TO_Table.Fields("[CONFIG]") = FR_Table.Fields("[CONFIG]")
                TO_Table.Fields("[OP_ID]") = FR_Table.Fields("[OP_ID]")
                TO_Table.Fields("[MACHINE_ID]") = FR_Table.Fields("[MACHINE_ID]")
                TO_Table.Fields("[DATE]") = FR_Table.Fields("[DATE]")
                TO_Table.Fields("[CODE]") = FR_Table.Fields("[CODE]")
                
                TO_Table.Fields("[PRODUCT]") = FR_Table.Fields("[PRODUCT]")
                TO_Table.Fields("[ATC PART]") = FR_Table.Fields("[ATC PART]")
                
'                TO_Table.Fields("[PRODUCT]") = Mid(FR_Table.Fields("[PRODUCT]"), 1, 12)
'                TO_Table.Fields("[ATC PART]") = Mid(FR_Table.Fields("[ATC PART]"), 1, 20)
                
                
                TO_Table.Fields("[START TIME]") = FR_Table.Fields("[START TIME]")
                TO_Table.Fields("[STOP TIME]") = FR_Table.Fields("[STOP TIME]")
                TO_Table.Fields("[TOTAL TIME]") = FR_Table.Fields("[TOTAL TIME]")
                TO_Table.Fields("[UNITS PRODUCED]") = FR_Table.Fields("[UNITS PRODUCED]")
                TO_Table.Fields("[DEFECTS]") = FR_Table.Fields("[DEFECTS]")
                TO_Table.Fields("[RESTOCK]") = FR_Table.Fields("[RESTOCK]")
                TO_Table.Update
                
                FR_Table.MoveNext
                
                cmdUpServer.Caption = iCount
                cmdUpServer.Refresh
                DoEvents
        Loop
End If

FR_Database.Close
TO_Database.Close

Dim sBuff As String
sBuff = Format(Timer - START_TIME, "0") & Seconds

cmdUpServer.Caption = "Update Records to Server"
MsgBox "Server Update Complete Add " & I, vbInformation, "ATC " & sBuff

End Sub

Private Sub Form_Load()

Caption = "Juarez Mexico Data Base Update            " & ATC_DWG & "    " & ATC_VERSION

DTPicker1.Value = Date
DTPicker2.Value = Date

DTPicker1.Value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.Value), "mm/dd/yyyy")
DTPicker2.Value = Format(DateAdd("d", 6, DTPicker1.Value), "mm/dd/yyyy")


End Sub
