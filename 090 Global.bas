Attribute VB_Name = "GLOBAL"
'   File      : 090 GLOBAL.BAS
'   SW Engr   : Roger Soulagnet
'   DWG NO    : 227-090
'   Date      : 10/18/2005
'   Program   : PLATING
'
'   Exe. File : 090 Plating.EXE

'090 Plating Chg 06/14/2010 Change to WorkSheet Date tracking Dept 500 Start Operation
'090 Plating Chg 06/16/2010 Excel Download Per Tank/Series Per Day
'090 Plating Chg 06/29/2010 frmWorkSheetDT.Show
'090 Plating Chg 07/05/2010 Unplanned Downtime CODE 101 Excel Download
'090 Plating Chg 09/21/2010 [DESCRIPTION] replacement
'090 Plating Chg 09/22/2010 [RECTIFIER] addition
'090 Plating Chg 09/23/2010 Table [MACHINE] appended
'090 Plating Chg 09/30/2010 Table [WORK SHEET PT] name change
'090 Plating Chg 10/19/2010 TERMINATION And PLATING.mdb
'090 Plating Chg 10/20/2010 Operator Form Master from Termination 097
'090 Plating Chg 10/25/2010 Work Sheet DT changed to [WORK SHEET]
'090 Plating chg 12/02/2010 [3] Selected EQ History Downtime WS and DATE_ID adjust
'090 Plating chg 12/07/2010 [4] Summary Planned and Unplanned DT all Departments
'090 Plating chg 12/28/2010 Data Base String Path
'090 Plating chg 02/11/2011 OLEAN 300N SERIES_ID
'090 Plating chg 02/23/2011 OLEAN Corrections
'090 Plating chg 02/24/2011 Plating Log Sheet Master.xls moved local
'090 Plating chg 02/25/2011 Olean Excel Corrections
'090 Plating chg 03/01/2011 Olean Excel Corrections Rectifier
'090 Plating chg 03/14/2011 COMPUTER CONFIG DB
'090 Plating chg 03/23/2011 WO SCHED MASTER
'090 Plating chg 07/15/2011 SET_ID PRINT OUT
'090 Plating chg 08/05/2011 Electrical Test PLATING_ID Report
'090 Plating chg 08/09/2011 ATC 115-170 CU2 (Strike) Hg 100/800 Series
'090 Plating chg 08/10/2011 Plating Tank 18 Copper MSA
'090 Plating chg 08/17/2011 PLATING_ID SORT SCHEME
'090 Plating chg 08/19/2011 PLATING_ID SORT Scheme Order Count
'090 Plating chg 08/22/2011 Validate Plating Group msgbox form
'090 Plating chg 08/23/2011 Cases Printed out on Reports
'090 Plating chg 08/24/2011 Real Time Summaries
'090 Plating chg 09/01/2011 Real Time Summaries Continue
'090 Plating chg 09/07/2011 PLATING TABLE add decimal place 10 - 20 pf
'090 Plating chg 09/08/2011 Amp Min Real Time Form
'090 Plating chg 09/09/2011 Amp Min Correction
'090 Plating chg 10/05/2011 Configuration Form
'090 Plating chg 10/06/2011 Excel Report Unhide
'090 Plating chg 10/19/2011 DataBase_Address Modes REM,LCL,FIL
'090 Plating chg 10/24/2011 Configuration Save and AutoSBE
'090 Plating chg 11/10/2011 Add [DEPT CODE] LOC_NY,LOC_JR : [MACHINE],[BARCODE] LOCATION_ID
'090 Plating chg 11/11/2011 New Configuration Form
'090 Plating chg 12/15/2011 Copper/Nickel Tables Form MSA/Pyro
'090 Plating chg 12/19/2011 IP Address
'090 Plating chg 12/21/2011 Location_ID
'090 Plating chg 01/09/2012 Plating Sort .DAT file
'090 Plating chg 01/18/2012 Location_ID displayed with Equipment Correct Equipment Addition
'090 Plating chg 01/19/2012 Use Location_ID in Create\Reveiw Sets
'090 Plating chg 02/03/2012 SET_ID on Excel Sheet
'090 Plating chg 02/17/2012 Location_ID Main for [DEPT CODE] and [MACHINE]
'090 Plating chg 02/28/2012 New Plating Parameters MSA 800C/800E
'090 Plating chg 03/14/2012 Excel Down Load Summary per DEFEECT_ID and Cancel CODE_ID 988
'090 Plating chg 03/18/2012 Excel Down Cancel CODE_ID 988,ETC Plating and Schedule Dates
'090 Plating chg 04/30/2012 Operator Form Update
'090 Plating chg 05/01/2012 Configuration Form
'090 Plating chg 05/04/2012 [NY 285-JR 286] [NY 287 JR 288]
'090 Plating chg 05/14/2012 Plating Log Sheet Corrections and Machine Selections
'090 Plating chg 06/11/2012 Add 296,297 Cancel Plating Summary Review
'090 Plating chg 07/03/2012 Lift Term Download
'090 Plating chg 07/06/2012 Lift Term Sorted Order
'090 Plating chg 07/10/2012 Update Schedule
'090 Plating chg 07/11/2012 Lift Term Duplicates
'090 Plating chg 07/23/2012 Encap = (Inspect and  Position 9 ATCPart "E")
'090 Plating chg 07/24/2012 '830C','830E'
'090 Plating chg 08/08/2012 Work Order Tank Information on Excel Download FIND_DEPT_ID = 1 Cancel ID 988
'090 Plating chg 10/01/2012 Plating EQ revision added Juarez Barrel Line
'090 Plating chg 10/03/2012 Update Dept Codes
'090 Plating chg 10/05/2012 Updater Tested
'090 Plating chg 10/11/2012 Juarez Dept Updated in Code
'090 Plating chg 10/18/2012 Data Base Mode 3
'090 Plating chg 10/25/2012 SQL Search And Review
'090 Plating chg 11/05/2012 100B Lift Term NY
'090 Plating chg 12/12/2012 NY/JR Optinal View
'090 Plating chg 12/20/2012 NY/JR  Table [DEPT_CODE]
'090 Plating chg 02/01/2013 800 MSA
'090 Plating chg 02/07/2013 Nominca
'090 Plating chg 02/19/2013 Correction Path and Lookup Function Exit
'090 Plating chg 03/08/2013 Add 100E Lift Term
'090 Plating chg 03/13/2013 Unlock Lot Number no Create and Review
'090 Plating chg 04/23/2013 Amp Time Correction
'090 Plating chg 05/02/2013 Plating Validation SBE 100/700
'090 Plating chg 05/05/2013 Case [0] Sortable Grouping
'090 Plating chg 05/31/2013 Tank/Case Per Day LOCATION_ID from NY REM
'090 Plating chg 08/07/2013 Main Screen Resizing
'090 Plating chg 08/27/2013 ConfigComputer_DB
'090 Plating chg 09/30/2013 ENABLE1_ATC_TABLES Correction
'090 Plating chg 01/09/2014 Suammary Correction for JR
'090 Plating chg 01/10/2014 TBL CALCULATION EQ
'090 Plating chg 02/19/2014 Expanded SQL
'090 Plating chg 03/21/2014 sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
'090 Plating chg 03/31/2014 Format(Time, "hh AM/PM") = "01 AM"
'090 Plating chg 04/02/2014 ConfigComputer_DB (2)
'090 Plating chg 04/28/2014 [DEPT_JR_ID]=" & DEPT_ID Plating Log Excel
'090 Plating chg 06/06/2014 Common Data Base Mode
'090 Plating chg 06/23/2014 TOTAL_QTY Surface Area Error
'090 Plating chg 07/09/2014 STRIKE1  [SK1 AMP]/NUMBER_HEADS
'090 Plating chg 07/14/2014 Excel Sheet Correction
'090 Plating chg 08/25/2014 TBL_EXECUTABLE Each Rectifier Strike
'090 Plating chg 09/09/2014 Plate Grouping Message
'090 Plating chg 10/28/2014 Configuration Form
'090 Plating chg 12/03/2014 Extended Voltages
'090 Plating chg 02/11/2015 [16] Set Count per EQ Base
'090 Plating chg 03/09/2015 CONFIG STATUS deleted
'090 Plating chg 06/30/2015 Plating Report .DAT and .TXT
'090 Plating chg 08/25/2015 Plating_Sort_DB
'090 Plating chg 08/28/2015 ExcelReportOlean
'090 Plating chg 11/10/2015 FROM [TBL SBE] adjust
'090 Plating chg 05/20/2016 Excel Work Order DB_OEE_VISUAL
'090 Plating chg 05/24/2016 "Mixed Tolerance Parts"
'090 Plating chg 10/03/2016 BASE Sum/SQ FT
'090 Plating chg 11/08/2016 800 A/B 600 S/F/L
'090 Plating chg 11/16/2016 EACH RECTIFER BARREL ONLY
'090 Plating chg 12/20/2016 L SERIES EXTEND
'090 Plating chg 01/26/2017 frmAlert.Show
'090 Plating chg 02/15/2017 "NO PROCESAR ESTA ORDEN"
'090 Plating chg 06/14/2017 DB_PLATING_TABLES
'090 Plating chg 06/29/2017 JAX Barrel/SBE
'090 Plating chg 07/27/2017 Configuration
'090 Plating chg 12/15/2017 Configuration Data Base Management
'090 Plating chg 01/31/2018 MAX_SET_NUM
'090 Plating chg 04/16/2018 Update Schedule
'090 Plating chg 06/05/2018 700 SERIES AND 100 SERIES SEPARATED
'090 Plating chg 08/30/2018 TYPE_CU = "PYRO"       17, 74, 75, 76
'090 Plating chg 12/05/2018 Series 600/800 Barrels
'090 Plating chg 04/30/2019 ATC Part Code less than 9 characters
'090 Plating chg 05/01/2019 ATCPart EXTEND SPACES
'090 Plating chg 07/06/2020 WORK ORDERS 10 CHAR
Option Explicit

Public Const TBL_ATC_DWG As String = "227-090"
Public Const TBL_NAME As String = "Plating Schedule"
Public Const TBL_EXECUTABLE As String = "090 Plating"

Public Const ATC_DWG As String = "DWG NO 227-090 REV A"
Public Const ATC_VERSION As String = "01/11/2021 T"

Public LOC_SOURCE_ID As String
Public DB_SOURCE_ID As Long
 
Public cCONFIGURATION As Integer

Public ENABLE1_ATC_TABLES As Integer
Public ENABLE2_DEPT_CODES As Integer
Public ENABLE3_CHEMISTRY As Integer
Public ENABLE4_EQ As Integer
Public ENABLE5_SP_TEST As Integer

Public TYPE_ID As String            'TANK/SBE
Public PROCESS_ID As String         'BASE/FINISH

Public ALERT_MESSAGE As String

Public BASE_ID As String
Public FINISH_ID As String

Public SHIFT_ID As String           'AM/PM

'Public CODE_ID As Long              'CODE_ID CODE is Department Code

Public GP_ID As Long

Public START_QTY_ID As Long
Public CASE_SIZE_ID As String       'A/B/C/E
'Public SERIES_ID As Integer         '0/1 [100/700,200/900]
Public SERIES_CASE_ID As String

Public TYPE_CU As String            'MSA,PYRO

Public NUMBER_HEADS As Integer
Public RECTIFIER As Integer
Public EQ_BASE_ID As Long
Public EQ_FINISH_ID As Long
Public GEAR_1_QTY As Long
Public GEAR_2_QTY As Long
Public GEAR_3_QTY As Long
Public GEAR_4_QTY As Long
Public GEAR_MAX_QTY As Long

Public TOTAL_QTY As Long
Public Shot_Qty As Long

Public SF As Double
Public SA As Double

Public PART_SA As Double
Public Media_SA As Double
Public Sum_SA As Double

Public ASF1 As Double
Public MIN1 As Integer
Public ASF2 As Double
Public MIN2 As Integer
Public ASF3 As Double
Public MIN3 As Integer

Public SKTASF1 As Double
Public SKTMIN1 As Integer
Public SKTASF2 As Double
Public SKTMIN2 As Integer

Public Const MAX_A_CASE As Long = 60000
Public Const MAX_B_CASE As Long = 10000
Public Const MAX_C_CASE As Long = 2000
Public Const MAX_E_CASE As Long = 1000

'Public DEPT_ID As Long              '537,524,530,525,532,285,540,544,546,529,535,541
Public SHOT_ID As Integer           '100/200
Public SPEED_ID As Integer          '100/200

'Public LETTER_ID As String
'Public SET_ID As Long               'SET ID
Public SET_NUMBER As Long           'SET NUMBER

Public FG_ID As Long        'FINSIHING GROUP ID
Public FLAG_WO_ID  As Long
Public FLAG_ID(1000) As Long


'   BIN TABLE LIMITS AND TOLERANCE CODES
Public gdBinLimit(10) As Double
Public gsBinTol(10) As String

Public FR_Database As Database
Public FR_WorkSpace As Workspace
Public FR_Table As Recordset

'Public FR2_Database As Database
'Public FR2_WorkSpace As Workspace
'Public FR2_Table As Recordset

Public FR3_Database As Database
Public FR3_WorkSpace As Workspace
Public FR3_Table As Recordset

Public FR4_Database As Database
Public FR4_WorkSpace As Workspace
Public FR4_Table As Recordset

Public TO_Database As Database
Public TO_WorkSpace As Workspace
Public TO_Table As Recordset

 
Public Sub Main()
  
If App.PrevInstance Then
    End
End If
   
Get_User
IP_ADDRESS = GetIPAddress

Configuration (FREAD)

' Copy from GV server the next databases,
' server Juarez = c:db/juarez
' copy the excel file plating log sheet master.xls c:atc\

Dim SourceFile As String
Dim DestinationFile As String
Dim FSO As New FileSystemObject

SourceFile = SERVER_DB_GR & "WO SCHED MASTER.mdb"
DestinationFile = "C:\DB\JR\WO SCHED MASTER.mdb"
FSO.CopyFile SourceFile, DestinationFile, True  '

SourceFile = SERVER_DB_GR & "PLATING JR.MDB"
DestinationFile = "C:\DB\JR\PLATING JR.MDB"
FSO.CopyFile SourceFile, DestinationFile, True

SourceFile = SERVER_DB_GR & "OEE SPM JR MASTER.mdb"
DestinationFile = "C:\DB\JR\OEE SPM JR MASTER.mdb"
FSO.CopyFile SourceFile, DestinationFile, True

SourceFile = SERVER_DB_GR & "ATC Electrical Test JR.MDB"
DestinationFile = "C:\DB\JR\ATC Electrical Test JR.MDB"
FSO.CopyFile SourceFile, DestinationFile, True

SourceFile = SERVER_DB_GR & "ATC Plating Tables.MDB"
DestinationFile = "C:\DB\JR\ATC Plating Tables.MDB"
FSO.CopyFile SourceFile, DestinationFile, True

MsgBox "databases copied from server"

Select Case Mid(IP_ADDRESS, 1, 8)
Case "10.0.38."
                    LOCATION_ID = "JR"
                    DataBase_MODE = DATABASE_MODE_REM_JUAREZ
                    DB_SOURCE_ID = 0
Case Else
                    LOCATION_ID = "NY"
                    DataBase_MODE = DATABASE_MODE_REM_NY
                    DB_SOURCE_ID = 1
End Select

LOCATION_ID = "JR"
DataBase_MODE = DATABASE_MODE_REM_JUAREZ
DB_SOURCE_ID = 0

LoveLetter

'
'FILE IS NOT NEEDED
'
Configuration (FWRITE)

ConfigComputer_DB (0)

DataBase_Address

Configuration (FWRITE)

'LOCATION_ID = "JR"
'RefreshDB_WO

Update_WO_Schedule

Select Case 0
Case 0
        frmMain.Show
Case 1
        ALERT_MESSAGE = "NO PROCESAR ESTA ORDEN"
        Select Case ALERT_MESSAGE
        Case "ok"

        Case Else
                frmAlert.Show vbModal
        End Select
Case 1
        frmSBECalculation.Show              'SBE Plating Calculations
Case 2
        frmSBEParameters.Show               'SBE Plating Parameter Tables
Case 3
        frmBarrelCalculation.Show           'Barrel Plating Calculations
Case 4
        frmBarrelParametersCU.Show          'Barrel Plating Parameter Tables
Case 5
        frmBarrelParametersNickel.Show      'Barrel Plating Parameter Tables
Case 6
        DB_PLATING_TERMINATION = SERVER_DB_NY & "PLATING JR.MDB"
        frmConfiguration.Show               'Configuration Plating and Enables
Case 7
        frmDept.Show                        'ATC Plating Dept Codes
Case 8
        frmEquipment.Show                   'ATC Plating Equipment
Case 9
        frmOperator.Show                    'ATC Operators Plating Department
Case 10
        frmPlatingCalculation.Show          'Plating Calculation
Case 11
        frmReal_Time.Show                   'Plating Real Time View
Case 12
        frmSetCreate.Show                   'Create Schedule/Grouping   Sets
Case 13
        frmSetReview.Show                   'Review Schedule/Grouping Sets
Case 14
        frmSummary.Show                     'Summary Review
Case 15
        frmSQL.Show                         'SQL Review
Case 16
        frmWorkSheetDT.Show                 'Plating Worksheet Down Time
Case 17
        frmWorkSheetR.Show                  'OEE Plating Worksheet Review
End Select
 
End Sub
'
'   Lot Decode  DWG 108-1192
'               MANUFACTURING LOT NUMBER
'               IDENTIFICATION SYSTEM PROCEDURE
'
'   RETURNS DESIGN VALUE
'
Function LotDV(ATC_PART As String) As Single
  
Select Case Mid$(ATC_PART, 1, 3)
Case "100", "180", "200", "700", "710", "800", "830", "900"
Case Else
        LotDV = 0
        Exit Function
End Select

Select Case Mid$(ATC_PART, 4, 1)
Case "A", "B", "C", "E", "R"
Case Else
        LotDV = 0
        Exit Function
End Select
     
'---  CONVERT TO A DOUBLE PF VALUE
If (Mid$(ATC_PART, 6, 1) = "R") Then
   LotDV = Val(Mid$(ATC_PART, 5, 1) & "." & Mid$(ATC_PART, 7, 1))
Else
   LotDV = Val(Mid$(ATC_PART, 5, 2) & "E" & Mid$(ATC_PART, 7, 1))
End If
               
End Function


Public Sub DataBase_Address()

'================================== WO SCHED MASTER
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
                DB_MASTER_SCHEDULE = SERVER_DB_NY & "WO SCHED MASTER.mdb"
Case DATABASE_MODE_LCL
                DB_MASTER_SCHEDULE = "C:\ATC\WO SCHED MASTER.mdb"
Case DATABASE_MODE_FIL
                DB_MASTER_SCHEDULE = ADDR_DB_MASTER_SCHED       'SERVER_ADDR_MAS
Case DATABASE_MODE_REM_JUAREZ
                DB_MASTER_SCHEDULE = "C:\ATC\WO SCHED MASTER.mdb"
End Select

'================================== PLATING DATABASE TABLES
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
                DB_PLATING_TERMINATION = SERVER_DB_NY & "TERMINATION And PLATING.MDB"
Case DATABASE_MODE_REM_JR
                DB_PLATING_TERMINATION = SERVER_DB_NY & "PLATING JR.MDB"
Case DATABASE_MODE_LCL
                DB_PLATING_TERMINATION = "C:\ATC\PLATING JR.MDB"
Case DATABASE_MODE_FIL
                DB_PLATING_TERMINATION = ADDR_DB_PLATING        'SERVER_ADDR_PLT
Case DATABASE_MODE_REM_JUAREZ
                DB_PLATING_TERMINATION = SERVER_DB_JR & "PLATING JR.MDB"
End Select
 
'================================== OEE VISUAL INSPECTION

Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
                DB_OEE_VISUAL = SERVER_DB_NY & "VISUAL INSPECTION.mdb"
Case DATABASE_MODE_LCL
                DB_OEE_VISUAL = "C:\ATC\VISUAL INSPECTION.mdb"
                DB_OEE_VISUAL = "C:\ATC\OEE SPM JR MASTER.mdb"
Case DATABASE_MODE_FIL
                DB_OEE_VISUAL = ADDR_DB_PLATING
Case DATABASE_MODE_REM_JUAREZ
                DB_OEE_VISUAL = SERVER_DB_JR & "OEE SPM JR MASTER.mdb"
End Select
 
'================================== PLATING SORT
'                                   REPORT DB_REPORT_ADDR & SET_ID & LETTER_ID & ".TXT" and ".DAT"

Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            DB_REPORT_ADDR = "\\NY-ENG\SPC Network\ATC Report\Plating\Plating "
Case DATABASE_MODE_LCL
            DB_REPORT_ADDR = "C:\ATC\Plating\"  'NA
Case DATABASE_MODE_FIL
            DB_REPORT_ADDR = "C:\ATC\Plating\"  'NA
Case DATABASE_MODE_REM_JUAREZ
            DB_REPORT_ADDR = "\\Juarezdc1\Public\ATC\REPORT\Plating "
End Select


'================================== EXCEL REPORT FOR PLATING SETUP AMPS AND AMP MIN

Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            PLATING_LOG_SHEET = "C:\ATC\Plating Log Sheet Master.xls"
Case DATABASE_MODE_LCL
            PLATING_LOG_SHEET = "C:\ATC\Plating Log Sheet Master.xls"
Case DATABASE_MODE_FIL
            PLATING_LOG_SHEET = ADDR_EXCEL_REPORT
Case DATABASE_MODE_REM_JUAREZ
            PLATING_LOG_SHEET = "C:\ATC\Plating Log Sheet Master.xls"
End Select
 
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            DB_PLATING_SORT_TABLE = SERVER_DB_NY & "ATC Electrical Test JR.MDB"
Case DATABASE_MODE_REM_JUAREZ
            DB_PLATING_SORT_TABLE = SERVER_DB_JR & "ATC Electrical Test JR.MDB"
End Select
 
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY, DATABASE_MODE_REM_JR
            DB_PLATING_TABLES = SERVER_DB_NY & "ATC Plating Tables.MDB"
Case DATABASE_MODE_REM_JUAREZ
            DB_PLATING_TABLES = SERVER_DB_JR & "ATC Plating Tables.MDB"
End Select
 
End Sub


Public Sub RefreshDB_WO()

On Error GoTo Network_Mode_Error

Dim SourceFile As String
Dim DestinationFile As String

Select Case LOCATION_ID
Case "JR"
            SourceFile = SERVER_DB_JR & "WO SCHED MASTER.MDB"
            DestinationFile = "C:\ATC\WO SCHED MASTER.MDB"
            Screen.MousePointer = vbHourglass
            
            FileCopy SourceFile, DestinationFile
            
            Screen.MousePointer = vbDefault
End Select

'MsgBox "Succesful", vbInformation, "ATC DataBase System"

Exit Sub
Network_Mode_Error:

Screen.MousePointer = vbDefault
'MsgBox "Unsuccesful", vbCritical, "ATC DataBase System"
  
Exit Sub

End Sub


Public Sub Configuration(iMode As Integer)

On Error GoTo ErrorConfig

Dim sFilename As String
sFilename = "C:\ATC\090 Configuration.TXT"

If Len(Dir$(sFilename)) = 0 Then
        'IF FILE DOES'NT EXSIST LOAD DEFAULTS
        iMode = FWRITE
        DataBase_MODE = 0
        DB_SOURCE_ID = 0
        LOCATION_ID = "NY"
        
        ADDR_DB_MASTER_SCHED = "\\NY-ENG\SPC Network\Data Base\WO SCHED MASTER.mdb"
        ADDR_DB_PLATING = "\\NY-ENG\SPC Network\Data Base\TERMINATION And PLATING.MDB"
        ADDR_DB_VINSPECTION = "\\NY-ENG\SPC Network\Data Base\VISUAL INSPECTION.MDB"
        ADDR_PLATING_REPORT = "\\NY-ENG\SPC Network\ATC Report\Plating\"
        ADDR_EXCEL_REPORT = "C:\ATC\Plating Log Sheet Master.xls"
        SERVER_PATH = "\\Juarezdc1\Public\ATC\XTR NY JR REVC"
                
        ENABLE1_ATC_TABLES = 1
        ENABLE2_DEPT_CODES = 1
        ENABLE3_CHEMISTRY = 1
        ENABLE4_EQ = 1
        ENABLE5_SP_TEST = 1
End If
 
Dim iFilenum As Integer
iFilenum = FreeFile

Dim sTemp As String
 
Select Case iMode
Case FREAD
            Open sFilename For Input Shared As iFilenum
            Input #iFilenum, sTemp: DataBase_MODE = InfoVal(sTemp)
            Input #iFilenum, sTemp: LOCATION_ID = InfoStr(sTemp)
            
            Input #iFilenum, sTemp: ADDR_DB_MASTER_SCHED = InfoStr(sTemp)
            Input #iFilenum, sTemp: ADDR_DB_PLATING = InfoStr(sTemp)
            Input #iFilenum, sTemp: ADDR_DB_VINSPECTION = InfoStr(sTemp)
            Input #iFilenum, sTemp: ADDR_PLATING_REPORT = InfoStr(sTemp)
            Input #iFilenum, sTemp: ADDR_EXCEL_REPORT = InfoStr(sTemp)
            Input #iFilenum, sTemp: SERVER_PATH = InfoStr(sTemp)
            
            Input #iFilenum, sTemp: DB_SOURCE_ID = InfoVal(sTemp)
            
            Input #iFilenum, sTemp: ENABLE1_ATC_TABLES = InfoVal(sTemp)
            Input #iFilenum, sTemp: ENABLE2_DEPT_CODES = InfoVal(sTemp)
            Input #iFilenum, sTemp: ENABLE3_CHEMISTRY = InfoVal(sTemp)
            Input #iFilenum, sTemp: ENABLE4_EQ = InfoVal(sTemp)
            Input #iFilenum, sTemp: ENABLE5_SP_TEST = InfoVal(sTemp)
            
            Close iFilenum
Case FWRITE
            Open sFilename For Output Shared As #iFilenum
            Print #iFilenum, "[01] DB [0:4][REM NY:REM JR]           ="; DataBase_MODE
            Print #iFilenum, "[02] LOCATION_ID   [NY:JR]             ="; LOCATION_ID
            
            Print #iFilenum, "[03] ADDR_DB_MASTER_SCHED              ="; ADDR_DB_MASTER_SCHED
            Print #iFilenum, "[04] ADDR_DB_PLATING                   ="; ADDR_DB_PLATING
            Print #iFilenum, "[05] ADDR_DB_VINSPECTION               ="; ADDR_DB_VINSPECTION
            Print #iFilenum, "[06] ADDR_PLATING_REPORT               ="; ADDR_PLATING_REPORT
            Print #iFilenum, "[07] ADDR_EXCEL_REPORT                 ="; ADDR_EXCEL_REPORT
            
            Print #iFilenum, "[08] SERVER_PATH                       ="; SERVER_PATH
            
            Print #iFilenum, "[09] DB_SOURCE_ID [View Option NY:JR]  ="; DB_SOURCE_ID
            Print #iFilenum, "[10] ENABLE1_ATC_TABLES                ="; ENABLE1_ATC_TABLES
            Print #iFilenum, "[11] ENABLE2_DEPT_CODES         [CMD]  ="; ENABLE2_DEPT_CODES
            
            Print #iFilenum, "[12] ENABLE3_CHEMISTRY          [NU]   ="; ENABLE3_CHEMISTRY
            Print #iFilenum, "[13] ENABLE4_EQ                 [NU]   ="; ENABLE4_EQ
            Print #iFilenum, "[14] ENABLE5_SP_TEST            [NA]   ="; ENABLE5_SP_TEST
            Close iFilenum
End Select


Exit Sub
ErrorConfig:

Select Case Err.Number

Case Else
    
    MsgBox "File Error " & sFilename, vbCritical, "SPC Casting"
    End
End Select

End Sub

