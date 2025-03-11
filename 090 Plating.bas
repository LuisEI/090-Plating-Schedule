Attribute VB_Name = "Plating"
Option Explicit

'
'   SORT TOLERANCES
'
Public Const TOLA As Double = 0.05      ' pf
Public Const TOLB As Double = 0.1       ' pf
Public Const TOLC As Double = 0.25      ' pf
Public Const TOLD As Double = 0.5       ' pf

Public Const TOLE As Double = 0.005     ' percent
Public Const TOLF As Double = 0.01      ' percent
Public Const TOLG As Double = 0.02      ' percent
Public Const TOLJ As Double = 0.05      ' percent
Public Const TOLK As Double = 0.1       ' percent
Public Const TOLM As Double = 0.2       ' percent
Public Const TOLN As Double = 0.3       ' percent

 
Public Sub BarrelCalculation()

SKTASF1 = 0
SKTMIN1 = 0
ASF1 = 0
MIN1 = 0
ASF2 = 0
MIN2 = 0
ASF3 = 0
MIN3 = 0

SKTASF2 = 0
SKTMIN2 = 0

'================================================================================
'   [1]  SURFACE AREA PART
'================================================================================
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TABLES)

sSQL = "SELECT [CASE],[PCS PER SIDE MAX],[SHOT],[SF] " & _
       "FROM [PCS PER SIDE] " & _
       "WHERE [CASE] ='" & CASE_SIZE_ID & "'"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

SF = FR_Table.Fields("[SF]")

SHOT_ID = 0
SPEED_ID = 0


'090 Plating chg 06/23/2014 TOTAL_QTY Surface Area Error

Select Case DEPT_ID
Case 555, 556, 557, 553, 554, 558  ' REPLATE
        sSQL = "SELECT * FROM [REPLATE BARREL] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[DV MIN]<=" & GEAR_1_QTY & " AND [DV MAX] >=" & GEAR_1_QTY
                                           
Case 524, 523, 537, 538 ' NICKEL HG
        sSQL = "SELECT * FROM [092 NICKEL HG] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[DV MIN]<=" & GEAR_1_QTY & " AND [DV MAX] >=" & GEAR_1_QTY

Case 285, 286, 525, 526, 530, 533, 532, 534 'NICKEL LW
        sSQL = "SELECT * FROM [092 NICKEL LW]" & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[DV MIN]<=" & GEAR_1_QTY & " AND [DV MAX] >=" & GEAR_1_QTY
                                                                                                 
Case 529, 528, 287, 288 ' NICKEL TIN
        sSQL = "SELECT * FROM [115 NICKEL TIN] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[DV MIN]<=" & GEAR_1_QTY & " AND [DV MAX] >=" & GEAR_1_QTY
                                           
Case 535, 536       ' NICKEL AU
        sSQL = "SELECT * FROM [088 NICKEL GOLD] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[DV MIN]<=" & GEAR_1_QTY & " AND [DV MAX] >=" & GEAR_1_QTY
                                           
Case 540, 539, 544, 551, 546, 552 'COPPER 1 / COPPER 2
        
        Select Case SERIES_ID
        Case 100, 800, 600, 700
                    sSQL = "SELECT * FROM [121 CU 1] " & _
                           "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                 "[DV MIN] <=" & GEAR_1_QTY & " AND " & _
                                 "[DV MAX] >=" & GEAR_1_QTY & " AND [TYPE_CU]='" & TYPE_CU & "'"
        
'chg 12/07/2017
        Case 200, 900
                    sSQL = "SELECT * FROM [121 CU 2 LW] " & _
                           "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                 "[DV MIN] <=" & GEAR_1_QTY & " AND " & _
                                 "[DV MAX] >=" & GEAR_1_QTY & " AND [TYPE_CU]='" & TYPE_CU & "'"
        End Select
        
Case 541, 549 'COPPER 2
        sSQL = "SELECT * FROM [121 CU 2 HG] " & _
                "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                      "[DV MIN]<=" & GEAR_1_QTY & " AND " & _
                      "[DV MAX] >=" & GEAR_1_QTY & " AND [TYPE_CU]='" & TYPE_CU & "'"
                                                            
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    SHOT_ID = FR_Table.Fields("[BARREL]")
    SPEED_ID = FR_Table.Fields("[MEDIA SPEED]")
End If


'================================================================================
'   [3]  MEDIA SURFACE AREA
'================================================================================

sSQL = "SELECT [600 SFL],[100 AB],[100 CE],[200 AB],[200 CE] " & _
       "FROM [SHOT] WHERE [SHOT_ID] = 1"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
              
Shot_Qty = 0

Select Case SHOT_ID
Case 100
        Select Case CASE_SIZE_ID
        Case "A", "B"
                    Shot_Qty = FR_Table.Fields("[100 AB]") * NUMBER_HEADS
        Case "C", "E"
                    Shot_Qty = FR_Table.Fields("[100 CE]") * NUMBER_HEADS
        Case "S", "F", "L"
                    Shot_Qty = FR_Table.Fields("[600 SFL]") * NUMBER_HEADS
        End Select
Case 200
        Select Case CASE_SIZE_ID
        Case "A", "B"
                    Shot_Qty = FR_Table.Fields("[200 AB]") * NUMBER_HEADS
        Case "C", "E"
                    Shot_Qty = FR_Table.Fields("[200 CE]") * NUMBER_HEADS
        End Select
End Select
  
'================================================================================
'   [4] MAIN TABLE
'================================================================================

Select Case DEPT_ID
Case 555, 556, 557, 553, 554, 558 ' REPLATE JUST FINISH
        sSQL = "SELECT * FROM [REPLATE BARREL] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND [BARREL] =" & SHOT_ID
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            ASF2 = FR_Table.Fields("[ASF]")
            MIN2 = FR_Table.Fields("[MIN]")
        End If
Case 524, 523, 537, 538 ' NICKEL HG
        sSQL = "SELECT * FROM [092 NICKEL HG] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND [BARREL] =" & SHOT_ID
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            ASF1 = FR_Table.Fields("[ASF NI]")
            MIN1 = FR_Table.Fields("[MIN NI]")
            ASF2 = FR_Table.Fields("[ASF HG]")
            MIN2 = FR_Table.Fields("[MIN HG]")
        End If

Case 285, 525, 526, 530, 533, 532, 534, 286 'NICKEL LW
        sSQL = "SELECT * FROM [092 NICKEL LW] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND [BARREL] =" & SHOT_ID
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            ASF1 = FR_Table.Fields("[ASF NI]")
            MIN1 = FR_Table.Fields("[MIN NI]")
            ASF2 = FR_Table.Fields("[ASF LW]")
            
            Select Case DEPT_ID
            Case 525, 526, 530, 533, 285, 286
                    MIN2 = FR_Table.Fields("[MIN LW W]")
            Case 532, 534
                    MIN2 = FR_Table.Fields("[MIN LW P]")
            End Select
            
        End If
Case 529, 528, 287, 288 ' NICKEL TIN
        sSQL = "SELECT * FROM [115 NICKEL TIN] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND [BARREL] =" & SHOT_ID
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            ASF1 = FR_Table.Fields("[ASF NI]")
            MIN1 = FR_Table.Fields("[MIN NI]")
            ASF2 = FR_Table.Fields("[ASF TIN]")
            MIN2 = FR_Table.Fields("[MIN TIN]")
        End If
Case 535, 536  ' NICKEL AU

        ' SPECIAL NOTE AU-STRIKE NO LOOK UP
        '                   TOTAL RUN TIME = 2 * SF
        '                   TIME IN MIN  = 2

        sSQL = "SELECT * FROM [088 NICKEL GOLD] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND [BARREL] =" & SHOT_ID
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            ASF1 = FR_Table.Fields("[ASF NI]")
            MIN1 = FR_Table.Fields("[MIN NI]")
            ASF2 = FR_Table.Fields("[ASF AU]")
            MIN2 = FR_Table.Fields("[MIN AU]")
            
            
            ASF3 = FR_Table.Fields("[ASF AU SK]")
            MIN3 = FR_Table.Fields("[MIN AU SK]")
            
        End If

Case 540, 539, 544, 551, 546, 552 'COPPER 1 / COPPER 2
        
        Select Case SERIES_ID
        Case 100, 800, 600, 700
                    sSQL = "SELECT * FROM [121 CU 1] " & _
                           "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                 "[BARREL] =" & SHOT_ID & " AND " & _
                                 "[TYPE_CU]='" & TYPE_CU & "'"
                                                       
                    Set FR_Table = FR_Database.OpenRecordset(sSQL)
                     
                    If (FR_Table.RecordCount <> 0) Then
                        
                        SKTASF1 = FR_Table.Fields("[ASF SK]")
                        SKTMIN1 = FR_Table.Fields("[MIN SK]")
                        
                        ASF1 = FR_Table.Fields("[ASF CU]")
                        MIN1 = FR_Table.Fields("[MIN CU]")
                        ASF2 = FR_Table.Fields("[ASF TIN]")
                        Select Case DEPT_ID
                        Case 540, 539, 544, 551
                                MIN2 = FR_Table.Fields("[MIN TIN]") '"W"
                        Case 546, 552
                                MIN2 = FR_Table.Fields("[MIN TIN PN]")
                        End Select
                    End If
'chg 12/07/2017
        
        Case 200, 900
                    sSQL = "SELECT * FROM [121 CU 2 LW] " & _
                           "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                 "[BARREL] =" & SHOT_ID & " AND " & _
                                 "[TYPE_CU]='" & TYPE_CU & "'"
                                                       
                    Set FR_Table = FR_Database.OpenRecordset(sSQL)
                     
                    If (FR_Table.RecordCount <> 0) Then
                                        
                        SKTASF1 = FR_Table.Fields("[ASF SK]")
                        SKTMIN1 = FR_Table.Fields("[MIN SK]")
                    
                        ASF1 = FR_Table.Fields("[ASF CU]")
                        MIN1 = FR_Table.Fields("[MIN CU]")
                        ASF2 = FR_Table.Fields("[ASF LW]")
                        Select Case DEPT_ID
                        Case 540, 539, 544, 551
                                MIN2 = FR_Table.Fields("[MIN LW]") '"W"
                        Case 546, 552
                                MIN2 = FR_Table.Fields("[MIN LP]")
                        End Select
                    End If
        
        End Select
                
Case 541, 549 'COPPER 2 HG
        
        sSQL = "SELECT * FROM [121 CU 2 HG] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                 "[BARREL] =" & SHOT_ID & " AND " & _
                                 "[TYPE_CU]='" & TYPE_CU & "'"
                                           
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            
            SKTASF1 = FR_Table.Fields("[ASF SK]")
            SKTMIN1 = FR_Table.Fields("[MIN SK]")
            
            ASF1 = FR_Table.Fields("[ASF CU]")
            MIN1 = FR_Table.Fields("[MIN CU]")
            
            ASF2 = FR_Table.Fields("[ASF HG]")
            MIN2 = FR_Table.Fields("[MIN HG]")
            
            'STRIKE
            
            ASF3 = FR_Table.Fields("[ASF LW]")
            MIN3 = FR_Table.Fields("[MIN LW]")
            
        End If

End Select

FR_Table.Close
FR_Database.Close

PART_SA = (SF * TOTAL_QTY) / 144
Media_SA = (Shot_Qty) / 144

SA = (SF * TOTAL_QTY + Shot_Qty) / 144

End Sub


Public Sub GearMax()

'================================================================================
'
'================================================================================
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TABLES)

Select Case DEPT_ID
Case 555, 556, 557, 553, 554, 558  ' REPLATE
        sSQL = "SELECT * FROM [REPLATE BARREL] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' ORDER BY [DV MAX] DESC "
                                                                
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If

Case 524, 523, 537, 538 ' NICKEL HG
        sSQL = "SELECT * FROM [092 NICKEL HG] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' ORDER BY [DV MAX] DESC "
       
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If

Case 285, 525, 526, 530, 533, 532, 534, 286 'NICKEL LW
        
        sSQL = "SELECT * FROM [092 NICKEL LW] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' ORDER BY [DV MAX] DESC "
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If
Case 529, 528, 287, 288 ' NICKEL TIN
        sSQL = "SELECT * FROM [115 NICKEL TIN] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' ORDER BY [DV MAX] DESC "
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If
Case 535, 536  ' NICKEL AU

        sSQL = "SELECT * FROM [088 NICKEL GOLD]  " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' ORDER BY [DV MAX] DESC "
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If

Case 540, 539, 544, 551, 546, 552 'COPPER 1 / COPPER 2
        
        Select Case SERIES_ID
        Case 100, 800, 600, 700
                        sSQL = "SELECT * FROM [121 CU 1] " & _
                               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                     "[TYPE_CU]='" & TYPE_CU & "' ORDER BY [DV MAX] DESC "
                                                            
                        Set FR_Table = FR_Database.OpenRecordset(sSQL)
                         
                        If (FR_Table.RecordCount <> 0) Then
                                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
                        End If
        Case 200, 900
                        sSQL = "SELECT * FROM [121 CU 2 LW] " & _
                               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                                     "[TYPE_CU]='" & TYPE_CU & "' ORDER BY [DV MAX] DESC "
                        
                        Set FR_Table = FR_Database.OpenRecordset(sSQL)
                         
                        If (FR_Table.RecordCount <> 0) Then
                                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
                        End If
        End Select
        
Case 541, 549 'COPPER 2
        sSQL = "SELECT * FROM [121 CU 2 HG] " & _
               "WHERE [CASE SERIES] ='" & SERIES_ID & CASE_SIZE_ID & "' AND " & _
                     "[TYPE_CU]='" & TYPE_CU & "' ORDER BY [DV MAX] DESC "
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
                GEAR_MAX_QTY = FR_Table.Fields("[DV MAX]")
        End If
End Select

FR_Table.Close
FR_Database.Close

End Sub


Public Sub SBE_Calculation()

SHOT_ID = 0
SKTASF1 = 0
SKTMIN1 = 0
SKTASF2 = 0
SKTMIN2 = 0
ASF1 = 0
MIN1 = 0
ASF2 = 0
MIN2 = 0
ASF3 = 0
MIN3 = 0

'================================================================================
'   [1]  SURFACE AREA PART
'================================================================================
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TABLES)

sSQL = "SELECT [CASE],[SF] FROM [PCS PER SIDE] " & _
       "WHERE [CASE] ='" & CASE_SIZE_ID & "'"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

TOTAL_QTY = TOTAL_QTY / NUMBER_HEADS

PART_SA = FR_Table.Fields("[SF]") * TOTAL_QTY / 144

'================================================================================
'   [2]  MEDIA SURFACE AREA
'================================================================================

sSQL = "SELECT [QTY],[MEDIA VOL],[MEDIA SF] " & _
       "FROM [TBL SBE 144] " & _
       "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & TOTAL_QTY & " " & _
       "ORDER BY [QTY] DESC"

Select Case CASE_SIZE_ID
Case "A", "B", "R"
    Select Case SERIES_ID
    Case 810
            sSQL = "SELECT [QTY],[MEDIA VOL],[MEDIA SF] " & _
                   "FROM [TBL SBE ABR JAX] " & _
                   "WHERE [CASE SIZE] ='" & CASE_SIZE_ID & "' AND [QTY]<=" & TOTAL_QTY & " " & _
                   "ORDER BY [QTY] DESC"
    End Select
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
        MsgBox "Quantity not Available", vbInformation, "ATC Data Base System"
        Exit Sub
End If

Media_SA = FR_Table.Fields("[MEDIA SF]")

SHOT_ID = FR_Table.Fields("[MEDIA VOL]")

Select Case DEPT_ID
Case 553, 554, 555, 556 'REPLATE
        Select Case SERIES_ID
        Case "100", "700"
                    Media_SA = 7.125
        Case "200", "900"
                    Media_SA = 9.5
        Case "600", "810"
                    Media_SA = 9.5
        End Select

Case Else
        Media_SA = FR_Table.Fields("[MEDIA SF]")
End Select

SA = PART_SA + Media_SA

'================================================================================
'   [4] MAIN TABLE
'================================================================================

Set FR_Database = OpenDatabase(DB_PLATING_TABLES)

' FIND LOOKUP MODE
sSQL = "SELECT first([MODE]) AS [SQL MODE] FROM [TBL SBE] " & _
       "WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] = " & SERIES_ID & " " & _
       "GROUP BY [CASE],[SERIES]"
                                       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim CAL_MODE As String

If (FR_Table.RecordCount = 0) Then
    Exit Sub
Else
    CAL_MODE = FR_Table.Fields("[SQL MODE]")
End If


'================================================================================
'   [3]  AMPS/MIN
'================================================================================
       
Select Case CAL_MODE
Case "NA"
        sSQL = "SELECT * FROM [TBL SBE] " & _
               "WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] = " & SERIES_ID
               
Case "DV"
        
        Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)
        'CHG 06/04/2019
        Dim LOT_CODE As String
        
        sSQL = "SELECT * FROM [GROUPING] WHERE [GP_ID]=" & GP_ID
        
        Set TO_Table = TO_Database.OpenRecordset(sSQL)
        If (TO_Table.RecordCount <> 0) Then
            LOT_CODE = Mid(TO_Table.Fields("[LOT NUM]"), 1, 1)
        Else
                MsgBox "CODE ERROR", vbCritical, ""
                Exit Sub
        End If
        TO_Database.Close

        Select Case CASE_SIZE_ID
        Case "L"
        
                sSQL = "SELECT * FROM [TBL SBE] " & _
                       "WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] = " & SERIES_ID & " AND " & _
                             "[DV MIN] <=" & DV_ID & " AND [DV MAX] >= " & DV_ID & " AND " & _
                             "[LOT CODE]='" & LOT_CODE & "'"
        
        Case Else
                        sSQL = "SELECT * FROM [TBL SBE] " & _
                       "WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] = " & SERIES_ID & " AND " & _
                             "[DV MIN] <=" & DV_ID & " AND [DV MAX] >= " & DV_ID
        
        End Select
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        TABLE_ID = FR_Table.Fields("[ID]")

Case "QTY"
        sSQL = "SELECT * FROM [TBL SBE] " & _
               "WHERE [CASE] ='" & CASE_SIZE_ID & "' AND [SERIES_TYPE] = " & SERIES_ID & " AND " & _
                     "[QTY MIN] <=" & TOTAL_QTY & " AND [QTY MAX] >= " & TOTAL_QTY

        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        TABLE_ID = FR_Table.Fields("[ID]")
        
End Select


Set FR_Table = FR_Database.OpenRecordset(sSQL)

Select Case DEPT_ID
Case 285, 287, 525, 526, 528, 529, 530, 533, 532, 534, 286, 288 'NICKEL BASE
        
        ASF1 = FR_Table.Fields("[ASF1]")
        MIN1 = FR_Table.Fields("[MIN1]")
        
End Select

Select Case DEPT_ID
Case 287, 288, 528, 529                         'TIN FINISH

        ASF2 = FR_Table.Fields("[ASF2]")
        MIN2 = FR_Table.Fields("[MIN2]")
        
Case 285, 286, 525, 526, 530, 532, 533, 534     'LW FINISH

        ASF2 = FR_Table.Fields("[ASF3]")
        MIN2 = FR_Table.Fields("[MIN3]")
        
Case 553, 554, 555, 556                        'FINISH REPLATE

        ASF2 = FR_Table.Fields("[ASF4]")
        MIN2 = FR_Table.Fields("[MIN4]")
        
Case Else

End Select

FR_Table.Close
FR_Database.Close

End Sub
'
'   INPUT SET_ID
'
Public Sub Set_Calculation()

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        DEPT_ID = FR_Table.Fields("[DEPT_ID]")
        CASE_SIZE_ID = Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)
        SERIES_ID = Val(Mid(FR_Table.Fields("[SERIES_ID]"), 1, 3))
        TYPE_ID = FR_Table.Fields("[TYPE_ID]")
        'chg 12/13/2011
        Select Case FR_Table.Fields("[EQ BASE]")
        Case 18, 73
                TYPE_CU = "MSA"
        Case 17, 74, 75, 76
                TYPE_CU = "PYRO"
        Case Else
                TYPE_CU = "NA"
        End Select
Else
        Exit Sub
End If
    
sSQL = "SELECT count([WORK ORDER])," & _
              "format(sum([QTY]),'###,####')," & _
              "sum([P1 BASE]) AS [SQL QTY1]," & _
              "sum([P2 BASE]) AS [SQL QTY2]," & _
              "sum([P3 BASE]) AS [SQL QTY3]," & _
              "sum([P4 BASE]) AS [SQL QTY4] " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "GROUP BY [SET_ID]"
    
Set FR_Table = FR_Database.OpenRecordset(sSQL)

GEAR_1_QTY = 0
GEAR_2_QTY = 0
GEAR_3_QTY = 0
GEAR_4_QTY = 0
TOTAL_QTY = 0
NUMBER_HEADS = 0

If (FR_Table.RecordCount = 0) Then
    Exit Sub
Else
    
    Dim I As Integer
    Dim sBuff As String
    
    For I = 1 To 4
        sBuff = "[SQL QTY" & I & "]"
        If (FR_Table.Fields(sBuff) <> 0) Then
            NUMBER_HEADS = NUMBER_HEADS + 1
        End If
    Next I
    
    GEAR_1_QTY = FR_Table.Fields("[SQL QTY1]")
    GEAR_2_QTY = FR_Table.Fields("[SQL QTY2]")
    GEAR_3_QTY = FR_Table.Fields("[SQL QTY3]")
    GEAR_4_QTY = FR_Table.Fields("[SQL QTY4]")
            
    TOTAL_QTY = FR_Table.Fields("[SQL QTY1]") + FR_Table.Fields("[SQL QTY2]") + FR_Table.Fields("[SQL QTY3]") + FR_Table.Fields("[SQL QTY4]")
    
End If

FR_Table.Close
FR_Database.Close

If (TOTAL_QTY = 0) Then
    Exit Sub
End If
  
Select Case TYPE_ID
Case "BARREL"
                BarrelCalculation
Case "SBE"
                SBE_Calculation
End Select
  
End Sub

'
'   Valid Part Sort checks for valid part format
'   Series , Cap Value Format , No Tolerance
'
Function ATCPartDV(sATCPart As String, sDesignValue As String) As Boolean
         
   
   '----- TEST FOR VALID SERIES DESIGN VALUE FORMAT
   Select Case Mid$(sATCPart, 1, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 2, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 3, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select

   '----- TEST FOR VALID DESIGN VALUE FORMAT

   Select Case Mid$(sATCPart, 5, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 6, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "R"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 7, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ATCPartDV = False
                Exit Function
   End Select

    Dim dDesignValue As Double
   '---  CONVERT TO A DOUBLE PF VALUE
   If (Mid$(sATCPart, 6, 1) = "R") Then
      dDesignValue = Val(Mid$(sATCPart, 5, 1) & "." & Mid$(sATCPart, 7, 1))
   Else
      dDesignValue = Val(Mid$(sATCPart, 5, 2) & "E" & Mid$(sATCPart, 7, 1))
   End If

   sDesignValue = Str$(dDesignValue)
   
   ATCPartDV = True

End Function


Public Function ValidATCPartCode(sATCPart As String)

If (Len(sATCPart) < 9) Then
    MsgBox "ATC Part Code less than 9 characters", vbCritical, "ATC Plating System"
    Exit Function
End If

sATCPart = sATCPart & " X"

Dim sSQL As String

'===========================================================================
'   LOOK UP ATC PART VALIDATION CODES FOR BASE AND FINISH
'===========================================================================

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim iCOL_BASE As Integer
Dim sValidTest_BASE As String

Dim iCOL_FINISH As Integer
Dim sValidTest_FINISH As String

If (FR_Table.RecordCount <> 0) Then
    
    BASE_ID = FR_Table.Fields("[BASE_ID]")
    iCOL_BASE = FR_Table.Fields("[BASE COL]")
    sValidTest_BASE = FR_Table.Fields("[BASE VALID TEST]")
    
    iCOL_FINISH = FR_Table.Fields("[FINISH COL]")
    sValidTest_FINISH = FR_Table.Fields("[FINISH VALID TEST]")

End If

'========================================================
'   ATC_PART_ID FOUND IN WORK ORDER SCHEDULE ?
'========================================================

Dim RESULT As Integer
Dim sMessage As String

'=================================================================
'   TEST VALIDATION PART CODES PROCESS BASE_ID
'=================================================================

Dim MYPOS As Integer
Dim SearchChar As String

Select Case BASE_ID
Case "Copper", "Nickel"

        SearchChar = Mid(sATCPart, iCOL_BASE, 1)
        Select Case Mid(sValidTest_BASE, 1, 2)
        Case "IN"
                    MYPOS = InStr(5, sValidTest_BASE, SearchChar, 1)
                    Select Case MYPOS
                    Case 0
                            RESULT = 0 'NOT FOUND
                    Case Else
                            RESULT = 1  'FOUND
                    End Select
        Case "NO"
                    MYPOS = InStr(9, sValidTest_BASE, SearchChar, 1)
                    Select Case MYPOS
                    Case 0
                            RESULT = 1 'NOT FOUND
                    Case Else
                            RESULT = 0  'FOUND
                    End Select
        End Select
        
        Select Case RESULT
        Case 1
                MsgBox "Valid ATC PART Base Process Code", vbInformation, "ATC Plating System"
        Case 0
                sMessage = "Department Does Not Match ATC Part Code " & vbNewLine
                sMessage = "Department " & DEPT_ID & " Base " & vbNewLine
                sMessage = sMessage & "Column " & iCOL_BASE & " " & vbNewLine
                sMessage = sMessage & " " & sValidTest_BASE & " " & vbNewLine
                sMessage = sMessage & "Input Code was " & SearchChar
                MsgBox sMessage, vbCritical, "ATC Plating System"
        End Select

Case Else

End Select

'=================================================================
'   TEST VALIDATION PART CODES PROCESS FINISH_ID
'=================================================================
        
SearchChar = Mid(sATCPart, iCOL_FINISH, 1)
Select Case Mid(sValidTest_FINISH, 1, 2)
Case "IN"
            MYPOS = InStr(5, sValidTest_FINISH, SearchChar, 1)
            Select Case MYPOS
            Case 0
                    RESULT = 0  'NOT FOUND
            Case Else
                    RESULT = 1  'FOUND
            End Select
Case "NO"
            MYPOS = InStr(9, sValidTest_FINISH, SearchChar, 1)
            Select Case MYPOS
            Case 0
                    RESULT = 1  'NOT FOUND
            Case Else
                    RESULT = 0  'FOUND
            End Select
End Select

Select Case RESULT
Case 1
        MsgBox "Valid ATC PART Base Process Code", vbInformation, "ATC Plating System"
Case 0
        sMessage = "Department Does Not Match ATC Part Code " & vbNewLine
        sMessage = "Department " & DEPT_ID & " Finish " & vbNewLine
        sMessage = sMessage & "Column " & iCOL_FINISH & " " & vbNewLine
        sMessage = sMessage & " " & sValidTest_FINISH
        MsgBox sMessage, vbCritical, "ATC Plating System"
End Select
 
End Function

Public Sub PrintGrouping()

Select Case LETTER_ID
Case "", " "
        MsgBox "Nothing Selected", vbCritical, "ATC Plating"
        Exit Sub
End Select

Screen.MousePointer = vbHourglass

Set_Calculation
             
Dim sBuff As String
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
             
sSQL = "SELECT count([WORK ORDER]) as [SQL COUNT WO]," & _
              "sum  ([P1 BASE])    as [SQL GH1]," & _
              "sum  ([P2 BASE])    as [SQL GH2]," & _
              "sum  ([P3 BASE])    as [SQL GH3]," & _
              "sum  ([P4 BASE])    as [SQL GH4]," & _
              "sum  ([QTY])        as [SQL SUM QTY] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND ucase([LETTER_ID])='" & LETTER_ID & "' " & _
       "GROUP BY [SET_ID],ucase([LETTER_ID]) "
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim WOCount As Integer
Dim SUMQTY As Long
Dim GH As String

If (FR_Table.RecordCount <> 0) Then
     WOCount = FR_Table.Fields("[SQL COUNT WO]")
     SUMQTY = FR_Table.Fields("[SQL SUM QTY]")
     
     If (FR_Table.Fields("[SQL GH1]") <> 0) Then
            GH = "1G/H1"
     End If
     If (FR_Table.Fields("[SQL GH2]") <> 0) Then
            GH = "1NG/H2"
     End If
     If (FR_Table.Fields("[SQL GH3]") <> 0) Then
            GH = "2G/H3"
     End If
     If (FR_Table.Fields("[SQL GH4]") <> 0) Then
            GH = "2NG/H4"
     End If
Else
    MsgBox "Not Available", vbCritical, "ATC Plating"
    FR_Table.Close
    FR_Database.Close
    Exit Sub
End If
                                                    
Dim sFilename As String
sFilename = DB_REPORT_ADDR & SET_ID & LETTER_ID & ".TXT"

Dim iFilenum As Integer
iFilenum = FreeFile

Open sFilename For Output Shared As #iFilenum

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim BASE_ID As String
Dim FINISH_ID As String
Dim DESCRIPTION As String

If (FR_Table.RecordCount <> 0) Then
    BASE_ID = FR_Table.Fields("[BASE_ID]")
    FINISH_ID = FR_Table.Fields("[FINISH_ID]")
    DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
End If

sSQL = "SELECT *  FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)
Print #iFilenum,
Print #iFilenum,
Print #iFilenum, Tab(20); "PLATING LOT GROUP WORK SHEET RUN : "; LETTER_ID
Print #iFilenum,
Print #iFilenum, Tab(20); "Electrical Test PLATING_ID : "; SET_ID & LETTER_ID
Print #iFilenum,
Print #iFilenum,

If (FR_Table.RecordCount <> 0) Then
                                                                           
    Print #iFilenum, Tab(5); "SET#";
    Print #iFilenum, Tab(15); "DATE";
    Print #iFilenum, Tab(25); "DEPT_ID";
    Print #iFilenum, Tab(35); "DESCRIPTION";
    Print #iFilenum, Tab(50); "PROCESS";
    Print #iFilenum, Tab(60); "SERIES";
    Print #iFilenum, Tab(70); "WO's";
    Print #iFilenum, Tab(75); "Gear/Head"
    Print #iFilenum,
    
    Print #iFilenum, Tab(5); FR_Table.Fields("[SET NUMBER]");
    Print #iFilenum, Tab(10); Format(FR_Table.Fields("[DATE_ID]"), "MM/DD/YYYY");
    Print #iFilenum, Tab(25); FR_Table.Fields("[DEPT_ID]");
    Print #iFilenum, Tab(35); DESCRIPTION;
    Print #iFilenum, Tab(50); FR_Table.Fields("[TYPE_ID]");
    Print #iFilenum, Tab(60); FR_Table.Fields("[SERIES_ID]");
    Print #iFilenum, Tab(70); WOCount;
    Print #iFilenum, Tab(75); GH
    Print #iFilenum, Tab(5); "===================================================================================="
    Print #iFilenum,
End If

Print #iFilenum, Tab(5); "W.O.#";
Print #iFilenum, Tab(15); "Item";
Print #iFilenum, Tab(25); "Run";
Print #iFilenum, Tab(30); "ATC Part#";
Print #iFilenum, Tab(50); "Qty";
Print #iFilenum, Tab(60); "Value";
Print #iFilenum, Tab(70); "Tol"
Print #iFilenum, Tab(5); "===================================================================================="

sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' ORDER BY [DV] ASC"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim sDesignValue As String

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        Print #iFilenum,
        If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
            'WO
            Print #iFilenum, Tab(5); Str(Mid(FR_Table.Fields("[WORK ORDER]"), 1, 6));
            'ITEM
            sBuff = Str(Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3)) & "." & Str(Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1))
            
            Print #iFilenum, Tab(15); Str(Mid(FR_Table.Fields("[WORK ORDER]"), 7, 3)) & "." & Str(Mid(FR_Table.Fields("[WORK ORDER]"), 10, 1));
            'RUN
            Print #iFilenum, Tab(25); Mid(FR_Table.Fields("[WORK ORDER]"), 11, 2);
        Else
            'LOT NUM
             Print #iFilenum, Tab(5); FR_Table.Fields("[WORK ORDER]");
        End If
        ' ATC PART NUMBER
        Print #iFilenum, Tab(30); FR_Table.Fields("[ATC PART]");
        
        ' TOTAL QUANTITY
        Print #iFilenum, Tab(50); Format(FR_Table.Fields("[P1 BASE]") + FR_Table.Fields("[P2 BASE]") + FR_Table.Fields("[P3 BASE]") + FR_Table.Fields("[P4 BASE]"), "###,###");
                
        ' DESIGN VALUE
        If (ATCPartDV(FR_Table.Fields("[ATC PART]"), sDesignValue) = True) Then
            Print #iFilenum, Tab(60); sDesignValue;
        Else
            Print #iFilenum, Tab(60); ""
        End If

        ' TOLERANCE
        Print #iFilenum, Tab(70); Mid(FR_Table.Fields("[ATC PART]"), 8, 1)
        
        FR_Table.MoveNext
    Loop
     
End If
Print #iFilenum,
Print #iFilenum,
Print #iFilenum, Tab(40); "Total Quantity : "; Format(SUMQTY, "#,###,##0")
Print #iFilenum,
Print #iFilenum,
                                                               
Print #iFilenum, Tab(30); "PLATING TABLE"
Print #iFilenum,
Print #iFilenum, Tab(30); "DV";
Print #iFilenum, Tab(40); "BIN LIMIT"
                                                               
Dim I As Integer
Dim SORT_TABLE As Boolean
Select Case Calculate_Plating_Table
Case 0
        SORT_TABLE = True
        sBuff = "Case [0] Sortable Grouping"
Case 1
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [1] Missing Lot Number and Part Number"
        sBuff = "Case [1] Missing Lot Number and Part Number"
Case 2
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [2] No Sort Required"
        sBuff = "Case [2] No Sort Required"
Case 3
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [3] Lot Sort Required"
        sBuff = "Case [3] Lot Sort Required"
Case 4
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [4] Too Many Bins Second Sort Required"
        sBuff = "Case [4] Too Many Bins Second Sort Required"
Case 5
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [5] WO Contains Lot Second Sort Required"
        sBuff = "Case [5] WO Contains Lot Second Sort Required"
        SORT_TABLE = True
Case 6
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [6] Second Sort Required"
        sBuff = "Case [6] Second Sort Required"
        SORT_TABLE = True
Case 7
        Print #iFilenum,
        Print #iFilenum, Tab(30); "Case [7] Not able to Create Sort Table"
        sBuff = "Case [7] Not able to Create Sort Table"
Case Else
End Select
                                                             
If (SORT_TABLE = True) Then
        For I = 0 To 9
            If (gdBinLimit(I + 1) <> 0) Then
                Select Case gdBinLimit(I + 1)
                Case 0 To 10
                        Print #iFilenum, Tab(40); Format(gdBinLimit(I + 1), "#,###,##0.00")
                Case 10 To 20
                        Print #iFilenum, Tab(40); Format(gdBinLimit(I + 1), "#,###,##0.0")
                Case Else
                        Print #iFilenum, Tab(40); Format(gdBinLimit(I + 1), "#,###,##0")
                End Select
            End If
            Print #iFilenum, Tab(30); gsBinTol(I + 1)
        Next I
End If
                                
'FR_Table.Close
'FR_Database.Close

Close iFilenum

If (SORT_TABLE = True) Then
    Dim sFileName2 As String
    sFileName2 = DB_REPORT_ADDR & SET_ID & LETTER_ID & ".DAT"
    iFilenum = FreeFile
    Open sFileName2 For Output Shared As #iFilenum
    For I = 0 To 9
        Print #iFilenum, gsBinTol(I + 1)
    Next I
    For I = 0 To 9
        If (gdBinLimit(I + 1) <> 0) Then
            Select Case gdBinLimit(I + 1)
            Case 0 To 10
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0.00")
            Case 10 To 20
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0.0")
            Case Else
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0")
            End Select
        Else
                    Print #iFilenum, "0"
        End If
    Next I
    Close iFilenum
    
    Plating_Sort_DB
    
End If

Screen.MousePointer = vbDefault

MsgBox sBuff, vbInformation, "ATC Plating System"

Dim iAns As Integer
iAns = MsgBox("Print Plating Lot Group Worksheet", vbYesNo, sFilename)
If (iAns = vbYes) Then
   'vbPRORLandscape vbPRORPortrait
    PrintFile sFilename, vbPRORPortrait, 12
End If

End Sub


Public Function ValidATCPartCodeTest(sATCPart As String)

If (Len(sATCPart) < 9) Then
    MsgBox "ATC Part Code less than 9 characters", vbCritical, "ATC Plating System"
    Exit Function
End If

Dim sSQL As String

'===========================================================================
'   LOOK UP ATC PART VALIDATION CODES FOR BASE AND FINISH
'===========================================================================

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
              
Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select
              
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim iCOL_BASE As Integer
Dim sValidTest_BASE As String

Dim iCOL_FINISH As Integer
Dim sValidTest_FINISH As String

If (FR_Table.RecordCount <> 0) Then
    
    BASE_ID = FR_Table.Fields("[BASE_ID]")
    iCOL_BASE = FR_Table.Fields("[BASE COL]")
    sValidTest_BASE = FR_Table.Fields("[BASE VALID TEST]")
    
    iCOL_FINISH = FR_Table.Fields("[FINISH COL]")
    sValidTest_FINISH = FR_Table.Fields("[FINISH VALID TEST]")
Else
    Exit Function
End If

'========================================================
'   ATC_PART_ID FOUND IN WORK ORDER SCHEDULE ?
'========================================================

Dim RESULT As Integer
Dim sMessage As String

'=================================================================
'   TEST VALIDATION PART CODES PROCESS BASE_ID
'=================================================================

Dim MYPOS As Integer
Dim SearchChar As String

ALERT_MESSAGE = "ok"

Select Case BASE_ID
Case "Copper", "Nickel"


        sATCPart = sATCPart & "             X" '05/01/2019

        SearchChar = Mid(sATCPart, iCOL_BASE, 1)
        Select Case Mid(sValidTest_BASE, 1, 2)
        Case "IN"
                    MYPOS = InStr(5, sValidTest_BASE, SearchChar, 1)
                    Select Case MYPOS
                    Case 0
                            RESULT = 0 'NOT FOUND
                    Case Else
                            RESULT = 1  'FOUND
                    End Select
        Case "NO"
                    MYPOS = InStr(9, sValidTest_BASE, SearchChar, 1)
                    Select Case MYPOS
                    Case 0
                            RESULT = 1 'NOT FOUND
                    Case Else
                            RESULT = 0  'FOUND
                    End Select
        End Select
        
        Select Case RESULT
        Case 1
              '  MsgBox "Valid ATC PART Base Process Code", vbInformation, "ATC Plating System"
        Case 0
                sMessage = "Department Does Not Match ATC Part Code " & vbNewLine
                sMessage = "Department " & DEPT_ID & " Base " & vbNewLine
                sMessage = sMessage & "Column " & iCOL_BASE & " " & vbNewLine
                sMessage = sMessage & " " & sValidTest_BASE & " " & vbNewLine
                sMessage = sMessage & "Input Code was " & SearchChar
                MsgBox sMessage, vbCritical, "ATC Plating System"
                
                ALERT_MESSAGE = "NO PROCESAR ESTA ORDEN"
        End Select

Case Else

End Select

'=================================================================
'   TEST VALIDATION PART CODES PROCESS FINISH_ID
'=================================================================
        
SearchChar = Mid(sATCPart, iCOL_FINISH, 1)
Select Case Mid(sValidTest_FINISH, 1, 2)
Case "IN"
            MYPOS = InStr(5, sValidTest_FINISH, SearchChar, 1)
            Select Case MYPOS
            Case 0
                    RESULT = 0  'NOT FOUND
            Case Else
                    RESULT = 1  'FOUND
            End Select
Case "NO"
            MYPOS = InStr(9, sValidTest_FINISH, SearchChar, 1)
            Select Case MYPOS
            Case 0
                    RESULT = 1  'NOT FOUND
            Case Else
                    RESULT = 0  'FOUND
            End Select
End Select

Select Case RESULT
Case 1
      '  MsgBox "Valid ATC PART Base Process Code", vbInformation, "ATC Plating System"
Case 0
        sMessage = "Department Does Not Match ATC Part Code " & vbNewLine
        sMessage = "Department " & DEPT_ID & " Finish " & vbNewLine
        sMessage = sMessage & "Column " & iCOL_FINISH & " " & vbNewLine
        sMessage = sMessage & " " & sValidTest_FINISH
        MsgBox sMessage, vbCritical, "ATC Plating System"
        
        ALERT_MESSAGE = "NO PROCESAR ESTA ORDEN"
End Select
 
Select Case ALERT_MESSAGE
Case "ok"

Case Else
        frmAlert.Show vbModal
End Select
 
 
End Function

Public Sub ValidateSeries()

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim SERIES_LID As Integer
Dim SERIES_CASE_LID As String
Dim SERIES_ID As String
Dim ATC_SERIES As String


SERIES_ID = FR_Table.Fields("[SERIES_ID]")

Select Case Mid(FR_Table.Fields("[SERIES_ID]"), 3, 1)
Case "U"
        SERIES_ID = Mid(FR_Table.Fields("[SERIES_ID]"), 1, 2) & "0" & Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)
Case Else


End Select

SERIES_LID = Mid(SERIES_ID, 1, 3)
SERIES_CASE_LID = SERIES_ID

sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
        
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim sBuff As String

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        Select Case SERIES_LID
        Case 200, 900
            '08/19/2019 change 3 char to 0
                ATC_SERIES = Mid(FR_Table.Fields("[ATC PART]"), 1, 1) & "00" & Mid(FR_Table.Fields("[ATC PART]"), 4, 1)
        
                If (ATC_SERIES <> SERIES_CASE_LID) Then
                    
                    sBuff = " W.O. " & FR_Table.Fields("[WORK ORDER]") & vbNewLine & FR_Table.Fields("[ATC PART]")
                    MsgBox "Error Only Series Selected Should Be In Set ", vbCritical, "ATC Plating" & sBuff
                    Exit Do
                    
                End If
        
        Case Else
               Select Case Mid(FR_Table.Fields("[ATC PART]"), 1, 3)
               Case 200, 900
                    
                    sBuff = " W.O. " & FR_Table.Fields("[WORK ORDER]") & vbNewLine & FR_Table.Fields("[ATC PART]")
                    MsgBox "Error  200 OR 900 Should Not Be Part of Set ", vbCritical, "ATC Plating" & sBuff
                    Exit Do
                    
               Case Else
                    
                    Select Case Mid(FR_Table.Fields("[ATC PART]"), 4, 1)
                    Case "A", "C", "E"
                            If (Mid(FR_Table.Fields("[ATC PART]"), 4, 1) <> Mid(SERIES_CASE_LID, 4, 1)) Then
                                
                                sBuff = " W.O. " & FR_Table.Fields("[WORK ORDER]") & vbNewLine & FR_Table.Fields("[ATC PART]")
                                MsgBox "Error Only Series Selected Should Be In Set ", vbCritical, "ATC Plating" & sBuff
                                Exit Do
                                
                           End If
                    Case "B", "T"
                            If (Mid(SERIES_CASE_LID, 4, 1) <> "B") Then
                                
                                sBuff = " W.O. " & FR_Table.Fields("[WORK ORDER]") & vbNewLine & FR_Table.Fields("[ATC PART]")
                                MsgBox "Error Only Series Selected Should Be In Set ", vbCritical, "ATC Plating" & sBuff
                                Exit Do
                                
                           End If
                    End Select
                    
               End Select
               
        End Select
        FR_Table.MoveNext
    Loop
End If

FR_Table.Close
FR_Database.Close

End Sub

Public Sub Report_SA()
             
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    FR_Table.Edit
    FR_Table.Fields("[PART_SA]") = Format(PART_SA, "0.0")          'PART SA
    FR_Table.Fields("[Media_SA]") = Format(Media_SA, "0.0")        'MEDIA SA
    'FR_Table.Fields("[SA]") = Format(SA, "0.0")                    'SQ FT
    FR_Table.Update
End If

FR_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
MsgBox "Excel Update Complete", vbInformation, "ATC Plating"

End Sub


Public Sub ExcelReport()

Dim sBuff As String

Set_Calculation
             
Dim STRIKE1 As String
Dim STRIKE2 As String
Dim STRIKE1_ID As String
Dim STRIKE2_ID As String
Dim BASE_ID As String
Dim FINISH_ID As String
Dim OVERPLATE As String
Dim DESCRIPTION As String

Dim sSQL As String

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [DEPT_ID],[DEPT_ID],[DESCRIPTION],[BASE_ID],[FINISH_ID]," & _
              "[TANK DWG],[SBE],[TANK],[STRIKE1],[STRIKE2],[OVERPLATE],[ACTIVE] " & _
       "FROM [DEPT CODE] ORDER BY [DEPT_ID]"

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    STRIKE1 = FR_Table.Fields("[STRIKE1]")      'Y/N
    STRIKE2 = FR_Table.Fields("[STRIKE2]")      'Y/N
    STRIKE1_ID = FR_Table.Fields("[STRIKE1_ID]")
    STRIKE2_ID = FR_Table.Fields("[STRIKE2_ID]")      'Y/N
    OVERPLATE = FR_Table.Fields("[OVERPLATE]")  'Y/N
    BASE_ID = FR_Table.Fields("[BASE_ID]")      '
    FINISH_ID = FR_Table.Fields("[FINISH_ID]")
    DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
End If

Dim wbWorld As Object, shtWorld As Object
Dim tSheet As Object

Set shtWorld = GetObject(PLATING_LOG_SHEET)
    
Set wbWorld = shtWorld.Application.Workbooks("Plating Log Sheet Master.xls")
Set tSheet = wbWorld.Sheets("New Format")   'WORK SHEET NAME
            
'chg 10/06/2011 ADDITION
'shtWorld.Application.Visible = True
tSheet.Parent.WINDOWS(1).Visible = True
    
Screen.MousePointer = vbHourglass
                
Dim iRow As Integer, iCol As Integer
                 
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT *  FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

'====================================================================
'
'   BASE OR FINISH?
'
If (FR_Table.RecordCount <> 0) Then
    EQ_BASE_ID = FR_Table.Fields("[EQ BASE]")
End If

sSQL = "SELECT * FROM [MACHINE] WHERE [NUMBER]=" & EQ_BASE_ID
Set TO_Table = TO_Database.OpenRecordset(sSQL)

RECTIFIER = 0
If (TO_Table.RecordCount <> 0) Then
    RECTIFIER = TO_Table.Fields("[RECTIFIER]")
End If
Select Case RECTIFIER
Case 0
    sSQL = "SELECT * FROM [MACHINE] WHERE [NUMBER]=" & EQ_FINISH_ID
    Set TO_Table = TO_Database.OpenRecordset(sSQL)
    If (TO_Table.RecordCount <> 0) Then
        RECTIFIER = TO_Table.Fields("[RECTIFIER]")
    End If
End Select
TO_Database.Close

'====================================================================

Dim ROWSTART As Integer

ROWSTART = 3

If (FR_Table.RecordCount <> 0) Then
                                                                       
    tSheet.Cells(ROWSTART, 2).Value = FR_Table.Fields("[SET NUMBER]")        'SET#
    tSheet.Cells(ROWSTART, 3).Value = Format(FR_Table.Fields("[DATE_ID]"), "MM/DD/YYYY")          'DATE
    tSheet.Cells(ROWSTART, 4).Value = FR_Table.Fields("[DEPT_ID]")           'DEPT_ID
    tSheet.Cells(ROWSTART, 5).Value = DESCRIPTION     'DESCRIPTION
    tSheet.Cells(ROWSTART, 6).Value = FR_Table.Fields("[TYPE_ID]")           'Process
    tSheet.Cells(ROWSTART, 7).Value = FR_Table.Fields("[SERIES_ID]")         'SERIES CASE
                                                                             
    tSheet.Cells(1, 10).Value = FR_Table.Fields("[SET_ID]")
   
    If (STRIKE1 = "Y") Then
        tSheet.Cells(5, 15).Value = STRIKE1_ID
    Else
        tSheet.Cells(5, 15).Value = ""
    End If
    tSheet.Cells(9, 15).Value = BASE_ID
    If (STRIKE2 = "Y") Then
        tSheet.Cells(14, 15).Value = STRIKE2_ID
    Else
        tSheet.Cells(14, 15).Value = ""
    End If
    tSheet.Cells(18, 15).Value = FINISH_ID
        
    tSheet.Cells(ROWSTART, 14).Value = SHOT_ID                         'SHOT
    tSheet.Cells(ROWSTART, 15).Value = Format(PART_SA, "0.0")          'PART SA
    tSheet.Cells(ROWSTART, 16).Value = Format(Media_SA, "0.0")         'MEDIA SA
    tSheet.Cells(ROWSTART, 17).Value = Format(SA, "0.0")               'SQ FT
               
   Select Case FR_Table.Fields("[TYPE_ID]")
   Case "BARREL", "SBE"
                '090 Plating chg 07/08/2014
                If (STRIKE1 = "Y") Then
                    tSheet.Cells(7, 14).Value = SKTASF1
                    tSheet.Cells(7, 15).Value = SKTMIN1
                    tSheet.Cells(7, 16).Value = Format(FR_Table.Fields("[SK1 AMP]"), "0.0")
                    tSheet.Cells(7, 17).Value = Format(FR_Table.Fields("[SK1 MIN]"), "0.0")
                
                    'Each Rectifier 1
                    tSheet.Cells(8, 16).Value = Format(FR_Table.Fields("[SK1 AMP]") / NUMBER_HEADS, "0.0")
                    tSheet.Cells(8, 17).Value = Format(FR_Table.Fields("[SK1 MIN]") / NUMBER_HEADS, "0.0")
                Else
                    tSheet.Cells(7, 14).Value = ""
                    tSheet.Cells(7, 15).Value = ""
                    tSheet.Cells(7, 16).Value = ""
                    tSheet.Cells(7, 17).Value = ""
                    tSheet.Cells(8, 16).Value = ""
                    tSheet.Cells(8, 17).Value = ""
                End If
                                    
                tSheet.Cells(12, 14).Value = " "
                tSheet.Cells(12, 16).Value = " "
                tSheet.Cells(12, 17).Value = " "
                                    
                'BASE PROCESS
                If (OVERPLATE = "N") Then
                    tSheet.Cells(11, 14).Value = ASF1
                    tSheet.Cells(11, 15).Value = MIN1
                    tSheet.Cells(11, 16).Value = Format(FR_Table.Fields("[BASE AMP]"), "0.0")
                    tSheet.Cells(11, 17).Value = Format(FR_Table.Fields("[BASE AMP MIN]"), "0.0")
                    tSheet.Cells(11, 18).Value = FR_Table.Fields("[EQ BASE]")
                    Select Case FR_Table.Fields("[TYPE_ID]")
                    Case "BARREL"
                        If (RECTIFIER = 1) Then
                            'Each Rectifier 2
                            tSheet.Cells(12, 14).Value = "Each Rectifier"
                            tSheet.Cells(12, 16).Value = Format(FR_Table.Fields("[BASE AMP]") / NUMBER_HEADS, "0.0")
                            tSheet.Cells(12, 17).Value = Format(FR_Table.Fields("[BASE AMP MIN]") / NUMBER_HEADS, "0.0")
                        Else
                            tSheet.Cells(12, 14).Value = " "
                            tSheet.Cells(12, 16).Value = " "
                            tSheet.Cells(12, 17).Value = " "
                        End If
                    End Select
                    
                Else
                    tSheet.Cells(12, 16).Value = ""
                    tSheet.Cells(12, 17).Value = ""
                    
                    tSheet.Cells(11, 14).Value = ""
                    tSheet.Cells(11, 15).Value = ""
                    tSheet.Cells(11, 16).Value = ""
                    tSheet.Cells(11, 17).Value = ""
                    tSheet.Cells(11, 18).Value = ""
                End If
                
                If (STRIKE2 = "Y") Then
                    tSheet.Cells(16, 14).Value = ASF3
                    tSheet.Cells(16, 15).Value = MIN3
                    tSheet.Cells(16, 16).Value = Format(FR_Table.Fields("[SK2 AMP]"), "0.0")
                    tSheet.Cells(16, 17).Value = Format(FR_Table.Fields("[SK2 MIN]"), "0.0")
                Else
                    tSheet.Cells(16, 14).Value = ""
                    tSheet.Cells(16, 15).Value = ""
                    tSheet.Cells(16, 16).Value = ""
                    tSheet.Cells(16, 17).Value = ""
                End If
                                                
                tSheet.Cells(17, 14).Value = " "
                tSheet.Cells(17, 16).Value = " "
                tSheet.Cells(17, 17).Value = " "
                tSheet.Cells(21, 14).Value = " "
                tSheet.Cells(21, 16).Value = " "
                tSheet.Cells(21, 17).Value = " "
                        
                Select Case FR_Table.Fields("[TYPE_ID]")
                Case "BARREL"
                    If (RECTIFIER = 1) Then
                        'Each Rectifier 3
                        tSheet.Cells(17, 14).Value = "Each Rectifier"
                        tSheet.Cells(17, 16).Value = Format(FR_Table.Fields("[SK2 AMP]") / NUMBER_HEADS, "0.0")
                        tSheet.Cells(17, 17).Value = Format(FR_Table.Fields("[SK2 MIN]") / NUMBER_HEADS, "0.0")
                    Else
                        tSheet.Cells(17, 14).Value = " "
                        tSheet.Cells(17, 16).Value = " "
                        tSheet.Cells(17, 17).Value = " "
                    End If
                End Select
                                                                                                                
                'FINISH PROCESS
                tSheet.Cells(20, 14).Value = ASF2
                tSheet.Cells(20, 15).Value = MIN2
                tSheet.Cells(20, 16).Value = Format(FR_Table.Fields("[FINISH AMP]"), "0.0")
                tSheet.Cells(20, 17).Value = Format(FR_Table.Fields("[FINISH AMP MIN]"), "0.0")
                tSheet.Cells(20, 18).Value = FR_Table.Fields("[EQ FINISH]")
                
                'chg 07/14/14
                
                Select Case FR_Table.Fields("[TYPE_ID]")
                Case "BARREL"
                    If (RECTIFIER = 1) Then
                        'Each Rectifier 4
                        tSheet.Cells(21, 14).Value = "Each Rectifier"
                        tSheet.Cells(21, 16).Value = Format(FR_Table.Fields("[FINISH AMP]") / NUMBER_HEADS, "0.0")
                        tSheet.Cells(21, 17).Value = Format(FR_Table.Fields("[FINISH AMP MIN]") / NUMBER_HEADS, "0.0")
                    Else
                        tSheet.Cells(21, 14).Value = " "
                        tSheet.Cells(21, 16).Value = " "
                        tSheet.Cells(21, 17).Value = " "
                    End If
                    
                End Select
                
    Case Else
    
                
    End Select

End If
        
    tSheet.Cells(23, 15).Value = ""
    tSheet.Cells(23, 16).Value = ""
    tSheet.Cells(23, 17).Value = ""
    tSheet.Cells(23, 18).Value = ""
    
    tSheet.Cells(25, 15).Value = ""
    tSheet.Cells(25, 16).Value = ""
    tSheet.Cells(25, 17).Value = ""
    tSheet.Cells(25, 18).Value = ""
                
    'BASE SPEED
    Select Case FR_Table.Fields("[TYPE_ID]")
    Case "BARREL"
            'tSheet.Cells(23, 15).Value = SPEED_ID
            'tSheet.Cells(23, 16).Value = SPEED_ID
            'tSheet.Cells(23, 17).Value = SPEED_ID
            'tSheet.Cells(23, 18).Value = SPEED_ID
            
            'tSheet.Cells(25, 15).Value = SPEED_ID
            'tSheet.Cells(25, 16).Value = SPEED_ID
            'tSheet.Cells(25, 17).Value = SPEED_ID
            'tSheet.Cells(25, 18).Value = SPEED_ID
           '----------Ana Chavez-------------- 1/25/2021
            tSheet.Cells(23, 15).Value = ""
            tSheet.Cells(23, 16).Value = ""
            tSheet.Cells(23, 17).Value = ""
            tSheet.Cells(23, 18).Value = ""
            
            tSheet.Cells(25, 15).Value = ""
            tSheet.Cells(25, 16).Value = ""
            tSheet.Cells(25, 17).Value = ""
            tSheet.Cells(25, 18).Value = ""
            '--------------------------------------
    Case "SBE"
    
        If (FR_Table.Fields("[SPEED]") <> 0) Then
            tSheet.Cells(23, 15).Value = Format(FR_Table.Fields("[SPEED]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 2]") <> 0) Then
            tSheet.Cells(23, 16).Value = Format(FR_Table.Fields("[SPEED 2]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 3]") <> 0) Then
            tSheet.Cells(23, 17).Value = Format(FR_Table.Fields("[SPEED 3]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 4]") <> 0) Then
            tSheet.Cells(23, 18).Value = Format(FR_Table.Fields("[SPEED 4]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 5]") <> 0) Then
            tSheet.Cells(25, 15).Value = Format(FR_Table.Fields("[SPEED]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 6]") <> 0) Then
            tSheet.Cells(25, 16).Value = Format(FR_Table.Fields("[SPEED 2]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 7]") <> 0) Then
            tSheet.Cells(25, 17).Value = Format(FR_Table.Fields("[SPEED 3]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 8]") <> 0) Then
            tSheet.Cells(25, 18).Value = Format(FR_Table.Fields("[SPEED 4]"), "0.0")
        End If

    End Select
    
    'BASE SERIALS
    'tSheet.Cells(24, 15).Value = FR_Table.Fields("[HEAD 1]")
    'tSheet.Cells(24, 16).Value = FR_Table.Fields("[HEAD 2]")
    'tSheet.Cells(24, 17).Value = FR_Table.Fields("[HEAD 3]")
    'tSheet.Cells(24, 18).Value = FR_Table.Fields("[HEAD 4]")
        
    'tSheet.Cells(26, 15).Value = FR_Table.Fields("[FN HEAD 1]")
    'tSheet.Cells(26, 16).Value = FR_Table.Fields("[FN HEAD 2]")
    'tSheet.Cells(26, 17).Value = FR_Table.Fields("[FN HEAD 3]")
    'tSheet.Cells(26, 18).Value = FR_Table.Fields("[FN HEAD 4]")
    
    ' ------------Ana Chavez-------1/25/2021
    tSheet.Cells(24, 15).Value = ""
    tSheet.Cells(24, 16).Value = ""
    tSheet.Cells(24, 17).Value = ""
    tSheet.Cells(24, 18).Value = ""
        
    tSheet.Cells(26, 15).Value = ""
    tSheet.Cells(26, 16).Value = ""
    tSheet.Cells(26, 17).Value = ""
    tSheet.Cells(26, 18).Value = ""
    
    ROWSTART = 5
    iRow = 0
    sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
           
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    
    If (FR_Table.RecordCount <> 0) Then
        
        Dim I As Integer
        For I = iRow + 1 To 30
            tSheet.Cells(I + ROWSTART, 1).Value = ""
            tSheet.Cells(I + ROWSTART, 2).Value = ""
            tSheet.Cells(I + ROWSTART, 3).Value = ""
            tSheet.Cells(I + ROWSTART, 4).Value = ""
            tSheet.Cells(I + ROWSTART, 5).Value = ""
            tSheet.Cells(I + ROWSTART, 6).Value = ""
            tSheet.Cells(I + ROWSTART, 7).Value = ""
            tSheet.Cells(I + ROWSTART, 8).Value = ""
            tSheet.Cells(I + ROWSTART, 9).Value = ""
            tSheet.Cells(I + ROWSTART, 10).Value = ""
            tSheet.Cells(I + ROWSTART, 11).Value = ""
        Next I
        
        Do Until FR_Table.EOF
            iRow = iRow + 1
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                
                sBuff = FR_Table.Fields("[WORK ORDER]")
                'WO
                tSheet.Cells(iRow + ROWSTART, 2).Value = Mid(sBuff, 1, 6)
                'ITEM
                tSheet.Cells(iRow + ROWSTART, 3).Value = Mid(sBuff, 7, 3) & "." & Mid(sBuff, 10, 1)
                'RUN
                tSheet.Cells(iRow + ROWSTART, 4).Value = Mid(sBuff, 11, 2)
            Else
                'LOT NUM
                 tSheet.Cells(iRow + ROWSTART, 2).Value = FR_Table.Fields("[WORK ORDER]")
            End If
            ' ATC PART NUMBER
            tSheet.Cells(iRow + ROWSTART, 5).Value = FR_Table.Fields("[ATC PART]")
            
            ' GEAR QTY
            If (FR_Table.Fields("[P1 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 6).Value = Format(FR_Table.Fields("[P1 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P2 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 7).Value = Format(FR_Table.Fields("[P2 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P3 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 8).Value = Format(FR_Table.Fields("[P3 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P4 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 9).Value = Format(FR_Table.Fields("[P4 BASE]"), "###,###")
            End If
        
            tSheet.Cells(iRow + ROWSTART, 10).Value = Format(FR_Table.Fields("[P1 BASE]") + FR_Table.Fields("[P2 BASE]") + FR_Table.Fields("[P3 BASE]") + FR_Table.Fields("[P4 BASE]"), "###,###")
            
            tSheet.Cells(iRow + ROWSTART, 11).Value = FR_Table.Fields("[LETTER_ID]")
            
            FR_Table.MoveNext
        Loop
        Beep
        
    End If
   
   '==================================================================
   '    FOOT SUMMARY
   '==================================================================
    sSQL = "SELECT count([WORK ORDER]) AS [SQL COUNT WO]," & _
                           "sum([QTY]) AS [SQL SUM QTY]," & _
                       "sum([P1 BASE]) AS [SQL SUM G1]," & _
                       "sum([P2 BASE]) AS [SQL SUM G2]," & _
                       "sum([P3 BASE]) AS [SQL SUM G3]," & _
                       "sum([P4 BASE]) AS [SQL SUM G4]," & _
                       "sum([P1 BASE])+sum([P2 BASE])+sum([P3 BASE])+sum([P4 BASE]) AS [SQL SUM TOTAL] " & _
           "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " GROUP BY [SET_ID]"
           
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    ROWSTART = 31
    If (FR_Table.RecordCount <> 0) Then
            tSheet.Cells(3, 8).Value = FR_Table.Fields("[SQL COUNT WO]")         ' WO's
            tSheet.Cells(ROWSTART, 6).Value = Format(FR_Table.Fields("[SQL SUM G1]"), "###,###")
            tSheet.Cells(ROWSTART, 7).Value = Format(FR_Table.Fields("[SQL SUM G2]"), "###,###")
            tSheet.Cells(ROWSTART, 8).Value = Format(FR_Table.Fields("[SQL SUM G3]"), "###,###")
            tSheet.Cells(ROWSTART, 9).Value = Format(FR_Table.Fields("[SQL SUM G4]"), "###,###")
            tSheet.Cells(ROWSTART, 10).Value = Format(FR_Table.Fields("[SQL SUM TOTAL]"), "###,###")
    End If
       
FR_Table.Close
FR_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Dim sFile As String
sFile = PLATING_LOG_SHEET
                                                                                                
shtWorld.SaveAs sFile

shtWorld.Application.Quit

Set shtWorld = Nothing

MsgBox "Excel Update Complete", vbInformation, "ATC Plating"

End Sub
Public Sub ExcelReportOlean()

Dim sBuff As String
             
Dim STRIKE1 As String
Dim STRIKE2 As String
Dim STRIKE1_ID As String
Dim STRIKE2_ID As String
Dim BASE_ID As String
Dim FINISH_ID As String
Dim OVERPLATE As String
Dim DESCRIPTION As String

Dim sSQL As String

Dim I As Integer

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT * FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        DEPT_ID = FR_Table.Fields("[DEPT_ID]")
        CASE_SIZE_ID = Mid(FR_Table.Fields("[SERIES_ID]"), 4, 1)
        SERIES_ID = Val(Mid(FR_Table.Fields("[SERIES_ID]"), 1, 3))
        TYPE_ID = FR_Table.Fields("[TYPE_ID]")
Else
        Exit Sub
End If

sSQL = "SELECT * FROM [TBL PLATING OLEAN] WHERE [ID] = 1"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    ASF1 = FR_Table.Fields("[BASE AMP]")
    MIN1 = FR_Table.Fields("[BASE MIN]")
    ASF2 = FR_Table.Fields("[FINISH AMP]")
    MIN2 = FR_Table.Fields("[FINISH MIN]")
End If
    
sSQL = "SELECT count([WORK ORDER])," & _
              "format(sum([QTY]),'###,####')," & _
              "sum([P1 BASE]) AS [SQL QTY1]," & _
              "sum([P2 BASE]) AS [SQL QTY2]," & _
              "sum([P3 BASE]) AS [SQL QTY3]," & _
              "sum([P4 BASE]) AS [SQL QTY4] " & _
       "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " " & _
       "GROUP BY [SET_ID]"
    
Set FR_Table = FR_Database.OpenRecordset(sSQL)

GEAR_1_QTY = 0
GEAR_2_QTY = 0
GEAR_3_QTY = 0
GEAR_4_QTY = 0
TOTAL_QTY = 0
NUMBER_HEADS = 0

If (FR_Table.RecordCount = 0) Then
    Exit Sub
Else
    For I = 1 To 4
        sBuff = "[SQL QTY" & I & "]"
        If (FR_Table.Fields(sBuff) <> 0) Then
            NUMBER_HEADS = NUMBER_HEADS + 1
        End If
    Next I
    GEAR_1_QTY = FR_Table.Fields("[SQL QTY1]")
    GEAR_2_QTY = FR_Table.Fields("[SQL QTY2]")
    GEAR_3_QTY = FR_Table.Fields("[SQL QTY3]")
    GEAR_4_QTY = FR_Table.Fields("[SQL QTY4]")
            
    TOTAL_QTY = FR_Table.Fields("[SQL QTY1]") + FR_Table.Fields("[SQL QTY2]") + FR_Table.Fields("[SQL QTY3]") + FR_Table.Fields("[SQL QTY4]")
End If
FR_Table.Close
FR_Database.Close

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT [DEPT_ID],[DEPT_ID],[DESCRIPTION],[BASE_ID],[FINISH_ID]," & _
              "[TANK DWG],[SBE],[TANK],[STRIKE1],[STRIKE2],[OVERPLATE],[ACTIVE] " & _
       "FROM [DEPT CODE] ORDER BY [DEPT_ID]"

Select Case LOCATION_ID
Case "NY"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_ID]=" & DEPT_ID
Case "JR"
        sSQL = "SELECT * FROM [DEPT CODE] WHERE [DEPT_JR_ID]=" & DEPT_ID
End Select


Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    STRIKE1 = FR_Table.Fields("[STRIKE1]")      'Y/N
    STRIKE2 = FR_Table.Fields("[STRIKE2]")      'Y/N
    STRIKE1_ID = FR_Table.Fields("[STRIKE1_ID]")
    STRIKE2_ID = FR_Table.Fields("[STRIKE2_ID]")      'Y/N
    OVERPLATE = FR_Table.Fields("[OVERPLATE]")  'Y/N
    BASE_ID = FR_Table.Fields("[BASE_ID]")      '
    FINISH_ID = FR_Table.Fields("[FINISH_ID]")
    DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
End If

Dim wbWorld As Object, shtWorld As Object
Dim tSheet As Object

Set shtWorld = GetObject(PLATING_LOG_SHEET)
    
Set wbWorld = shtWorld.Application.Workbooks("Plating Log Sheet Master.xls")
Set tSheet = wbWorld.Sheets("New Format")
'chg 10/06/2011 ADDITION
'shtWorld.Application.Visible = True
tSheet.Parent.WINDOWS(1).Visible = True
            
Screen.MousePointer = vbHourglass
                
Dim iRow As Integer, iCol As Integer
                 
Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

sSQL = "SELECT *  FROM [SCHEDULE SETS] WHERE [SET_ID]=" & SET_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

'====================================================================
If (FR_Table.RecordCount <> 0) Then
    EQ_BASE_ID = FR_Table.Fields("[EQ BASE]")
End If

sSQL = "SELECT * FROM [MACHINE] WHERE [NUMBER]=" & EQ_BASE_ID
Set TO_Table = TO_Database.OpenRecordset(sSQL)


'chg 03/01/2011

RECTIFIER = 1
If (TO_Table.RecordCount <> 0) Then
    RECTIFIER = TO_Table.Fields("[RECTIFIER]")
End If
TO_Database.Close
'====================================================================

Dim ROWSTART As Integer

ROWSTART = 3

If (FR_Table.RecordCount <> 0) Then
                                                                       
    tSheet.Cells(ROWSTART, 2).Value = FR_Table.Fields("[SET NUMBER]")        'SET#
    tSheet.Cells(ROWSTART, 3).Value = Format(FR_Table.Fields("[DATE_ID]"), "MM/DD/YYYY")          'DATE
    tSheet.Cells(ROWSTART, 4).Value = FR_Table.Fields("[DEPT_ID]")           'DEPT_ID
    tSheet.Cells(ROWSTART, 5).Value = DESCRIPTION     'DESCRIPTION
    tSheet.Cells(ROWSTART, 6).Value = FR_Table.Fields("[TYPE_ID]")           'Process
    tSheet.Cells(ROWSTART, 7).Value = FR_Table.Fields("[SERIES_ID]")         'SERIES CASE
                                                                             
    If (STRIKE1 = "Y") Then
        tSheet.Cells(5, 15).Value = STRIKE1_ID
    Else
        tSheet.Cells(5, 15).Value = ""
    End If
    tSheet.Cells(9, 15).Value = BASE_ID
    If (STRIKE2 = "Y") Then
        tSheet.Cells(14, 15).Value = STRIKE2_ID
    Else
        tSheet.Cells(14, 15).Value = ""
    End If
    tSheet.Cells(18, 15).Value = FINISH_ID
        
    SHOT_ID = 200
    PART_SA = 0
    Media_SA = 0
    SA = 0
        
    tSheet.Cells(ROWSTART, 14).Value = SHOT_ID                         'SHOT
    tSheet.Cells(ROWSTART, 15).Value = Format(PART_SA, "0.0")          'PART SA
    tSheet.Cells(ROWSTART, 16).Value = Format(Media_SA, "0.0")         'MEDIA SA
    tSheet.Cells(ROWSTART, 17).Value = Format(SA, "0.0")               'SQ FT
               
   Select Case FR_Table.Fields("[TYPE_ID]")
   Case "BARREL", "SBE"
                
                If (STRIKE1 = "Y") Then
                    tSheet.Cells(7, 14).Value = SKTASF1
                    tSheet.Cells(7, 15).Value = SKTMIN1
                    tSheet.Cells(7, 16).Value = Format(FR_Table.Fields("[SK1 AMP]"), "0.0")
                    tSheet.Cells(7, 17).Value = Format(FR_Table.Fields("[SK1 MIN]"), "0.0")
                Else
                    tSheet.Cells(7, 14).Value = ""
                    tSheet.Cells(7, 15).Value = ""
                    tSheet.Cells(7, 16).Value = ""
                    tSheet.Cells(7, 17).Value = ""
                End If
                tSheet.Cells(12, 14).Value = " "
                tSheet.Cells(12, 16).Value = " "
                tSheet.Cells(12, 17).Value = " "
                
                'BASE PROCESS
                If (OVERPLATE = "N") Then
                    tSheet.Cells(11, 14).Value = "" 'ASF1
                    tSheet.Cells(11, 15).Value = MIN1
                    tSheet.Cells(11, 16).Value = Format(FR_Table.Fields("[BASE AMP]"), "0.0")    'chg
                    tSheet.Cells(11, 17).Value = Format(FR_Table.Fields("[BASE AMP MIN]"), "0.0")   'chg
                    tSheet.Cells(11, 18).Value = FR_Table.Fields("[EQ BASE]")
                    
                    Select Case FR_Table.Fields("[TYPE_ID]")
                    Case "BARREL"
                        If (RECTIFIER = 1) Then
                            tSheet.Cells(12, 14).Value = "Each Rectifier"
                            tSheet.Cells(12, 16).Value = Format(FR_Table.Fields("[BASE AMP]") / NUMBER_HEADS, "0.0") 'chg
                            tSheet.Cells(12, 17).Value = Format(FR_Table.Fields("[BASE AMP MIN]") / NUMBER_HEADS, "0.0") 'chg
                        End If
                    End Select
                    
                Else
                    tSheet.Cells(12, 16).Value = ""
                    tSheet.Cells(12, 17).Value = ""
                    
                    tSheet.Cells(11, 14).Value = ""
                    tSheet.Cells(11, 15).Value = ""
                    tSheet.Cells(11, 16).Value = ""
                    tSheet.Cells(11, 17).Value = ""
                    tSheet.Cells(11, 18).Value = ""
                End If
                
                If (STRIKE2 = "Y") Then
                    tSheet.Cells(16, 14).Value = ASF3
                    tSheet.Cells(16, 15).Value = MIN3
                    tSheet.Cells(16, 16).Value = Format(FR_Table.Fields("[SK2 AMP]"), "0.0")
                    tSheet.Cells(16, 17).Value = Format(FR_Table.Fields("[SK2 MIN]"), "0.0")
                Else
                    tSheet.Cells(16, 14).Value = ""
                    tSheet.Cells(16, 15).Value = ""
                    tSheet.Cells(16, 16).Value = ""
                    tSheet.Cells(16, 17).Value = ""
                End If
                
                   'FINISH PROCESS
                tSheet.Cells(20, 14).Value = "" 'ASF2
                tSheet.Cells(20, 15).Value = MIN2
                tSheet.Cells(20, 16).Value = Format(FR_Table.Fields("[FINISH AMP]"), "0.0")
                tSheet.Cells(20, 17).Value = Format(FR_Table.Fields("[FINISH AMP MIN]"), "0.0")
                tSheet.Cells(20, 18).Value = FR_Table.Fields("[EQ FINISH]")
                
                '090 Plating chg 08/28/2015
                
                Select Case FR_Table.Fields("[TYPE_ID]")
                Case "BARREL"
                    If (RECTIFIER = 1) Then
                        'Each Rectifier 4
                        tSheet.Cells(21, 14).Value = "Each Rectifier"
                        tSheet.Cells(21, 16).Value = Format(FR_Table.Fields("[FINISH AMP]") / NUMBER_HEADS, "0.0")
                        tSheet.Cells(21, 17).Value = Format(FR_Table.Fields("[FINISH AMP MIN]") / NUMBER_HEADS, "0.0")
                    Else
                        tSheet.Cells(21, 14).Value = " "
                        tSheet.Cells(21, 16).Value = " "
                        tSheet.Cells(21, 17).Value = " "
                    End If
                End Select
                                                                                                                                                   
    Case Else
                    
    End Select

End If
        
    tSheet.Cells(23, 15).Value = ""
    tSheet.Cells(23, 16).Value = ""
    tSheet.Cells(23, 17).Value = ""
    tSheet.Cells(23, 18).Value = ""
    
    tSheet.Cells(25, 15).Value = ""
    tSheet.Cells(25, 16).Value = ""
    tSheet.Cells(25, 17).Value = ""
    tSheet.Cells(25, 18).Value = ""
        
    'BASE SPEED
    Select Case FR_Table.Fields("[TYPE_ID]")
    Case "BARREL"
            tSheet.Cells(23, 15).Value = SPEED_ID
            tSheet.Cells(23, 16).Value = SPEED_ID
            tSheet.Cells(23, 17).Value = SPEED_ID
            tSheet.Cells(23, 18).Value = SPEED_ID
            
            tSheet.Cells(25, 15).Value = SPEED_ID
            tSheet.Cells(25, 16).Value = SPEED_ID
            tSheet.Cells(25, 17).Value = SPEED_ID
            tSheet.Cells(25, 18).Value = SPEED_ID
    Case "SBE"
    
        If (FR_Table.Fields("[SPEED]") <> 0) Then
            tSheet.Cells(23, 15).Value = Format(FR_Table.Fields("[SPEED]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 2]") <> 0) Then
            tSheet.Cells(23, 16).Value = Format(FR_Table.Fields("[SPEED 2]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 3]") <> 0) Then
            tSheet.Cells(23, 17).Value = Format(FR_Table.Fields("[SPEED 3]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 4]") <> 0) Then
            tSheet.Cells(23, 18).Value = Format(FR_Table.Fields("[SPEED 4]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 5]") <> 0) Then
            tSheet.Cells(25, 15).Value = Format(FR_Table.Fields("[SPEED]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 6]") <> 0) Then
            tSheet.Cells(25, 16).Value = Format(FR_Table.Fields("[SPEED 2]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 7]") <> 0) Then
            tSheet.Cells(25, 17).Value = Format(FR_Table.Fields("[SPEED 3]"), "0.0")
        End If
        If (FR_Table.Fields("[SPEED 8]") <> 0) Then
            tSheet.Cells(25, 18).Value = Format(FR_Table.Fields("[SPEED 4]"), "0.0")
        End If

    End Select
    
    'BASE SERIALS
    tSheet.Cells(24, 15).Value = FR_Table.Fields("[HEAD 1]")
    tSheet.Cells(24, 16).Value = FR_Table.Fields("[HEAD 2]")
    tSheet.Cells(24, 17).Value = FR_Table.Fields("[HEAD 3]")
    tSheet.Cells(24, 18).Value = FR_Table.Fields("[HEAD 4]")
        
    tSheet.Cells(26, 15).Value = FR_Table.Fields("[FN HEAD 1]")
    tSheet.Cells(26, 16).Value = FR_Table.Fields("[FN HEAD 2]")
    tSheet.Cells(26, 17).Value = FR_Table.Fields("[FN HEAD 3]")
    tSheet.Cells(26, 18).Value = FR_Table.Fields("[FN HEAD 4]")
    
    ROWSTART = 5
    iRow = 0
    sSQL = "SELECT * FROM [GROUPING] WHERE [SET_ID]=" & SET_ID
           
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    
    If (FR_Table.RecordCount <> 0) Then
        For I = iRow + 1 To 30
            tSheet.Cells(I + ROWSTART, 1).Value = ""
            tSheet.Cells(I + ROWSTART, 2).Value = ""
            tSheet.Cells(I + ROWSTART, 3).Value = ""
            tSheet.Cells(I + ROWSTART, 4).Value = ""
            tSheet.Cells(I + ROWSTART, 5).Value = ""
            tSheet.Cells(I + ROWSTART, 6).Value = ""
            tSheet.Cells(I + ROWSTART, 7).Value = ""
            tSheet.Cells(I + ROWSTART, 8).Value = ""
            tSheet.Cells(I + ROWSTART, 9).Value = ""
            tSheet.Cells(I + ROWSTART, 10).Value = ""
            tSheet.Cells(I + ROWSTART, 11).Value = ""
        Next I
    
        Do Until FR_Table.EOF
            iRow = iRow + 1
            If (Len(FR_Table.Fields("[WORK ORDER]")) = 12) Then
                'WO
                
                sBuff = FR_Table.Fields("[WORK ORDER]")
                
                tSheet.Cells(iRow + ROWSTART, 2).Value = Mid(sBuff, 1, 6)
                'ITEM
                tSheet.Cells(iRow + ROWSTART, 3).Value = Mid(sBuff, 7, 3) & "." & Mid(sBuff, 10, 1)
                'RUN
                tSheet.Cells(iRow + ROWSTART, 4).Value = Mid(sBuff, 11, 2)
            Else
                'LOT NUM
                 tSheet.Cells(iRow + ROWSTART, 2).Value = FR_Table.Fields("[WORK ORDER]")
            End If
            ' ATC PART NUMBER
            tSheet.Cells(iRow + ROWSTART, 5).Value = FR_Table.Fields("[ATC PART]")
            
            ' GEAR QTY
            If (FR_Table.Fields("[P1 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 6).Value = Format(FR_Table.Fields("[P1 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P2 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 7).Value = Format(FR_Table.Fields("[P2 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P3 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 8).Value = Format(FR_Table.Fields("[P3 BASE]"), "###,###")
            End If
            If (FR_Table.Fields("[P4 BASE]") <> 0) Then
                tSheet.Cells(iRow + ROWSTART, 9).Value = Format(FR_Table.Fields("[P4 BASE]"), "###,###")
            End If
        
            tSheet.Cells(iRow + ROWSTART, 10).Value = Format(FR_Table.Fields("[P1 BASE]") + FR_Table.Fields("[P2 BASE]") + FR_Table.Fields("[P3 BASE]") + FR_Table.Fields("[P4 BASE]"), "###,###")
            
            tSheet.Cells(iRow + ROWSTART, 11).Value = FR_Table.Fields("[LETTER_ID]")
            
            FR_Table.MoveNext
        Loop
        Beep
        
    End If
   
   '==================================================================
   '    FOOT SUMMARY
   '==================================================================
    sSQL = "SELECT count([WORK ORDER]) AS [SQL COUNT WO]," & _
                    "sum([QTY])        AS [SQL SUM QTY]," & _
                    "sum([P1 BASE])    AS [SQL SUM G1]," & _
                    "sum([P2 BASE])    AS [SQL SUM G2]," & _
                    "sum([P3 BASE])    AS [SQL SUM G3]," & _
                    "sum([P4 BASE])    AS [SQL SUM G4] " & _
           "FROM [GROUPING] WHERE [SET_ID]=" & SET_ID & " GROUP BY [SET_ID]"
           
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    ROWSTART = 31
    If (FR_Table.RecordCount <> 0) Then
            tSheet.Cells(3, 8).Value = FR_Table.Fields("[SQL COUNT WO]")         ' WO's
            tSheet.Cells(ROWSTART, 6).Value = Format(FR_Table.Fields("[SQL SUM G1]"), "###,###")
            tSheet.Cells(ROWSTART, 7).Value = Format(FR_Table.Fields("[SQL SUM G2]"), "###,###")
            tSheet.Cells(ROWSTART, 8).Value = Format(FR_Table.Fields("[SQL SUM G3]"), "###,###")
            tSheet.Cells(ROWSTART, 9).Value = Format(FR_Table.Fields("[SQL SUM G4]"), "###,###")
            tSheet.Cells(ROWSTART, 10).Value = Format(FR_Table.Fields("[SQL SUM QTY]"), "###,###")
    End If
       
FR_Table.Close
FR_Database.Close

Screen.MousePointer = vbDefault
                                                                                                                         
Dim sFile As String
sFile = PLATING_LOG_SHEET
                                                                                                
shtWorld.SaveAs sFile

shtWorld.Application.Quit

Set shtWorld = Nothing

MsgBox "Excel Update Complete", vbInformation, "ATC Plating"

End Sub

Public Sub Get_DV()

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Dim sSQL As String
sSQL = "SELECT [GP_ID],[SET_ID],[WORK ORDER],[LOT NUM],[ATC PART],[LETTER_ID],[DV] " & _
       "FROM [GROUPING] WHERE   [LETTER_ID]<> ' ' AND [DV] = 0 " & _
       "ORDER BY [GP_ID] DESC"

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
   Do Until FR_Table.EOF
        FR_Table.Edit
        FR_Table.Fields("[DV]") = LotDV(FR_Table.Fields("[ATC PART]"))
        FR_Table.Update
        FR_Table.MoveNext
   Loop
End If
FR_Table.Close
FR_Database.Close

End Sub

Public Function Calculate_Plating_Table() As Integer

Dim DV(10) As Double
Dim COUNT As Integer
Dim BIN_COUNT As Integer

Calculate_Plating_Table = 0

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
                                
sSQL = "SELECT * FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND " & _
             "[LETTER_ID]='" & LETTER_ID & "' AND " & _
             "[LOT NUM]='LOT NUM'"
     
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        'Missing Lot Number
        Calculate_Plating_Table = 1
        Exit Function
End If
                
sSQL = "SELECT count([DV]) AS [SQL COUNT] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "GROUP BY [SET_ID]& [LETTER_ID],[DV]"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
'
'   GET NUMBER OF BINS REQUIRED 2-4 RANGE OR 5-10
'
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        FR_Table.MoveNext
    Loop
End If

Dim I As Integer
For I = 0 To 10
    gdBinLimit(I) = 0
    gsBinTol(I) = ""
Next I

Select Case COUNT
Case 1
        '
        ' No Sort Required
        '
        Calculate_Plating_Table = 2
               
        sSQL = "SELECT count([DV]) AS [SQL COUNT] " & _
               "FROM [GROUPING] " & _
               "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
               "GROUP BY [SET_ID]& [LETTER_ID],mid([ATC PART],5,4)"
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
        '
        ' TEST FOR SAME VALUE AND TOLERANCE
        '
        COUNT = 0
        If (FR_Table.RecordCount <> 0) Then
            Do Until FR_Table.EOF
                COUNT = COUNT + 1
                FR_Table.MoveNext
            Loop
        End If
        Select Case COUNT
        Case 1
                Calculate_Plating_Table = 2  ' No Sort Required
        Case Else
                Calculate_Plating_Table = 3  ' Lot Sort Required
        End Select
                
        Exit Function
Case 2 To 10

Case Else
      ' Too Many Bins Second Sort Required not able to generate Table
        Calculate_Plating_Table = 4
        Exit Function
End Select

sSQL = "SELECT first([DV])               AS [SQL DV]," & _
             "MAX(MID([ATC PART],8,1))   AS [SQL TOL] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "GROUP BY mid([ATC PART],5,3) " & _
       "ORDER BY first([DV]) ASC"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    I = 1
    Select Case COUNT
    Case 1 To 4
                    Do Until FR_Table.EOF
                        Select Case FR_Table.Fields("[SQL DV]")
                        Case 0.1 To 1
                                    gdBinLimit(I) = 0.01
                                    gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.2
                        Case 1 To 10
                                    Select Case FR_Table.Fields("[SQL TOL]")
                                    Case "A"    '.05 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.3
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.3
                                    Case "B"    '.1 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.3
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.3
                                    Case "C"    '.25 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.35
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.35
                                    Case "D"    '.5 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.8
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.8
                                    Case Else
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 1
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 1
                                    End Select
                        Case Else
                                  Select Case FR_Table.Fields("[SQL TOL]")
                                  Case "F" 'F 1%  - 0.04
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.96
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.04
                                  Case "G" 'G 2%  - 0.05
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.95
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.05
                                  Case "J"  '5%   - 0.08
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.92
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.08
                                  Case "K"  '10%  - 0.15
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.85
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.15
                                  Case Else
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.8
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.2
                                  End Select
                        End Select
                        
                        gsBinTol(I) = FR_Table.Fields("[SQL DV]") & FR_Table.Fields("[SQL TOL]")
                        FR_Table.MoveNext
                        BIN_COUNT = I
                        I = I + 2
                    Loop
    Case 5 To 9
                    Do Until FR_Table.EOF

                        Select Case FR_Table.Fields("[SQL DV]")
                        Case 0.1 To 1
                                    gdBinLimit(I) = 0.01
                                    gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.2
                        Case 1 To 10
                                    Select Case FR_Table.Fields("[SQL TOL]")
                                    Case "A"    '.05 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.3
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.3
                                    Case "B"    '.1 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.3
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.3
                                    Case "C"    '.25 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.35
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.35
                                    Case "D"    '.5 PF
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 0.7
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 0.7
                                    Case Else
                                            gdBinLimit(I) = FR_Table.Fields("[SQL DV]") - 1
                                            gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") + 1
                                    End Select
                        Case Else
                                  Select Case FR_Table.Fields("[SQL TOL]")
                                  Case "F" 'F 1%  - 0.04
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.96
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.04
                                  Case "G" 'G 2%  - 0.05
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.95
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.05
                                  Case "J"  '5%   - 0.08
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.92
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.08
                                  Case "K"  '10%  - 0.15
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.85
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.15
                                  Case Else
                                        gdBinLimit(I) = FR_Table.Fields("[SQL DV]") * 0.8
                                        gdBinLimit(I + 1) = FR_Table.Fields("[SQL DV]") * 1.2
                                  End Select
                        End Select
                                                                        
                        gsBinTol(I) = FR_Table.Fields("[SQL DV]") & FR_Table.Fields("[SQL TOL]")
                        
                        FR_Table.MoveNext
                        BIN_COUNT = I
                        I = I + 1
                    Loop
    End Select
End If

Select Case COUNT
Case 2 To 10

Case Else
        Calculate_Plating_Table = 4
        Exit Function
End Select

For I = 1 To BIN_COUNT
    If (gdBinLimit(I) <> 0) Then
        If gdBinLimit(I + 1) > gdBinLimit(I) Then
        
        Else
            Calculate_Plating_Table = 7     ' Not Able to generate Table
        End If
    End If
Next I
'
'   Test for 5 Lot Sort Required
'
sSQL = "SELECT * " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND " & _
           "[LETTER_ID]='" & LETTER_ID & "' AND " & _
    "len(trim([WORK ORDER]))=10 "
                   
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
        Calculate_Plating_Table = 5
        Exit Function
End If
'
'   6 Second Sort Required
'
' COMPARE [GROUP BY 7 COUNT] TO [GROUP BY 8 COUNT]

Dim COUNT7 As Integer
Dim COUNT8 As Integer
                
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

    Dim sFileName2 As String
    sFileName2 = DB_REPORT_ADDR & SET_ID & LETTER_ID & ".DAT"
    Dim iFilenum As Integer
    iFilenum = FreeFile
    Open sFileName2 For Output Shared As #iFilenum
    For I = 0 To 9
        Print #iFilenum, gsBinTol(I + 1)
    Next I
    For I = 0 To 9
        If (gdBinLimit(I + 1) <> 0) Then
            Select Case gdBinLimit(I + 1)
            Case 0 To 10
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0.00")
            Case 10 To 20
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0.0")
            Case Else
                    Print #iFilenum, Format(gdBinLimit(I + 1), "######0")
            End Select
        Else
                    Print #iFilenum, "0"
        End If
    Next I
    Close iFilenum

If (COUNT7 = COUNT8) Then
    
Else
        Calculate_Plating_Table = 6
        Exit Function
End If
        

End Function

Public Sub Plating_Sort_DB()


On Error GoTo Error_Plating_Sort_DB

  Screen.MousePointer = vbHourglass

    Dim sSQL   As String
    Dim sBuff   As String
    Dim I As Integer
    Set FR_Database = OpenDatabase(DB_PLATING_SORT_TABLE)
    
    sSQL = "SELECT * FROM [9 Bin Special Sort] WHERE [Table Name]='Set " & SET_ID & LETTER_ID & "'"
    
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    
    If (FR_Table.RecordCount = 0) Then
             FR_Table.AddNew
             FR_Table.Fields("[Table Name]") = "Set " & SET_ID & LETTER_ID
             FR_Table.Fields("[TABLE_ID]") = FR_Table.Fields("[TABLE ID]")
             FR_Table.Fields("[Table Type]") = 4
             FR_Table.Fields("[DF Hi]") = 199
             FR_Table.Fields("[DF Lo]") = -199
             FR_Table.Fields("[Test Frequency]") = "1KHZ"
             FR_Table.Fields("[Active]") = vbTrue
             For I = 0 To 8
                     sBuff = "[Tol Code " & I & "]"
                     If Len(gsBinTol(I + 1)) > 1 Then
                        FR_Table.Fields(sBuff) = gsBinTol(I + 1)
                     Else
                        FR_Table.Fields(sBuff) = "NA"
                     End If
             Next I
             For I = 0 To 9
                    sBuff = "[Bin Limit " & I & "]"
                    If (gdBinLimit(I + 1) <> 0) Then
                        Select Case gdBinLimit(I + 1)
                        Case 0 To 10
                                FR_Table.Fields(sBuff) = Format(gdBinLimit(I + 1), "######0.00")
                        Case 10 To 20
                                FR_Table.Fields(sBuff) = Format(gdBinLimit(I + 1), "######0.0")
                        Case Else
                                FR_Table.Fields(sBuff) = Format(gdBinLimit(I + 1), "######0")
                        End Select
                    Else
                        FR_Table.Fields(sBuff) = 0
                    End If
            Next I
            FR_Table.Fields("[DATE_ID]") = Date
            FR_Table.Update
    End If
    
    FR_Database.Close

    Screen.MousePointer = vbDefault
    
    Exit Sub

Error_Plating_Sort_DB:
  Screen.MousePointer = vbDefault
    
End Sub
