VERSION 5.00
Begin VB.Form frmMSGBOX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "090 Message Box.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMessage1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Missing Lot Number and ATC Part"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   7065
   End
End
Attribute VB_Name = "frmMSGBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Caption = "Plate Grouping Message     " & ATC_DWG & "    " & ATC_VERSION

Set FR_Database = OpenDatabase(DB_PLATING_TERMINATION)
Set TO_Database = OpenDatabase(DB_PLATING_TERMINATION)

Dim sSQL As String
                                
sSQL = "SELECT * FROM [GROUPING] " & _
       "WHERE [SET_ID] =" & SET_ID & " AND " & _
          "[LETTER_ID] ='" & LETTER_ID & "' AND " & _
            "[LOT NUM] ='LOT NUM'"
     
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
        lblMessage1.Caption = "Missing Lot Number and ATC Part"
        Exit Sub
End If
                
'===================================================================
'   [1] TEST FOR MISSING LOT NUMBERS
'===================================================================

sSQL = "SELECT FIRST([DV]) AS [SQL FIRST]," & _
              "COUNT([DV]) AS [SQL COUNT] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "GROUP BY [SET_ID]& [LETTER_ID],[DV] HAVING COUNT([DV])>1"
Set FR_Table = FR_Database.OpenRecordset(sSQL)

'===================================================================
'   [2] GROUPING [DV] COUNT GREATER THAN 1
'===================================================================

Dim COUNT As Integer
If (FR_Table.RecordCount <> 0) Then

'    MsgBox "Yes GROUPING DV COUNT GREATER THAN 1"
    
    Do Until FR_Table.EOF
            sSQL = "SELECT COUNT([DV]) AS [SQL COUNT] " & _
                   "FROM [GROUPING] " & _
                   "WHERE [SET_ID]=" & SET_ID & " AND " & _
                      "[LETTER_ID]='" & LETTER_ID & "' AND " & _
                      "[DV]=" & FR_Table.Fields("[SQL FIRST]") & " " & _
                   "GROUP BY [SET_ID]& [LETTER_ID],[LOT NUM] HAVING COUNT([DV])>1"
            Set TO_Table = TO_Database.OpenRecordset(sSQL)
        
            COUNT = 0
            Do Until TO_Table.EOF
                    COUNT = COUNT + 1
                    TO_Table.MoveNext
            Loop
            Select Case COUNT
            Case 1
                    
            Case Else
                    lblMessage1.Caption = "DV " & FR_Table.Fields("[SQL FIRST]") & " Mixed Lots"
                    Exit Sub
            End Select
            FR_Table.MoveNext
    Loop
End If


sSQL = "SELECT mid([ATC PART],1,8) AS [SQL ATC PART]," & _
              "mid([ATC PART],8,1) AS [SQL TOL]," & _
              "[DV]                AS [SQL DV] " & _
       "FROM [GROUPING] " & _
       "WHERE [SET_ID]=" & SET_ID & " AND [LETTER_ID]='" & LETTER_ID & "' " & _
       "ORDER BY  mid([ATC PART],1,8) "
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim ATC_TOL As String
Dim ATC_DV As Double

If (FR_Table.RecordCount <> 0) Then
        
    ATC_TOL = FR_Table.Fields("[SQL TOL]")
    ATC_DV = FR_Table.Fields("[SQL DV]")
    Do
        FR_Table.MoveNext
        If FR_Table.EOF Then Exit Do
        
        If ATC_DV = FR_Table.Fields("[SQL DV]") Then
            If ATC_TOL = FR_Table.Fields("[SQL TOL]") Then
                'OK
            Else
                lblMessage1.Caption = "Mixed Tolerance Parts"
                Exit Sub
            End If
        End If
        ATC_TOL = FR_Table.Fields("[SQL TOL]")
        ATC_DV = FR_Table.Fields("[SQL DV]")
    Loop
End If

lblMessage1.Caption = "Plating Group Pass"

End Sub
