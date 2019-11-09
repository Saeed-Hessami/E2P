Attribute VB_Name = "eHubtoP6"
Option Explicit
Dim iTotalNewItems As Integer
Dim iTotalUpdatedItems As Integer
Dim iTotalScopeUpItems As Integer
Sub SetupEnvironment()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If MsgBox("Have you downloaded and stored the data file?", vbYesNo, "Check before you go!") = vbNo Then GoTo EndNow
    Call OpeneHubFile
    MsgBox "That is it. In summary:" & vbNewLine & vbNewLine & "Total new eHub activities = " & iTotalNewItems & _
        vbNewLine & "Total eHub activities updated = " & iTotalUpdatedItems & vbNewLine & _
        "Total scope up items = " & iTotalScopeUpItems & vbNewLine & vbNewLine & _
        "You can close this file.", vbOK, "Summary Report!"
EndNow:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
Sub OpeneHubFile()
Dim varEhubFile As Variant
Dim varP6File As Variant
Dim wbMacro As Workbook
Dim wbP6 As Workbook
Dim wbeHubData As Workbook
Dim sStartTime As Single
Dim sEndtime As Single
Dim iQuestion As Integer
Dim seHubSimpleFolderLocation As String
Dim fP6ImportFileName
Dim feHubImportFileName

    Set wbMacro = ActiveWorkbook    'this workbook containing the script

    MsgBox "Open the P6 Exported file.", vbOK, "P6 Data File!"
    
    varP6File = Application.GetOpenFilename()
    If varP6File <> False Then
        Workbooks.Open (varP6File)
        Set wbP6 = ActiveWorkbook
    Else
        Exit Sub
    End If

    MsgBox "Open the eHub Exported file.", vbOK, "eHub Data File!"

    varEhubFile = Application.GetOpenFilename()
    If varEhubFile <> False Then
        Workbooks.Open (varEhubFile)
        Set wbeHubData = ActiveWorkbook
    Else
        Exit Sub
    End If
    
    'Starting the timer to calculate the time to execute the code to create the file, not including file save activities.
    sStartTime = Timer()
    
    Call PopulateP6(wbP6, wbMacro, wbeHubData)    'run the main procedure
    
    iQuestion = MsgBox("The procedure finished in " & Format(Timer() - sStartTime, "000.00") & " seconds." & _
        vbNewLine & vbNewLine & "Do you want to save the resulting imported files?", vbYesNo)
    If iQuestion = vbNo Then GoTo GoHere
    
    seHubSimpleFolderLocation = wbMacro.Sheets("Rules").Range("O2").Value
    
    'Save files
    fP6ImportFileName = Application.GetSaveAsFilename(, filefilter:="Excel Files, *.xlsx")
    If fP6ImportFileName = False Then
        MsgBox "P6 Import file not saved!"
        GoTo ContinueHere
    End If
    With wbP6
        .SaveAs Filename:=fP6ImportFileName, FileFormat:=xlWorkbookDefault
        .Close
    End With
    
ContinueHere:

    MsgBox "Saving the eHub Import file..."
    feHubImportFileName = Application.GetSaveAsFilename(, filefilter:="Excel Files, *.csv")
    If feHubImportFileName = False Then
        MsgBox "eHub Master Import file not saved!"
        GoTo GoHere
    End If
    With wbeHubData
        .SaveAs Filename:=feHubImportFileName, FileFormat:=xlCSV
        Call CreateeHubImportFile(wbMacro, wbeHubData)
        .SaveAs Filename:=seHubSimpleFolderLocation & "eHub-Import-File-" & Format(Now(), "ddMMMyyyy"), FileFormat:=xlCSV
        .Close
    End With

GoHere:

End Sub
'Sub PopulateP6(wbP6 As Workbook, wbMacro As Workbook, wbeHubData As Workbook)
Sub PopulateP6()
Dim wseHub As Worksheet 'eHub export worksheet
Dim wsP6 As Worksheet   'P6 TASK worksheet
Dim wsRules As Worksheet    'Macro Rules worksheet
Dim varP6Data() As Variant 'P6 data array
Dim vareHubData() As Variant 'eHub data array
Dim varP6eHubMapping() As Variant 'eHub to P6 Mapping table array
Dim sIssueType As String
Dim sFSA As String
Dim sCoid As String
Dim sP6WbsCode As String
Dim P6WbsArray() As String
Dim eHubActArray() As String
Dim vareHubActivityElement As Variant
Dim sP6ActivityId As String
Dim P6ActivityArray() As String
Dim sRulesStatus As String
Dim rLtsHolidays As Range
Dim fP6ImportFileName
Dim feHubImportFileName
Dim StartColumn As Integer
Dim EndColumn As Integer
Dim StartRow As Integer
Dim EndRow As Integer
Dim leHubDataCounter As Long
Dim lP6DataCounter As Long
Dim lRulesPosition As Long
Dim lLastRow As Long
Dim lScopeUpCounter As Long
Dim sP6Coid As String
Dim iActualStartOffset As Integer
Dim bFirstActualStart As Boolean
Dim leHubCounter As Long
Dim iCounter As Integer
Dim seHubCoid As String
Dim seHubFSA As String
Dim iQuestion As Integer

'**********FOR TESTING ONLY***********************************************************
Dim wbP6 As Workbook, wbMacro As Workbook, wbeHubData As Workbook
    Set wbP6 = Application.Workbooks("P6-Import-File-04Nov2019(TMPL)-3-sh.xlsx")
    Set wbMacro = Application.Workbooks("Ehub to P6 - Macro Enabled - Rev 45.xlsm")
    Set wbeHubData = Application.Workbooks("eHub-Export-04Nov2019.csv")
'*************************************************************************************

    iTotalNewItems = 0
    iTotalUpdatedItems = 0
    iTotalScopeUpItems = 0
    lP6DataCounter = 0
    Set wseHub = wbeHubData.Sheets(1)
    Set wsP6 = wbP6.Sheets("TASK")
    Set wsRules = wbMacro.Sheets("Rules")
    Set rLtsHolidays = wsRules.Range("w2").CurrentRegion
    Set rLtsHolidays = rLtsHolidays.Resize(rLtsHolidays.Rows.Count - 1).Offset(1)

    varP6Data = wsP6.Range("A1").CurrentRegion    'P6 data table array
    vareHubData = wseHub.Range("A1").CurrentRegion    'eHub data table array
    varP6eHubMapping = wsRules.Range("A1").CurrentRegion    'P6 to eHub Mapping table array
   
    'check to ensure there are correct number of columns in the eHub data file
    If Not UBound(vareHubData, 2) = GetMaxNumber(varP6eHubMapping, 4) Then
        MsgBox "eHub Export file does not have the expected number of columns!", vbOK, "eHub Data File Not Correct!"
        Exit Sub
    End If
    'check to ensure the columns in the eHub data file matches the elements in the Rules P6 to eHub Mapping table
    For iCounter = LBound(varP6eHubMapping, 1) + 1 To UBound(varP6eHubMapping, 1)
        If Not varP6eHubMapping(iCounter, 2) = vareHubData(1, varP6eHubMapping(iCounter, 4)) Then
            MsgBox "Found an issue with eHub Data. Column " & varP6eHubMapping(iCounter, 3) & " of eHub Data is missing or moved to a different column."
            Exit Sub
        End If
    Next iCounter
    
    For lP6DataCounter = 3 To UBound(varP6Data, 1)  'Reading P6 Data
        If varP6Data(lP6DataCounter, 5) = Empty Then 'No eHub Issue key
            sP6WbsCode = varP6Data(lP6DataCounter, 3)   'Decoding the P6 WBS Code
            P6WbsArray() = Split(sP6WbsCode, ".")
            If UBound(P6WbsArray) < 4 Then GoTo NextP6Data  'not an activity, it must be a summary or a milestone, check the next P6 activity
            sFSA = P6WbsArray(3)    'P6 FSA Number
            sCoid = P6WbsArray(1)   'P6 COID
            sP6WbsCode = P6WbsArray(4)   'P6 WBS Activity
            lRulesPosition = posInArray(sP6WbsCode, varP6eHubMapping, 2, 5)    'lookup P6 WBS code in the Rules table
            If lRulesPosition = 0 Then GoTo NextP6Data  'P6 Activity not found in the Mapping table go to next P6 activity
            eHubActArray = Split(varP6eHubMapping(lRulesPosition, 1), ",")  'splitting the Issue Type from the rules table to a separate element
            For Each vareHubActivityElement In eHubActArray 'process each Issue Type element in the rules table
                leHubDataCounter = posInArray(vareHubActivityElement, vareHubData, 2, 1)    'find the location of the Issue Type element in eHub data
                If leHubDataCounter = 0 Then GoTo NextElement 'P6 WBS Code (Issue Type Element) not found in the eHub data, go to the next P6 activity
                For leHubCounter = leHubDataCounter To UBound(vareHubData, 1) 'loop through eHub data for the matching Elements
                    seHubCoid = vareHubData(leHubCounter, 8)
                    'get FSA if available otherwise get Feeder number from eHub data
                    If IsEmpty(vareHubData(leHubCounter, 9)) Then
                        seHubFSA = vareHubData(leHubCounter, 10)
                    Else
                        seHubFSA = vareHubData(leHubCounter, 9)
                    End If
                    If varP6eHubMapping(lRulesPosition, 11) = "Yes" And varP6eHubMapping(lRulesPosition, 1) = vareHubActivityElement Then
                        seHubFSA = sFSA
                    End If
                    If seHubCoid & seHubFSA & LCase(vareHubData(leHubCounter, 1)) = sCoid & sFSA & LCase(Trim(vareHubActivityElement)) Then  'found a match between P6 and eHub data
                        vareHubData(leHubCounter, 5) = varP6Data(lP6DataCounter, 3) 'record P6 ID in eHub data
                        varP6Data(lP6DataCounter, 5) = vareHubData(leHubCounter, 2) 'record eHub ID in P6 data
                        sP6ActivityId = varP6Data(lP6DataCounter, 1)   'Identifying P6 Activity ID
                        P6ActivityArray() = Split(sP6ActivityId, ".")   'Decoding the P6 Activity ID
                        sP6ActivityId = P6ActivityArray(UBound(P6ActivityArray))    'Identifying P6 Activity Abbreviation
                        lRulesPosition = posInArray(sP6ActivityId, varP6eHubMapping, 2, 7)    'lookup P6 Activity in the Rules table
                        If lRulesPosition = 0 Then
                            lRulesPosition = 1
                            GoTo LoopAgain    'Activity id not in the Rules table loop again
                        End If
                        bFirstActualStart = False
                        Do While varP6eHubMapping(lRulesPosition, 7) = sP6ActivityId And lRulesPosition <= UBound(varP6eHubMapping, 1)
                            sRulesStatus = Trim(varP6eHubMapping(lRulesPosition, 10))   'identify the mapping status
                            Select Case sRulesStatus
                                Case Is = "Actual Start"
                                    If bFirstActualStart = False Then 'this is a workaround for when eHub sequence is not followed
                                        iActualStartOffset = varP6eHubMapping(lRulesPosition, 4)
                                        bFirstActualStart = True
                                    End If
                                    If bFirstActualStart And Not vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then 'this is P6 Actual start, first Actual Start
                                        varP6Data(lP6DataCounter, 7) = vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) 'set P6 data Actual Start date equal to eHub date
                                        varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9)  'recording % complete
                                    ElseIf Not vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then
                                        varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9) 'recording % complete
                                        'workaround to fix the issue of eHub having a mid activity started
                                        If vareHubData(leHubCounter, iActualStartOffset) = Empty Then
                                            varP6Data(lP6DataCounter, 7) = WorksheetFunction.WorkDay(Now(), RoundUp(varP6Data(lP6DataCounter, 6) / 100 * -1 * varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                                        End If
                                    End If
                                Case Is = "Actual Finish"
                                    If varP6eHubMapping(lRulesPosition, 9) = 100 And Not vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then '100% complete
                                        varP6Data(lP6DataCounter, 6) = 100    'set P6 Activity % complete to 100%
                                        varP6Data(lP6DataCounter, 8) = vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) 'set P6 data Actual Finish date equal to eHub date
                                    ElseIf Not vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then
                                        varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9) 'setting % complete tovalue from the script table
                                    End If
                                Case Is = "Forecast Start"
                                    vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = WorksheetFunction.WorkDay(varP6Data(lP6DataCounter, 10), _
                                        Fix((varP6eHubMapping(lRulesPosition, 9) / 100 * varP6Data(lP6DataCounter, 4)) - varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                                Case Is = "Forecast Finish"
                                    If varP6eHubMapping(lRulesPosition, 9) = 100 Then '100% complete
                                        vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = varP6Data(lP6DataCounter, 10)
                                    Else
                                        vareHubData(leHubCounter, varP6eHubMapping(lRulesPosition, 4)) = WorksheetFunction.WorkDay(varP6Data(lP6DataCounter, 10), _
                                        Fix((varP6eHubMapping(lRulesPosition, 9) / 100 * varP6Data(lP6DataCounter, 4)) - varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                                    End If
                            End Select
                            lRulesPosition = lRulesPosition + 1
                            If lRulesPosition > UBound(varP6eHubMapping, 1) Then Exit Do
                        Loop
                        iTotalNewItems = iTotalNewItems + 1
                    Else
                        leHubCounter = posInArray(vareHubActivityElement, vareHubData, leHubCounter + 1, 1) 'find the location of the next Element in eHub data
                        If leHubCounter = 0 Then
                            Exit For 'P6 WBS Code (Issue Type Element) not found in the eHub data, go to the next P6 activity
                        Else
                            leHubCounter = leHubCounter - 1
                        End If
                    End If
                    If lRulesPosition > UBound(varP6eHubMapping, 1) Then Exit For
LoopAgain:
                Next leHubCounter
NextElement:
            Next vareHubActivityElement
        Else    'found eHub Issue Key
            sP6WbsCode = varP6Data(lP6DataCounter, 3)   'Decoding the P6 WBS Code
            P6WbsArray() = Split(sP6WbsCode, ".")
            If UBound(P6WbsArray) < 4 Then GoTo NextP6Data  'not an activity, it must be a summary or a milestone, check the next P6 activity
            sP6ActivityId = varP6Data(lP6DataCounter, 1)   'Decoding the P6 Activity ID
            P6ActivityArray() = Split(sP6ActivityId, ".")
            sP6ActivityId = P6ActivityArray(UBound(P6ActivityArray))
            lRulesPosition = posInArray(sP6ActivityId, varP6eHubMapping, 2, 7)    'lookup P6 Activity in the Rules table
            If lRulesPosition = 0 Then GoTo NextP6Data    'Activity id not in the mapping table loop again
            leHubDataCounter = posInArray(varP6Data(lP6DataCounter, 5), vareHubData, 2, 2)
            If leHubDataCounter = 0 Then    'this is a workaround for the cases where the eHub Issue Kay has been deleted after it was added to P6
                iQuestion = MsgBox("Found the following eHub Issue id that has since been deleted from the eHub Data." _
                    & vbNewLine & vbNewLine & varP6Data(lP6DataCounter, 5) & vbNewLine & vbNewLine & _
                    "Do you want to continue?", vbYesNo)
                Select Case iQuestion
                    Case Is = 6
                        MsgBox "eHub Issue id '" & varP6Data(lP6DataCounter, 5) & "' is now deleted from the P6 '" & varP6Data(lP6DataCounter, 1) & "' Activity ID."
                        varP6Data(lP6DataCounter, 5) = ""
                        GoTo NextP6Data
                    Case Is = 7
                        MsgBox "P6 data Activity ID '" & varP6Data(lP6DataCounter, 1) & "' will need to be corrected. Otherwise the same error message will be encountered next time the script is executed."
                        Exit Sub
                End Select
            End If
            If vareHubData(leHubDataCounter, 5) = Empty Then vareHubData(leHubDataCounter, 5) = sP6WbsCode  'added this code to cover the scenrio where eHub doesn't yet have P6 ID reference number.
            bFirstActualStart = False
            Do While varP6eHubMapping(lRulesPosition, 7) = sP6ActivityId
                sRulesStatus = Trim(varP6eHubMapping(lRulesPosition, 10))
                Select Case sRulesStatus
                    Case Is = "Actual Start"
                        If bFirstActualStart = False Then 'this is a workaround for when eHub sequence is not followed
                            iActualStartOffset = varP6eHubMapping(lRulesPosition, 4)
                            bFirstActualStart = True
                        End If
                        If bFirstActualStart And Not vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then 'this is P6 Actual start, first Actual Start
                            varP6Data(lP6DataCounter, 7) = vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) 'set P6 data Actual Start date equal to eHub date
                            varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9)  'recording % complete
                        ElseIf Not vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then
                            varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9) 'recording % complete
                            'workaround to fix the issue of eHub having a mid activity started
                            If vareHubData(leHubDataCounter, iActualStartOffset) = Empty Then
                                varP6Data(lP6DataCounter, 7) = WorksheetFunction.WorkDay(Now(), RoundUp(varP6Data(lP6DataCounter, 7) / 100 * -1 * varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                            End If
                        End If
                    Case Is = "Actual Finish"
                        If varP6eHubMapping(lRulesPosition, 9) = 100 And Not vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then '100% complete
                            varP6Data(lP6DataCounter, 6) = 100    'set P6 Activity % complete to 100%
                            varP6Data(lP6DataCounter, 8) = vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) 'set P6 data Actual Finish date equal to eHub date
                        ElseIf Not vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = Empty Then
                            varP6Data(lP6DataCounter, 6) = varP6eHubMapping(lRulesPosition, 9) 'setting % complete tovalue from the script table
                        End If
                    Case Is = "Forecast Start"
                        vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = WorksheetFunction.WorkDay(varP6Data(lP6DataCounter, 10), _
                        Fix((varP6eHubMapping(lRulesPosition, 9) / 100 * varP6Data(lP6DataCounter, 4)) - varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                    Case Is = "Forecast Finish"
                        If varP6eHubMapping(lRulesPosition, 9) = 100 Then '100% complete
                            vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = varP6Data(lP6DataCounter, 10)
                        Else
                            vareHubData(leHubDataCounter, varP6eHubMapping(lRulesPosition, 4)) = WorksheetFunction.WorkDay(varP6Data(lP6DataCounter, 10), _
                            Fix((varP6eHubMapping(lRulesPosition, 9) / 100 * varP6Data(lP6DataCounter, 4)) - varP6Data(lP6DataCounter, 4)), rLtsHolidays)
                        End If
                End Select
                lRulesPosition = lRulesPosition + 1
                If lRulesPosition > UBound(varP6eHubMapping, 1) Then Exit Do
            Loop
            iTotalUpdatedItems = iTotalUpdatedItems + 1
        End If
NextP6Data:
    Next lP6DataCounter
    
'Scope up
    sCoid = ""
    For lP6DataCounter = 3 To UBound(varP6Data, 1)  'lLastRow
        If varP6Data(lP6DataCounter, 5) = Empty Then GoTo NextScopeCounter
        P6WbsArray() = Split(varP6Data(lP6DataCounter, 3), ".")
        sP6Coid = P6WbsArray(1)   'P6 COID
        If Not sP6Coid = sCoid Then
            sCoid = P6WbsArray(1)   'P6 COID
            For leHubDataCounter = 2 To UBound(vareHubData, 1)
                If vareHubData(leHubDataCounter, 8) = sCoid And vareHubData(leHubDataCounter, 5) = Empty Then
                    varP6Data = ReDimPreserve(varP6Data, UBound(varP6Data, 1) + 1, UBound(varP6Data, 2))
                    For lScopeUpCounter = 1 To UBound(varP6Data, 2)
                        Select Case lScopeUpCounter
                            Case Is = 1
                                varP6Data(UBound(varP6Data, 1), lScopeUpCounter) = vareHubData(leHubDataCounter, 1) & "-" & vareHubData(leHubDataCounter, 6)
                            Case Is = 2
                                varP6Data(UBound(varP6Data, 1), lScopeUpCounter) = "Not Started"
                            Case Is = 3
                                varP6Data(UBound(varP6Data, 1), lScopeUpCounter) = "IMS." & sP6Coid & ".Scope.Up"
                            Case Is = 5
                                varP6Data(UBound(varP6Data, 1), lScopeUpCounter) = vareHubData(leHubDataCounter, 2)
                        End Select
                    Next lScopeUpCounter
                    iTotalScopeUpItems = iTotalScopeUpItems + 1
                End If
            Next leHubDataCounter
        End If
NextScopeCounter:
    Next lP6DataCounter
    
'Scope Down
    'the place holder for the scope down use case
UpdateNow:
'Update the data files
    StartColumn = varP6eHubMapping(posInArray("Created", varP6eHubMapping, 2, 2), 4)
    EndRow = UBound(vareHubData, 1)
    EndColumn = UBound(vareHubData, 2)
    StartRow = 2
    With wseHub
        .Range(.Cells(StartRow, StartColumn), .Cells(EndRow, EndColumn)).NumberFormat = "dd-mmm-yyyy"
        .Range("A1").CurrentRegion = vareHubData
    End With
    StartColumn = 7
    EndRow = UBound(varP6Data, 1)
    EndColumn = UBound(varP6Data, 2)
    StartRow = 3
    StartColumn = 7
    StartRow = 3
    With wsP6
        .Activate
        .Range("A1").CurrentRegion.Select
        Selection.Resize(UBound(varP6Data, 1), UBound(varP6Data, 2)).Select
        EndRow = UBound(varP6Data, 1)
        EndColumn = UBound(varP6Data, 2) - 1
        .Range(.Cells(StartRow, StartColumn), .Cells(EndRow, EndColumn)).NumberFormat = "dd-mmm-yyyy"
        Selection = varP6Data
        .Range("A1").Select
    End With
End Sub

Sub CreateeHubImportFile(wbMacro As Workbook, wbeHubData As Workbook)
Dim wseHub As Worksheet 'eHub export worksheet
Dim wsRules As Worksheet    'Macro Rules worksheet
Dim varP6eHubMapping() As Variant 'eHub to P6 Mapping table array
Dim lMappingCounter As Long
Dim leHubCounter As Long

    Set wseHub = wbeHubData.Sheets(1)
    Set wsRules = wbMacro.Sheets("Rules")
    varP6eHubMapping = wsRules.Range("A1").CurrentRegion    'P6 to eHub Mapping table array
    leHubCounter = varP6eHubMapping(2, 4) - 1
    For lMappingCounter = 2 To UBound(varP6eHubMapping, 1)
        If varP6eHubMapping(lMappingCounter, 10) = "Forecast Start" Or varP6eHubMapping(lMappingCounter, 10) = "Forecast Finish" Or varP6eHubMapping(lMappingCounter, 10) = "Header" Then
            wseHub.Cells(1, varP6eHubMapping(lMappingCounter, 4)).Interior.ColorIndex = 6 'EntireColumn.Delete
        End If
    Next lMappingCounter
    Do While wseHub.Range("a1").Offset(0, leHubCounter).Value <> ""
        With wseHub.Range("a1").Offset(0, leHubCounter)
            If Not .Interior.ColorIndex = 6 Then
                .EntireColumn.Delete
                leHubCounter = leHubCounter - 1
            End If
        leHubCounter = leHubCounter + 1
        End With
    Loop
    wseHub.Rows(1).Interior.ColorIndex = 0
End Sub
Public Function posInArray(ByVal itemSearched As Variant, ByVal aArray As Variant, ByVal RowToStart As Long, ByVal ColumnToCheck As Long) As Long

Dim pos As Long, item As Variant, i As Long

posInArray = 0
If IsArray(aArray) Then
    If Not IsEmpty(aArray) Then
        pos = 1
        If itemSearched <> "" Then
            For i = RowToStart To UBound(aArray, 1)
                If LCase(Trim(itemSearched)) = LCase(aArray(i, ColumnToCheck)) Then
                    posInArray = i
                    Exit Function
                End If
            Next i
        Else
            For i = RowToStart To UBound(aArray, 1)
                If Not aArray(i, ColumnToCheck) = Empty Then
                    posInArray = i
                    Exit Function
                End If
            Next i
        End If
        posInArray = 0
    End If
End If

End Function
Public Function RoundUp(ByVal Value As Double)
    If Int(Value) = Value Then
        RoundUp = Value
    ElseIf Value - Int(Value) > 0.45 Then
        RoundUp = Int(Value) + 1
    Else
        RoundUp = Int(Value)
    End If
End Function
Public Function RoundNumber(ByVal Value As Double, ByVal UpDown As Boolean) 'true is up false is down
    If UpDown = True Then 'round-up
        If Int(Value) = Value Then
            RoundNumber = Value
        Else
            RoundNumber = Int(Value) + 1
        End If
    Else    'round-down
        RoundNumber = Int(Value)
    End If
End Function
Public Function GetMaxNumber(ByRef arrValues() As Variant, iColNumber As Integer) As Long

Dim i As Long

For i = LBound(arrValues) To UBound(arrValues)
    If IsNumeric(arrValues(i, iColNumber)) Then
        If arrValues(i, iColNumber) > GetMaxNumber Then GetMaxNumber = arrValues(i, iColNumber)
    End If
Next i

End Function
Public Function ReDimPreserve(aArrayToPreserve As Variant, nNewFirstUBound As Variant, nNewLastUBound As Variant) As Variant
    Dim nFirst As Long
    Dim nLast As Long
    Dim nOldFirstUBound As Long
    Dim nOldLastUBound As Long

    ReDimPreserve = False
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(1 To nNewFirstUBound, 1 To nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function

Sub ClearCreateP6()
Dim wb As Workbook
Dim wsRules As Worksheet
Dim lRulesLastRow As Long

Set wb = ActiveWorkbook
Set wsRules = wb.Sheets("Rules")

lRulesLastRow = wsRules.Cells(Rows.Count, "L").End(xlUp).Row
wsRules.Range(wsRules.Range("L2"), wsRules.Range("L2").Offset(lRulesLastRow, 3)).Clear

End Sub

