Attribute VB_Name = "UpdateSupplies"
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object

Sub AtualizarMapa(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

    ErrorSection = "Initialization"

    Dim temp As Double
    temp = Timer

    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim Gerador As String
    Dim IsMapa As Boolean
    Dim MapaFound As Boolean
    Dim response As VbMsgBoxResult
    Dim CurrentCol As Long
    
    ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit                 ' Exit the function or sub
        End If
    Loop
    
    ErrorSection = "WorkbookSearch"

    MapaFound = False
    
    ReDim PEPList(0)
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
        ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsMapa = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                ErrorSection = "WorksheetSearchFor-" & wb.Name
                If InStr(1, LCase(ws.Name), "mapa de suprimentos", vbTextCompare) > 0 Then
                    IsMapa = True
                    
                    CurrentCol = 9
                    
                    Do While Not IsEmpty(ws.Cells(1, CurrentCol))
                        ErrorSection = "While-" & CurrentCol
                        Gerador = ws.Cells(1, CurrentCol)
                        PEP = ws.Cells(1, CurrentCol + 1)
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1) ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                        
                        Application.StatusBar = "Trabalhando em " & PEP
        
                        If Not UpdateMapa(ws, Gerador, CurrentCol) Or Not UpdateCJI3(wb, PEP, CurrentCol) Then
                            MsgBox "Não foi possível atualizar Mapa de Suprimentos de " & vbCrLf & PEP, vbInformation
                        End If
                        
                        CurrentCol = CurrentCol + 4
                    Loop
                    Exit For
                End If
            Next ws
        End If
    
        If Not IsMapa Then
            GoTo NextWorkbook
        Else
            MapaFound = True
        End If
        
        ' UpdateCover wb, wsCJI3
        
        Application.StatusBar = False

        'ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If Not MapaFound Then
        MsgBox "Nenhum Mapa de Suprimentos foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Mapa de Suprimentos atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
    Debug.Print "Atualizar Mapa - Total execution time: "; Timer - temp
    
CleanExit:

    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarMapa"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Function UpdateMapa(wsMapa As Worksheet, Gerador As String, CurrentCol As Long) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateMapa Start"

    Dim exportWb As Workbook
    'Dim wbIter As Workbook ' Iterator for workbooks
    Dim Workbook As Workbook
    'Dim wsMapa As Worksheet
    Dim exportWs As Worksheet
    Dim Row As Long
    Dim exportWbName As String
    Dim exportWbPath As String
    'Dim StartDate As String
    'Dim EndDate As String
    'Dim ordem As String
    'Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    Dim wsMapaLR As Long
    Dim exportWsLR As Long
    'Dim currentRows As Long
    'Dim requiredRows As Long
    'Dim foundCell As Range
    'Dim Gerador As String
    
    ' Find wsMapa last row and save to wsMapaLR
    wsMapaLR = wsMapa.Cells(wsMapa.Rows.Count, "A").End(xlUp).Row
    
    If wsMapaLR < 4 Then
        wsMapaLR = 4
    End If
    
    ' Name of the workbook to find
    exportWbName = "CS11-" & Gerador

    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count
    
Debug.Print "Setup time: " & Timer - temp
temp = Timer

ErrorSection = "SAPNavigation"

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncs11"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = Gerador
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = "1341"
    session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = "1"
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "BEST"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler

    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text

    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportWbPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[0]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop

    ' If exported file not found, open it
    If Not found Then
        Set exportWb = Workbooks.Open(exportWbPath & "\" & exportWbName)
    End If

ErrorSection = "ExportFormating"

    Set exportWs = exportWb.Sheets(1)
    
    ' Find exportWs last row and save to exportWsLR (using column C as reference)
    exportWsLR = exportWs.Cells(exportWs.Rows.Count, "C").End(xlUp).Row
    
    ' Delete dummy itens
    For Row = 2 To exportWsLR
        If exportWs.Cells(Row, 8).Value <> "" Then
            exportWs.Cells(Row, 8).EntireRow.Delete
            Row = Row - 1
        End If
    Next Row
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer

ErrorSection = "PasteData"

    ' Clear, copy and paste values and formats without breaking formulas and headers
    If wsMapa.AutoFilterMode Then wsMapa.AutoFilter.ShowAllData ' Clear any applied filters
    
    ' Compare groups line by line.
    ' Assumption: wsMapa data starts at row 5 and exportWs data starts at row 2.
    Dim wsMapaCurrentRow As Long, exportWsCurrentRow As Long
    Dim groupExportStart As Long, groupExportEnd As Long, groupExportCount As Long
    Dim groupMapaStart As Long, groupMapaEnd As Long, groupMapaCount As Long
    Dim lastGroupMapaStart As Long, lastGroupMapaEnd As Long
    Dim wsMapaGroupFound As Boolean
    
    ' Set current row as the first row
    wsMapaCurrentRow = 4
    exportWsCurrentRow = 2
    
    ' Set last group as the header row
    lastGroupMapaStart = wsMapaCurrentRow - 1
    lastGroupMapaEnd = wsMapaCurrentRow - 1
    
    ' Strikethrough cells from wsMapa
    wsMapa.Range("A" & wsMapaCurrentRow & ":C" & wsMapaLR).Font.Strikethrough = True
    
    Do While exportWsCurrentRow <= exportWsLR
ErrorSection = "PasteDataWhile-" & exportWsCurrentRow
        ' Identify group start in exportWs: a row with 0 in column C
        If Trim(exportWs.Cells(exportWsCurrentRow, "E").Value) = "0" Or exportWs.Cells(exportWsCurrentRow, "E").Value = 0 Then
ErrorSection = "ExportLimits-" & exportWsCurrentRow
            groupExportStart = exportWsCurrentRow
            groupExportEnd = groupExportStart
            ' Determine the end of this exportWs group:
            Do While groupExportEnd + 1 <= exportWsLR And Not (Trim(exportWs.Cells(groupExportEnd + 1, "E").Value) = "0" Or exportWs.Cells(groupExportEnd + 1, "E").Value = 0)
                groupExportEnd = groupExportEnd + 1
            Loop
            groupExportCount = groupExportEnd - groupExportStart + 1
            
ErrorSection = "MatchGroup-" & exportWsCurrentRow
            ' Look for a matching group in wsMapa
            wsMapaGroupFound = False
            For wsMapaCurrentRow = lastGroupMapaEnd To wsMapaLR
                If wsMapaCurrentRow <= wsMapaLR And (Trim(wsMapa.Cells(wsMapaCurrentRow, CurrentCol).Value) = "0" Or wsMapa.Cells(wsMapaCurrentRow, CurrentCol).Value = 0) And wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = Trim(Replace(exportWs.Cells(exportWsCurrentRow, "D").Value, Gerador, "")) Then
                    groupMapaStart = wsMapaCurrentRow
                    groupMapaEnd = groupMapaStart
                    ' Determine the end of this exportWs group:
                    Do While groupMapaEnd + 1 <= wsMapaLR And Not (Trim(wsMapa.Cells(groupMapaEnd + 1, CurrentCol).Value) = "0" Or wsMapa.Cells(groupMapaEnd + 1, CurrentCol).Value = 0)
                        groupMapaEnd = groupMapaEnd + 1
                    Loop
                    groupMapaCount = groupMapaEnd - groupMapaStart + 1
                    wsMapaGroupFound = True
                    Exit For
                End If
            Next wsMapaCurrentRow
            
            If Not wsMapaGroupFound Then
ErrorSection = "CreateGroup-" & exportWsCurrentRow
                ' Define the start
                wsMapaCurrentRow = lastGroupMapaEnd
                exportWsCurrentRow = groupExportStart
                groupMapaStart = lastGroupMapaEnd + 1
            
                Do While exportWsCurrentRow <= groupExportEnd
ErrorSection = "CreateGroupWhile-" & exportWsCurrentRow
                    ' They are different: insert a row below copying the existing row and fill the new row green
                    wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                    ' Option 1: Copy the original row as base
                    wsMapa.Rows(wsMapaCurrentRow).Copy
                    wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                    Application.CutCopyMode = False
                    ' Replace compared columns with exportWs values
                    wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                    wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = Trim(Replace(exportWs.Cells(exportWsCurrentRow, "D").Value, Gerador, ""))
                    wsMapa.Cells(wsMapaCurrentRow + 1, CurrentCol).Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                    ' Remove strikethrough cells from wsMapa
                    wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow, "A"), wsMapa.Cells(wsMapaCurrentRow, "C")).Font.Strikethrough = False
                    groupMapaEnd = wsMapaCurrentRow + 1 ' Adjust the last group row marker after inserting a row
                    wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
            ElseIf groupExportCount = 1 And groupExportCount < groupMapaCount Then
ErrorSection = "IfExportGroupSmaller-" & exportWsCurrentRow
                ' Reorder the groups to keep the same order as exportWs
                If groupMapaStart <> lastGroupMapaEnd + 1 Then
                    wsMapa.Rows(groupMapaStart & ":" & groupMapaEnd).Cut
                    wsMapa.Rows(lastGroupMapaEnd + 1).Insert Shift:=xlDown
                    Application.CutCopyMode = False
                    ' Update groupMapaStart and groupMapaEnd based on the new location.
                    groupMapaStart = lastGroupMapaEnd + 1
                    groupMapaEnd = groupMapaStart + groupMapaCount - 1
                End If
            
                ' Define the start
                wsMapaCurrentRow = groupMapaStart
                exportWsCurrentRow = groupExportStart
                ' This means a lone header wasn't found on wsMapa, so a header must be found without messing with the groupMapa found
                ' They are different: insert a row below copying the existing row and fill the new row green
                wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                ' Option 1: Copy the original row as base
                wsMapa.Rows(wsMapaCurrentRow).Copy
                wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                Application.CutCopyMode = False
                ' Replace compared columns with exportWs values
                wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = Trim(Replace(exportWs.Cells(exportWsCurrentRow, "D").Value, Gerador, ""))
                wsMapa.Cells(wsMapaCurrentRow + 1, CurrentCol).Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                ' Fill the new row (green) for columns A to C
                wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow, "A"), wsMapa.Cells(wsMapaCurrentRow, "C")).Font.Strikethrough = False
                groupMapaEnd = wsMapaCurrentRow ' Adjust the last group row marker after inserting a row
                wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                wsMapaCurrentRow = wsMapaCurrentRow + 1
                exportWsCurrentRow = exportWsCurrentRow + 1
            Else
ErrorSection = "IfExportGroupBigger-" & exportWsCurrentRow
                ' Reorder the groups to keep the same order as exportWs
                If groupMapaStart <> lastGroupMapaEnd + 1 Then
                    wsMapa.Rows(groupMapaStart & ":" & groupMapaEnd).Cut
                    wsMapa.Rows(lastGroupMapaEnd + 1).Insert Shift:=xlDown
                    Application.CutCopyMode = False
                    ' Update groupMapaStart and groupMapaEnd based on the new location.
                    groupMapaStart = lastGroupMapaEnd + 1
                    groupMapaEnd = groupMapaStart + groupMapaCount - 1
                End If
            
                ' Define the start
                wsMapaCurrentRow = groupMapaStart
                exportWsCurrentRow = groupExportStart
                
                ' Compare line by line the values from wsMapa column A and C and exportWs column C and E.
                ' Assumption: wsMapa data starts at row 5 and exportWs data starts at row 2.
                Do While wsMapaCurrentRow <= groupMapaEnd
ErrorSection = "IfExportGroupBiggerWhile-" & exportWsCurrentRow
                    If exportWsCurrentRow <= groupExportEnd Then
                        ' Compare wsMapa col A with exportWs col C and wsMapa col C with exportWs col E
                        If Trim(wsMapa.Cells(wsMapaCurrentRow, "A").Value) <> Trim(exportWs.Cells(exportWsCurrentRow, "C").Value) Or _
                           Trim(wsMapa.Cells(wsMapaCurrentRow, CurrentCol).Value) <> Trim(exportWs.Cells(exportWsCurrentRow, "E").Value) Then
                            ' They are different: insert a row below copying the existing row and fill the new row green
                            wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                            ' Option 1: Copy the original row as base
                            wsMapa.Rows(wsMapaCurrentRow).Copy
                            wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                            Application.CutCopyMode = False
                            ' Replace compared columns with exportWs values
                            wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                            wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = Trim(Replace(exportWs.Cells(exportWsCurrentRow, "D").Value, Gerador, ""))
                            wsMapa.Cells(wsMapaCurrentRow + 1, CurrentCol).Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                            ' Fill the new row (green) for columns A to C
                            wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow + 1, "A"), wsMapa.Cells(wsMapaCurrentRow + 1, "C")).Font.Strikethrough = False
                            groupMapaEnd = groupMapaEnd + 1 ' Adjust the last group row marker after inserting a row
                            wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                            wsMapaCurrentRow = wsMapaCurrentRow + 1 ' Skip the newly inserted row
                        Else
                            ' Remove strikethrough cells from wsMapa
                            wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow, "A"), wsMapa.Cells(wsMapaCurrentRow, "C")).Font.Strikethrough = False
                        End If
                    End If
                    
                    ' Move pointer to next group element in both sheets
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
                
                ' Continue adding inexistent values if groupMapaCount < groupExportCount
                Do While exportWsCurrentRow <= groupExportEnd
ErrorSection = "IfExportSheetBiggerWhile-" & exportWsCurrentRow
                    wsMapaCurrentRow = groupMapaEnd
                    If exportWsCurrentRow <= groupExportEnd Then
                        ' They are different: insert a row below copying the existing row and fill the new row green
                        wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                        ' Option 1: Copy the original row as base
                        wsMapa.Rows(wsMapaCurrentRow).Copy
                        wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                        Application.CutCopyMode = False
                        ' Replace compared columns with exportWs values
                        wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                        wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = Trim(Replace(exportWs.Cells(exportWsCurrentRow, "D").Value, Gerador, ""))
                        wsMapa.Cells(wsMapaCurrentRow + 1, CurrentCol).Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                        ' Fill the new row (green) for columns A to C
                        wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow + 1, "A"), wsMapa.Cells(wsMapaCurrentRow + 1, "C")).Font.Strikethrough = False
                        groupMapaEnd = groupMapaEnd + 1 ' Adjust the last group row marker after inserting a row
                        wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                        wsMapaCurrentRow = wsMapaCurrentRow + 1 ' Skip the newly inserted row
                    End If
                    
                    ' Move pointer to next group element in both sheets
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
            End If
            
            ' Move exportRow pointer past this group.
            exportWsCurrentRow = groupExportEnd + 1
        Else
            exportWsCurrentRow = exportWsCurrentRow + 1
        End If
        
        lastGroupMapaStart = groupMapaStart
        lastGroupMapaEnd = groupMapaEnd
    Loop

ErrorSection = "Ending"

    UpdateMapa = True

    ' Close the exported workbook without saving changes
    exportWb.Close SaveChanges:=False

    ' Delete the exported workbook file
    On Error Resume Next ' In case the file is not found or cannot be deleted
    Kill exportWbPath & "\" & exportWbName
    On Error GoTo ErrorHandler

    ' Cleanup
    Application.CutCopyMode = False
    
Debug.Print "Project Review Mapa de Suprimentos Sheet update: " & Timer - temp
temp = Timer

CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateMapa"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function UpdateCJI3(wb As Workbook, PEP As String, CurrentCol As Long) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateCJI3 Start"

    Dim exportWb As Workbook
    Dim Workbook As Workbook
    Dim wsCJI3 As Worksheet
    Dim exportWs As Worksheet
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim EndDate As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    
    On Error Resume Next
    Set wsCJI3 = wb.Sheets("CJI3")
    On Error GoTo ErrorHandler
    
    ' Check if the "CJI3" sheet exists
    If wsCJI3 Is Nothing Then
        UpdateCJI3 = False
        Exit Function
    End If
    
    ' Name of the workbook to find
    exportWbName = "CJI3-" & PEP
    
    ' Set end date
    EndDate = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy")
    
    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count

ErrorSection = "SAPNavigation"

Debug.Print "Setup time: " & Timer - temp
temp = Timer

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncji3"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
    session.findById("wnd[1]").sendVKey 0
    On Error GoTo ErrorHandler
    
    ' Clear other fields
    session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PROJN-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-HIGH").Text = ""
    
    ' Search PEP
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = PEP
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").Text = "01.11.2000"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").Text = EndDate
    session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/CUSTO_CIDIO"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").SetFocus
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo EmptyCJI3
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    On Error GoTo ErrorHandler
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

    If False Then
EmptyCJI3:
        On Error GoTo ErrorHandler
ErrorSection = "EmptyCJI3"
        UpdateCJI3 = False
        GoTo CleanExit
    End If
    
ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop
    
    ' Validate if the workbook was opened successfully
    If exportWb Is Nothing Then
        wsCJI3.UsedRange.ClearContents
        UpdateCJI3 = False
        Exit Function
    End If
    
    Set exportWs = exportWb.Sheets(1)
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer
    
ErrorSection = "PasteData"

    ' Clear, copy and paste data from exportWs to wsCJI3
    If wsCJI3.AutoFilterMode Then wsCJI3.AutoFilter.ShowAllData ' Clear any applied filters
    wsCJI3.UsedRange.ClearContents
    exportWs.UsedRange.Copy
    wsCJI3.UsedRange.PasteSpecial
    
    ' Ensure columns A and B are converted to numbers
    With wsCJI3
        .Columns("A:B").NumberFormat = "0"  ' Set format to number
        .Columns("A:B").Value = .Columns("A:B").Value  ' Convert text to numbers
    End With
    
    ' Cleanup
    Application.CutCopyMode = False
    exportWb.Close False  ' Close the exported workbook without saving

Debug.Print "Project Review CJI3 Sheet update: " & Timer - temp
temp = Timer
    
    UpdateCJI3 = True
    
CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateCJI3"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function SetupSAPScripting() As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo ErrorHandler
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo ErrorHandler
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    On Error GoTo ErrorHandler
    Set Connection = SAPApplication.Children(0)
    Set session = Connection.Children(0)
    On Error GoTo ErrorHandler
    
    SetupSAPScripting = True
    
    If False Then
ErrorHandler:
    SetupSAPScripting = False
    End If
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function


