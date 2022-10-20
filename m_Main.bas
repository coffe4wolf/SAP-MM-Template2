Attribute VB_Name = "m_Main"
Option Explicit

Public Wb                           As Workbook
Public wsCreateMM                   As Object
Public wsSettings                   As Worksheet
Public wsCategoriesMaster           As Worksheet
Public wsRussianCations             As Worksheet
Public wsPurchasingGroupsMaster     As Worksheet
Public wsContractorsMaster          As Worksheet
Public wsBulkImport                 As Worksheet
Public wsCableTypeMaster            As Worksheet

Public categoryWsCategoriesMaster   As Range
Public categoryCreateMMRange        As Range
Public lrWsCategoriesMaster         As Long

' Sheet categories master
Public Const columnWithMaterialGroup = "C"
Public Const columnWithClass = "D"
Public Const columnWithAttribute = "E"
Public Const columnWithAttributeValue = "F"
Public Const columnWithShortAttribute = "G"

' Sheet create MM
Public Const shortDescriptionRangeAddress = "E11"
Public Const fullDescriptionRangeAddress = "E14"
Public Const techDescrupionRangeAddress = "E27"
Public Const shortDescriptionEngAddress = "E40"
Public Const attrubiteValuesAddress = "B12:B700"

Public Const sheetProtectionPassword = "1234"

' Sheet settings
Public Const specialClassAttributesAddress = "C2"

Public Const rowToStartDrawAttributes = 12



Sub test()

    Call init

    'Application.EnableEvents = True
    
    Call lockCreateMMRanges
    
    'Debug.Print InStr(0, wsCreateMM.Range("B14").value, wsCableTypeMaster.Range("D7").value)

End Sub

Sub main()

    Call init

    Call initCreateMMSheet
    
    'Call generateShortDescrption
    

End Sub

Sub getCableTypeDescription(attributesStartRange As String, ws As Worksheet)

    
    Dim lrWs                    As Long
    Dim lrWsCableTypes          As Long
    Dim rowsCounter             As Long
    Dim cll                     As Range
    Dim createMMCll             As Range
    Dim attrStartRange          As Range
    Dim attributesRange         As Range
    
    Dim attributeCableType      As String: attributeCableType = "attributeCableType"
    Dim attributeFireSafery     As String: attributeFireSafery = "attributeFireSafety"
    
    Dim cableTypeMasterColumn   As String: cableTypeMasterColumn = "A"
    Dim fireSafetyMasterColumn  As String: fireSafetyMasterColumn = "D"
    
    Dim setCableTypeValue       As String
    
    Dim fireSafetyDescription   As String: fireSafetyDescription = ""
    Dim resultDescription       As String: resultDescription = ""
    
    
    
    lrWs = GetBorders("LR", ws.Name, ThisWorkbook)
    lrWsCableTypes = GetBorders("LR", wsCableTypeMaster.Name)
    
    
    
    ' if it cable.
    If wsCreateMM.category_ComboBox.value = wsSettings.Range("categoryCable").value Then
    
        
        Set attrStartRange = ws.Range(attributesStartRange)
        Set attributesRange = ws.Range(attributesStartRange & ":" & ColumnLetter(attrStartRange) & lrWs)
        
            ' Search setted cable type value.
            For Each createMMCll In attributesRange
            
                ' If we met cable type.
                If createMMCll.value = wsSettings.Range("attributeCableType").value Then
                
                    ' Search that type on cableTypesMaster sheet.
                    rowsCounter = 2
                    Do While Trim(wsCableTypeMaster.Cells(rowsCounter, 1).value) <> ""
                                            
                        If UCase(wsCableTypeMaster.Cells(rowsCounter, 1).value) = UCase(createMMCll.Offset(0, 1).value) Then
                            
                            resultDescription = resultDescription & createMMCll.value & ":" & vbCrLf & wsCableTypeMaster.Cells(rowsCounter, 2).value & vbCrLf
                            Exit Do
                            
                        Else
                        
                            rowsCounter = rowsCounter + 1
                            
                        End If
                        
                    Loop
                    
                ' If we met cable fire safety.
                ElseIf createMMCll.value = wsSettings.Range("attributeFireSafety").value Then
                
                    ' Search that type on cableTypesMaster sheet.
                    rowsCounter = 2
                    Do While Trim(wsCableTypeMaster.Cells(rowsCounter, 4).value) <> ""
                                            
                        If InStr(1, createMMCll.Offset(0, 1).value, wsCableTypeMaster.Cells(rowsCounter, 4).value) > 0 Then
                            
                            fireSafetyDescription = fireSafetyDescription & wsCableTypeMaster.Cells(rowsCounter, 5).value & vbCrLf
                            
                        End If
                        
                        rowsCounter = rowsCounter + 1
                        
                    Loop
                    
                    If fireSafetyDescription <> "" Then fireSafetyDescription = wsSettings.Range("attributeFireSafety").value & ":" & vbCrLf & fireSafetyDescription & vbCrLf
                
                End If
        
        
            Next createMMCll
    
    
    End If
    

    
    wsCreateMM.Range(fullDescriptionRangeAddress).Offset(1, 1).value = wsCreateMM.Range(fullDescriptionRangeAddress).Offset(1, 1).value & vbCrLf & vbCrLf & fireSafetyDescription & resultDescription



End Sub


Sub addMaterialToBulkImport()



    Call ImprovePerformance(True)
    Call init
    
    
    
    Dim lrWsBulkImport As Long
    
    
    
    With wsBulkImport
    
        lrWsBulkImport = GetBorders("LR", wsBulkImport.Name) + 1
        
        .Range("A" & lrWsBulkImport).value = lrWsBulkImport - 10                                                                                                                            ' Position number
        .Range("B" & lrWsBulkImport).value = priorityFromTextToValue(wsCreateMM.priority_ComboBox.value)                                                                                    ' Priority
        .Range("C" & lrWsBulkImport).value = wsCreateMM.Range(shortDescriptionRangeAddress).Offset(1, 1).value                                                                              ' Short description
        .Range("D" & lrWsBulkImport).value = wsCreateMM.Range(fullDescriptionRangeAddress).Offset(1, 1).value & vbCrLf & vbCrLf & wsCreateMM.Range(techDescrupionRangeAddress).Offset(1, 1).value    ' Full description
        .Range("E" & lrWsBulkImport).value = wsCreateMM.Range(shortDescriptionEngAddress).Offset(1, 1).value                                                                                ' Short description eng
        .Range("F" & lrWsBulkImport).value = wsCreateMM.unitOfEntry_TextBox.value                                                                                                           ' Unit of entry
        .Range("G" & lrWsBulkImport).value = wsCreateMM.chooseSupplier_TextBox.value                                                                                                        ' Contractors code
        .Range("H" & lrWsBulkImport).value = wsCreateMM.article_TextBox.value                                                                                                               ' Article
        .Range("I" & lrWsBulkImport).value = wsCreateMM.maxPrice_TextBox.value                                                                                                              ' Max price
        .Range("J" & lrWsBulkImport).value = wsCreateMM.materialGroupCode_ComboBox.value                                                                                                    ' Material group code
        .Range("K" & lrWsBulkImport).value = wsCreateMM.purchasingGroup_ComboBox.value                                                                                                      ' Purchasing group code
        .Range("L" & lrWsBulkImport).value = checkBoxToText(wsCreateMM.criticalPart_CheckBox)                                                                                               ' Critical equipment
        .Range("M" & lrWsBulkImport).value = "" ' Serialisation Profile
        .Range("N" & lrWsBulkImport).value = "" ' Shelf life
        .Range("O" & lrWsBulkImport).value = "" ' Storage Condition
        .Range("P" & lrWsBulkImport).value = "" ' Temperature Condition"
        .Range("Q" & lrWsBulkImport).value = "" ' Container Requirements
        .Range("R" & lrWsBulkImport).value = "" ' Category group
        .Range("S" & lrWsBulkImport).value = "" ' Group
        .Range("T" & lrWsBulkImport).value = "" ' Material type
        .Range("U" & lrWsBulkImport).value = checkBoxToText(wsCreateMM.batchManagement_CheckBox)                'Batch Management
    
    End With
    
    

    Call ImprovePerformance(True)
    
    

End Sub

Sub clearEntryData()

    
    
    Call ImprovePerformance(True)
    
    Call init

    Dim lrWsCreateMM As Long
    
    Call unprotectSheet(wsCreateMM, sheetProtectionPassword)

    With wsCreateMM
    
        lrWsCreateMM = GetBorders("LR", .Name)
    
        .Range("A11:G700").EntireRow.Delete
        
        wsCreateMM.category_ComboBox.value = ""
        wsCreateMM.materialGroup_ComboBox.value = ""
        wsCreateMM.class_ComboBox.value = ""
        wsCreateMM.priority_ComboBox = ""
        wsCreateMM.criticalPart_CheckBox = False
        wsCreateMM.batchManagement_CheckBox = False
        wsCreateMM.purchasingGroup_ComboBox.value = ""
        wsCreateMM.materialGroupCode_ComboBox.value = ""
        wsCreateMM.priority_ComboBox.value = ""
        wsCreateMM.article_TextBox.value = ""
        wsCreateMM.chooseSupplier_TextBox.value = ""
        wsCreateMM.maxPrice_TextBox.value = ""
        wsCreateMM.criticalPart_CheckBox.value = False
        wsCreateMM.batchManagement_CheckBox.value = False
        wsCreateMM.unitOfEntry_TextBox.value = ""
    
    End With

    Call lockCreateMMRanges

    Call ImprovePerformance(True)


End Sub

Function getShortAttributeValue(attribValue As String, attrib As String, materialGroupValue As String, classValue As String) As String

    Call ImprovePerformance(True)

    Call init
    
    Dim lrWsCategoriesMaster     As Long
    Dim rowToStart               As Long
    Dim rowsCounter              As Long
    
    ' Search in categoriesMaster.
    rowToStart = 2
    lrWsCategoriesMaster = GetBorders("LR", wsCategoriesMaster.Name, Wb)
    
    
    
    With wsCategoriesMaster
    
        For rowsCounter = rowToStart To lrWsCategoriesMaster
        
            If .Range(columnWithAttributeValue & rowsCounter).value = attribValue _
                And .Range(columnWithAttribute & rowsCounter).value = attrib _
                And .Range(columnWithMaterialGroup & rowsCounter).value = materialGroupValue _
                And .Range(columnWithClass & rowsCounter).value = classValue Then
                
                getShortAttributeValue = .Range(columnWithShortAttribute & rowsCounter).value
                Exit For
            
            End If
            
        
        Next rowsCounter
    
    End With
    
    
    
    Call ImprovePerformance(False)

End Function

Sub generateDescription(Optional rowToStartRead As Long = 12, Optional rangeWithShortDescriptionTemplate As String = "F11", Optional rangeWithFullDescriptionTemplate As String = "F14", Optional rangeTechDescription As String = "F28")

    Call ImprovePerformance(True)

    Call init
    
    Call unprotectSheet(wsCreateMM, sheetProtectionPassword)
    
    
    Dim columnWithAttributes                As String: columnWithAttributes = "A"
    Dim shortDescription                    As String
    Dim fullDescription                     As String
    Dim techDescription                     As String
    Dim attr                                As String
    Dim val                                 As String
    
    
    
    With wsCreateMM
    
        shortDescription = .Range(rangeWithShortDescriptionTemplate).value
        fullDescription = .Range(rangeWithFullDescriptionTemplate).value
        techDescription = ""
        
        .Range(rangeWithShortDescriptionTemplate).Offset(1, -1).Clear
    
        Do While Not IsEmpty(.Range(columnWithAttributes & rowToStartRead))
        
            attr = .Range(columnWithAttributes & rowToStartRead).value
            
            ' Check attribute is required.
            If .Range(columnWithAttributes & rowToStartRead).Offset(0, 1).value = wsRussianCations.Range("notRequired") Or Trim(.Range(columnWithAttributes & rowToStartRead).Offset(0, 1).value) = "" Then
            
                shortDescription = Replace(shortDescription, "[" & .Range(columnWithAttributes & rowToStartRead).value & "]", "", , , vbTextCompare)
                 
                fullDescription = Replace(fullDescription, "[" & .Range(columnWithAttributes & rowToStartRead).value & "]", "")
            
            Else
            
                ' Insert short attribute value into short description if it exists.
                If Not IsEmpty(.Range(columnWithAttributes & rowToStartRead).Offset(0, 2).value) Then
                
                    shortDescription = UCase(Replace(shortDescription, "[" & .Range(columnWithAttributes & rowToStartRead).value & "]", .Range(columnWithAttributes & rowToStartRead).Offset(0, 2).value, , , vbTextCompare))
                 
                Else
                
                    shortDescription = UCase(Replace(shortDescription, "[" & .Range(columnWithAttributes & rowToStartRead).value & "]", .Range(columnWithAttributes & rowToStartRead).Offset(0, 1).value, , , vbTextCompare))
                    
                End If
            
                
                fullDescription = Replace(fullDescription, "[" & .Range(columnWithAttributes & rowToStartRead).value & "]", .Range(columnWithAttributes & rowToStartRead).Offset(0, 1).value)
            
            End If
            
            ' Technical description.
            techDescription = techDescription & .Range(columnWithAttributes & rowToStartRead).value & ": " & .Range(columnWithAttributes & rowToStartRead).Offset(0, 1).value & vbCrLf
            
            rowToStartRead = rowToStartRead + 1
        
        Loop
        
        ' Check short decription no longer 40 chars.
        If Len(shortDescription) <= 40 Then
        
            .Range(rangeWithShortDescriptionTemplate).Offset(1, 0).value = shortDescription
            
        Else
            
            .Range(rangeWithShortDescriptionTemplate).Offset(1, 0).value = Left(shortDescription, 40)
            .Range(rangeWithShortDescriptionTemplate).Offset(1, -1).value = wsRussianCations.Range("shortDescriptionTrimmedAlert").value
            .Range(rangeWithShortDescriptionTemplate).Offset(1, -1).Font.ColorIndex = 3
            
        End If
        .Range(rangeWithFullDescriptionTemplate).Offset(1, 0).value = fullDescription
        .Range(rangeTechDescription).value = techDescription
    
    End With
    
    
    Call protectSheet(wsCreateMM, sheetProtectionPassword)
    
    Call ImprovePerformance(False)

End Sub

Sub initCreateMMSheet()

    Call ImprovePerformance(True)
        
        Dim categoryDataValidationFormula As String: categoryDataValidationFormula = ""
        Dim value As Variant
        
        wsCreateMM.category_ComboBox.Clear
        
        ' Get categories list for category cell.
        For Each value In SelectionToDictionary(wsCategoriesMaster.Range("A2:A" & CStr(lrWsCategoriesMaster)), Wb).Keys()
        
            wsCreateMM.category_ComboBox.AddItem value
            
        Next value
        
        ' Init material code combobox.
        wsCreateMM.materialGroupCode_ComboBox.List = wsPurchasingGroupsMaster.Range("materialGroupCodes").value
        
        ' Init purchasing code combobox.
        For Each value In SelectionToDictionary(wsPurchasingGroupsMaster.Range("purchasingGroupCodes"), Wb).Keys()
            wsCreateMM.purchasingGroup_ComboBox.AddItem value
        Next value
        
        ' Init priority combobox.
        wsCreateMM.priority_ComboBox.List = wsSettings.Range("priority").value
        
        
    Call ImprovePerformance(False)
    
End Sub

Sub init()

    Call ImprovePerformance(True)

    ' Init workbooks and worksheets.
    Set Wb = ThisWorkbook
    Set wsCreateMM = Wb.Sheets("Create MM")
    Set wsSettings = Wb.Sheets("settings")
    Set wsCategoriesMaster = Wb.Sheets("categoriesMaster")
    Set wsRussianCations = Wb.Sheets("russianCaptions")
    Set wsPurchasingGroupsMaster = Wb.Sheets("purchasingGroupsMaster")
    Set wsContractorsMaster = Wb.Sheets("contractorsMaster")
    Set wsBulkImport = Wb.Sheets("Bulk import to SAP")
    Set wsCableTypeMaster = Wb.Sheets("cableTypesMaster")
    
    lrWsCategoriesMaster = GetBorders("LR", wsCategoriesMaster.Name, Wb)
    
    ' Init variables.
    Set categoryCreateMMRange = wsCreateMM.Range("createMM_category")
    
    
    'If wsBulkImport.ProtectContents = True Then wsBulkImport.Unprotect sheetProtectionPassword
    
    'ThisWorkbook.Styles("Normal").Font.Size = 12
    
    'If wsBulkImport.ProtectContents = False Then wsBulkImport.Protect sheetProtectionPassword
    
    Call ImprovePerformance(False)

End Sub

Sub setValueToCombobox(combobox As Object, filterValue As String, filterColumn As String, valueColumn As String, ws As Worksheet)



    Dim rowToStart As Integer: rowToStart = 2
    Dim rowsCounter As Long
    
    Dim valuesDict          As New Dictionary
    Dim value               As Variant
    Dim stringWithValues    As String
    Dim emptyResultValue    As String
    
    
    
    With ws
    
        For rowsCounter = rowToStart To lrWsCategoriesMaster
        
            If .Range(filterColumn & rowsCounter).value = filterValue Then
            
                    valuesDict(.Range(valueColumn & rowsCounter).value) = Empty
                    
            End If
            
        Next rowsCounter
        
    End With
    
    
    
    For Each value In valuesDict.Keys()
        combobox.value = value
    Next value
    
    
    
End Sub

Sub addValuesToCombobox(combobox As combobox, filterValue As String, filterColumn As String, valueColumn As String, ws As Worksheet)



    Dim rowToStart As Integer: rowToStart = 2
    Dim rowsCounter As Long
    
    Dim valuesDict          As New Dictionary
    Dim value               As Variant
    Dim stringWithValues    As String
    Dim emptyResultValue    As String
    
    
    
    With ws
    
        For rowsCounter = rowToStart To lrWsCategoriesMaster
        
            If .Range(filterColumn & rowsCounter).value = filterValue Then
            
                    valuesDict(.Range(valueColumn & rowsCounter).value) = Empty
                    
            End If
            
        Next rowsCounter
        
    End With
    
    
    
    For Each value In valuesDict.Keys()
        combobox.AddItem value
    Next value
    
    
    
End Sub

Function concatRangeValues(filterValue As String, dataRange As Range, Optional delimiter As String = ",", Optional ws As Worksheet) As String



    If ws Is Nothing Then Set ws = wsCategoriesMaster
    
    
    
    Dim row             As Range
    Dim cellsCounter    As Integer
    Dim dropDownList    As String: dropDownList = ""
    
    
    
    For Each row In dataRange.Rows
        
        For cellsCounter = 1 To dataRange.Rows.Columns.Count
        
            dropDownList = dropDownList & row.Cells(1, cellsCounter).value & " | "
            
        Next cellsCounter
        
        dropDownList = Left(dropDownList, Len(dropDownList) - 3) & vbCrLf
        
    Next row
    
    concatRangeValues = Left(dropDownList, Len(dropDownList) - 1)
    
    

End Function

Function priorityFromTextToValue(priorityText As String) As String



    Call init
    
    
    
    If priorityText = wsSettings.Range("piorityHigh") Then
        priorityFromTextToValue = "02"
    ElseIf priorityText = wsSettings.Range("piorityNormal") Then
        priorityFromTextToValue = "01"
    Else
        priorityFromTextToValue = "ERROR"
    End If
    
    
    
End Function

Function checkBoxToText(chkbox As Object)



    If chkbox.value = True Then
        checkBoxToText = "X"
    ElseIf chkbox.value = False Then
        chkbox = ""
    End If



End Function

Sub drawClassAttributes(class As String, rowToStart As Integer, shortDescriptionRangeAddress As String, fullDescriptionAddress As String, techDescriptionAddress As String, shortDescriptionEngAddress As String)



    Call ImprovePerformance(True)
    Call init



    Dim rowsCounter         As Long
    Dim rowsOutputCounter    As Long: rowsOutputCounter = rowToStart
    
    Dim valuesDict          As New Dictionary
    Dim value               As Variant
    
    Dim lrWsCreateMM        As Long
    
    
    
    'Draw header.
    wsCreateMM.Range("A" & rowToStart - 1).value = wsRussianCations.Range("attribute").value
    wsCreateMM.Range("B" & rowToStart - 1).value = wsRussianCations.Range("attributeValue").value
    wsCreateMM.Range("C" & rowToStart - 1).value = wsRussianCations.Range("attributeShortValue").value
    wsCreateMM.Range("A" & rowToStart - 1 & ":C" & rowToStart - 1).Font.Bold = True
    
    
    
    ' Clear area for drawing.
    wsCreateMM.Range("A" & rowToStart & ":B1000").EntireRow.Delete
    wsCreateMM.Range("A" & rowToStart & ":B1000").NumberFormat = "@"
    
    
    
    ' Gather attributes.
    For rowsCounter = 2 To lrWsCategoriesMaster
    
        If wsCategoriesMaster.Range("D" & rowsCounter).value = class Then
        
            valuesDict(wsCategoriesMaster.Range("E" & rowsCounter).value) = Empty
            
        End If
        
    Next rowsCounter
    
    
    
    ' Draw attributes to CreateMM sheet.
    For Each value In valuesDict.Keys()
    
        wsCreateMM.Range("A" & rowsOutputCounter).value = value
        rowsOutputCounter = rowsOutputCounter + 1
    
    Next value
    
    
    
    lrWsCreateMM = GetBorders("LR", wsCreateMM.Name, Wb)
    
    
    
    Dim attrList As String
    ' Attribute's drop down lists.
    For rowsOutputCounter = rowToStart To lrWsCreateMM
    
        For rowsCounter = 2 To lrWsCategoriesMaster
        
            If class = wsCategoriesMaster.Range("D" & rowsCounter).value And _
                wsCreateMM.Range("A" & rowsOutputCounter).value = wsCategoriesMaster.Range("E" & rowsCounter).value Then
            
                attrList = attrList & CStr(wsCategoriesMaster.Range("F" & rowsCounter).value) & ","
            
            End If
            
        Next rowsCounter
        
        ' Set drop down lists for atrributes valeus.
        If Len(attrList) > 1 Then
        
            wsCreateMM.Range("B" & rowsOutputCounter).Validation.Delete
            'wsCreateMM.Range("B" & rowsOutputCounter).NumberFormat = "@"
            attrList = Left(attrList, Len(attrList) - 1)
            wsCreateMM.Range("B" & rowsOutputCounter).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=attrList
            wsCreateMM.Range("B" & rowsOutputCounter).Validation.ShowError = False
            
        End If
        
        attrList = ""
        
    Next rowsOutputCounter
    
    
    
    ' Get short description header.
    wsCreateMM.Range(shortDescriptionRangeAddress).value = wsRussianCations.Range("shortDescriptionText").value
    wsCreateMM.Range(shortDescriptionRangeAddress).Font.Bold = True
    wsCreateMM.Range(shortDescriptionRangeAddress).VerticalAlignment = xlVAlignTop
    
    
    
    ' Get Short description template text.
    For rowsCounter = 2 To lrWsCategoriesMaster

        If wsCategoriesMaster.Range("D" & rowsCounter).value = class Then
            ' Here!
            wsCreateMM.Range(shortDescriptionRangeAddress).Offset(0, 1).value = wsCategoriesMaster.Range("J" & rowsCounter).value
            Exit For

        End If

    Next rowsCounter
    
    ' Short description Template formatting.
    'wsCreateMM.Range(shortDescriptionRangeAddress).Offset(0, 1).Font.Italic = True
    
    ' Short description value field formatting.
    wsCreateMM.Range(shortDescriptionRangeAddress).Offset(1, 1).Borders.LineStyle = xlContinuous
    wsCreateMM.Range(shortDescriptionRangeAddress).Offset(1, 1).VerticalAlignment = xlVAlignTop
    
    
    ' Get full description Header.
    wsCreateMM.Range(fullDescriptionAddress).value = wsRussianCations.Range("fullDescriptionText").value
    wsCreateMM.Range(fullDescriptionAddress).Font.Bold = True
    
    ' Get full description Template.
    wsCreateMM.Range(fullDescriptionAddress).Offset(0, 1).Font.Italic = True
    
    For rowsCounter = 2 To lrWsCategoriesMaster

        If wsCategoriesMaster.Range("D" & rowsCounter).value = class Then

            wsCreateMM.Range(fullDescriptionAddress).Offset(0, 1).value = wsCategoriesMaster.Range("K" & rowsCounter).value
            Exit For

        End If

    Next rowsCounter
    
    ' Format full description value field.
    wsCreateMM.Range(wsCreateMM.Range(fullDescriptionAddress).Offset(1, 1), wsCreateMM.Range(fullDescriptionAddress).Offset(11, 1)).Merge
    wsCreateMM.Range(wsCreateMM.Range(fullDescriptionAddress).Offset(1, 1), wsCreateMM.Range(fullDescriptionAddress).Offset(11, 1)).Borders.LineStyle = xlContinuous
    wsCreateMM.Range(fullDescriptionAddress).Offset(1, 1).WrapText = True
    wsCreateMM.Range(fullDescriptionAddress).Offset(1, 1).ColumnWidth = 80
    wsCreateMM.Range(fullDescriptionAddress).Offset(1, 1).VerticalAlignment = xlVAlignTop



    ' Technical description Header.
    wsCreateMM.Range(techDescriptionAddress).value = wsRussianCations.Range("techDescriptionText").value
    wsCreateMM.Range(techDescriptionAddress).Font.Bold = True

    ' Full description value field formatting.
    wsCreateMM.Range(wsCreateMM.Range(techDescriptionAddress).Offset(1, 1), wsCreateMM.Range(techDescriptionAddress).Offset(11, 1)).Merge
    wsCreateMM.Range(wsCreateMM.Range(techDescriptionAddress).Offset(1, 1), wsCreateMM.Range(techDescriptionAddress).Offset(11, 1)).VerticalAlignment = xlVAlignTop
    wsCreateMM.Range(wsCreateMM.Range(techDescriptionAddress).Offset(1, 1), wsCreateMM.Range(techDescriptionAddress).Offset(11, 1)).Borders.LineStyle = xlContinuous
    wsCreateMM.Range(wsCreateMM.Range(techDescriptionAddress).Offset(1, 1), wsCreateMM.Range(techDescriptionAddress).Offset(11, 1)).WrapText = True



    ' Short description eng.
    ' Header.
    wsCreateMM.Range(shortDescriptionEngAddress).value = wsRussianCations.Range("shortDescriptionEngText").value
    wsCreateMM.Range(shortDescriptionEngAddress).Font.Bold = True
    
    ' Field for text.
    wsCreateMM.Range(shortDescriptionEngAddress).Offset(1, 1).VerticalAlignment = xlVAlignTop
    wsCreateMM.Range(shortDescriptionEngAddress).Offset(1, 1).Borders.LineStyle = xlContinuous


    Call ImprovePerformance(False)
    
    

End Sub

Sub SaveTemplateToSeparateFile(soruceSheet As String)



    Call ImprovePerformance(True)
    
    Call init

    Dim targetWb        As Workbook
    Dim currentWb       As Workbook
    
    Dim choiceFileDialog As Integer
    Dim resultExtension  As Integer
    Dim shtIndex         As Integer
    
    Dim chosenExtension As String
    Dim pathToFile      As String
    Dim fullPathToSave  As Variant
    
    Dim defaultSheetName    As String
    Dim templateSheetName   As String
    Dim sourceSheetName     As String
    
    Dim templateLastRow     As Long



    Set currentWb = ThisWorkbook
    
    
    
    templateSheetName = soruceSheet
    sourceSheetName = soruceSheet
    
    
    
    ' Call FileDialog to choose location for file saving.
    fullPathToSave = Application.GetSaveAsFilename(InitialFileName:="", _
                                                    FileFilter:="Excel Workbook (*.xlsx),*.xlsx," + _
                                                                "Excel Binary Workbook (*.xlsb),*.xlsb,")
                                                                
                                                                
    'Interrupt sub if user pressed Cancel or X in FileDialog.
    If fullPathToSave = False Then
        MsgBox "Path is not chosen."
        Exit Sub
    End If



    chosenExtension = RxMatch(fullPathToSave, "\.[\w]+$")       'Get chosen excel workbook's extension.
    fullPathToSave = RxReplace(fullPathToSave, "\.[\w]+$", "")  'Cut extension from full path to saving workbook.
    
    
    
    Select Case chosenExtension
        Case ".xlsx"
            'You want to save Excel 2007-2016 file
            resultExtension = xlWorkbookDefault
        Case ".xlsb"
            'You want ta save Excel 2007-2016 BINARY file
            resultExtension = xlExcel12
    End Select
    
    
    
    ' Save new Workbook to specified folder with specified in FilDialog name.
    Workbooks.Add
    Set targetWb = ActiveWorkbook
    targetWb.SaveAs FileName:=fullPathToSave, FileFormat:=resultExtension
    
    
    
    ' Adding temporary list to have a possibility delete default
    ' workbook sheet.
    With targetWb
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "EmptyList"
    End With
    
    
    
    ' Detect defualt sheet
    If SheetExists("Sheet1", targetWb) = True Then
        defaultSheetName = "Sheet1"
    ElseIf SheetExists("Лист1", targetWb) = True Then
        defaultSheetName = "Лист1"
    End If
    
    
    
    ' Delete default sheet
    targetWb.Sheets(defaultSheetName).Delete
    
    
    
    ' Copy sheets from source workbook to target.
    shtIndex = 1
    Dim ws As Worksheet
    For Each ws In currentWb.Worksheets
        If ws.Name = templateSheetName Then
            currentWb.Activate
            'Ws.Visible = xlSheetVisible
            ws.Copy After:=targetWb.Sheets("EmptyList")
            'Ws.Visible = xlSheetVeryHidden
        End If
        shtIndex = shtIndex + 1
    Next ws
    
    
    
    templateLastRow = GetBorders("LR", sourceSheetName, currentWb)
    
    
    
    ' Copy data from curent WB to target.
    currentWb.Sheets(sourceSheetName).Range("A11:U" & templateLastRow).Copy targetWb.Sheets(templateSheetName).Range("A11:U" & templateLastRow)
    targetWb.Sheets(templateSheetName).Columns("N:T").EntireColumn.Delete
    'targetWb.Sheets(templateSheetName).Range("A11:Q" & templateLastRow).PasteSpecial Paste:=xlPasteFormats
    
    
    
    ' Handle cases when no sheets was cpoied.
    If targetWb.Sheets.Count <= 1 Then
        MsgBox ("Copy error. There are no sheets to copy.")
        Exit Sub
    End If
    
    
    
    If SheetExists("EmptyList", targetWb) = True Then: targetWb.Sheets("EmptyList").Delete
    
    
    
    targetWb.Save
    targetWb.Close
    
    
    
    Call ImprovePerformance(False)
    
    
    
    MsgBox (wsRussianCations.Range("saveTemplateSuccess").value)
    
    
    
End Sub

Sub ClearBulkImportSheet()


    Call ImprovePerformance(True)

    Call init
    
    
    
    Dim lrWsBulkImport As Long
    
    
    Call unprotectSheet(wsBulkImport, sheetProtectionPassword)
    
    
    With wsBulkImport

        .Range("A11:A1000").EntireRow.Delete
     
    End With
    
    
    Call protectSheet(wsBulkImport, sheetProtectionPassword)
    
    
    Call ImprovePerformance(False)


End Sub


Sub deleteRowsBulkImport()

    
    Call init
    
    
    
    Call unprotectSheet(wsBulkImport, sheetProtectionPassword)
    
    
    
    If Selection.Worksheet.Name = wsBulkImport.Name Then
    
        If Selection.row > 10 Then
            
            Selection.EntireRow.Delete
        
        End If
    
    End If



    Call protectSheet(wsBulkImport, sheetProtectionPassword)
    


End Sub



Sub lockCreateMMRanges()



    Call init
    
    
    Call unprotectSheet(wsCreateMM, sheetProtectionPassword)
    
    
    With wsCreateMM
    
       
    
        .Range("F12").Locked = False
        
        .Range("F15:F25").Locked = False
        
        .Range("F28:F38").Locked = False
        
        .Range("F41").Locked = False
        
        .Range("B12:B700").Locked = False
        
        wsCreateMM.Protect password:=sheetProtectionPassword, DrawingObjects:=True, Contents:=True
        
        
        
    End With
    
    
End Sub
