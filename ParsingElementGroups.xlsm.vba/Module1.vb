Attribute VB_Name = "Module1"
Option Explicit

Public Sub Process()
    
    Const elementsInGroupCount As Integer = 4
    
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    
    Set sourceWorksheet = ThisWorkbook.Sheets("original")
    Set destinationWorksheet = ThisWorkbook.Sheets("new")
    
    ClearSheetContents destinationWorksheet
    
    Dim sourceRowCount As Integer
    Dim destinationRowCount As Integer
    
    destinationRowCount = 2
    For sourceRowCount = 2 To sourceWorksheet.UsedRange.Rows.Count
        
        Dim currentMember As Member
        Set currentMember = New Member
        currentMember.MemberNumber = sourceWorksheet.Cells(sourceRowCount, 1).Value
        currentMember.NationalInsuranceNumber = sourceWorksheet.Cells(sourceRowCount, 2).Value
        
        Dim currentGroupCount As Integer
        For currentGroupCount = 0 To 10000
            Dim currentGroup As Range
            Set currentGroup = sourceWorksheet.Cells(sourceRowCount, 3).Offset(0, currentGroupCount * (elementsInGroupCount + 1))
                        
            If currentGroup.Value = "" Then
                Exit For
            End If
            
            Dim currentElementCount As Integer
            For currentElementCount = 1 To elementsInGroupCount
                Dim currentElement As PensionElement
                Set currentElement = New PensionElement
                
                currentElement.ElementDate = currentGroup.Value
                currentElement.Value = currentGroup.Offset(0, currentElementCount).Value
                currentElement.ElementID = currentGroup.Offset(-(currentGroup.Row - 1), currentElementCount).Value
                
                currentMember.Elements.Add currentElement
            Next
        Next
        
        outputMember destinationWorksheet, currentMember
    Next
    
End Sub

Private Sub outputMember(outputSheet As Worksheet, outputMember As Member)
    
    Dim destinationRow As Integer
    destinationRow = outputSheet.UsedRange.Rows.Count + 1
    
    Dim currentElement As PensionElement
    For Each currentElement In outputMember.Elements
        outputSheet.Cells(destinationRow, 1).Value = outputMember.MemberNumber
        outputSheet.Cells(destinationRow, 2).Value = currentElement.ElementID
        outputSheet.Cells(destinationRow, 3).Value = currentElement.ElementType
        outputSheet.Cells(destinationRow, 4).Value = currentElement.ElementDate
        outputSheet.Cells(destinationRow, 5).Value = currentElement.Value
        outputSheet.Cells(destinationRow, 6).Value = outputMember.NationalInsuranceNumber
        destinationRow = destinationRow + 1
    Next
    
End Sub

Private Sub ClearSheetContents(sheet As Worksheet)

    sheet.Range(sheet.Cells(2, 1), sheet.Cells(WorksheetFunction.Max(2, sheet.UsedRange.Rows.Count), 1)).EntireRow.Delete True

End Sub
