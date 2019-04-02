Attribute VB_Name = "WorksheetFunctions"
Public Function GetTaxYear(ByVal DateAsText As String)
    
    If Len(DateAsText) <> 10 Then
        GetTaxYear = "Date Must Be 10 Characters"
        Exit Function
    End If
    
    Dim Day As Integer
    Dim Month As Integer
    Dim Year As Integer
    
    Day = Left(DateAsText, 2)
    Month = Mid(DateAsText, 4, 2)
    Year = Right(DateAsText, 4)
    
    If Month < 4 Or (Month = 4 And Day < 6) Then
        GetTaxYear = Year - 1
    Else
        GetTaxYear = Year
    End If
    
End Function
