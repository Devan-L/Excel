Private Function pNINOValidation(ByVal iNINO As String)

    'Presence Check
    If iNINO = "" Then
        pNINOValidation = "Missing NINO"
        Exit Function
    End If



    'Test length
        If Len(iNINO) > 9 Then pNINOValidation = "NINO: Too Long"
        If Len(iNINO) < 8 Then pNINOValidation = "NINO: Too Short"
    If pNINOValidation <> "" Then Exit Function
    
    'Test format AANNNNNNA
        If Not Left(iNINO, 1) Like "[a-z,A-Z]" Or _
            Not Right(Left(iNINO, 2), 1) Like "[a-z,A-Z]" Or _
            Not Right(Left(iNINO, 3), 1) Like "[0-9]" Or _
            Not Right(Left(iNINO, 4), 1) Like "[0-9]" Or _
            Not Right(Left(iNINO, 5), 1) Like "[0-9]" Or _
            Not Right(Left(iNINO, 6), 1) Like "[0-9]" Or _
            Not Right(Left(iNINO, 7), 1) Like "[0-9]" Or _
            Not Right(Left(iNINO, 8), 1) Like "[0-9]" Then pNINOValidation = "NINO: Invalid Format"
    
        
        'Optional 9th character
        If Len(iNINO) = 9 Then
            If Not Right(iNINO, 1) Like "[a-z,A-Z]" Then pNINOValidation = "NINO: Invalid Format"
        End If
    If pNINOValidation <> "" Then Exit Function


    'Test First Character
        If Left(iNINO, 1) = "D" Or Left(iNINO, 1) = "F" Or Left(iNINO, 1) = "I" Or Left(iNINO, 1) = "Q" Or Left(iNINO, 1) = "U" Or Left(iNINO, 1) = "V" Then pNINOValidation = "NINO: Invalid Character Present"

    'Test Second Character
        If Left(iNINO, 1) = "D" Or Left(iNINO, 1) = "F" Or Left(iNINO, 1) = "I" Or Left(iNINO, 1) = "O" Or Left(iNINO, 1) = "Q" Or Left(iNINO, 1) = "U" Or Left(iNINO, 1) = "V" Then pNINOValidation = "NINO: Invalid Character Present"

    'Test First Two Characters
        If Left(iNINO, 2) = "BG" Or Left(iNINO, 1) = "GB" Or Left(iNINO, 1) = "KN" Or Left(iNINO, 1) = "NT" Or Left(iNINO, 1) = "ZZ" Then pNINOValidation = "NINO: Invalid Character Present"


If pNINOValidation <> "" Then Exit Function
pNINOValidation = "NINO: Pass"

End Function

Private Function pPostcodeValidation(ByVal iPostcode As String)
    
    Dim iOutwardCode As String
    Dim iInwardCode As String
    
    
    'Presence Check
    If iPostcode = "" Then
        pPostcodeValidation = "Postcode: Missing"
        Exit Function
    End If
    
    On Error Resume Next
    
    iOutwardCode = Left(iPostcode, WorksheetFunction.Find(" ", iPostcode) - 1)
    iInwardCode = Right(iPostcode, Len(iPostcode) - WorksheetFunction.Find(" ", iPostcode))
    
    On Error GoTo 0
    
    'Overall length
        If Len(iPostcode) > 8 Then pPostcodeValidation = "Postcode: Too Long"
        If Len(iPostcode) < 6 Then pPostcodeValidation = "Postcode: Too Short"
    If pPostcodeValidation <> "" Then Exit Function
    
    'Outward Code length
        If Len(iOutwardCode) > 4 Then pPostcodeValidation = "Postcode: Outward Code too long"
        If Len(iOutwardCode) < 2 Then pPostcodeValidation = "Postcode: Outward Code too short"
    If pPostcodeValidation <> "" Then Exit Function
    
    'Outward code format checks:
    
        'AN
        If Len(iOutwardCode) = 2 Then
            If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And _
                Right(Left(iOutwardCode, 2), 1) Like "[0-9]") Then
                pPostcodeValidation = "Postcode: Invalid Format for Outward Code"
                Exit Function
            End If
        End If
        
        'ANN / AAN /ANA
        If Len(iOutwardCode) = 3 Then
            If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[0-9]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]") Or _
                (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]") Or _
                (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[0-9]" And Right(Left(iOutwardCode, 3), 1) Like "[a-z,A-Z]") Then
                pPostcodeValidation = "Postcode: Invalid Format for Outward Code"
                Exit Function
            End If
        End If
        
        'AANA / AANN
        If Len(iOutwardCode) = 4 Then
            If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]" And Right(Left(iOutwardCode, 4), 1) Like "[0-9]") Or _
                (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]" And Right(Left(iOutwardCode, 4), 1) Like "[a-z,A-Z]") Then
                pPostcodeValidation = "Postcode: Invalid Format for Outward Code"
                Exit Function
            End If
        End If
                 

    'Inward Code Format Check (NAA)
    If (Not Left(iInwardCode, 1) Like "[0-9]" Or Not Right(Left(iInwardCode, 2), 1) Like "[a-z,A-Z]" Or Not Right(Left(iInwardCode, 3), 1) Like "[a-z,A-Z]") Then
        pPostcodeValidation = "Postcode: Invalid Format for Inward Code"
        Exit Function
    End If



    'Check Outward Code Banned Letters
        'First Alpha Check
            If Left(iOutwardCode, 1) = "Q" Or Left(iOutwardCode, 1) = "V" Or Left(iOutwardCode, 1) = "X" Then pPostcodeValidation = "Postcode: Invalid Alpha Character in Outward Code"
        
        'Second Alpha Check
            If Right(Left(iOutwardCode, 2), 1) = "I" Or Right(Left(iOutwardCode, 1), 1) = "J" Or Right(Left(iOutwardCode, 1), 1) = "Z" Then pPostcodeValidation = "Postcode: Invalid Alpha Character in Outward Code"
        
        'Third Position Check
            If Not IsNumeric(Right(Left(iOutwardCode, 3), 1)) Then
                If Right(Left(iOutwardCode, 3), 1) <> "A" And Right(Left(iOutwardCode, 3), 1) <> "B" And Right(Left(iOutwardCode, 3), 1) <> "C" And Right(Left(iOutwardCode, 3), 1) <> "D" And Right(Left(iOutwardCode, 3), 1) <> "E" And Right(Left(iOutwardCode, 3), 1) <> "F" And Right(Left(iOutwardCode, 3), 1) <> "G" And Right(Left(iOutwardCode, 3), 1) <> "H" And Right(Left(iOutwardCode, 3), 1) <> "J" And Right(Left(iOutwardCode, 3), 1) <> "K" And Right(Left(iOutwardCode, 3), 1) <> "P" And Right(Left(iOutwardCode, 3), 1) <> "R" And Right(Left(iOutwardCode, 3), 1) <> "S" And Right(Left(iOutwardCode, 3), 1) <> "T" And Right(Left(iOutwardCode, 3), 1) <> "U" Then pPostcodeValidation = "Postcode: Invalid Alpha Character in Outward Code"
            End If



    'Inward Code Banned Letter Check
    If iInwardCode Like "*C*" Or iInwardCode Like "*I*" Or iInwardCode Like "*K*" Or iInwardCode Like "*M*" Or iInwardCode Like "*O*" Or iInwardCode Like "*V*" Then pPostcodeValidation = "Postcode: Invalid Inward Code"
    
    


If pPostcodeValidation <> "" Then Exit Function
pPostcodeValidation = "Postcode: Pass"

End Function

Private Function pTitleGenderValidation(ByVal iTitle As String, ByVal iGender As String)
    
    
    'Title presence
    If iTitle = "" Then
        pTitleGenderValidation = "Missing Title"
        Exit Function
    End If
        
    'Gender presence
    If iGender = "" Then
        pTitleGenderValidation = "Missing Gender"
        Exit Function
    End If
            
    'Gender Valid
    If iGender <> "M" And iGender <> "Male" And iGender <> "F" And iGender <> "Female" Then
        pTitleGenderValidation = "Gender: Invalid Value"
        Exit Function
    End If
        
    'Male Title with Female Check
    If (iTitle = "MR" Or iTitle = "SIR" Or iTitle = "Rev" Or iTitle = "Lord" Or iTitle = "CAPT") And (iGender = "F" Or iGender = "Female") Then
        pTitleGenderValidation = "Gender/Title: Inconsistent Gender and Title"
        Exit Function
    End If
    
    'Female Title with Male Check
    If (iTitle = "MS" Or iTitle = "Miss" Or iTitle = "MRS") And (iGender = "M" Or iGender = "Male") Then
        pTitleGenderValidation = "Gender/Title: Inconsistent Gender and Title"
        Exit Function
    End If
    
    'Misc other titles
    'Dr / Cdr / EUR.ING
    
If pTitleGenderValidation <> "" Then Exit Function
pTitleGenderValidation = "Gender/Title: Pass"
End Function

Private Function pSurnameValidation(ByVal iSurname As String)

    'Presence Check
    If iSurname = "" Then
        pSurnameValidation = "Missing Surname"
        Exit Function
    End If


    'Format Check:
            If iSurname Like "*[0-9,!,£,$,%,^,&,*,(,),+,=,_,`,¬,\,|,<,>,?,/,'.,:,#,~,[,{,}]*" Then
                pSurnameValidation = "Surname: Invalid Format"
                Exit Function
            End If


        
If pSurnameValidation <> "" Then Exit Function
pSurnameValidation = "Surname: Pass"

End Function

Private Function pDOBValidation(ByVal istrDOB As String)
    
    Dim iDOB As Date

    'Presence Check
    If istrDOB = "" Then
        pDOBValidation = "Missing DOB"
        Exit Function
    End If

    On Error Resume Next
    iDOB = DateValue(istrDOB)
    On Error GoTo 0

    'Format Check:
    If Not IsDate(iDOB) Or iDOB = 0 Then
        pDOBValidation = "DOB: Invalid Format for Date"
        Exit Function
    End If


    'Future Date Check:
    If iDOB > Now Then
        pDOBValidation = "DOB: Future Date"
        Exit Function
    End If
        
If pDOBValidation <> "" Then Exit Function
pDOBValidation = "DOB: Pass"

End Function
