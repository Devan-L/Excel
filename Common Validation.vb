Private Function IsNINOValid(ByVal NINO As String)

    'Presence Check
    If NINO = "" Then
        IsNINOValid = "Missing NINO"
        Exit Function
    End If



    'Test length
	If Len(NINO) > 9 Then IsNINOValid = "NINO: Too Long"
	If Len(NINO) < 8 Then IsNINOValid = "NINO: Too Short"
    If IsNINOValid <> "" Then Exit Function
    
    'Test format AANNNNNNA
	If Not Left(NINO, 1) Like "[a-z,A-Z]" Or _
		Not Right(Left(NINO, 2), 1) Like "[a-z,A-Z]" Or _
		Not Right(Left(NINO, 3), 1) Like "[0-9]" Or _
		Not Right(Left(NINO, 4), 1) Like "[0-9]" Or _
		Not Right(Left(NINO, 5), 1) Like "[0-9]" Or _
		Not Right(Left(NINO, 6), 1) Like "[0-9]" Or _
		Not Right(Left(NINO, 7), 1) Like "[0-9]" Or _
		Not Right(Left(NINO, 8), 1) Like "[0-9]" Then IsNINOValid = "NINO: Invalid Format"
    
        
        'Optional 9th character
	If Len(NINO) = 9 Then
		If Not Right(NINO, 1) Like "[a-z,A-Z]" Then IsNINOValid = "NINO: Invalid Format"
	End If
    If IsNINOValid <> "" Then Exit Function


    'Test First Character
	If Left(NINO, 1) = "D" Or Left(NINO, 1) = "F" Or Left(NINO, 1) = "I" Or Left(NINO, 1) = "Q" Or Left(NINO, 1) = "U" Or Left(NINO, 1) = "V" Then IsNINOValid = "NINO: Invalid Character Present"

    'Test Second Character
	If Left(NINO, 1) = "D" Or Left(NINO, 1) = "F" Or Left(NINO, 1) = "I" Or Left(NINO, 1) = "O" Or Left(NINO, 1) = "Q" Or Left(NINO, 1) = "U" Or Left(NINO, 1) = "V" Then IsNINOValid = "NINO: Invalid Character Present"

    'Test First Two Characters
	If Left(NINO, 2) = "BG" Or Left(NINO, 1) = "GB" Or Left(NINO, 1) = "KN" Or Left(NINO, 1) = "NT" Or Left(NINO, 1) = "ZZ" Then IsNINOValid = "NINO: Invalid Character Present"


	If IsNINOValid <> "" Then Exit Function
	IsNINOValid = "NINO: Pass"

End Function

Private Function IsPostcodeValid(ByVal Postcode As String)
    
    Dim iOutwardCode As String
    Dim iInwardCode As String
    
    
    'Presence Check
    If Postcode = "" Then
        IsPostcodeValid = "Postcode: Missing"
        Exit Function
    End If
    
    On Error Resume Next
    
    iOutwardCode = Left(Postcode, WorksheetFunction.Find(" ", Postcode) - 1)
    iInwardCode = Right(Postcode, Len(Postcode) - WorksheetFunction.Find(" ", Postcode))
    
    On Error GoTo 0
    
    'Overall length
	If Len(Postcode) > 8 Then IsPostcodeValid = "Postcode: Too Long"
	If Len(Postcode) < 6 Then IsPostcodeValid = "Postcode: Too Short"
    If IsPostcodeValid <> "" Then Exit Function
    
    'Outward Code length
        If Len(iOutwardCode) > 4 Then IsPostcodeValid = "Postcode: Outward Code too long"
        If Len(iOutwardCode) < 2 Then IsPostcodeValid = "Postcode: Outward Code too short"
    If IsPostcodeValid <> "" Then Exit Function
    
    'Outward code format checks:
    
	'AN
	If Len(iOutwardCode) = 2 Then
		If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And _
			Right(Left(iOutwardCode, 2), 1) Like "[0-9]") Then
			IsPostcodeValid = "Postcode: Invalid Format for Outward Code"
			Exit Function
		End If
	End If
        
	'ANN / AAN /ANA
	If Len(iOutwardCode) = 3 Then
		If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[0-9]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]") Or _
			(Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]") Or _
			(Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[0-9]" And Right(Left(iOutwardCode, 3), 1) Like "[a-z,A-Z]") Then
			IsPostcodeValid = "Postcode: Invalid Format for Outward Code"
			Exit Function
		End If
	End If
	
	'AANA / AANN
	If Len(iOutwardCode) = 4 Then
		If Not (Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]" And Right(Left(iOutwardCode, 4), 1) Like "[0-9]") Or _
			(Left(iOutwardCode, 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 2), 1) Like "[a-z,A-Z]" And Right(Left(iOutwardCode, 3), 1) Like "[0-9]" And Right(Left(iOutwardCode, 4), 1) Like "[a-z,A-Z]") Then
			IsPostcodeValid = "Postcode: Invalid Format for Outward Code"
			Exit Function
		End If
	End If
			 

    'Inward Code Format Check (NAA)
    If (Not Left(iInwardCode, 1) Like "[0-9]" Or Not Right(Left(iInwardCode, 2), 1) Like "[a-z,A-Z]" Or Not Right(Left(iInwardCode, 3), 1) Like "[a-z,A-Z]") Then
        IsPostcodeValid = "Postcode: Invalid Format for Inward Code"
        Exit Function
    End If



    'Check Outward Code Banned Letters
	'First Alpha Check
	If Left(iOutwardCode, 1) = "Q" Or Left(iOutwardCode, 1) = "V" Or Left(iOutwardCode, 1) = "X" Then IsPostcodeValid = "Postcode: Invalid Alpha Character in Outward Code"
	
	'Second Alpha Check
	If Right(Left(iOutwardCode, 2), 1) = "I" Or Right(Left(iOutwardCode, 1), 1) = "J" Or Right(Left(iOutwardCode, 1), 1) = "Z" Then IsPostcodeValid = "Postcode: Invalid Alpha Character in Outward Code"
	
	'Third Position Check
	If Not IsNumeric(Right(Left(iOutwardCode, 3), 1)) Then
		If Right(Left(iOutwardCode, 3), 1) <> "A" And Right(Left(iOutwardCode, 3), 1) <> "B" And Right(Left(iOutwardCode, 3), 1) <> "C" And Right(Left(iOutwardCode, 3), 1) <> "D" And Right(Left(iOutwardCode, 3), 1) <> "E" And Right(Left(iOutwardCode, 3), 1) <> "F" And Right(Left(iOutwardCode, 3), 1) <> "G" And Right(Left(iOutwardCode, 3), 1) <> "H" And Right(Left(iOutwardCode, 3), 1) <> "J" And Right(Left(iOutwardCode, 3), 1) <> "K" And Right(Left(iOutwardCode, 3), 1) <> "P" And Right(Left(iOutwardCode, 3), 1) <> "R" And Right(Left(iOutwardCode, 3), 1) <> "S" And Right(Left(iOutwardCode, 3), 1) <> "T" And Right(Left(iOutwardCode, 3), 1) <> "U" Then IsPostcodeValid = "Postcode: Invalid Alpha Character in Outward Code"
	End If

    'Inward Code Banned Letter Check
    If iInwardCode Like "*C*" Or iInwardCode Like "*I*" Or iInwardCode Like "*K*" Or iInwardCode Like "*M*" Or iInwardCode Like "*O*" Or iInwardCode Like "*V*" Then IsPostcodeValid = "Postcode: Invalid Inward Code"
    
	
	If IsPostcodeValid <> "" Then Exit Function
	IsPostcodeValid = "Postcode: Pass"

End Function

Private Function IsTitleGenderValid(ByVal Title As String, ByVal Gender As String)
    
    
    'Title presence
    If Title = "" Then
        IsTitleGenderValid = "Missing Title"
        Exit Function
    End If
        
    'Gender presence
    If Gender = "" Then
        IsTitleGenderValid = "Missing Gender"
        Exit Function
    End If
            
    'Gender Valid
    If Gender <> "M" And Gender <> "Male" And Gender <> "F" And Gender <> "Female" Then
        IsTitleGenderValid = "Gender: Invalid Value"
        Exit Function
    End If
        
    'Male Title with Female Check
    If (Title = "MR" Or Title = "SIR" Or Title = "Rev" Or Title = "Lord" Or Title = "CAPT") And (Gender = "F" Or Gender = "Female") Then
        IsTitleGenderValid = "Gender/Title: Inconsistent Gender and Title"
        Exit Function
    End If
    
    'Female Title with Male Check
    If (Title = "MS" Or Title = "Miss" Or Title = "MRS") And (Gender = "M" Or Gender = "Male") Then
        IsTitleGenderValid = "Gender/Title: Inconsistent Gender and Title"
        Exit Function
    End If
    
    'Misc other titles
    'Dr / Cdr / EUR.ING
    
	If IsTitleGenderValid <> "" Then Exit Function
	IsTitleGenderValid = "Gender/Title: Pass"

	End Function

Private Function IsSurnameValid(ByVal Surname As String)

    'Presence Check
    If Surname = "" Then
        IsSurnameValid = "Missing Surname"
        Exit Function
    End If


    'Format Check:
	If Surname Like "*[0-9,!,£,$,%,^,&,*,(,),+,=,_,`,¬,\,|,<,>,?,/,'.,:,#,~,[,{,}]*" Then
		IsSurnameValid = "Surname: Invalid Format"
		Exit Function
	End If


        
	If IsSurnameValid <> "" Then Exit Function
	IsSurnameValid = "Surname: Pass"

End Function

Private Function IsDOBValid(ByVal DOB As String)
	''' Dates in culture insensitive format will still pass checks!
	
    Dim ParsedDOB As Date

    'Presence Check
    If DOB = "" Then
        IsDOBValid = "Missing DOB"
        Exit Function
    End If

    On Error Resume Next
    ParsedDOB = DateValue(DOB)
    On Error GoTo 0

    'Format Check:
    If Not IsDate(ParsedDOB) Or ParsedDOB = 0 Then
        IsDOBValid = "DOB: Invalid Format for Date"
        Exit Function
    End If

    'Future Date Check:
    If ParsedDOB > Now Then
        IsDOBValid = "DOB: Future Date"
        Exit Function
    End If
        
	If IsDOBValid <> "" Then Exit Function
	IsDOBValid = "DOB: Pass"

End Function
