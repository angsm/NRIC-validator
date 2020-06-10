Attribute VB_Name = "Validate_FIN_NRIC"
Sub Validate_FIN_NRIC()
    Dim isLocal     As Boolean
    Dim isForeign     As Boolean
    Dim colArr
    Dim msgStore() As String
    Dim startRow As Integer
    
    ReDim Preserve msgStore(0)
    startRow = 28
    msgStore(0) = "Kindly check the invalid FIN/NRIC below"
    
    'Update columns that needs to be formatted
    colArr = Array("D")
    
    For Each col In colArr 'Iterate colArr letters
        Range(col & startRow).Select
        ' Set Do loop to stop when an empty cell is reached.
        Do Until IsEmpty(ActiveCell)
            
            If is_valid(ActiveCell.Value) = False Then
                ReDim Preserve msgStore(UBound(msgStore) + 1)
                    msgStore(UBound(msgStore)) = ActiveCell.Value() + " -- " + ActiveCell.Address
            End If
            
            ActiveCell.Offset(1, 0).Select 'Move to next cell down
        Loop
    Next col 'Move to next column
    
    If (UBound(msgStore) - LBound(msgStore) + 1) <= 1 Then 'count in the first static text
        MsgBox ("All FIN/NRIC are valid :)")
    Else
        MsgBox Join(msgStore, vbCrLf)
    End If

End Sub

Function is_valid(ByVal nricOrFin As String) As Boolean

    is_valid = is_nric_valid(nricOrFin)
    If is_valid = False Then
        is_valid = is_fin_valid(nricOrFin)
    Else
        is_valid = True
    End If

End Function

Function is_nric_valid(ByVal fin As String) As Boolean

    'Check if activecell length is empty
    If Trim(fin & vbNullString) = vbNullString Then
        is_nric_valid = False
    End If
    
    'Check length of NRIC
    If Len(fin) <> 9 Then
        is_nric_valid = False
        Exit Function
    End If
    
    'Extract first and last character
    first = Left(fin, 1)
    last = Right(fin, 1)
    
    'First character is S or T
    If first <> "S" And first <> "T" Then
        is_nric_valid = False
        Exit Function
    End If
    
    'Extract only the numerics
    If IsNumeric(Mid(fin, 2, 7)) Then
        numeric = CLng(Mid(fin, 2, 7))
    Else
        is_nric_valid = False
        Exit Function
    End If
    
    Dim postfixes
    If first = "S" Then
        postfixes = Array("J", "Z", "I", "H", "G", "F", "E", "D", "C", "B", "A")
    Else
        postfixes = Array("G", "F", "E", "D", "C", "B", "A", "J", "Z", "I", "H")
    End If
    
    is_nric_valid = check_mod_11(last, numeric, postfixes)
        

End Function

Function is_fin_valid(ByVal fin As String) As Boolean

    'Check if activecell length is empty
    If Trim(fin & vbNullString) = vbNullString Then
        is_fin_valid = False
        Exit Function
    End If
    
    'Check length of NRIC
    If Len(fin) <> 9 Then
        is_fin_valid = False
        Exit Function
    End If
    
    'Extract first and last character
    first = Left(fin, 1)
    last = Right(fin, 1)
    
    'First character is F or G
    If first <> "F" And first <> "G" Then
        is_fin_valid = False
        Exit Function
    End If
    
    'Extract only the numerics
    If IsNumeric(Mid(fin, 2, 7)) Then
        numeric = CLng(Mid(fin, 2, 7))
    Else
        is_fin_valid = False
        Exit Function
    End If
    
    Dim postfixes
    If first = "F" Then
        postfixes = Array("X", "W", "U", "T", "R", "Q", "P", "N", "M", "L", "K")
    Else
        postfixes = Array("R", "Q", "P", "N", "M", "L", "K", "X", "W", "U", "T")
    End If
    
    is_fin_valid = check_mod_11(last, numeric, postfixes)
        
End Function

Function check_mod_11(ByVal last As String, ByVal numeric As Long, ByVal postfixes As Variant) As String

    Dim total As Integer
    Dim count As Integer
    Dim weights()
    weights = Array(2, 7, 6, 5, 4, 3, 2)
    
    total = 0
    count = 0
    
    'Mod and divide until numeric is zero
    Do Until (numeric = 0)
        total = total + (numeric Mod 10) * weights((UBound(weights) - LBound(weights) + 1) - (1 + count))
       
        count = count + 1
        numeric = Application.WorksheetFunction.Floor_Math(numeric / 10)
      
    Loop

    'If last character is sample as caluclated character, then IC is valid
    If StrComp(last, postfixes(total Mod 11)) = 0 Then
        check_mod_11 = True
    Else
        check_mod_11 = False
    End If
        

End Function
