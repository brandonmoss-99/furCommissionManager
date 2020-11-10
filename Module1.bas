Attribute VB_Name = "Module1"
'Written by Brandon Moss
Sub RectangleRoundedCorners1v2_Click()
    'input mode
    If Worksheets("Frontend").Range("D3") = "Input" Then
        Dim CommLastRow As Long, CustLastRow As Long
        'find next free row in CommRecord and CustRecord
        CommLastRow = Worksheets("CommRecord").Range("A" & Rows.Count).End(xlUp).Row + 1
        CustLastRow = Worksheets("CustRecord").Range("A" & Rows.Count).End(xlUp).Row + 1
        
        'If Required fields aren't empty
        If Not IsEmpty(Worksheets("Frontend").Cells(5, 4).Value) And Not IsEmpty(Worksheets("Frontend").Cells(6, 4).Value) And Not IsEmpty(Worksheets("Frontend").Cells(7, 4).Value) And Not IsEmpty(Worksheets("Frontend").Cells(8, 4).Value) Then
            'holds the CustID for a matched name input or the error if doesn't exist in CustRecord
            Dim matchedCustID
            'set to the CustID for a matched name input or the error if doesn't exist in CustRecord
            matchedCustID = Application.Index(Sheets("CustRecord").Range("A:A"), Application.Match(Sheets("Frontend").Cells(5, 4).Value, Sheets("CustRecord").Range("B:B"), 0))
            
            'If the input name wasn't found in CustRecord
            If IsError(matchedCustID) = True Then
                'make new row for new name, and create ID for that row
                If CustLastRow > 2 Then
                    Worksheets("CustRecord").Cells(CustLastRow, 1).Value = Worksheets("CustRecord").Cells(CustLastRow - 1, 1).Value + 1
                Else
                    Worksheets("CustRecord").Cells(CustLastRow, 1).Value = 1
                End If
                'add inputted name (CustName) to CustRecord associated with new CustID just created
                Worksheets("CustRecord").Cells(CustLastRow, 2).Value = Worksheets("Frontend").Cells(5, 4).Value
                matchedCustID = Worksheets("CustRecord").Cells(CustLastRow, 1).Value
            End If
            
            'make new row for new commission in CommRecord
            If CommLastRow > 2 Then
                Worksheets("CommRecord").Cells(CommLastRow, 1).Value = Worksheets("CommRecord").Cells(CommLastRow - 1, 1).Value + 1
            Else
                Worksheets("CommRecord").Cells(CommLastRow, 1).Value = 1
            End If
            
            'Add CustID of inputted name to CommRecord
            Worksheets("CommRecord").Cells(CommLastRow, 2).Value = matchedCustID
            
            'fill in CommRecord
            Dim i As Integer
            For i = 6 To 13
            If IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False Then
                Worksheets("CommRecord").Cells(CommLastRow, i - 3).Value = Worksheets("Frontend").Cells(i, 4).Value
            End If
            Next i
            
            'fill in CustRecord
            'holds the location in CustRecord where ID wanting to add/update contact info is
            Dim CustIDLocation
            'find and store the location of the ID wanting to add/update contact info for
            CustIDLocation = Application.Match(matchedCustID, Sheets("CustRecord").Range("A:A"), 0)
            For i = 15 To 18
            'if frontend has contact info input and there's no contact info in CustRecord for the given CustID
            If IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) = True Then
                Worksheets("CustRecord").Cells(CustIDLocation, i - 12).Value = Worksheets("Frontend").Cells(i, 4).Value
            'if frontend has contact info input but there's already info in CustRecord for the given CustID
            ElseIf IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) = False Then
                dataMismatch = MsgBox("Oops, " + Worksheets("Frontend").Cells(i, 3).Text + " for: " + Worksheets("CustRecord").Cells(CustIDLocation, 2).Text + " is already set as: " + Worksheets("CustRecord").Cells(CustIDLocation, i - 12) + vbNewLine + "Change to: " + Worksheets("Frontend").Cells(i, 4).Text + "?", vbYesNo, "Input: Data mismatch/Redundant Entry")
                
                'User presses yes to MsgBox
                If dataMismatch = 6 Then
                    Worksheets("CustRecord").Cells(CustIDLocation, i - 12).Value = Worksheets("Frontend").Cells(i, 4).Value
                End If
            End If
            Next i
            
            'let user know new commission added
            MsgBox "New Commission ID: " + Worksheets("CommRecord").Cells(CommLastRow, 1).Text + " created for: " + Worksheets("Frontend").Cells(5, 4).Text, vbInformation, "New Commission Confirmation"
            Worksheets("Frontend").Range("D5", "D18").Clear
        Else
            'Give user an error that required fields are empty
            MsgBox Worksheets("Frontend").Cells(5, 3).Text + "/" + Worksheets("Frontend").Cells(6, 3).Text + "/" + Worksheets("Frontend").Cells(7, 3).Text + "/" + Worksheets("Frontend").Cells(8, 3).Text + " inputs are required!", vbCritical
        End If
    End If
    
    
    
    'Update User mode
    If Worksheets("Frontend").Range("D3") = "Update User" Then
        'if required fields filled in
        If (Not IsEmpty(Worksheets("Frontend").Cells(5, 4).Value)) And (Not IsEmpty(Worksheets("Frontend").Cells(15, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(16, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(17, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(18, 4).Value)) Then
            
            'set to the CustID for a matched name input or the error if doesn't exist in CustRecord
            matchedCustID = Application.Index(Sheets("CustRecord").Range("A:A"), Application.Match(Sheets("Frontend").Cells(5, 4).Value, Sheets("CustRecord").Range("B:B"), 0))
            
            'If the input name was found in CustRecord
            If IsError(matchedCustID) = False Then
                
                CustIDLocation = Application.Match(matchedCustID, Sheets("CustRecord").Range("A:A"), 0)
                For i = 15 To 18
                    'if frontend has contact info input and there's no contact info in CustRecord for the given CustID
                    If IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) = True Then
                        Worksheets("CustRecord").Cells(CustIDLocation, i - 12).Value = Worksheets("Frontend").Cells(i, 4).Value
                    'if frontend has contact info input but there's already info in CustRecord for the given CustID
                    ElseIf IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) = False Then
                        dataMismatch = MsgBox("Confirm " + Worksheets("Frontend").Cells(i, 3).Text + " update for: " + Worksheets("CustRecord").Cells(CustIDLocation, 2).Text + vbNewLine + "From: '" + Worksheets("CustRecord").Cells(CustIDLocation, i - 12) + "'" + vbNewLine + "To: '" + Worksheets("Frontend").Cells(i, 4).Text + "'?", vbYesNo, "Input: Data mismatch/Redundant Entry")
                        
                        'User presses yes to MsgBox
                        If dataMismatch = 6 Then
                            Worksheets("CustRecord").Cells(CustIDLocation, i - 12).Value = Worksheets("Frontend").Cells(i, 4).Value
                        End If
                    End If
                Next i
                
                'Let user know updated successfully
                MsgBox "Info Updated Successfully!", vbInformation, "Update User Success"
                Worksheets("Frontend").Range("D5", "D18").Clear
                
            Else 'if input name wasn't found in CustRecord
                MsgBox "Couldn't find Customer: " + Worksheets("Frontend").Cells(5, 4).Text + " in Customer Record!" + vbNewLine + "Please check spelling & try again!", vbCritical, "Couldn't match info"
            End If
            
        Else 'if missing information
            MsgBox Worksheets("Frontend").Cells(3, 4).Text + " is missing info. Please add the required info", vbCritical, "Missing info"
        End If
    End If
    
    
    
    
    'Update Commission mode
    If Worksheets("Frontend").Range("D3") = "Update Commission" Then
        'if required fields filled in
        If (Not IsEmpty(Worksheets("Frontend").Cells(5, 4).Value)) And (Not IsEmpty(Worksheets("Frontend").Cells(6, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(7, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(8, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(9, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(10, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(11, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(12, 4).Value) Or Not IsEmpty(Worksheets("Frontend").Cells(13, 4).Value)) Then
    
            'holds the CommID for a matched name input or the error if doesn't exist in CustRecord
            Dim CommIDLocation
            CommIDLocation = Application.Match(Sheets("Frontend").Cells(5, 4).Value, Sheets("CommRecord").Range("A:A"), 0)
            
            'If the input name was found in CustRecord
            If IsError(CommIDLocation) = False Then
                
                For i = 6 To 13
                    'if frontend has comm info input and there's no comm info in CommRecord for the given CommID
                    If IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CommRecord").Cells(CommIDLocation, i - 3)) = True Then
                        Worksheets("CommRecord").Cells(CommIDLocation, i - 3).Value = Worksheets("Frontend").Cells(i, 4).Value
                    'if frontend has comm info input but there's already info in CommRecord for the given CommID
                    ElseIf IsEmpty(Worksheets("Frontend").Cells(i, 4).Value) = False And IsEmpty(Worksheets("CommRecord").Cells(CommIDLocation, i - 3)) = False Then
                        dataMismatch = MsgBox("Confirm " + Worksheets("Frontend").Cells(i, 3).Text + " update for: " + Worksheets("CommRecord").Cells(CommIDLocation, 1).Text + vbNewLine + "From: '" + Worksheets("CommRecord").Cells(CommIDLocation, i - 3).Text + "'" + vbNewLine + "To: '" + Worksheets("Frontend").Cells(i, 4).Text + "'?", vbYesNo, "Input: Data mismatch/Redundant Entry")
                        
                        'User presses yes to MsgBox
                        If dataMismatch = 6 Then
                            Worksheets("CommRecord").Cells(CommIDLocation, i - 3).Value = Worksheets("Frontend").Cells(i, 4).Value
                        End If
                    End If
                Next i
                
                'Let user know updated successfully
                MsgBox "Info Updated Successfully!", vbInformation, "Update Commission Success"
                Worksheets("Frontend").Range("D5", "D18").Clear
                
            Else 'if input name wasn't found in CommRecord
                MsgBox "Couldn't find Comm: " + Worksheets("Frontend").Cells(5, 4).Text + " in Commission Record!" + vbNewLine + "Please check spelling & try again!", vbCritical, "Couldn't match info"
            End If
                
        Else 'if missing information
            MsgBox Worksheets("Frontend").Cells(3, 4).Text + " is missing info. Please add the required info", vbCritical, "Missing info"
        End If
    End If
    
    'Search User Mode
    If Worksheets("Frontend").Range("D3") = "Search User" Then
        'if required cells filled in
        If Not IsEmpty(Worksheets("Frontend").Cells(5, 4).Value) Then
        'set to the CustID for a matched name input or the error if doesn't exist in CustRecord
            matchedCustID = Application.Index(Sheets("CustRecord").Range("A:A"), Application.Match(Sheets("Frontend").Cells(5, 4).Value, Sheets("CustRecord").Range("B:B"), 0))
            'If the input name was found in CustRecord
            If IsError(matchedCustID) = False Then
                'get location in CustRecord of input name
                CustIDLocation = Application.Match(matchedCustID, Sheets("CustRecord").Range("A:A"), 0)
                'return ID of name searched
                Worksheets("Frontend").Cells(6, 4) = Worksheets("CustRecord").Cells(CustIDLocation, 1).Value
                'return how many times the name searched has commissioned before
                Dim pastCommTotal
                pastCommTotal = Application.CountIf(Worksheets("CommRecord").Range("B:B"), matchedCustID)
                Worksheets("Frontend").Cells(7, 4).Value = pastCommTotal
                
                'return avg experience for name searched
                Dim totalExperience
                totalExperience = Application.SumIf(Worksheets("CommRecord").Range("B:B"), matchedCustID, Worksheets("CommRecord").Range("J:J"))
                Worksheets("Frontend").Cells(8, 4).Value = totalExperience / pastCommTotal

                'return social media of name searched
                For i = 15 To 18
                    If Not IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) Then
                        Worksheets("Frontend").Cells(i, 4) = Worksheets("CustRecord").Cells(CustIDLocation, i - 12).Value
                    End If
                Next i
            Else 'if name couldn't be found in CustRecord
                MsgBox "Couldn't find Customer: " + Worksheets("Frontend").Cells(5, 4).Text + " in Customer Record!" + vbNewLine + "Please check spelling & try again!", vbCritical, "Couldn't match info"
            End If
        Else 'if not all required cells filled in
            MsgBox Worksheets("Frontend").Cells(3, 4).Text + " is missing info. Please add the required info", vbCritical, "Missing info"
        End If
    End If
    
    
    'Search Commission Mode
    If Worksheets("Frontend").Range("D3") = "Search Commission" Then
        'if required cells filled in
        If Not IsEmpty(Worksheets("Frontend").Cells(5, 4).Value) Then
            CommIDLocation = Application.Match(Sheets("Frontend").Cells(5, 4).Value, Sheets("CommRecord").Range("A:A"), 0)
            'If the input ID was found in CommRecord
            If IsError(CommIDLocation) = False Then
                'return comm info of ID searched
                For i = 6 To 13
                    If Not IsEmpty(Worksheets("CommRecord").Cells(CommIDLocation, i - 3)) Then
                        Worksheets("Frontend").Cells(i, 4) = Worksheets("CommRecord").Cells(CommIDLocation, i - 3).Value
                    End If
                Next i
                
                
                'return name & socials of commissioner of ID searched
                Dim TempID
                TempID = Application.Index(Worksheets("CommRecord").Range("B:B"), Application.Match(Worksheets("Frontend").Cells(5, 4).Value, Worksheets("CommRecord").Range("A:A"), 0))
                CustIDLocation = Application.Match(TempID, Worksheets("CustRecord").Range("A:A"), 0)
                
                For i = 14 To 18
                    If Not IsEmpty(Worksheets("CustRecord").Cells(CustIDLocation, i - 12)) Then
                        Worksheets("Frontend").Cells(i, 4) = Worksheets("CustRecord").Cells(CustIDLocation, i - 12)
                    End If
                Next i
                
            Else 'if name couldn't be found in CommRecord
                MsgBox "Couldn't find Commission: " + Worksheets("Frontend").Cells(5, 4).Text + " in Commission Record!" + vbNewLine + "Please check spelling & try again!", vbCritical, "Couldn't match info"
            End If
        Else 'if not all required cells filled in
            MsgBox Worksheets("Frontend").Cells(3, 4).Text + " is missing info. Please add the required info", vbCritical, "Missing info"
        End If
    
    
    End If
    
    
End Sub
