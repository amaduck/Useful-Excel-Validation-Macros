
Function checkColumnHeadings(sheetToCheck, rowToCheck, firstColumn, headingsArray)
    
    ' Checks the column headings in a specified sheet against a passed array
    ' Useful check in case sheet changed from original configuration
    ' Works on contiguous columns - can enter blank value to bridge gaps

    ' Returns a boolean based on whether user wants to continue or not
    ' To return boolean simply based on matching, remove if notMatching statement entirely, and replace with checkColumnHeadings = notMatching
    
    checkColumnHeadings = True

    expectedColumnNames = headingsArray
    Dim actualColumnNames(9) As String
        
    Sheets(sheetToCheck).Select
    
    columnNumber = firstColumn
    notMatching = False
    For Each ColumnName In expectedColumnNames
        If Cells(rowToCheck, columnNumber) <> ColumnName Then notMatching = True
        actualColumnNames(columnNumber - 1) = Cells(rowToCheck, columnNumber)
        columnNumber = columnNumber + 1
    Next
    
    If notMatching Then
        expectedString = ""
        actualString = ""
        comparisonString = "Expected" & vbTab & vbTab & vbTab & "Actual" & vbTab & vbTab & vbTab & "Match?" & vbNewLine
        For x = 0 To 9
            
            headingsMatch = "Yes"
            If expectedColumnNames(x) <> actualColumnNames(x) Then headingsMatch = "No"
            
            ' Spacing for three columns, string lengths > 27 will affect formatting
            firstStringLength = Len(expectedColumnNames(x))
            tabRepeats = 3 - Int(firstStringLength / 9)
            firstSpacing = ""
            For increaseSpacing = 1 To tabRepeats
                firstSpacing = firstSpacing & vbTab
            Next increaseSpacing
            secondStringLength = Len(actualColumnNames(x))
            tabRepeats = 3 - Int(secondStringLength / 9)
            secondSpacing = ""
            For increaseSpacing = 1 To tabRepeats
                secondSpacing = secondSpacing & vbTab
            Next increaseSpacing
            
            comparisonString = comparisonString & expectedColumnNames(x) & firstSpacing & actualColumnNames(x) & secondSpacing & headingsMatch & vbNewLine

        Next x
        
        continue = MsgBox("Column headings in SheetList have been changed. If the column contents haven't been changed, it's ok to run the macro. " & _
            "Otherwise, the macro will not work correctly. To avoid this message in future, change the expectedColumnNames values in function checkColumnHeadings. " _
            & "Expected and actual values shown below for comparison. " _
            & vbNewLine & vbNewLine & comparisonString _
            & vbNewLine & vbNewLine & "Do you want to continue with run?", vbYesNo, "Warning!")
        If continue <> vbYes Then checkColumnHeadings = False
    End If

End Function
