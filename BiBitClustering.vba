Sub BiBitCreateModules()
'Creates the potential modules about twice as fast, but creates a lot of equivalent (but not identical duplicates).
'The extra duplicates are removed with the built in remove duplicates method and a custom remove equivalents method.
'Back to back identical testArray's are checked for and skipped to help shorten run time and limit the number of duplicates.
'If all testArrays are stored then all duplicates could be ignored, but the memory cost would be much higher. (Memory seems to be a limiting factor on the current system, so this was not done.)
'If memory is no longer a factor and the data set is very large, then ignoring all duplicates would be usefull, since excel can only hold about 1 Million rows in a sheet (modSheet with the created modules could exceed this with large datasets.)
Set MatrixSheet = ActiveSheet
'Uncomment this to make msgbox and if statement to make module placement optional.
'placeMods = MsgBox("Do you want the program to break the data into modules? (This will replace the original data with *'s and create lists of courses to be modules. THIS CANNOT BE UNDONE. It is recommended to cancel and create a duplicate of the data just in case.)", vbYesNoCancel, "Easy Mode?")
'If placeMods = vbCancel Then
'    Exit Sub
'End If
box = MsgBox("This Macro assumes your data has headers for both columns and rows. Analysis begins at cell 2,2; however the next boxes need the absolute row and column ends of the data to be analyzed." & vbCrLf & "Does your data start at 2,2?", vbYesNoCancel, "Formatting Requirements")
If box = vbCancel Then
    Exit Sub
ElseIf box = vbNo Then
    resumePrevious = MsgBox("Are you resuming a previous create module macro? Only answer yes if the current sheet is a list of potential modules of the form 'Courses 1 2 3 4 Employees 2 3 4 5'.", vbYesNoCancel, "Resume Previous?")
    If resumePrevious = vbYes Then
        Set modSheet = ActiveSheet
        GoTo RemoveDuplicates
    Else
        Exit Sub
    End If
End If
r = InputBox("Input the last row number to analyze.", "Last Row Number", 2)
If r = "" Then
    Exit Sub
End If
c = InputBox("Input the last column number to analyze. The number is require not the letters shown, so column AA is actually 27.", "Last Column Number", 2)
If r = "" Or c = "" Then
    Exit Sub
End If
lastRow = CInt(r)
lastCol = CInt(c)
'Collapses the columns and rows into 16 bit binary integers
' and puts the collapsed matrix into a new worksheet
startTime = Timer
Set modSheet = Sheets.Add
Set CollapseRow = Sheets.Add
Set CollapseCol = Sheets.Add
Dim intArray() As Long
Dim test() As String
Dim Val As Long
Dim currentRow As Integer
Dim tmod As String
Dim size, numEmps As Integer
currentRow = 1
For i = 2 To lastCol
    intArray = ColBinToInt(i)
    j = 1
    For Each m In intArray
        CollapseRow.Cells(j, i - 1) = m
        j = j + 1
    Next
Next
For i = 2 To lastRow
    intArray = RowBinToInt(i)
    j = 1
    For Each m In intArray
        CollapseCol.Cells(i - 1, j) = m
        j = j + 1
    Next
Next

'Compare rows to each other. Find the common columns between two rows, then search every other row for the same columns in common.
'Problem comparing two rows, since the rows were compressed into batches of 16.
'Need to find a way to compare within a batch.
'Set k and k1 to cycle through original rows, and create binary numbers from those rows. ~ Test array is a binary array, across columns.
'Test each row bitwise against the test array column. Doesn't take advantage of the compression; could make use of column compression to speed this up a little
'Uses MatrixSheet instead of Collapse due to above comments. Indices start at 2 because of this too.
rowTime = Timer
For k = 1 To CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).row - 1
    ReDim testArray(CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).Column) As Long
    ReDim compArray(CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).Column) As Long
    For k1 = k + 1 To CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).row
        tmod = ""
        tmod = "Courses " & k + 1 & " " & k1 + 1
        size = 2
        For L = 1 To CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).Column 'MatrixSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
            testArray(L - 1) = CollapseCol.Cells(k, L) And CollapseCol.Cells(k1, L)
        Next
        x = True
        For c = LBound(testArray) To UBound(testArray)
            If compArray(c) <> testArray(c) Then
                x = False
                Exit For
            End If
        Next
        If x = True Then
            GoTo RowsNext
        End If
        compArray = testArray
        For a = 1 To CollapseCol.UsedRange.SpecialCells(xlCellTypeLastCell).row
            If a <> k And a <> k1 Then
                winner = True
                For L = 1 To UBound(testArray) 'change start to 2 if reverting.
                    'If winner And (testArray(L - 2) = MatrixSheet.Cells(a, L) Or (testArray(L - 2) = 0 And MatrixSheet.Cells(a, L) <> 0)) Then
                    If winner And ((testArray(L - 1) And CollapseCol.Cells(a, L)) = testArray(L - 1)) Then
                        winner = True
                    Else
                        winner = False
                    End If
                Next
                If winner Then
                    tmod = tmod & " " & a + 1
                    size = size + 1
                End If
            End If
        Next
       
        'New code to be tested, may speed up removing duplicate modules by 200x.
        'The new code block may not parse rows and columns the same depending on which it is comparing at the time. Employees and Courses may have different numbers depending on which are compared.
        test() = Split(Trim(tmod))
        QuickSort test(), 1, UBound(test)
        contents = test(0)
        current = 0
        Val = 0
        For Z = 1 To UBound(test)
            If current = (test(Z) - 2) \ 16 Then
                Val = 2 ^ ((test(Z) - 2) Mod 16) + Val
            Else
                If Val <> 0 Then
                    contents = contents & " " & Val
                End If
                current = (test(Z) - 2) \ 16
                Val = 2 ^ ((test(Z) - 2) Mod 16) + (current * 1000000) 'The current * 1 million allows for the block of 16 to be distiguished without losing the Val data.
            End If
        Next
        If Val <> 0 Then
            contents = contents & " " & Val
        End If
        tmod = contents
       
        'End of new code block.
       
        tmod = tmod + " Employees"
        numEmps = 0
'        For L = 0 To UBound(testArray)
'            If testArray(L) = 1 Then
'                tmod = tmod & " " & (L + 2)
'                numEmps = numEmps + 1
'            End If
'        Next
       
' Uncomment the block of code below to revert changes if the "new code block" is removed
'        For L = 0 To UBound(testArray)
'            For N = 0 To 15
'                If testArray(L) And (2 ^ N) Then
'                    tmod = tmod & " " & ((16 * L) + 16 - N + 1) 'fine
'                    numEmps = numEmps + 1
'                End If
'            Next
'        Next

        'New code block for faster remove duplicates
        For L = 0 To UBound(testArray)
            If testArray(L) <> 0 Then
                tmod = tmod & " " & testArray(L) + L * 1000000
                For N = 0 To 15
                    If testArray(L) And (2 ^ N) Then
                        numEmps = numEmps + 1
                    End If
                Next
            End If
        Next
       
        'End of new code block

        size = size * numEmps
        If size = 0 Then
            GoTo RowsNext
        End If
        modSheet.Unprotect
        modSheet.Cells(currentRow, 1).Value = tmod
        modSheet.Cells(currentRow, 2).Value = size
        modSheet.Cells(currentRow, 3).Value = numEmps
        currentRow = currentRow + 1
RowsNext:
    Next
    Debug.Print "Finished comparing row " & k & " out of " & lastRow & " at " & Timer - startTime
Next
Debug.Print "Finished comparing all rows in " & Timer - rowTime & "."
modSheet.Range("A:C").RemoveDuplicates (1)
currentRow = modSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row + 1
'Compares the columns to each other based on similarities in the rows.
colTime = Timer
For k = 1 To CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).Column - 1
    ReDim testArray(CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).row) As Long
    ReDim compArray(CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).row) As Long
    For k1 = k + 1 To CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).Column
        tmod = ""
        tmod = " Employees " & (k + 1) & " " & (k1 + 1)
        numEmps = 2
        For L = 1 To CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).row
            testArray(L - 1) = CollapseRow.Cells(L, k) And CollapseRow.Cells(L, k1)
        Next
        x = True
        For c = LBound(testArray) To UBound(testArray)
            If compArray(c) <> testArray(c) Then
                x = False
                Exit For
            End If
        Next
        If x = True Then
            GoTo ColumnsNext
        End If
        compArray = testArray
        For a = 1 To CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).Column
            If a <> k And a <> k1 Then
                winner = True
                For L = 1 To UBound(testArray)
                    If winner And ((testArray(L - 1) And CollapseRow.Cells(L, a)) = testArray(L - 1)) Then
                        winner = True
                    Else
                        winner = False
                    End If
                Next
                If winner Then
                    tmod = tmod & " " & (a + 1) 'fine
                    numEmps = numEmps + 1
                End If
            End If
        Next
       
        size = 0
       
       
'Uncomment this code block if removing the "new code block"
'        For L = 0 To UBound(testArray)
'            For N = 0 To 15
'                If testArray(L) And (2 ^ N) Then
'                    tmod = " " & ((16 * L) + N + 2) & tmod 'fine
'                    numEmps = numEmps + 1
'                End If
'            Next
'        Next
       
        'New code block to speed up removing duplicates.
       
        test() = Split(Trim(tmod))
        QuickSort test(), 1, UBound(test)
        contents = test(0)
        current = 0
        Val = 0
        For Z = 1 To UBound(test)
            If current = (test(Z) - 2) \ 16 Then
                Val = 2 ^ (15 - ((test(Z) - 2) Mod 16)) + Val
            Else
                If Val <> 0 Then
                    contents = contents & " " & Val
                End If
                current = (test(Z) - 2) \ 16
                Val = 2 ^ (15 - ((test(Z) - 2) Mod 16)) + (current * 1000000) 'The current * 1 million allows for the block of 16 to be distiguished without losing the Val data.
            End If
        Next
            If Val <> 0 Then
                contents = contents & " " & Val
            End If
        contents = " " & contents
       
        For L = 0 To UBound(testArray)
            If testArray(L) <> 0 Then
                tempcontents = tempcontents & " " & testArray(L) + L * 1000000
                For N = 0 To 15
                    If testArray(L) And (2 ^ N) Then
                        size = size + 1
                    End If
                Next
            End If
        Next
        tmod = tempcontents & contents
        tempcontents = ""
        'End of new code block
       
        tmod = "Courses" & tmod
        size = size * numEmps
        If size = 0 Then
            GoTo ColumnsNext
        End If
        modSheet.Unprotect
        modSheet.Cells(currentRow, 1).Value = tmod
        modSheet.Cells(currentRow, 2).Value = size
        modSheet.Cells(currentRow, 3).Value = numEmps ' an error here, maybe fixed
        currentRow = currentRow + 1
ColumnsNext:
    Next
    Debug.Print "Finished comparing column " & k & " out of " & CollapseRow.UsedRange.SpecialCells(xlCellTypeLastCell).Column - 1; " at "; Timer - startTime
Next
Debug.Print "Finished comparing all columns in " & Timer - colTime & "."
'a = modSheet.Range(Cells(1, 1), Cells(mods.Count, 2)).Sort(Columns(2), xlDescending)
modSheet.Columns(1).AutoFit
modSheet.Activate
a = Range(Cells(1, 1), Cells(currentRow, 3)).Sort(Columns(2), xlDescending)
ActiveWorkbook.Save

RemoveDuplicates:
Dim startSize, endSize As Integer
Do
    dupTime = Timer
    Do
        startSize = Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
        modSheet.Range("A:C").RemoveDuplicates (1)
        'RemoveEquivalents 'Not needed with the "new code blocks"
        endSize = Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
    Loop While startSize <> endSize
    rTime = Timer - dupTime
    Debug.Print "Inner loop finished in " & rTime
    pickTime = Timer
    tTime = tTime + rTime
    PickModules
    endSize = Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
    tTime = tTime + Timer - pickTime
    Debug.Print "This Iteration fininshed in " & tTime
Loop While startSize <> endSize
Debug.Print "Remove Duplicates finished in " & tTime
endTime = Timer - startTime
Debug.Print endTime
Application.DisplayAlerts = False
MatrixSheet.Copy After:=MatrixSheet
Set MatrixCopySheet = ActiveSheet
Application.DisplayAlerts = True
'Uncomment below if statement and first dialog box above to make module placement optional.
'If placeMods = vbYes Then
    PlaceModules MatrixCopySheet, modSheet, lastCol
'End If
Application.DisplayAlerts = False
CollapseRow.Delete
CollapseCol.Delete
modSheet.Delete
Application.DisplayAlerts = True
End Sub

Function ColBinToInt(col) As Long()

'Binary is bottom up, so the first row is the least significant (2^0) and the last row is the most significant (2^15)
Dim BinStr As String
ReDim DecArray(lastRow \ 16) As Long
Dim y As Integer
Dim Dec, temp As Long
For i = 2 To lastRow
    If MatrixSheet.Cells(i, col) <> 1 Then
        BinStr = "0" & BinStr
    Else
        BinStr = "1" & BinStr
    End If
    If (i - 1) Mod 16 = 0 Then
        Dec = 0
        temp = 0
        For x = 1 To 16
            temp = Val(Right(BinStr, 1))
            BinStr = Left(BinStr, 16 - x)
            If temp <> "0" Then
                Dec = Dec + (2 ^ (x - 1))
            End If
        Next
        DecArray(y) = Dec
        y = y + 1
    End If
Next
If BinStr <> "" Then
    Dec = 0
    temp = 0
    Dim miss As Integer
    miss = 16 - Len(BinStr)
    BinStr = WorksheetFunction.Rept("0", miss) + BinStr 'switched BinStr to after 0's to account for bottum up approach.
    For x = 1 To 16
        temp = Val(Right(BinStr, 1))
        BinStr = Left(BinStr, 16 - x)
        If temp <> "0" Then
            Dec = Dec + (2 ^ (x - 1))
        End If
    Next
    DecArray(y) = Dec
    y = y + 1
End If
ColBinToInt = DecArray
End Function

Function RowBinToInt(row) As Long()

'Binary is left to right, so the first column is the most significant (2^15) and the last column is the most significant (2^0)
Dim BinStr As String
ReDim DecArray(lastCol \ 16) As Long
Dim y As Integer
Dim Dec, temp As Long
For i = 2 To lastCol
    If MatrixSheet.Cells(row, i) <> 1 Then
        BinStr = BinStr & "0"
    Else
        BinStr = BinStr & "1"
    End If
    If (i - 1) Mod 16 = 0 Then
        Dec = 0
        temp = 0
        For x = 1 To 16
            temp = Val(Right(BinStr, 1))
            BinStr = Left(BinStr, 16 - x)
            If temp <> "0" Then
                Dec = Dec + (2 ^ (x - 1))
            End If
        Next
        DecArray(y) = Dec
        y = y + 1
    End If
Next
If BinStr <> "" Then
    Dec = 0
    temp = 0
    Dim miss As Integer
    miss = 16 - Len(BinStr)
    BinStr = BinStr + WorksheetFunction.Rept("0", miss)
    For x = 1 To 16
        temp = Val(Right(BinStr, 1))
        BinStr = Left(BinStr, 16 - x)
        If temp <> "0" Then
            Dec = Dec + (2 ^ (x - 1))
        End If
    Next
    DecArray(y) = Dec
    y = y + 1
End If
RowBinToInt = DecArray
End Function


Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
'Taken from http://stackoverflow.com/questions/152319/vba-array-sort-function
'Modified to interpret the numeric strings as numbers with casting
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = CInt(vArray((inLow + inHi) \ 2))

  While (tmpLow <= tmpHi)

     While (CInt(vArray(tmpLow)) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < CInt(vArray(tmpHi)) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = CInt(vArray(tmpLow))
        vArray(tmpLow) = CInt(vArray(tmpHi))
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub