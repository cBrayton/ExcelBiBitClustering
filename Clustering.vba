Public mods As New Collection
Private lastRow As Integer
Private lastCol As Integer
Private startCol As Integer
Dim MatrixSheet As Worksheet
'Private N As Integer
Private seen() As Boolean 'Boolean array listing the columns seen
'Write to File not implemented yet. Code is in, but conditionals need to be added to switch based on boolean.
Private writeToFile As Boolean 'Boolean to determine whether to write to a file or put the modules straight into an excel sheet
'If writeToFile is true then the program will run until it completes and make an arbitrarily large csv file.
'If writeToFile is false then the program will output the created modules to an excel sheet and crash after just over 1 million modules.

'This method is now worse than BiBit CreateModules, and they have the same functionality.
Sub Main()
Dim mmod As pMod
Dim modSheet As Worksheet
Set MatrixSheet = ActiveSheet
placeMods = MsgBox("Do you want the program to break the data into modules? (This will replace the original data with *'s and create lists of courses to be modules. THIS CANNOT BE UNDONE.)", vbYesNoCancel, "Easy Mode?")
If placeMods = vbCancel Then
    Exit Sub
End If
box = MsgBox("This Macro assumes your data has headers for both columns and rows. Analysis begins at cell 2,2; however the next boxes need the absolute row and column ends of the data to be analyzed.", vbOKCancel, "Formatting Requirments")
If box = vbCancel Then
    Exit Sub
End If
'N = 0
'===========================================
' These values need to be updated based on the dataset before
' the program is run.
' ^ this is now done automatically through message boxes.
'===========================================
'lastRow = 34 'Uncommment these lines to manually assign values
'lastCol = 33
r = InputBox("Input the last row number to analyze.", "Last Row Number", 1)
If r = "" Then
    Exit Sub
End If
c = InputBox("Input the last column number to analyze. The number is require not the letters shown.", "Last Column Number", 1)
If c = "" Then
    Exit Sub
End If
lastRow = CInt(r)
lastCol = CInt(c)
'===========================================
' These values need to be updated based on the dataset before
' the program is run.
'===========================================
startTime = Timer
ReDim seen(lastCol)
seen(2) = True
'==============================
'The two blocks below change the output type of the program.
'The upperblock writes the output to a file and the lower block
' adds the input to a new excel worksheet in the current book.
'The file can get arbitrarily large (although will need to be opened
' in something other than excel, and the new worksheet will crash
' if the dataset is too large.
'===============================
'Dim FilePath As String
'FilePath = Application.DefaultFilePath & "\ModulesCreated.csv"
'Open FilePath For Output As #1
Set mmod = CreateModules(2)
'N = N + 1
'Write #1, mmod.PrintMod(); mmod.Size(); N
'=====================================
On Error Resume Next
mods.Add mmod, mmod.PrintMod()
On Error GoTo 0
'Set currentSht = Sheets.Add
''currentSht.name = "Compiled Comments"
Set modSheet = Sheets.Add
For i = 1 To mods.Count
    modSheet.Unprotect
    modSheet.Cells(i, 1).Value = mods(i).PrintMod()
    modSheet.Unprotect
    modSheet.Cells(i, 2).Value = mods(i).size()
Next
a = Range(Cells(1, 1), Cells(mods.Count, 2)).Sort(Columns(2), xlDescending)
Columns(1).AutoFit
'=======================================

'=======================================
'This portion of the code calls functions/subs
' from the other modules to completely finish
' the module creation process (create all possible,
' remove overlaps, then extract and group the modules
' from the source matrix).
'========================================
'Dim modSheet As Worksheet
'Set modSheet = Sheets.Add
'With modSheet.QueryTables.Add("TEXT;" & FilePath, modSheet.Cells(1, 1))
'      .FieldNames = True
'      .RowNumbers = False
'      .FillAdjacentFormulas = False
'      .RefreshOnFileOpen = False
'      .BackgroundQuery = True
'      .RefreshStyle = xlInsertDeleteCells
'      .SavePassword = False
'      .SaveData = True
'      .AdjustColumnWidth = True
'      .TextFilePromptOnRefresh = False
'      .TextFilePlatform = xlMacintosh
'      .TextFileStartRow = 1
'      .TextFileParseType = xlDelimited
'      .TextFileTextQualifier = xlTextQualifierDoubleQuote
'      .TextFileConsecutiveDelimiter = False
'      .TextFileTabDelimiter = True
'      .TextFileSemicolonDelimiter = False
'      .TextFileCommaDelimiter = False
'      .TextFileSpaceDelimiter = False
'      .TextFileOtherDelimiter = ","
'      .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
'      .Refresh BackgroundQuery:=False
'End With
modSheet.Activate
Dim startSize, endSize As Integer
Do
startSize = Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
modSheet.Range("A:B").RemoveDuplicates (1)
PickModules
endSize = Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
Loop While startSize <> endSize
endTime = Timer - startTime
Debug.Print endTime
If placeMods = vbYes Then
    PlaceModules MatrixSheet, modSheet, lastCol
End If
End Sub

Function CreateModules(Optional startCol)
'v Could be really usefull or break everything
'Application.Volatile
Dim bmod As pMod
Set bmod = New pMod
If IsMissing(startCol) Then
    startCol = 2
End If
For i = 2 To lastRow
    If Cells(i, startCol) = 1 Then
        bmod.CoursesIn (i)
    End If
Next
If bmod.ySize <> 0 Then
    bmod.EmployeesIn (startCol)
'    seen(startCol) = True
End If
Set bmod = expandModules(bmod)
Set CreateModules = bmod

End Function

Function expandModules(basemod As pMod)
Dim rCount As Integer
Dim rmod As New pMod
Set rmod = New pMod
Dim miss() As Integer
ReDim miss(lastRow)
Dim misses As Integer
Dim endCol As Integer
If basemod.EmployeesOut(0) > basemod.EmployeesOut(basemod.xSize() - 1) Then
    GoTo Part2
End If
startCol = basemod.EmployeesOut(basemod.xSize() - 1)
'Need a base case to control for final col searched
'Need a control for looping if first rows weren't searched
For j = startCol + 1 To lastCol
'    If seen(j) Then
'        GoTo NextIteration1
'    Else
    rCount = 0
    misses = 0
    ReDim miss(lastRow)
    For Each crs In basemod.y 'i = 2 To lastRow
        If crs > 0 Then
            If Cells(crs, j) = 1 Then
                rCount = rCount + 1
            Else
                miss(misses) = crs
                misses = misses + 1
            End If
        End If
    Next
    If rCount = 0 Then
       
    ElseIf rCount = basemod.ySize() And misses = 0 Then
        basemod.EmployeesIn (j)
        GoTo NextIteration1
    ElseIf basemod.ySize() <= 2 Then
        GoTo NextIteration1
    ElseIf rCount = 0 Then
        GoTo NextIteration1
    ElseIf rCount < basemod.ySize() Then
  '      rmod.CopyMod (basemod)
        Set rmod = basemod.CopyMod()
        For Each cr In miss
            If cr <> 0 Then
                rmod.RemoveMod (cr)
            End If
        Next
'        rmod.EmployeesIn (j)
        Set rmod = expandModules(rmod)
'        N = N + 1
'        Write #1, rmod.PrintMod(); rmod.Size(); N
        On Error Resume Next
        mods.Add rmod, rmod.PrintMod
        On Error GoTo 0
    End If
NextIteration1:
    If rCount < Application.WorksheetFunction.CountIf(Range(Cells(2, j), Cells(lastRow, j)), 1) And Not seen(j) Then
        seen(j) = True
        Set rmod = CreateModules(j)
'        N = N + 1
'        Write #1, rmod.PrintMod(); rmod.Size(); N
        On Error Resume Next
        mods.Add rmod, rmod.PrintMod
        On Error GoTo 0
    'ElseIf rCount = Application.WorksheetFunction.CountIf(Range(Cells(2, j), Cells(lastRow, j)), 1) Then
     '   seen(j) = True
'    End If
    End If
Next
'if all columns done then return else goto column one and work to start colum (pX(0))
Part2:
If basemod.EmployeesOut(0) = 2 Then
    Set expandModules = basemod
    Exit Function
Else
    If basemod.EmployeesOut(basemod.xSize() - 1) < lastCol Then
        startCol = basemod.EmployeesOut(basemod.xSize() - 1)
    Else
        startCol = 1
    End If
    endCol = basemod.EmployeesOut(0) - 1
    For j = startCol + 1 To endCol
'        If seen(j) Then
'            GoTo NextIteration2
'        Else
        rCount = 0
        misses = 0
        ReDim miss(lastRow)
        For Each crs In basemod.y 'i = 2 To lastRow
            If crs > 0 Then
                If Cells(crs, j) = 1 Then
                    rCount = rCount + 1
                Else
                    miss(misses) = crs
                    misses = misses + 1
                End If
            End If
        Next
        If rCount = 0 Then
       
        ElseIf rCount = basemod.ySize() And misses = 0 Then
            basemod.EmployeesIn (j)
            GoTo NextIteration2
        ElseIf basemod.ySize() <= 2 Or basemod.xSize() <= 2 Then
            GoTo NextIteration2
        ElseIf rCount < basemod.ySize() Then
  '      rmod.CopyMod (basemod)
            Set rmod = basemod.CopyMod()
            For Each cr In miss
                rmod.RemoveMod (cr)
            Next
'            rmod.EmployeesIn (j)
            Set rmod = expandModules(rmod)
'            N = N + 1
'            Write #1, rmod.PrintMod(); rmod.Size(); N
            On Error Resume Next
            mods.Add rmod, rmod.PrintMod
            On Error GoTo 0
        End If
'        If rCount = Application.WorksheetFunction.CountIf(Range(Cells(2, j), Cells(lastRow, j)), 1) Then
'            seen(j) = True
'        End If
'        End If
'        If rCount < Application.WorksheetFunction.CountIf(Columns(, j), 1) Then
'            rmod = CreateModules(j)
'            mods.Add rmod
'        End If
NextIteration2:
    Next
    Set expandModules = basemod
End If

'=====
   
'            For Each crs In tmod.pY
'                If crs = i And tmod.pX(tmod.xSize - 1) <> j Then
'                    tmod.Employees (j)
'                    break
'                End If
'            Next
'        End If
'        ' Need a remove employee method incase rCount is
'    Next
'    If rCount > tmod.ySize() Then
'            rmod = CreateModules(j)
'            mods.Add (rmod)
'    ElseIf rCount < tmod.ySize() Then
'        For Each crs In tmod.pY
'            If Cells(crs, j) = "" Or Cells(crs, j) = Null Then
'                rmod = Copy(tmod)
'                rmod.RemoveMod (crs)
'                rmod = expandModules(rmod)
'                mods.Add (rmod)
'            End If
'        Next
'    End If
'Next
End Function
