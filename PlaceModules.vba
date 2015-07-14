Option Explicit
Sub PlaceModules(source, mSheet, numRoles)

Dim rowCount, startRow, numCourses As Integer
Dim rng As Range
Dim modSheet As Worksheet
Dim sourceSheet As Worksheet
Dim destSheet As Worksheet
'Dim course As Range
Dim i, j, k, N, current, crs, emp As Integer
Dim dmod() As String
Dim a As Boolean
Dim color As Boolean
'===================================
'Change these sheet numbers to match the current data set.
'modSheet is the sheet generated by Createmodules, and trimmed by PickModules
'sourceSheet is the original data sheet that the modules are generated from
'===================================
'modSheet = Sheet36.Name
'sourceSheet = Sheet41.Name
'destSheet = Sheet40.Name
Set modSheet = mSheet
Set sourceSheet = source
'===================================
Set destSheet = Sheets.Add
rowCount = 2
'numRoles = Worksheets(sourceSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).Column
startRow = 2
Worksheets(sourceSheet.Name).Rows(1).Copy
Worksheets(destSheet.Name).Rows(1).PasteSpecial (xlPasteValues)
Worksheets(destSheet.Name).Rows(1).PasteSpecial (xlPasteFormats)
color = True
'Loops through all the modules in the modSheet created by CreateModules and reduced by PickModules
For i = 1 To Worksheets(modSheet.Name).UsedRange.SpecialCells(xlCellTypeLastCell).row
    dmod = Split(Trim(Worksheets(modSheet.Name).Cells(i, 1)))
    numCourses = 1
    'Counts the number of courses in the module
    Do While dmod(numCourses + 1) <> "Employees"
        numCourses = numCourses + 1
    Loop
    'Loops through all the courses in the module, and copies the assignments for each course to a new worksheet.
    For j = 1 To numCourses
'        course = Worksheets(sourceSheet.Name).Rows(dmod(j))
       
        'New code to decompress courses
        Dim courses As Integer
        Dim cmod(16) As Integer
        courses = 0
        current = dmod(j) \ 1000000
        dmod(j) = dmod(j) Mod 1000000
        For N = 0 To 15
            If dmod(j) And (2 ^ N) Then
                cmod(courses) = ((16 * current) + N + 2) 'fine
                courses = courses + 1
            End If
        Next
        Dim modStart As Integer
        modStart = rowCount
        For crs = 0 To courses - 1
            Worksheets(sourceSheet.Name).Rows(cmod(crs)).Copy
            Worksheets(destSheet.Name).Rows(rowCount).PasteSpecial (xlPasteValues)
            Worksheets(destSheet.Name).Rows(rowCount).PasteSpecial (xlPasteFormats)
            rowCount = rowCount + 1
        Next
        'End of New Code block
       
'        Worksheets(sourceSheet.Name).Rows(dmod(j)).Copy (Worksheets(destSheet.Name).Rows(rowCount))
        'Loops through all the employees in the module and replaces them with * in the original matrix and changes them to 2's in the module matrix.
        For k = numCourses + 2 To UBound(dmod)
           
            'New code to decompress employees
            Dim numEmps As Integer
            Dim tmod(16) As Integer
            numEmps = 0
'            current = dmod(k) \ 1000000
'            dmod(k) = dmod(k) Mod 1000000
            For N = 0 To 15
                If dmod(k) Mod 1000000 And (2 ^ N) Then
                    tmod(numEmps) = ((16 * (dmod(k) \ 1000000)) + 16 - N + 1) 'fine
                    numEmps = numEmps + 1
                End If
            Next
            rowCount = modStart
            For crs = 0 To courses - 1
                For emp = 0 To numEmps - 1
                'Weird issue here where the dest sheet is populating correctly, but the sourceSheet isn't.
                'Probably something with cmod(crs), but not sure what. Only first row out of each 16 is updating.
                    Worksheets(sourceSheet.Name).Cells(cmod(crs), tmod(emp)).Value = "*"
                    Worksheets(destSheet.Name).Cells(rowCount, tmod(emp)) = "2"
                Next
                rowCount = rowCount + 1
            Next
            'End of New code block
                       
'            Worksheets(sourceSheet.Name).Cells(CInt(dmod(j)), CInt(dmod(k))).Value = "*"
'            Worksheets(destSheet.Name).Cells(rowCount, CInt(dmod(k))) = "2"
        Next
'Below line not needed with compressed data
'        rowCount = rowCount + 1
    Next
    destSheet.Activate
    'Replaces the 1's in the module matrix with *'s then turns the 2's into 1's to represent the assigned employees.
    'The conditionals are required to avoid an excel notification if nothing was replaced.
    If Not Range(Cells(startRow, 2), Cells(rowCount - 1, numRoles)).Find("1") Is Nothing Then
        a = Range(Cells(startRow, 2), Cells(rowCount - 1, numRoles)).Replace("1", "*", xlWhole)
    End If
    If Not Range(Cells(startRow, 2), Cells(rowCount - 1, numRoles)).Find("2") Is Nothing Then
        a = Range(Cells(startRow, 2), Cells(rowCount - 1, numRoles)).Replace("2", "1", xlWhole)
    End If
    'Colors the modules in the module sheet with alternating color schemes.
    'Change RGB values in the statemens to change the colors.
    If color Then
        Range(Cells(startRow, 1), Cells(rowCount - 1, numRoles)).Interior.color = RGB(130, 130, 230)
    Else
        Range(Cells(startRow, 1), Cells(rowCount - 1, numRoles)).Interior.color = RGB(130, 230, 130)
    End If
    rowCount = rowCount + 1
    startRow = rowCount
    color = Not color
Next

End Sub

Sub MoveModule()
Dim a As Boolean
Sheet53.Activate
If Not Cells(2, 2).Find("2") Is Nothing Then
    a = Cells(2, 2).Replace("2", "1", xlWhole)
End If
End Sub