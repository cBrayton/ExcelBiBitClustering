Sub PickModules()

lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
'Loops through all the rows in the module sheet created by CreateModules and runs adjustLower off the largest module.
For r = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(r, 1).Value = "" Then
         'Exit sub works here because the duplicates are removed in CreateModules (see comment block below).
'        'Exit sub definitely works, but goto DupCheck may be better to use.
'        'GoTo DupCheck
        Exit Sub
    End If
    m = adjustLower(TakeMax(r), Range(Cells(r, 1), Cells(lastRow, 3)))
Next
'This block is unneccessary due to the remove duplicates loop in CreateModules where this is called.
''Loops through all rows and deletes any row that is identical to the row above it.
''This will miss some duplicates.
'For r = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
'DupCheck:
'    If Cells(r, 1).Value = Cells(r + 1, 1).Value Then
'        Rows(r + 1).Delete
'        GoTo DupCheck
'    End If
'Next
End Sub

Function TakeMax(r)
'This function finds the largest module in modSheet, and returns an array of strings representing the modules.
'The array format is: # of Courses, row #'s of courses ..., "Employees" , column #'s of employees.
' E.G.: 4 1 2 3 4 Employees 2 3 5 6 7
Dim max() As String
lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
a = Range(Cells(1, 1), Cells(lastRow, 3)).Sort(Columns(2), xlDescending)
'If the row is blank then exit.
If Cells(r, 2).Value = "" Then
    Exit Function
End If
'If there is a tie for the largest module then the tying modules are compared using SampleChange.
'The comparison simulates the results of adjusting the lower modules to fit each of the potential largest modules.
'The winner is the module that leaves behind the largest module after being removed.
'This method ensures that the largest mutually exclusive modules are taken. (It doesn't work perfectly since it won't check if there is no tie, but this is a rare/difficult case to check.)
If Cells(r, 2).Value = Cells(r + 1, 2).Value Then
    Dim temp, winner, winnerLoc As Integer
    Dim tempString(2) As String
    Dim x As Integer
    x = r + 1
    Do While Cells(r, 2).Value = Cells(x + 1, 2).Value
        x = x + 1
    Loop
    winner = 0
    winnerLoc = r
    For i = r To x
        temp = sampleChange(Split(Trim(Cells(i, 1).Value)), Range(Cells(r, 1), Cells(x, 3)))
        If temp > winner Then
            winner = temp
            winnerLoc = i
        End If
    Next
    'This stores the winning module in max, and switches it with the module in the first location.
    tempString(0) = Cells(winnerLoc, 1).Value
    tempString(1) = Cells(winnerLoc, 2).Value
    tempString(2) = Cells(winnerLoc, 3).Value
    max = Split(Trim(Cells(winnerLoc, 1).Value))
    Cells(winnerLoc, 1).Value = Cells(r, 1).Value
    Cells(winnerLoc, 2).Value = Cells(r, 2).Value
    Cells(winnerLoc, 3).Value = Cells(r, 3).Value
    Cells(r, 1).Value = tempString(0)
    Cells(r, 2).Value = tempString(1)
    Cells(r, 3).Value = tempString(2)
'If there is no tie then the largest module is used.
Else
    max = Split(Trim(Cells(r, 1).Value))
End If
Dim numCourses As Integer
numCourses = 1
Do While max(numCourses + 1) <> "Employees"
    numCourses = numCourses + 1
Loop
max(0) = numCourses 'the first index of max indicates the number of courses. ie 2 then max(1) and max(2) are the course numbers and max(4) to max(end) are the employees. With the compressed system this doesn't indicate the number of courses but still functions similarly
TakeMax = max

End Function

Function adjustLower(max, rng)
'This function adjusts all the modules in rng so they don't overlap the module represented by max.
'The module returns the size of the second largest modules (the largest is the unchanged max). This return value seems unneccesary and this could be turned to a sub. It was originally a function because it is nearly identical to the function SampleChange below.
Dim contents As String
Dim tempCourses As Integer
Dim courseLoc As New Collection
Dim employeeLoc As New Collection
For r = rng.Rows.Count To 1 Step -1 'Runs bottom to top so rows aren't skipped when deleting.
    Set row = Range(rng.Cells(r, 1), rng.Cells(r, 3))
    If row.Cells(1, 1) = rng.Cells(1, 1) Then
        GoTo nextRow
    End If
    contents = ""
    Set courseLoc = New Collection
    Set employeeLoc = New Collection
    Dim test() As String
    test() = Split(Trim(row.Cells(1, 1)))
    Dim i, testLen As Integer
    testLen = UBound(test)
    If testLen = -1 Then
        Exit For
    End If
    i = 1
    tempCourses = 1
    Do While test(i) <> "Employees" 'Gets the course matches
        For j = 1 To max(0)
            'If test(i) = max(j) Then
            '   courseLoc.Add (i)
            'End If
            If test(i) \ 1000000 = max(j) \ 1000000 Then
                current = test(i) \ 1000000
                Match = (test(i) - (current * 1000000)) And (max(j) - (current * 1000000))
                courseLoc.Add (Match), CStr(i)
            End If
        Next
        If courseLoc.Count <> i Then
            courseLoc.Add 0, CStr(i)
        End If
        i = i + 1
    Loop
    tempCourses = i - 1
    i = i + 1
    For i = i To UBound(test) 'Gets the employee matches
        For j = max(0) + 2 To UBound(max)
            'If test(i) = max(j) Then
            'employeeLoc.Add (i)
            'End If
            If test(i) \ 1000000 = max(j) \ 1000000 Then
                current = test(i) \ 1000000
                Match = test(i) - (current * 1000000) And max(j) - (current * 1000000)
                employeeLoc.Add (Match), CStr(i)
            End If
        Next
        If employeeLoc.Count <> i - (tempCourses + 1) Then
            employeeLoc.Add 0, CStr(i)
        End If
    Next
' The first commented out if block may be obsolete here since the modules are now sorted.
'    If employeeLoc.Count + courseLoc.Count + 1 < UBound(max) Then
    numCourses = 0
    numEmployees = 0
    For Each crs In courseLoc
        For N = 0 To 15
            If crs And (2 ^ N) Then
                numCourses = numCourses + 1
            End If
        Next
    Next
    For Each emp In employeeLoc
        For N = 0 To 15
            If emp And (2 ^ N) Then
                numEmployees = numEmployees + 1
            End If
        Next
    Next
    If numCourses = 0 Or numEmployees = 0 Then
        GoTo nextRow
    ElseIf numEmployees * (row.Cells(1, 2).Value / row.Cells(1, 3).Value) > numCourses * row.Cells(1, 3).Value Then
        For i = 1 To tempCourses
            temp = test(i) \ 1000000
            test(i) = (test(i) - (temp * 1000000)) And Not courseLoc.Item(i)
            If test(i) = 0 Then
                test(i) = ""
            Else
                test(i) = test(i) + (temp * 1000000)
            End If
        Next
        row.Cells(1, 2).Value = row.Cells(1, 2).Value - (numCourses * row.Cells(1, 3).Value)
    Else
        For i = tempCourses + 2 To UBound(test)
            temp = test(i) \ 1000000
            test(i) = (test(i) - (temp * 1000000)) And Not employeeLoc.Item(CStr(i))
            If test(i) = 0 Then
                test(i) = ""
            Else
            test(i) = test(i) + (temp * 1000000)
            End If
        Next
        row.Cells(1, 2).Value = row.Cells(1, 2).Value - (numEmployees * (row.Cells(1, 2).Value / row.Cells(1, 3).Value))
        row.Cells(1, 3).Value = row.Cells(1, 3).Value - numEmployees
    End If
'These conditions work but need to be largely redone for the New code blocks and compression.
'        If employeeLoc.Count * tempCourses > courseLoc.Count * (UBound(test) - (tempCourses + 1)) Then
'            For Each crs In courseLoc
'                test(crs) = ""
'            Next
'            row.Cells(1, 2).Value = row.Cells(1, 2).Value - (courseLoc.Count * (UBound(test) - (tempCourses + 1)))
'        Else
'            For Each emp In employeeLoc
'                test(emp) = ""
'            Next
'            row.Cells(1, 2).Value = row.Cells(1, 2).Value - (employeeLoc.Count * tempCourses)
'        End If

'    Else
'        row.Delete 'change to row.delete from temp = row.cells(1,2).value; to remove equivalent but not identical duplicates
'        GoTo nextRow
'    End If

    For k = 1 To UBound(test)
        If test(k) <> "" Then
            contents = contents & " " & test(k)
        End If
    Next
    contents = "Courses" & contents
    row.Cells(1, 1).Value = contents
    sample = Split(Trim(row.Cells(1, 1).Value))
    If sample(1) = "Employees" Or sample(UBound(sample)) = "Employees" Then
        row.Delete 'Reruns a row when deleting.
    End If
nextRow:
Next
'a = rng.Sort(Columns(2), xlAscending)
'i = 1
'Do While rng.Rows(i).Cells(1, 2).Value = 0
'    sample = Split(Trim(rng.Rows(i).Cells(1, 1).Value))
'    If sample(1) = "Employees" Or sample(UBound(sample)) = "Employees" Then
'        rng.Rows(i).Delete
'        i = i - 1
'    End If
'    i = i + 1
'Loop
a = rng.Sort(Columns(2), xlDescending)
adjustLower = rng.Cells(2, 2)
End Function

Function sampleChange(max, rng)
Dim contents As String
Dim temp, winner As Integer
Dim tempCourses As Integer
If UBound(max) = -1 Then
    Exit Function
End If
winner = 0
For Each row In rng.Rows
    contents = ""
    Dim courseLoc As New Collection
    Dim employeeLoc As New Collection
    Set courseLoc = New Collection
    Set employeeLoc = New Collection
    test = Split(Trim(row.Cells(1, 1)))
    If UBound(test) = -1 Then
        Exit For
    End If
    Dim numCourses As Integer
    numCourses = 1
    Do While max(numCourses + 1) <> "Employees"
        numCourses = numCourses + 1
    Loop
    max(0) = numCourses 'the first index of max indicates the number of courses. ie 2 then max(1) and max(2) are the course numbers and max(4) to max(end) are the employees.
    Dim i As Integer
    i = 1
    tempCourses = 1 - 1
    Do While test(i) <> "Employees" 'Gets the course matches
        For j = 1 To max(0)
            'If test(i) = max(j) Then
            '   courseLoc.Add (i)
            'End If
            If test(i) \ 1000000 = max(j) \ 1000000 Then
                current = test(i) \ 1000000
                Match = test(i) - (current * 1000000) And max(j) - (current * 1000000)
                courseLoc.Add (Match), CStr(i)
            End If
        Next
        i = i + 1
    Loop
    tempCourses = i - 1
    i = i + 1
    For i = i To UBound(test) 'Gets the employee matches
        For j = max(0) + 2 To UBound(max)
            'If test(i) = max(j) Then
            'employeeLoc.Add (i)
            'End If
            If test(i) \ 1000000 = max(j) \ 1000000 Then
                current = test(i) \ 1000000
                Match = test(i) - (current * 1000000) And max(j) - (current * 1000000)
                employeeLoc.Add (Match), CStr(i)
            End If
        Next
    Next
'    Do While test(i) <> "Employees" 'Gets the course matches
'        For j = 1 To max(0)
'            If test(i) = max(j) Then
'            courseLoc.Add (i)
'            End If
'        Next
'        i = i + 1
'    Loop
'    tempCourses = i
'    i = i + 1
'    For i = i To UBound(test) 'Gets the employee matches
'        For j = max(0) + 2 To UBound(max)
'            If test(i) = max(j) Then
'            employeeLoc.Add (i)
'            End If
'        Next
'    Next

    numCourses = 0
    numEmployees = 0
    For Each crs In courseLoc
        For N = 0 To 15
            If crs And (2 ^ N) Then
                numCourses = numCourses + 1
            End If
        Next
    Next
    For Each emp In employeeLoc
        For N = 0 To 15
            If emp And (2 ^ N) Then
                numEmployees = numEmployees + 1
            End If
        Next
    Next
    If numEmployees * (row.Cells(1, 2).Value / row.Cells(1, 3).Value) > numCourses * row.Cells(1, 3).Value Then
        temp = row.Cells(1, 2).Value - (numCourses * row.Cells(1, 3).Value)
    Else
        temp = row.Cells(1, 2).Value - (numEmployees * (row.Cells(1, 2).Value / row.Cells(1, 3).Value))
    End If
   
   
'    If employeeLoc.Count + courseLoc.Count + 1 < UBound(max) Then
'        If employeeLoc.Count * tempCourses > courseLoc.Count * (UBound(test) - (tempCourses + 1)) Then
'            temp = row.Cells(1, 2).Value - courseLoc.Count * (UBound(test) - (tempCourses + 1))
'        Else
'            temp = row.Cells(1, 2).Value - employeeLoc.Count * tempCourses
'        End If
'    Else
'        temp = row.Cells(1, 2).Value
'    End If
    If temp > winner Then
        winner = temp
    End If
Next
sampleChange = winner
End Function

Sub RemoveEquivalents()

For L = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(L, 1) <> "" Then
        If Cells(L, 2) = Cells(L + 1, 2) Then
            Dim sample() As String
            sample() = Split(Trim(Cells(L, 1)))
            Dim test() As String
            test() = Split(Trim(Cells(L + 1, 1)))
            Match = True
            N = 0
            For j = 0 To UBound(sample)
                i = 0
                Do While test(i) <> "Employees" And N = j
                    If test(i) = sample(j) Then
                        N = N + 1
                        Exit Do
                    End If
                i = i + 1
                Loop
                If sample(j) = "Employees" Then
                    Do While test(i) <> "" And N = j
                        If test(i) = sample(j) Then
                            N = N + 1
                            Exit Do
                        End If
                    i = i + 1
                    Loop
                End If
            Next
            If N = UBound(sample) Then
            ActiveSheet.Rows(L + 1).Delete
            End If
        End If
    End If
Next

End Sub