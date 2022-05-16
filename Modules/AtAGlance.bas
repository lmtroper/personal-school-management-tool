Sub Show_AtAGlanceUF()

AtAGlance_UserForm.Show

End Sub

Sub AtAGlance_Output() 'Outputs and organizes assessments and tasks that fall within a specific date range

Application.ScreenUpdating = False

'Declares worksheets to reference and retrieve data from
Dim deliverables As Worksheet, tests As Worksheet, atAGlance As Worksheet, task As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set atAGlance = ThisWorkbook.Sheets("At_A_Glance")
Set task = ThisWorkbook.Sheets("Task_List")
Set tests = ThisWorkbook.Sheets("Tests")

'Declares Tables in the Task_List Sheet
Dim tblTasks As Long, tblDeliverables As Long, tblQuizzes As Long, tblMidterms As Long, tblFinals As Long
tblDeliverables = atAGlance.ListObjects("Table91117").ListRows.Count 'Number of rows in Deliverables Table
tblQuizzes = atAGlance.ListObjects("Table911").ListRows.Count 'Number of rows in Quizzes Table
tblMidterms = atAGlance.ListObjects("Table91118").ListRows.Count 'Number of rows in Midterms Table
tblFinals = atAGlance.ListObjects("Table9111819").ListRows.Count 'Number of rows in Finals Table

Dim i As Integer 'Used for Looping

Dim iRow As Long, iRow2 As Long
iRow = [Counta(Deliverables!B:B)] 'Counts the amount of non-empty entries in column B of Deliverables Sheet
iRow2 = [Counta(Tests!B:B)] 'Counts the amount of non-empty entries in column B of Tests Sheet

With deliverables
    For i = 2 To iRow  'Loop through each assessment from the first non-header entry to the last entry
    
        'Determine if the deadline of the assessment falls between the specified start and end date
        If .Cells(i, "G") >= CDate(AtAGlance_UserForm.txt_startDate) And .Cells(i, "G") <= CDate(AtAGlance_UserForm.txt_endDate) Then
                For k = 1 To 5
                    If .Cells(i, "B") = atAGlance.Cells(k + 9, "E") Then
                        If atAGlance.Cells(k + 9, "F") <> "" Then
                            atAGlance.Cells(k + 9, "F") = atAGlance.Cells(k + 9, "F").Value + (deliverables.Cells(i, "F"))
                        Else
                            atAGlance.Cells(k + 9, "F") = (deliverables.Cells(i, "F"))
                        End If
                    Exit For
                    End If
                Next k
                
            'If the table length is equal to 1 and the first entry of the table is blank
            'Then add the assessment data into the first row
             If tblDeliverables = 1 And atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "A").Text = "" Then
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(i, "B").Text
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(i, "C").Text
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(i, "G").Text
        
             Else
                 atAGlance.ListObjects("Table91117").ListRows.Add 'Add a new row to prevent re-writing data
                 tblDeliverables = atAGlance.ListObjects("Table91117").ListRows.Count 'Re-count the rows in the table
                
                'Add the data into the last row of the table (the row that just got added)
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(tblDeliverables, "A").Value = deliverables.Cells(i, "B").Text
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(tblDeliverables, "B").Value = deliverables.Cells(i, "C").Text
                 atAGlance.ListObjects("Table91117").DataBodyRange.Cells(tblDeliverables, "C").Value = deliverables.Cells(i, "G").Text
             End If
        End If
    Next i
End With

With tests
    For i = 2 To iRow2 'Loop through each assessment from the first non-header entry to the last entry
       
       'Determine if the deadline of the assessment falls between the specified start and end date
       If .Cells(i, "H") >= CDate(AtAGlance_UserForm.txt_startDate) And .Cells(i, "H") <= CDate(AtAGlance_UserForm.txt_endDate) Then
                
                'If the course is the same as the course on the At A Glance Page
                'E.g. If the test is for Math 205 and Math 205 is on the At A Glance Page
                'For adding up the amount of hours per course
                For k = 1 To 5 ' 5 courses
                    If .Cells(i, "B") = atAGlance.Cells(k + 9, "E") Then
        
                         If atAGlance.Cells(k + 9, "F") <> "" Then
                            atAGlance.Cells(k + 9, "F") = atAGlance.Cells(k + 9, "F").Value + (deliverables.Cells(i, "F"))
                        Else
                            atAGlance.Cells(k + 9, "F") = (deliverables.Cells(i, "F"))
                        End If
                        
                    Exit For
                    End If
                Next k
            
            If .Cells(i, "A").Text = "Quiz" Then
                
                For k = 1 To 5
                    If .Cells(i, "B") = atAGlance.Cells(k + 9, "E") Then
                        If atAGlance.Cells(k + 9, "F") <> "" Then
                            atAGlance.Cells(k + 9, "F") = atAGlance.Cells(k + 9, "F").Value + (deliverables.Cells(i, "F"))
                        Else
                            atAGlance.Cells(k + 9, "F") = (deliverables.Cells(i, "F"))
                        End If
                    Exit For
                    End If
                Next k
                 
                'If the table length is equal to 1 and the first entry of the table is blank
                'Then add the assessment data into the first row
                 If tblQuizzes = 1 And atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "A").Text = "" Then
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "A").Value = .Cells(i, "B").Text
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "B").Value = .Cells(i, "C").Text
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "C").Value = .Cells(i, "H").Text
            
                 Else
                     atAGlance.ListObjects("Table911").ListRows.Add 'Add a new row to prevent re-writing data
                     tblQuizzes = atAGlance.ListObjects("Table911").ListRows.Count 'Re-count the rows in the table
                    
                    'Add the data into the last row of the table (the row that just got added)
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(tblQuizzes, "A").Value = tests.Cells(i, "B").Text
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(tblQuizzes, "B").Value = tests.Cells(i, "C").Text
                     atAGlance.ListObjects("Table911").DataBodyRange.Cells(tblQuizzes, "C").Value = tests.Cells(i, "H").Text
                 End If
            
            ElseIf .Cells(i, "A").Text = "Midterm" Or .Cells(i, "A").Text = "Unit/Term Test" Then
                 
                'If the table length is equal to 1 and the first entry of the table is blank
                'Then add the assessment data into the first row
                 If tblMidterms = 1 And atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "A").Text = "" Then
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "A").Value = .Cells(i, "B").Text
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "B").Value = .Cells(i, "C").Text
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "C").Value = .Cells(i, "H").Text
            
                 Else
                     atAGlance.ListObjects("Table91118").ListRows.Add 'Add a new row to prevent re-writing data
                     tblMidterms = atAGlance.ListObjects("Table91118").ListRows.Count 'Re-count the rows in the table
    
                    'Add the data into the last row of the table (the row that just got added)
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(tblMidterms, "A").Value = tests.Cells(i, "B").Text
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(tblMidterms, "B").Value = tests.Cells(i, "C").Text
                     atAGlance.ListObjects("Table91118").DataBodyRange.Cells(tblMidterms, "C").Value = tests.Cells(i, "H").Text
                 End If
            
            ElseIf .Cells(i, "A").Text = "Final" Then
            
                 
                'If the table length is equal to 1 and the first entry of the table is blank
                'Then add the assessment data into the first row
                 If tblFinals = 1 And atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "A").Text = "" Then
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "A").Value = .Cells(i, "B").Text
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "B").Value = .Cells(i, "C").Text
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "C").Value = .Cells(i, "H").Text
            
                 Else
                     atAGlance.ListObjects("Table9111819").ListRows.Add 'Add a new row to prevent re-writing data
                     tblFinals = atAGlance.ListObjects("Table9111819").ListRows.Count 'Re-count the rows in the table
                    
                    'Add the data into the last row of the table (the row that just got added)
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(tblFinals, "A").Value = tests.Cells(i, "B").Text
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(tblFinals, "B").Value = tests.Cells(i, "C").Text
                     atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(tblFinals, "C").Value = tests.Cells(i, "H").Text
                 End If
            End If
        End If
    Next i
End With


With task

'Declare the tables in Task_List to retrieve its data
Dim tblMeetingRow As Long, tblEventRow As Long, tblThingsToDoRows As Long
tblMeetingRow = task.ListObjects("Table6").ListRows.Count 'Number of rows in Meetings Table
tblEventRow = task.ListObjects("Table4").ListRows.Count 'Number of rows in Events Table
tblThingsToDoRows = task.ListObjects("Table8").ListRows.Count 'Number of rows in Things to Do Table
tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Number of rows in Tasks Table


    For i = 1 To tblMeetingRow 'Loop through rows in Meetings Table
    tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows to prevent re-writing of data
    
       'Determine if the deadline of the assessment falls between the specified start and end date
       If task.ListObjects("Table6").DataBodyRange.Cells(i, "C").Value >= CDate(AtAGlance_UserForm.txt_startDate) And task.ListObjects("Table6").DataBodyRange.Cells(i, "C").Value <= CDate(AtAGlance_UserForm.txt_endDate) Then
            
            'If the table length is equal to 1 and the first entry of the table is blank
            'Then add the assessment data into the first row of the Tasks Table
            If tblTasks = 1 And atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Text = "" Then
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "A").Value = task.ListObjects("Table6").DataBodyRange.Cells(i, "A").Value
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "B").Value = "Meeting"
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "C").Value = task.ListObjects("Table6").DataBodyRange.Cells(i, "C").Value
                
            Else
                atAGlance.ListObjects("Table12").ListRows.Add 'Add a new row to prevent re-writing data
                tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows in the table
                
                'Add the data into the last row of the Tasks table (the row that just got added)
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "A").Value = task.ListObjects("Table6").DataBodyRange.Cells(i, "A").Value
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "B").Value = "Meeting"
                atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "C").Value = task.ListObjects("Table6").DataBodyRange.Cells(i, "C").Value
            End If
        End If
    Next i


  For i = 1 To tblEventRow 'Loop through rows in Events Table
  tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows to prevent re-writing of data
  
       'Determine if the deadline of the assessment falls between the specified start and end date
       If .ListObjects("Table4").DataBodyRange.Cells(i, "C").Value >= CDate(AtAGlance_UserForm.txt_startDate) And task.ListObjects("Table4").DataBodyRange.Cells(i, "C").Value <= CDate(AtAGlance_UserForm.txt_endDate) Then

            'If the table length is equal to 1 and the first entry of the table is blank
            'Then add the assessment data into the first row of the Tasks Table
             If tblTasks = 1 And atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Text = "" Then
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Value = task.ListObjects("Table4").DataBodyRange.Cells(i, "A").Value
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "B").Value = "Event"
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "C").Value = task.ListObjects("Table4").DataBodyRange.Cells(i, "C").Value
        
             Else
                 
                 atAGlance.ListObjects("Table12").ListRows.Add 'Add a new row to prevent re-writing data
                 tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows in the table
                 
                'Add the data into the last row of the Tasks table (the row that just got added)
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "A").Value = task.ListObjects("Table4").DataBodyRange.Cells(i, "A").Value
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "B").Value = "Event"
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "C").Value = task.ListObjects("Table4").DataBodyRange.Cells(i, "C").Value
             End If
        End If
    Next i
    

    For i = 1 To tblThingsToDoRows 'Loop through rows in Things to Do Table
    tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows to prevent re-writing of data
    
       'Determine if the deadline of the assessment falls between the specified start and end date
       If .ListObjects("Table8").DataBodyRange.Cells(i, "C").Value >= CDate(AtAGlance_UserForm.txt_startDate) And task.ListObjects("Table8").DataBodyRange.Cells(i, "C").Value <= CDate(AtAGlance_UserForm.txt_endDate) Then
             
            'If the table length is equal to 1 and the first entry of the table is blank
            'Then add the assessment data into the first row of the Tasks Table
             If tblTasks = 1 And atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Text = "" Then
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Value = .ListObjects("Table8").DataBodyRange.Cells(i, "A").Value
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "B").Value = "Thing To Do"
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "C").Value = .ListObjects("Table8").DataBodyRange.Cells(i, "C").Value
        
             Else
                 atAGlance.ListObjects("Table12").ListRows.Add 'Add a new row to prevent re-writing data
                 tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the rows in the table
                
                'Add the data into the last row of the Tasks table (the row that just got added)
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "A").Value = task.ListObjects("Table8").DataBodyRange.Cells(i, "A").Value
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "B").Value = "Thing To Do"
                 atAGlance.ListObjects("Table12").DataBodyRange.Cells(tblTasks, "C").Value = task.ListObjects("Table8").DataBodyRange.Cells(i, "C").Value
             End If
             
        End If
    Next i
    
End With

'Adds up the total of hours for each course
atAGlance.Cells(16, "F") = atAGlance.Cells(14, "F").Value + atAGlance.Cells(13, "F").Value + atAGlance.Cells(12, "F").Value + atAGlance.Cells(11, "F").Value + atAGlance.Cells(10, "F").Value

Application.ScreenUpdating = True

End Sub

Sub Clear_AtAGlance_Sheet() 'Clears all the values of the tables on the Sheet

Application.ScreenUpdating = False
'Declare worksheet to reference and retrieve data
Dim atAGlance As Worksheet
Set atAGlance = ThisWorkbook.Sheets("At_A_Glance")

Dim i As Integer 'Used for looping

'Declare the tables in the At A Glance Sheet
Dim tblTasks As Long, tblDeliverables As Long, tblQuizzes As Long, tblMidterms As Long, tblFinals As Long
tblDeliverables = atAGlance.ListObjects("Table91117").ListRows.Count 'Number of rows in Deliverables Table
tblQuizzes = atAGlance.ListObjects("Table911").ListRows.Count 'Number of rows in Quizzes Table
tblMidterms = atAGlance.ListObjects("Table91118").ListRows.Count 'Number of rows in Midterms Table
tblFinals = atAGlance.ListObjects("Table9111819").ListRows.Count 'Number of rows in Finals Table
tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Number of rows in Tasks Table
    
    For i = 1 To tblDeliverables 'Loop through the rows of the table
        tblDeliverables = atAGlance.ListObjects("Table91117").ListRows.Count 'Re-count the table rows each loop
        
        If tblDeliverables = 1 Then
            'If there is only one row, clear the values
            atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "A").Value = ""
            atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "B").Value = ""
            atAGlance.ListObjects("Table91117").DataBodyRange.Cells(1, "C").Value = ""
            
        Else
             atAGlance.ListObjects("Table91117").ListRows(tblDeliverables).Delete 'Delete the entire row from the table
        End If

    Next

    For i = 1 To tblQuizzes 'Loop through the rows of the table
        tblQuizzes = atAGlance.ListObjects("Table911").ListRows.Count 'Re-count the table rows each loop
        If tblQuizzes = 1 Then
            'If there is only one row, clear the values
            atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "A").Value = ""
            atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "B").Value = ""
            atAGlance.ListObjects("Table911").DataBodyRange.Cells(1, "C").Value = ""
            
        Else
            atAGlance.ListObjects("Table911").ListRows(tblQuizzes).Delete 'Delete the entire row from the table
        End If
    Next

    For i = 1 To tblMidterms 'Loop through the rows of the table
        tblMidterms = atAGlance.ListObjects("Table91118").ListRows.Count 'Re-count the table rows each loop
        If tblMidterms = 1 Then
            atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "A").Value = ""
            atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "B").Value = ""
            atAGlance.ListObjects("Table91118").DataBodyRange.Cells(1, "C").Value = ""
            
        Else
            atAGlance.ListObjects("Table91118").ListRows(tblMidterms).Delete 'Delete the entire row from the table
                    
        End If
    Next

    For i = 1 To tblFinals 'Loop through the rows of the table
    tblFinals = atAGlance.ListObjects("Table9111819").ListRows.Count 'Re-count the table rows each loop
        
        If tblFinals = 1 Then
            'If there is only one row, clear the values
            atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "A").Value = ""
            atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "B").Value = ""
            atAGlance.ListObjects("Table9111819").DataBodyRange.Cells(1, "C").Value = ""
            
        Else
            atAGlance.ListObjects("Table9111819").ListRows(tblFinals).Delete 'Delete the entire row from the table
                    
        End If
        
    Next i

    For i = 1 To tblTasks 'Loop through the rows of the table
    tblTasks = atAGlance.ListObjects("Table12").ListRows.Count 'Re-count the table rows each loop
        
        If tblTasks = 1 Then
            'If there is only one row, clear the values
            atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "A").Value = ""
            atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "B").Value = ""
            atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "C").Value = ""
            atAGlance.ListObjects("Table12").DataBodyRange.Cells(1, "D").Value = ""
            
        Else
            atAGlance.ListObjects("Table12").ListRows(tblTasks).Delete 'Delete the entire row from the table
                    
        End If

    Next i

With atAGlance

'Clears the hours
For i = 1 To 7
    .Cells(9 + i, "F").ClearContents
Next i

End With


Application.ScreenUpdating = True

End Sub
