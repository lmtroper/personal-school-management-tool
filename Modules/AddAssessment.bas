Sub Show_AddAssessmentUF() 'Linked to button on sheet

AddAssessment_UserForm.Show

End Sub

Sub Submit_DeliverableInfo()

Dim iRow As Long
iRow = [Counta(Deliverables!A:A)] + 1

Dim upcoming As Worksheet, deliverables As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set deliverables = ThisWorkbook.Sheets("Deliverables")

With deliverables
    'Deliverable Type
    .Cells(iRow, 1) = AddDeliverable_UserForm.cmbDeliverableType.Value
    
    'Determine which type of deliverable it is
    If AddDeliverable_UserForm.cmbADExistingCourses.Text <> "" Then
        .Cells(iRow, 2) = AddDeliverable_UserForm.cmbADExistingCourses.Value
    Else
        .Cells(iRow, 2) = AddDeliverable_UserForm.txtADCourseName.Value + " " + AddDeliverable_UserForm.txtADCourseCode.Value
                
    End If
                
    'Submits info into the Deliverable Sheet
    .Cells(iRow, 3) = AddDeliverable_UserForm.txtDeliverableName.Value
    .Cells(iRow, 4) = AddDeliverable_UserForm.txtADWeight.Value
    .Cells(iRow, 5) = AddDeliverable_UserForm.txtADAntGrade.Value
    .Cells(iRow, 6) = AddDeliverable_UserForm.txtADAntStudyHours.Text
    .Cells(iRow, 7) = AddDeliverable_UserForm.txtDeadline.Value
End With

With upcoming
Dim tblDeliverable As Long
tblDeliverable = upcoming.ListObjects("Table1").ListRows.Count 'Number of rows in Deliverables Table
    
    'If the table length is equal to 1 and the first entry of the table is blank
    'Then add the UserForm data into the first row
    If tblDeliverable = 1 And upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "A").Value = " - " Then
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "B").Value = AddDeliverable_UserForm.cmbDeliverableType.Value
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "C").Value = AddDeliverable_UserForm.txtDeliverableName.Value
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "D").Value = AddDeliverable_UserForm.txtADWeight.Value
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "E").Value = AddDeliverable_UserForm.txtADAntGrade.Value
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "F").Value = AddDeliverable_UserForm.txtADAntStudyHours.Value
        upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "G").Value = AddDeliverable_UserForm.txtDeadline.Value
        
    
    Else
        upcoming.ListObjects("Table24").ListRows.Add 'Add a new row to prevent re-writing data
        tblDeliverable = upcoming.ListObjects("Table24").ListRows.Count 'Re-count the rows in the table
        
            'Add the data into the last row of the table (the row that just got added)
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "B").Value = AddDeliverable_UserForm.cmbDeliverableType.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "C").Value = AddDeliverable_UserForm.txtDeliverableName.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "D").Value = AddDeliverable_UserForm.txtADWeight.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "E").Value = AddDeliverable_UserForm.txtADAntGrade.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "F").Value = AddDeliverable_UserForm.txtADAntStudyHours.Text
            upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "G").Value = AddDeliverable_UserForm.txtDeadline.Value
    
    End If
    
    'Determine which type of deliverable it is
    If AddDeliverable_UserForm.cmbADExistingCourses.Text <> "" Then
        upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "A").Value = AddDeliverable_UserForm.cmbADExistingCourses.Value
    Else
        upcoming.ListObjects("Table24").DataBodyRange.Cells(tblDeliverable, "A").Value = AddDeliverable_UserForm.txtADCourseName.Value + " " + AddDeliverable_UserForm.txtADCourseCode.Value
        
    End If
                
End With

End Sub

Sub Submit_TestInfo() 'Submits info about a new test into the Test Sheet (hidden sheet) and Test Table on Upcoming Assessments

Dim iRow As Long
iRow = [Counta(Tests!A:A)] + 1

Dim upcoming As Worksheet, tests As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set tests = ThisWorkbook.Sheets("Tests")

'Determine which type of test it is
With tests
    If AddTest_UserForm.optQuiz.Value = True Then
        .Cells(iRow, 1) = "Quiz"
    ElseIf AddTest_UserForm.optMidterm.Value = True Then
        .Cells(iRow, 1) = "Midterm"
    ElseIf AddTest_UserForm.optUnitTest.Value = True Then
        .Cells(iRow, 1) = "Unit/Term Test"
    ElseIf AddTest_UserForm.optFinal = True Then
        .Cells(iRow, 1) = "Final"
    End If
    
    'Determines if it is an existing course from the combo-box or a new course in the textboxes
    If AddTest_UserForm.cmbATExistingCourses.Text <> "" Then
        .Cells(iRow, 2) = AddTest_UserForm.cmbATExistingCourses.Value
    Else
        .Cells(iRow, 2) = AddTest_UserForm.txtATCourseName.Value + " " + AddTest_UserForm.txtATCourseCode.Value
                
    End If
               
    'Fills in values on the Test Sheet
    .Cells(iRow, 3) = AddTest_UserForm.txtTestName.Value
    .Cells(iRow, 4) = AddTest_UserForm.txtLengthHours.Value
    .Cells(iRow, 5) = AddTest_UserForm.txtATWeight.Value
    .Cells(iRow, 6) = AddTest_UserForm.txtATAntGrade.Value
    .Cells(iRow, 7) = AddTest_UserForm.txtATAntStudyHours.Text
    .Cells(iRow, 8) = AddTest_UserForm.txtDate.Value
End With


With upcoming
Dim tblTest As Long
tblTest = upcoming.ListObjects("Table1").ListRows.Count 'Number of rows in the Tests Table

    'If the table length is equal to 1 and the first entry of the table is blank
    'Then add the UserForm data into the first row
    If tblTest = 1 And upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "A").Value = " - " Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "C").Value = AddTest_UserForm.txtTestName.Value
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "D").Value = AddTest_UserForm.txtLengthHours.Value
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "E").Value = AddTest_UserForm.txtATWeight.Value
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "F").Value = AddTest_UserForm.txtATAntGrade.Value
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "G").Value = AddTest_UserForm.txtATAntStudyHours.Text
        upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "H").Value = AddTest_UserForm.txtDate.Value
        
    
    Else
        upcoming.ListObjects("Table1").ListRows.Add 'Add a new row to prevent re-writing data
        tblTest = upcoming.ListObjects("Table1").ListRows.Count 'Re-count the rows in the table
        
            'Add the data into the last row of the table (the row that just got added)
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "C").Value = AddTest_UserForm.txtTestName.Value
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "D").Value = AddTest_UserForm.txtLengthHours.Value
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "E").Value = AddTest_UserForm.txtATWeight.Value
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "F").Value = AddTest_UserForm.txtATAntGrade.Value
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "G").Value = AddTest_UserForm.txtATAntStudyHours.Text
            upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "H").Value = AddTest_UserForm.txtDate.Value
    
    End If
    
    'If the user enters existing course or new course
    If AddTest_UserForm.cmbATExistingCourses.Text <> "" Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "A").Value = AddTest_UserForm.cmbATExistingCourses.Value
    Else
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "A").Value = AddTest_UserForm.txtATCourseName.Value + " " + AddTest_UserForm.txtATCourseCode.Value
    End If
    
    'Determines the type of test it is
    If AddTest_UserForm.optQuiz.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "B").Value = "Quiz"
    ElseIf AddTest_UserForm.optMidterm.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "B").Value = "Midterm"
    ElseIf AddTest_UserForm.optUnitTest.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "B").Value = "Unit/Term Test"
    ElseIf AddTest_UserForm.optFinal = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(tblTest, "B").Value = "Final"
    End If
                
End With

End Sub
