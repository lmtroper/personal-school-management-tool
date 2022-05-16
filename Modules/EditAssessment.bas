Sub Show_EditUF()

EditAssessment_UserForm1.Show

End Sub


Sub Submit_TestEditData() 'Re-writes the original assessment data with the edited data

Dim upcoming As Worksheet, tests As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set tests = ThisWorkbook.Sheets("Tests")

indexNum = EditAssessment_UserForm1.cmbTests.ListIndex 'Element number of value of drop-box

'With Tests Sheet:
With tests
    'Determine type of test
    If EditTest_UserForm.optQuiz.Value = True Then
        .Cells(indexNum + 2, 1) = "Quiz"
    ElseIf EditTest_UserForm.optMidterm.Value = True Then
        .Cells(indexNum + 2, 1) = "Midterm"
    ElseIf EditTest_UserForm.optUnitTest.Value = True Then
        .Cells(indexNum + 2, 1) = "Unit/Term Test"
    ElseIf EditTest_UserForm.optFinal.Value = True Then
        .Cells(indexNum + 2, 1) = "Final"
    End If
                 
    'Re-write over original data
    .Cells(indexNum + 2, 2) = EditTest_UserForm.cmbExistingCourses2.Value
    .Cells(indexNum + 2, 3) = EditTest_UserForm.txtAssessmentName2.Value
    .Cells(indexNum + 2, 4) = EditTest_UserForm.txtLengthHours2.Value
    .Cells(indexNum + 2, 5) = EditTest_UserForm.txtWeight2.Value
    .Cells(indexNum + 2, 6) = EditTest_UserForm.txtAntGrade2.Value
    .Cells(indexNum + 2, 7) = EditTest_UserForm.txtAntStudyHours2.Text
    .Cells(indexNum + 2, 8) = EditTest_UserForm.txtDeadline2.Value
End With


With upcoming
    'Re-writes the values in the table
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "A").Value = EditTest_UserForm.cmbExistingCourses2.Value
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "C").Value = EditTest_UserForm.txtAssessmentName2.Value
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "D").Value = EditTest_UserForm.txtLengthHours2.Value
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "E").Value = EditTest_UserForm.txtWeight2.Value
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "F").Value = EditTest_UserForm.txtAntGrade2.Value
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "G").Value = EditTest_UserForm.txtAntStudyHours2.Text
    upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "H").Value = EditTest_UserForm.txtDeadline2.Value

    'Determines the test type
    If EditTest_UserForm.optQuiz.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "B").Value = "Quiz"
    ElseIf EditTest_UserForm.optMidterm.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "B").Value = "Midterm"
    ElseIf EditTest_UserForm.optUnitTest.Value = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "B").Value = "Unit/Term Test"
    ElseIf EditTest_UserForm.optFinal = True Then
        upcoming.ListObjects("Table1").DataBodyRange.Cells(indexNum + 1, "B").Value = "Final"
    End If
                
End With

End Sub

Sub Submit_DeliverableEditData()

Dim upcoming As Worksheet, deliverables As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set deliverables = ThisWorkbook.Sheets("Deliverables")

Dim indexNum As Long
indexNum = EditAssessment_UserForm1.cmbDeliverables.ListIndex

    'With deliverables sheet:
    With deliverables
        'Re-write over orginal data with edited data
        .Cells(indexNum + 2, 1) = EditDeliverable_UserForm.cmbDeliverableType
        .Cells(indexNum + 2, 2) = EditDeliverable_UserForm.cmbExistingCourses3.Value
        .Cells(indexNum + 2, 3) = EditDeliverable_UserForm.txtAssessmentName3.Value
        .Cells(indexNum + 2, 4) = EditDeliverable_UserForm.txtWeight3.Value
        .Cells(indexNum + 2, 5) = EditDeliverable_UserForm.txtAntGrade3.Value
        .Cells(indexNum + 2, 6) = EditDeliverable_UserForm.txtAntStudyHours3.Text
        .Cells(indexNum + 2, 7) = EditDeliverable_UserForm.txtDeadline3.Value
    End With


    With upcoming
    
            'Re-write over original data in the Deliverables Table with edited data
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "A").Value = EditDeliverable_UserForm.cmbExistingCourses3.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "B").Value = EditDeliverable_UserForm.cmbDeliverableType.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "C").Value = EditDeliverable_UserForm.txtAssessmentName3.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "D").Value = EditDeliverable_UserForm.txtWeight3.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "E").Value = EditDeliverable_UserForm.txtAntGrade3.Value
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "F").Value = EditDeliverable_UserForm.txtAntStudyHours3.Text
            upcoming.ListObjects("Table24").DataBodyRange.Cells(indexNum + 1, "G").Value = EditDeliverable_UserForm.txtDeadline3.Value
    End With

End Sub
