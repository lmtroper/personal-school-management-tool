Sub Show_CompleteUF()

CompleteAssessment_UserForm1.Show

End Sub


Sub Submit_CompleteForm2() 'Remove the assessment from Upcoming Assessments and transfers data to Completed Assessments Sheet

Dim indexNum As Long, tblSemesterRow As Long, studyHours As Variant
Dim percentperHour As Variant

'Declare the worksheets that will be referenced
Dim upcoming As Worksheet, complete As Worksheet, deliverables As Worksheet, tests As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set complete = ThisWorkbook.Sheets("Completed_Assessments")
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")


Dim iRow As Long
iRow = [Counta(Completed_Assessments!A:A)] + 1 'Counts the amount of non-empty entries in column A of Completed Assessments Sheet

'If the assessment is a deliverable (user selected a value on the deliverable combo-box):
If CompleteAssessment_UserForm1.cmbDeliverables.Value <> "" Then

    indexNum = CompleteAssessment_UserForm1.cmbDeliverables.ListIndex 'the index number of the value selected in combo-box
    
    'If the user receives a suggestion for how many more hours they should have studied
    If CompleteAssessment_UserForm2.txtSuggestHours.Value <> "" Then
        studyHours = CompleteAssessment_UserForm2.txtSuggestHours.Value
    Else
        studyHours = " - "
    End If
    
    'Fils in detais into Complete Assessments Sheet
    'Populates a specific table depending on the academic semester (i.e. time-frame)
    With complete
        If Date > 2021 - 1 - 1 Then '1A Semester
            tblSemesterRow = complete.ListObjects("Table2827").ListRows.Count 'Counts number of rows in table
            
            'If the table length is equal to 1 and the first entry of the table is blank
            'Then add the UserForm data into the first row
            If tblSemesterRow = 1 And complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
        
            Else
                 complete.ListObjects("Table2827").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2827").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If
            
        ElseIf Date >= 2021 - 1 - 11 And Date <= 2021 - 8 - 31 Then
            tblSemesterRow = complete.ListObjects("Table28").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table28").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table28").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If
        
        
        ElseIf Date >= 2021 - 9 - 1 And Date <= 2021 - 4 - 30 Then
            tblSemesterRow = complete.ListObjects("Table2821").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2821").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2821").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If

        ElseIf Date >= 2022 - 5 - 1 And Date <= 2022 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2822").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2822").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2822").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If

        ElseIf Date >= 2023 - 1 - 1 And Date <= 2023 - 8 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2823").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2823").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2823").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If


        ElseIf Date >= 2023 - 9 - 1 And Date <= 2024 - 4 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2824").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2824").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2824").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If

    
        ElseIf Date >= 2024 - 5 - 1 And Date <= 2024 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2825").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2825").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2825").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If

        ElseIf Date > 2024 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2826").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
        
            Else
                 complete.ListObjects("Table2826").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2826").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "A").Value = deliverables.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "B").Value = deliverables.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "C").Value = deliverables.Cells(indexNum + 2, 4).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "E").Value = deliverables.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "G").Value = deliverables.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "K").Value = deliverables.Cells(indexNum + 2, 7).Text
            End If
        End If
    End With
               
    'Removes the specific deliverable from the Deliverable Sheet
    With deliverables
        .Rows(indexNum + 2).Delete
    End With
             
    'Removes the deliverable from the table on the upcoming assessments sheet
    With upcoming
        Dim tblDeliverable As Long
        tblDeliverable = upcoming.ListObjects("Table24").ListRows.Count 'Number of rows in Deliverables Table
        
        'If length = 1 then fill in the top row with "-"
        If tblDeliverable = 1 Then
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "A").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "B").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "C").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "D").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "E").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "F").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "G").Value = " - "
            
                Range("Table24").HorizontalAlignment = xlCenter
                                
        Else
                upcoming.ListObjects("Table24").ListRows(indexNum + 1).Delete 'Delete the entire row from the table
                
        End If
    End With
    
End If

'If the user selects a test to complete
If CompleteAssessment_UserForm1.cmbTests.Value <> "" Then
    indexNum = CompleteAssessment_UserForm1.cmbTests.ListIndex
    Dim tblTests As Long
    tblTests = upcoming.ListObjects("Table1").ListRows.Count 'Number of rows in Tests Table
    
    If CompleteAssessment_UserForm2.txtSuggestHours.Value <> "" Then
        studyHours = CompleteAssessment_UserForm2.txtSuggestHours.Value
    Else
        studyHours = " - "
    End If
    
    
    With complete
        If Date > 2021 - 1 - 1 Then
            tblSemesterRow = complete.ListObjects("Table2827").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2827").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2827").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2827").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2827").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If
            
        ElseIf Date >= 2021 - 1 - 11 And Date <= 2021 - 8 - 31 Then
            tblSemesterRow = complete.ListObjects("Table28").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table28").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table28").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table28").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table28").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
                 
            End If
        
        
        ElseIf Date >= 2021 - 9 - 1 And Date <= 2021 - 4 - 30 Then
            tblSemesterRow = complete.ListObjects("Table2821").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2821").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2821").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2821").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2821").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If

        ElseIf Date >= 2022 - 5 - 1 And Date <= 2022 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2822").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2822").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2822").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2822").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2822").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If

        ElseIf Date >= 2023 - 1 - 1 And Date <= 2023 - 8 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2823").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2823").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2823").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2823").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2823").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If


        ElseIf Date >= 2023 - 9 - 1 And Date <= 2024 - 4 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2824").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2824").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2824").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2824").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2824").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If

    
        ElseIf Date >= 2024 - 5 - 1 And Date <= 2024 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2825").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2825").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2825").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2825").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2825").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If

        ElseIf Date > 2024 - 12 - 31 Then
            tblSemesterRow = complete.ListObjects("Table2826").ListRows.Count
        
            If tblSemesterRow = 1 And complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "A").Value = "-" Then
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "J").Value = studyHours
                 complete.ListObjects("Table2826").DataBodyRange.Cells(1, "K").Value = tests.Cells(indexNum + 2, 8).Text
        
            Else
                 complete.ListObjects("Table2826").ListRows.Add 'Add a new row to prevent re-writing data
                 tblSemesterRow = complete.ListObjects("Table2826").ListRows.Count 'Re-count the rows in the table
                 
                 'Add the data into the last row of the table (the row that just got added)
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "A").Value = tests.Cells(indexNum + 2, 2).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "B").Value = tests.Cells(indexNum + 2, 3).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "C").Value = tests.Cells(indexNum + 2, 5).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "D").Value = CompleteAssessment_UserForm2.txtClassAvg.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "E").Value = tests.Cells(indexNum + 2, 6).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "F").Value = CompleteAssessment_UserForm2.txtActualGrade.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "G").Value = tests.Cells(indexNum + 2, 7).Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "H").Value = CompleteAssessment_UserForm2.txtActualHoursStudy.Value
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "I").Value = CompleteAssessment_UserForm2.txtCommentSection.Text
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "J").Value = studyHours
                 complete.ListObjects("Table2826").DataBodyRange.Cells(tblSemesterRow, "K").Value = tests.Cells(indexNum + 2, 8).Text
            End If
        End If
    End With
                
    'Removes test from the hidden Test sheet
    With tests
        .Rows(indexNum + 2).Delete
    End With
         
    'Removes test from the Tests table on the Upcoming Assessments Sheet
    With upcoming
        If tblTests = 1 Then
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "A").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "B").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "C").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "D").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "E").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "F").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "G").Value = " - "
                upcoming.ListObjects("Table1").DataBodyRange.Cells(1, "H").Value = " - "
                                
        Else
                upcoming.ListObjects("Table1").ListRows(indexNum + 1).Delete 'Delete the entire row from the table
                
        End If
    End With
    
End If

End Sub
