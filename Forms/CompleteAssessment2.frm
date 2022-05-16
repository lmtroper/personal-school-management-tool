Private Sub cmbCommentHelp_Click() 'Provides user with sample questions to answer and make comments on

    MsgBox ("Here are some example questions you can comment on:" & vbCr & vbCr & "What you did well?" & vbCr & "What you did not do so well?" & vbCr & _
    "What you would do better if you had a do-over?" & vbCr & "Where did you study?" & vbCr & "How much sleep did you get?")

End Sub

Private Sub cmbReset_Click() 'Clears values in UserForm

    txtActualGrade.Text = ""
    txtActualHoursStudy.Text = ""
    txtClassAvg.Text = ""
    txtCommentSection.Text = ""

End Sub

Private Sub cmbSubmit_Click()

'Check for required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    If txtActualGrade.Text = "" Then
        msgValue = MsgBox("Please enter your grade")
            
    ElseIf Not IsNumeric(txtActualGrade.Text) Then
        msgValue = MsgBox("Please enter a valid grade")
            
    ElseIf txtActualHoursStudy.Text = "" Then
        msgValue = MsgBox("Please enter the amount of hours you studied for")
            
    ElseIf Not IsNumeric(txtActualHoursStudy.Text) Then
        msgValue = MsgBox("Please enter a valid number of study hours")
            
    ElseIf Not IsNumeric(txtClassAvg.Text) And txtClassAvg <> "" Then
        msgValue = MsgBox("Please enter a valid grade for the class average")

    Else
        'User Confirmation before assessment is removed and placed into Completed_Assessments sheet
        msgValue = MsgBox("Do you want to update this task as completed?" & vbCr & vbCr & "(The assessment will be moved to the Completed Assessments Sheet)", vbYesNo + vbInformation, "Confirmation")
        
        If msgValue = vbNo Then Exit Sub
            
        Call Submit_CompleteForm2
        Unload Me
        Unload CompleteAssessment_UserForm1
    
    End If
    
End Sub

Private Sub cmbSubmit_Enter()

'Check for required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    If txtActualGrade.Value = "" Then
        msgValue = MsgBox("Please enter your grade")
            
    ElseIf Not IsNumeric(txtActualGrade.Value) Or txtActualGrade.Value < 0 Then
        msgValue = MsgBox("Please enter a valid grade")
            
    ElseIf txtActualHoursStudy.Text = "" Then
        msgValue = MsgBox("Please enter the amount of hours you studied for")
            
    ElseIf Not IsNumeric(txtActualHoursStudy.Text) Or txtActualHoursStudy.Value < 0 Then
        msgValue = MsgBox("Please enter a valid number of study hours")
            
    ElseIf Not IsNumeric(txtClassAvg.Text) And txtClassAvg <> "" Then
        msgValue = MsgBox("Please enter a valid grade for the class average")

    Else
        'User Confirmation before assessment is removed and placed into Completed_Assessments sheet
        msgValue = MsgBox("Do you want to update this task as completed?" & vbCr & vbCr & "(The assessment will be moved to the Completed Assessments Sheet)", vbYesNo + vbInformation, "Confirmation")
        
        If msgValue = vbNo Then Exit Sub
            
        Call Submit_CompleteForm2
        Unload Me
        Unload CompleteAssessment_UserForm1
    
    End If


End Sub

Private Sub cmdSubmit_Click()

Call Submit_CompleteForm2
Unload Me
Unload CompleteAssessment_UserForm1

End Sub


Private Sub txtActualHoursStudy_Change()

Dim deliverables As Worksheet, tests As Worksheet, upcoming As Worksheet

Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")

Dim studyHours As Variant, percentperHour As Variant, indexNum As Long
indexNum = CompleteAssessment_UserForm1.cmbDeliverables.ListIndex
indexNum2 = CompleteAssessment_UserForm1.cmbTests.ListIndex

'If the user is completing a deliverable:
If CompleteAssessment_UserForm1.cmbDeliverables.Value <> "" Then

'actgrade is what the user actual receives, antgrade is the users expected grade
actgrade = txtActualGrade.Value
antgrade = deliverables.Cells(indexNum + 2, 5).Value

'difference between actgrade and antgrade
difference = antgrade - actgrade

    'If the user inputs an actual grade and # of hours studying that does not equal 0
    If txtActualGrade.Value <> "" And Me.txtActualHoursStudy <> "" And Me.txtActualHoursStudy <> 0 Then
    
        'If difference is greater than 0, actual grade is lower than expected grade
        'We want an algorithm to suggest how many more hours they should have spent studying
        'Using linear approximation
        If difference > 0 Then
            Me.txtSuggestHours.Enabled = True
            Me.txtSuggestHours.BackColor = vbWhite
            
            'Find percent per hour by dividing actual grade by number of hours studied
            percentperHour = CompleteAssessment_UserForm2.txtActualGrade.Value / CompleteAssessment_UserForm2.txtActualHoursStudy.Value
            
            '(Expected Grade - Actual Grade)/percent per hour
            studyHours = ((deliverables.Cells(indexNum + 2, 5).Value) - CompleteAssessment_UserForm2.txtActualGrade.Value) / percentperHour
            
            txtSuggestHours = Round(studyHours, 1)
        
        'If difference is less than or equal to 0, user's actual grade is equal or greater than expected
        'Does not need a suggestion so enable the text box
        ElseIf difference <= 0 Then
            Me.txtSuggestHours.Value = ""
            Me.txtSuggestHours.Enabled = False
            Me.txtSuggestHours.BackColor = vbGrey
            
        End If
    End If

'If the user is completing a test:
ElseIf CompleteAssessment_UserForm1.cmbTests.Value <> "" Then

actgrade = txtActualGrade.Value
antgrade = tests.Cells(indexNum2 + 2, 6).Value
difference = antgrade - actgrade

    If txtActualGrade.Value <> "" And Me.txtActualHoursStudy <> "" And Me.txtActualHoursStudy <> 0 Then
    
        If difference > 0 Then
            Me.txtSuggestHours.Value = ""
            Me.txtSuggestHours.Enabled = True
            Me.txtSuggestHours.BackColor = vbWhite
            percentperHour = CompleteAssessment_UserForm2.txtActualGrade.Value / CompleteAssessment_UserForm2.txtActualHoursStudy.Value
            
            studyHours = ((tests.Cells(indexNum2 + 2, 7).Value) - CompleteAssessment_UserForm2.txtActualGrade.Value) / percentperHour
            txtSuggestHours = Round(studyHours, 1)
            
        ElseIf difference <= 0 Then
            Me.txtSuggestHours.Value = ""
            Me.txtSuggestHours.Enabled = False
            Me.txtSuggestHours.BackColor = vbGrey
            
        End If
    End If
End If
    
End Sub


Private Sub UserForm_Initialize()

'The deliverable or test selected is placed in a textbox at the top
If CompleteAssessment_UserForm1.cmbDeliverables.Value <> "" Then
     txtAssessment.Value = CompleteAssessment_UserForm1.cmbDeliverables.Value
        
    ElseIf CompleteAssessment_UserForm1.cmbTests.Value <> "" Then
        txtAssessment.Value = CompleteAssessment_UserForm1.cmbTests.Value
End If
        
End Sub
