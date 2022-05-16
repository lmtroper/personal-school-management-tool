Sub Show_DeleteUF()

DeleteAssessment_UserForm.Show

End Sub

Sub Delete_Assessment() 'Delete an Assessment from the Workbook

Dim startCell As Range, indexNum As Long

'Declare the worksheets that will be referenced
Dim upcoming As Worksheet, complete As Worksheet, deliverables As Worksheet, tests As Worksheet
Set upcoming = ThisWorkbook.Sheets("Upcoming_Assessments")
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")

'If user selects a deliverable to delete:
If DeleteAssessment_UserForm.cmbDeliverables.Value <> "" Then

    indexNum = DeleteAssessment_UserForm.cmbDeliverables.ListIndex 'Element number of value selected on combo-box
    Dim tblDeliverable As Long
    tblDeliverable = upcoming.ListObjects("Table24").ListRows.Count 'Number of rows in Deliverable Table
    
    'Delete deliverable from Deliverables Sheet
    With deliverables
        .Rows(indexNum + 2).Delete
    End With
              
   'Delete deliverable from Deliverables Table on Upcoming Assessments Sheet
    With upcoming
        If tblDeliverable = 1 Then
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "A").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "B").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "C").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "D").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "E").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "F").Value = " - "
                upcoming.ListObjects("Table24").DataBodyRange.Cells(1, "G").Value = " - "

                                
        Else
                upcoming.ListObjects("Table24").ListRows(indexNum + 1).Delete 'Delete the entire row from the table
                
        End If
    End With
    
End If

    
'If user selects a test to delete:
If DeleteAssessment_UserForm.cmbTests.Value <> "" Then
    indexNum = DeleteAssessment_UserForm.cmbTests.ListIndex 'Element number of value selected on combo-box
    Dim tblTests As Long
    tblTests = upcoming.ListObjects("Table1").ListRows.Count 'Number of rows in Tests Table
            
    'Delete test from Test Sheet
    With tests
        .Rows(indexNum + 2).Delete
    End With
           
    'Delete test from Deliverables Table on Upcoming Assessments Sheet
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
