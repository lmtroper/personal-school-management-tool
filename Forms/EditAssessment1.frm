Private Sub UserForm_Initialize() 'Resets the list-box with an updated list of assessments

Dim deliverables As Worksheet, tests As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbTests.Clear
cmbDeliverables.Clear

For i = 2 To iRow
    cmbDeliverables.AddItem deliverables.Cells(i, 2) & ": " & deliverables.Cells(i, 3)
Next
        
For i = 2 To iRow2
    cmbTests.AddItem tests.Cells(i, 2) & ": " & tests.Cells(i, 3)
Next

cmbTests.Value = ""
cmbDeliverables.Value = ""

End Sub

Private Sub cmbDeliverables_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbDeliverables.SpecialEffect = fmSpecialEffectEtched
cmbTests.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub cmbTests_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbTests.SpecialEffect = fmSpecialEffectEtched
cmbDeliverables.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbDeliverables.SpecialEffect = fmSpecialEffectFlat
cmbTests.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub cmbDeliverables_Change()

'If User selects a value in the Deliverables combo-box, erase any values in the Tests combo-box
'Ensures User only deletes one assessment

If Me.cmbDeliverables.Value <> "" Then
    Me.cmbTests.Enabled = False
End If

End Sub
Private Sub cmbTests_Change()

'If User selects a value in the Tests combo-box, erase any values in the Deliverables combo-box
'Ensures User only deletes one assessment

If Me.cmbTests.Value <> "" Then
    Me.cmbDeliverables.Enabled = False
End If

End Sub

Private Sub cmdSelect_Click()

EditAssessment_UserForm1.Hide

'If user selects a deliverable to edit:
If cmbDeliverables.Value <> "" Then
    EditDeliverable_UserForm.Show
End If

'If user selects a test to edit:
If cmbTests.Value <> "" Then
    EditTest_UserForm.Show
End If

End Sub

Private Sub cmdClear_Click()

cmbTests.Value = ""
cmbTests.Enabled = True
cmbDeliverables.Enabled = True
cmbDeliverables.Value = ""

End Sub
