Private Sub UserForm_Initialize()
    ' Populate dependent relationship dropdown
    cmbDependentRel.AddItem "Husband"
    cmbDependentRel.AddItem "Wife"
    cmbDependentRel.AddItem "Son"
    cmbDependentRel.AddItem "Daughter"
    cmbDependentRel.AddItem "Father"
    cmbDependentRel.AddItem "Mother"
    
    cmbDependentPlan.AddItem "Critical Guard"
    cmbDependentPlan.AddItem "Critical Illness"
    cmbDependentPlan.AddItem "Dental"
    cmbDependentPlan.AddItem "Health ProtectorGuard 1"
    cmbDependentPlan.AddItem "Health ProtectorGuard 2"
    cmbDependentPlan.AddItem "Health ProtectorGuard 3"
    cmbDependentPlan.AddItem "HealthiestYou"
    cmbDependentPlan.AddItem "Mental"
    cmbDependentPlan.AddItem "New Benefits"
    cmbDependentPlan.AddItem "VisionWise"
End Sub
Private Sub btnLoad_Click()
    If cmbDependentPlan.Value = "Dental" Or cmbDependentPlan.Value = "HealthiestYou" Or cmbDependentPlan.Value = "New Benefits" Or cmbDependentPlan.Value = "New Benefits" Then
        labelTest.Caption = "Plan does NOT need Form." & vbNewLine & "1. Go to StepWise and Calculate new estimated rate quote."
    Else
        labelTest.Caption = "Needs Form. Send via e-mail."
    End If
End Sub
