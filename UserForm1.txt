Private Sub UserForm_Initialize()
    ' Populate concern dropdown
    cmbConcern.AddItem "Adding Dependent"
    cmbConcern.AddItem "Appeals & Grievances"
    cmbConcern.AddItem "Cancellation"
    cmbConcern.AddItem "Reactivation"
    cmbConcern.AddItem "Reapplication"
    cmbConcern.AddItem "Removing Dependent"
    cmbConcern.AddItem "Reinstatement"
    cmbConcern.AddItem "Request VOC"
    cmbConcern.AddItem "Withdrawal"
End Sub
Private Sub CheckCICStatus()
    If checkBoxFName.Value And checkBoxLName.Value And checkBoxDOB.Value And checkBoxAddress.Value Then
        labelCICVerificationStatus.Caption = "Verified"
        labelCICVerificationStatus.ForeColor = vbGreen
        btnLoad.Enabled = True
    Else
        labelCICVerificationStatus.Caption = "NOT Verified"
        labelCICVerificationStatus.ForeColor = vbRed
        btnLoad.Enabled = False
    End If
End Sub
Private Sub checkBoxFName_Click()
    Call CheckCICStatus
End Sub
Private Sub checkBoxLName_Click()
    Call CheckCICStatus
End Sub
Private Sub checkBoxDOB_Click()
    Call CheckCICStatus
End Sub
Private Sub checkBoxAddress_Click()
    Call CheckCICStatus
End Sub
Private Sub txtName_AfterUpdate()
    If Len(txtName.Value) > 0 Then
        checkBoxFName.Value = True
    Else
        checkBoxFName.Value = False
    End If
End Sub
Private Sub checkBoxLockSubID_Click()
    If checkBoxLockSubID.Value Then
        txtSubID.Locked = True
    Else
        txtSubID.Locked = False
    End If
End Sub
Private Sub btnLoad_Click()
    Dim callerName As String
    Dim concern As String
    Dim scriptTemplate As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    If Trim(txtName.Value) = "" Then
        callerName = "Maam/Sir"
    Else
        callerName = txtName.Value
    End If
    
    concern = cmbConcern.Value
    
    If concern = "Adding Dependent" Then
        AddDependentForm.Show
    Else
        Set ws = ThisWorkbook.Sheets("Scripts")
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Search for concern in Scripts sheet
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = concern Then
                scriptTemplate = ws.Cells(i, 2).Value
                Exit For
            End If
        Next i
        
        ' Replace placeholder with caller name
        scriptTemplate = Replace(scriptTemplate, "{Name}", callerName)
        
        ' Display script
        txtScript.Value = scriptTemplate
    End If
End Sub
