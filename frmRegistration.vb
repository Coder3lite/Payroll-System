Private Sub cmdBack_Click()
    frmRegistration.Hide
    frmMain.Show
End Sub

Private Sub cmdSubmit_Click()
    If (txtEmpC.Text <> "" And txtLName.Text <> "" And txtFName.Text <> "" And txtPWord1.Text <> "" And txtPWord2.Text <> "") And (txtPWord1.Text = txtPWord2.Text) Then
        x = MsgBox("Registration Complete!!", vbInformation + vbOKOnly)
            
            frmMain.AdoReg.Recordset.AddNew
            frmMain.AdoReg.Recordset.Fields!EmpNum = txtEmpC.Text
            frmMain.AdoReg.Recordset.Fields!LastName = txtLName.Text
            frmMain.AdoReg.Recordset.Fields!FirstName = txtFName.Text
            frmMain.AdoReg.Recordset.Fields!PWord = txtPWord1.Text
            frmMain.AdoReg.Recordset.Fields!atStatus = False
            frmMain.AdoReg.Recordset.Update
            
            frmMain.AdoDisplay.Recordset.AddNew
            frmMain.AdoDisplay.Recordset.Fields!eCode = txtEmpC.Text
            frmMain.AdoDisplay.Recordset.Fields!lastN = txtLName.Text
            frmMain.AdoDisplay.Recordset.Fields!firstN = txtFName.Text
            frmMain.AdoDisplay.Recordset.Update
            
            If vbOKOnly = clicked Then
                txtEmpC.Text = ""
                txtLName.Text = ""
                txtFName.Text = ""
                txtPWord1.Text = ""
                txtPWord2.Text = ""
                frmRegistration.Hide
                frmMain.Show
            End If
    Else
        a = MsgBox("Some field might be empty or Password do not match!", vbCritical + vbOKOnly)
            If vbOKOnly = clicked Then
                txtPWord1.Text = ""
                txtPWord2.Text = ""
            End If
    End If
End Sub
