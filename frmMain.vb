

Private Sub cmdRegister_Click()
    frmMain.Hide
    frmRegistration.Show
End Sub

Private Sub cmdSubmit_Click()
Dim s, m, h As Integer
Dim sx, mx, hx As Integer
Dim sy, my, hy As Integer

    If AdoReg.Recordset.EOF = True Then
        a = MsgBox("The code you entered was either incorrect or not yet registered!", vbOKOnly + vbInformation)
        txtEmpCode.Text = ""
        txtPassWrd.Text = ""
    ElseIf txtPassWrd.Text = "" Or txtEmpCode.Text = "" Then
        d = MsgBox("Please enter required information!", vbOKOnly + vbCritical)
    ElseIf txtPassWrd.Text <> AdoReg.Recordset.Fields!PWord Then
        b = MsgBox("Password Incorrect!", vbOKOnly + vbCritical)
        txtEmpCode.Text = ""
        txtPassWrd.Text = ""
    Else
        If AdoReg.Recordset.Fields!atStatus = False Then
            AdoReg.Recordset.Fields!atStatus = True
            AdoReg.Recordset.Update
            AdoDisplay.Recordset.Fields!timei = TimeValue(Now)
            AdoDisplay.Recordset.Fields!timeo = Null
            AdoDisplay.Recordset.Fields!rday = MonthName(Month(Now)) & " " & Day(Now)
            AdoDisplay.Recordset.Update
            c = MsgBox("Time In Successful!", vbOKOnly + vbInformation)
        Else
            sx = Second(AdoDisplay.Recordset.Fields!timei)
            mx = Minute(AdoDisplay.Recordset.Fields!timei)
            hx = Hour(AdoDisplay.Recordset.Fields!timei)
            AdoReg.Recordset.Fields!atStatus = False
            AdoReg.Recordset.Update
            AdoDisplay.Recordset.Fields!timeo = TimeValue(Now)
            sy = Second(AdoDisplay.Recordset.Fields!timeo)
            my = Minute(AdoDisplay.Recordset.Fields!timeo)
            hy = Hour(AdoDisplay.Recordset.Fields!timeo)
            
            If sy >= sx Then
                s = sy - sx
            Else
                my = my - 1
                s = (sy + 60) - sx
            End If
            
            If s <= 9 Then
                    s = "0" & s
            End If
            
            If my >= sx Then
                m = my - mx
            Else
                hy = hy - 1
                m = (my + 60) - mx
            End If
            
            If m <= 9 Then
                m = "0" & m
            End If
            
            h = hy - hx
            
            If h <= 9 Then
                h = "0" & h
            End If
            
            AdoDisplay.Recordset.Fields!numhours = h & ":" & m & ":" & s
            AdoDisplay.Recordset.Update
            c = MsgBox("Time Out Successful!", vbOKOnly + vbInformation)
        End If

        txtEmpCode.Text = ""
        txtPassWrd.Text = ""
    End If
End Sub


Private Sub Form_Load()
    lblDay.Caption = WeekdayName(Weekday(Now))
    lblYear.Caption = Year(Now)
    lblMonthDay.Caption = MonthName(Month(Now)) & " " & Day(Now)
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = Time
End Sub

Private Sub txtEmpCode_Change()
    AdoReg.RecordSource = "select * from tblAttendance where EmpNum ='" & txtEmpCode.Text & "'"
    AdoReg.Refresh
    AdoDisplay.RecordSource = "select * from tblDisplay where eCode like'%" & txtEmpCode.Text & "%'"
    AdoDisplay.Refresh
End Sub
