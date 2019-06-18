Private Sub cb_forget_Click()
    Unload login
    forget.Show
End Sub

Private Sub cb_register_Click()
    Unload login
    register.Show
End Sub

Private Sub cb_login_Click()
    Dim rownum As Integer
    Dim key As Integer
    rownum = Cells(Rows.count, 1).End(xlUp).Row
    key = -1
    Worksheets("會員資料").Activate
    For i = 1 To rownum
        If unsecret(Cells(i, 1)) = tb_id.Text Then key = i
    Next
    If tb_id.Text = "" Or tb_password.Text = "" Then
        MsgBox ("帳號或密碼不得為空！")
    Else
        If key <> -1 Then
            If tb_password.Text = unsecret(Cells(key, 2)) Then
                tb_id.Text = ""
                tb_password.Text = ""
                Unload Me
                MsgBox (Cells(key, 3).Value & "，歡迎回來！")
                tuser = key
                tusername = Cells(key, 3).Value
                main.Show
            Else
                tb_password.Text = ""
                MsgBox ("密碼錯誤！")
            End If
        Else
            tb_id.Text = ""
            MsgBox ("無此用戶！")
        End If
    End If
End Sub
