Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Public Class Form1
    Dim CON1 As New SqlConnection
    Dim da, da_s As New SqlDataAdapter
    Dim ds, ds_s As New DataSet
    Dim com1, com2, com3 As SqlCommand
    Dim read As SqlCommand
    Dim serv, us, pas, daba, comstrl As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        On Error GoTo b
        If TextBox1.Text.Trim.Length < 4 Then
            MsgBox(" ابتدا فایل مورد نظر را انتخاب کنید", MsgBoxStyle.Critical, "خطا")
            TextBox1.Focus()
            Exit Sub
        End If
        daba = Cb1.Text
        comstrl = "Data Source=" & serv & ";Initial Catalog=" & daba & ";User ID=" & us & ";Password=" & pas & ""
        CON1 = New SqlConnection(comstrl)
        CON1.Open()
        com1 = New SqlCommand("ALTER DATABASE " & daba & " SET SINGLE_USER WITH ROLLBACK IMMEDIATE; use master RESTORE DATABASE " & daba & " FROM DISK ='" & TextBox1.Text & "' with replace ; ALTER DATABASE  " & daba & " SET  MULTI_USER WITH NO_WAIT ", CON1)
        com1.CommandTimeout = 0
        com1.ExecuteNonQuery()
        CON1.Close()
        MsgBox(" اطلاعات بازیابی شد", MsgBoxStyle.Information, "پیام")
        TextBox1.Text = ""
        Exit Sub
b:
        If Err.Number = 5 Then
            MsgBox(" وضعیت اتصال را بررسی کنید و سپس نام دیتابیس مورد نظر را انتخاب کنید", MsgBoxStyle.Critical, "خطا")
            TextBox1.Focus()
            Exit Sub
        End If
        MsgBox(Err.Description)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button6.Click
        On Error GoTo c
        Dim AdrBa As String
        FD.ShowDialog()
        AdrBa = FD.SelectedPath
        TextBox2.Text = AdrBa
        Exit Sub
c:
        MsgBox(Err.Description)
    End Sub
    Sub restoreback()
        Dim comstr As String = "Data Source=" & Cb1.Text & ";Initial Catalog=" & Cb1.Text & ";User ID=" & TextBox4.Text & ";Password=" & TextBox3.Text & ""
        CON1 = New SqlConnection(comstr)
        CON1.Open()
        da.SelectCommand = New SqlCommand()
        da.SelectCommand.Connection = CON1
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.CommandText = "Sp_Restorwithoff"
        da.SelectCommand.Parameters.AddWithValue("@patchFile", TextBox1.Text)
        ds.Clear()
        da.Fill(ds, "Sp_Restorwithoff")
        CON1.Close()

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error GoTo d
        If TextBox5.Text.Trim.Length < 0 Then
            MsgBox(" نام سرور را وارد کنید", MsgBoxStyle.Critical, "خطا")
            TextBox4.Focus()
            Exit Sub
        End If

        If TextBox4.Text.Trim.Length < 0 Then
            MsgBox(" نام کاربری را وارد کنید", MsgBoxStyle.Critical, "خطا")
            TextBox4.Focus()
            Exit Sub
        End If
        serv = TextBox5.Text
        us = TextBox4.Text
        pas = TextBox3.Text
        Cb1.Text = ""
        Dim comstr As String = "Data Source=" & TextBox5.Text & ";Initial Catalog=master ;User ID=" & TextBox4.Text & ";Password=" & TextBox3.Text & ""
        Loaddb()
        Exit Sub
d:
        If Err.Number = 5 Then
            MsgBox(" نام کاربری/نام سرور/کلمه عبور صحیح نیست .اتصال برقرار نشد", MsgBoxStyle.Critical, "خطا")
            TextBox4.Focus()
            Exit Sub
        End If
        MsgBox(Err.Description)
    End Sub
    Sub Loaddb()
        Dim sdr As SqlDataReader
        Dim rs As String
        Dim cont, i As Integer
        Dim comstr As String = "Data Source=" & TextBox5.Text & ";Initial Catalog=" & Cb1.Text & ";User ID=" & TextBox4.Text & ";Password=" & TextBox3.Text & ""
        CON1 = New SqlConnection(comstr)
        CON1.Open()
        ds_s.Clear()
        Cb1.DataBindings.Clear()
        da_s.SelectCommand = New SqlCommand("EXEC sp_databases", CON1)
        da_s.Fill(ds_s, "ip")
        Cb1.DataBindings.Add(New Binding("datasource", ds_s, "ip"))
        Cb1.DisplayMember = "database_Name"
        Cb1.FormattingEnabled = True
        Cb1.ValueMember = "remarks"
        CON1.Close()

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error GoTo s
        If TextBox2.Text.Trim.Length < 3 Then
            MsgBox(" مسیر ذخیره فایل را مشخص کنید", MsgBoxStyle.Critical, "خطا")
            TextBox2.Focus()
            Exit Sub
        End If
        daba = Cb1.Text
        comstrl = "Data Source=" & serv & ";Initial Catalog=" & daba & ";User ID=" & us & ";Password=" & pas & ""
        CON1 = New SqlConnection(comstrl)
        CON1.Open()
        com1 = New SqlCommand("BACKUP DATABASE " & daba & " TO DISK ='" & TextBox2.Text & "" & Dp1.Text & "'  WITH COMPRESSION ", CON1)
        com3 = New SqlCommand("Exec sp_configure 'backup compression default',1 ; Reconfigure with Override ;", CON1)
        com3.ExecuteNonQuery()
        com1.ExecuteNonQuery()

        CON1.Close()
        MsgBox(" تسخه پشتیبان تهیه شد", MsgBoxStyle.Information, "پیام")
        TextBox2.Text = ""
        Exit Sub
s:
        If Err.Number = 5 Then
            MsgBox(" وضعیت اتصال را بررسی کنید و سپس نام دیتابیس مورد نظر را انتخاب کنید", MsgBoxStyle.Critical, "خطا")
            TextBox1.Focus()
            Exit Sub
        End If
        MsgBox(Err.Description)
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim adrRe As String
        op.ShowDialog()
        adrRe = op.FileName
        TextBox1.Text = adrRe
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox5.Focus()
    End Sub
    Private Sub TextBox2_MouseHover(sender As Object, e As EventArgs) Handles TextBox2.MouseHover
        TextBox6.Visible = True
    End Sub

    Private Sub TextBox2_MouseLeave(sender As Object, e As EventArgs) Handles TextBox2.MouseLeave
        TextBox6.Visible = False
    End Sub
End Class
