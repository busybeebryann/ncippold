Imports System.Data.SqlClient
Imports System.Data
Public Class Add_Employee
    Dim age As Integer
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles bd.ValueChanged
        age = Date.Now.Year - bd.Value.Year
        If Date.Now.Month < bd.Value.Month Then
            age = age - 1
            lblage.Text = age
        Else
            lblage.Text = age
        End If
        lblage.BackColor = Color.White
    End Sub
    Dim gender As String = "none"
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Private Sub Add_Employee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pos()
        pos2()
        dept()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        gender = "Male"
        RadioButton1.BackColor = Color.White
        RadioButton2.BackColor = Color.White
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        gender = "Female"
        RadioButton1.BackColor = Color.White
        RadioButton2.BackColor = Color.White
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles em.TextChanged
        If em.Text.Contains("@") And em.Text.Contains(".com") Or em.Text.Contains(".net") Then
            em.ForeColor = Color.Green
        Else
            em.ForeColor = Color.Red
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFileDialog1.Filter = "image file (*.jpg, *.bmp, *.png) | *.jpg; *.bmp; *.png| all files (*.*) | *.* "
        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        insert()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim addemp As New Add_Employee
        addemp.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Public Sub pos()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_jobpos"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                position.Items.Add(acsdr("job_pos"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Dim totalcount As String
    Dim yearnow As String

    Public Sub pos2()
        yearnow = Now.Year
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT COUNT(*) as total from tbl_user"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                totalcount = acsdr("total")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
            totalcount = Val(totalcount) + 1
            If totalcount < 10 Then
                ei.Text = "PL" + yearnow + "000" + totalcount
            ElseIf totalcount > 10 And totalcount < 100 Then
                ei.Text = "PL" + yearnow + "00" + totalcount
            ElseIf totalcount >= 100 Then
                ei.Text = "PL" + yearnow + "0" + totalcount
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub dept()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_department"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                department.Items.Add(acsdr("dept_name"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub insert()
        If (ei.Text <> "" And fn.Text <> "" And ln.Text <> "" And mi.Text <> "" And gender <> "" And bd.Text <> "" And lblage.Text <> "Please Select Birthdate." And cn.Text <> "" And em.Text <> "" And add.Text <> "" And position.Text <> "" And department.Text <> "" And dh.Text <> "" And ms.Text <> "" And dp.Text <> "" And ss.Text <> "" And ph.Text <> "" And pi.Text <> "" And tn.Text <> "") Then
            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO tbl_user ([emp_id],[first_name],[last_name],[middle_name],[gender],[b_date],[age],[c_no],[email],[address],[job_pos],[dept],[date_hired],[marital_status],[dependents],[SSS],[philhealth],[pagibig],[TAX],[hr_pr_day],[basic_pay],[educ_loan],[province]) VALUES ('" + ei.Text + "','" + fn.Text + "','" + ln.Text + "','" + mi.Text + "','" + gender + "','" + bd.Text + "','" + lblage.Text + "','" + cn.Text + "','" + em.Text + "','" + add.Text + "','" + position.Text + "','" + department.Text + "','" + dh.Text + "','" + ms.Text + "','" + dp.Text + "','" + ss.Text + "','" + ph.Text + "','" + pi.Text + "','" + tn.Text + "','" + ComboBox1.Text + "','" + Bpay.Text + "','" + "Yes" + "','" + ComboBox3.Text + "')"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
                'MsgBox(SqlQuery)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try



            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO tbl_logindetails ([user_name],[emp_id],[pass_word],[access_lvl]) VALUES ('" + ei.Text + "','" + ei.Text + "','p@$$w0rd','2')"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
                PictureBox1.Image.Save(Application.StartupPath & "\images\" + ei.Text + ".jpg", Imaging.ImageFormat.Png)
                MessageBox.Show("Registration Successful", "Registration Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim addemp As New Add_Employee
                addemp.Show()
                Me.Close()

            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try

        Else
            MessageBox.Show("Please Answer All the fields with color red marker!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If fn.Text = "" Then
                fn.BackColor = Color.IndianRed
            End If
            If mi.Text = "" Then
                mi.BackColor = Color.IndianRed
            End If
            If ln.Text = "" Then
                ln.BackColor = Color.IndianRed
            End If
            If gender = "" Then
                RadioButton1.BackColor = Color.IndianRed
                RadioButton2.BackColor = Color.IndianRed
            End If
            If ei.Text = "" Then
                ei.BackColor = Color.IndianRed
            End If
            If lblage.Text = "Please Select Birthdate." Then
                lblage.BackColor = Color.IndianRed
            End If
            If ms.Text = "" Then
                ms.BackColor = Color.IndianRed
            End If
            If dp.Text = "" Then
                dp.BackColor = Color.IndianRed
            End If
            If cn.Text = "" Then
                cn.BackColor = Color.IndianRed
            End If
            If em.Text = "" Then
                em.BackColor = Color.IndianRed
            End If
            If add.Text = "" Then
                add.BackColor = Color.IndianRed
            End If
            If position.Text = "" Then
                position.BackColor = Color.IndianRed
            End If
            If department.Text = "" Then
                department.BackColor = Color.IndianRed
            End If
            If ComboBox1.Text = "" Then
                ComboBox1.BackColor = Color.IndianRed
            End If
            If tn.Text = "" Then
                tn.BackColor = Color.IndianRed
            End If
            If ss.Text = "" Then
                ss.BackColor = Color.IndianRed
            End If
            If ph.Text = "" Then
                ph.BackColor = Color.IndianRed
            End If
            If pi.Text = "" Then
                pi.BackColor = Color.IndianRed
            End If
        End If

    End Sub
    Private Sub fn_TextChanged(sender As Object, e As EventArgs) Handles fn.TextChanged
        fn.BackColor = Color.White
    End Sub

    Private Sub ln_TextChanged(sender As Object, e As EventArgs) Handles ln.TextChanged
        ln.BackColor = Color.White
    End Sub

    Private Sub mi_TextChanged(sender As Object, e As EventArgs) Handles mi.TextChanged
        mi.BackColor = Color.White
    End Sub

    Private Sub ms_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ms.SelectedIndexChanged
        ms.BackColor = Color.White
    End Sub

    Private Sub add_TextChanged(sender As Object, e As EventArgs) Handles add.TextChanged
        add.BackColor = Color.White
    End Sub

    Private Sub cn_TextChanged(sender As Object, e As EventArgs) Handles cn.TextChanged
        cn.BackColor = Color.White
    End Sub

    Private Sub dp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dp.SelectedIndexChanged
        dp.BackColor = Color.White
    End Sub

    Private Sub pi_TextChanged(sender As Object, e As EventArgs) Handles pi.TextChanged
        pi.BackColor = Color.White

        If pi.Text.Length = 4 Then
            pi.Text = pi.Text + "-"
            pi.Select(pi.Text.Length, 5)
        End If

        If pi.Text.Length = 9 Then
            pi.Text = pi.Text + "-"
            pi.Select(pi.Text.Length, 10)
        End If

        If pi.Text.Length = 14 Then
            pi.ForeColor = Color.Green

        Else
            pi.ForeColor = Color.Red
        End If

    End Sub
    Private Sub ph_TextChanged(sender As Object, e As EventArgs) Handles ph.TextChanged
        ph.BackColor = Color.White

        If ph.Text.Length = 2 Then
            ph.Text = ph.Text + "-"
            ph.Select(ph.Text.Length, 3)
        End If

        If ph.Text.Length = 13 Then
            ph.Text = ph.Text + "-"
            ph.Select(ph.Text.Length, 14)
        End If

        If ph.Text.Length = 15 Then
            ph.ForeColor = Color.Green
        Else
            ph.ForeColor = Color.Red
        End If

    End Sub
    Private Sub ss_TextChanged(sender As Object, e As EventArgs) Handles ss.TextChanged
        ss.BackColor = Color.White

        If ss.Text.Length = 2 Then
            ss.Text = ss.Text + "-"
            ss.Select(ss.Text.Length, 3)
        End If

        If ss.Text.Length = 13 Then
            ss.Text = ss.Text + "-"
            ss.Select(ss.Text.Length, 14)
        End If

        If ss.Text.Length = 15 Then
            ss.ForeColor = Color.Green
        Else
            ss.ForeColor = Color.Red
        End If


    End Sub
    Private Sub tn_TextChanged(sender As Object, e As EventArgs) Handles tn.TextChanged
        tn.BackColor = Color.White

        If tn.Text.Length = 3 Then
            tn.Text = tn.Text + "-"
            tn.Select(tn.Text.Length, 4)
        End If

        If tn.Text.Length = 7 Then
            tn.Text = tn.Text + "-"
            tn.Select(tn.Text.Length, 8)
        End If

        If tn.Text.Length = 11 Then
            tn.ForeColor = Color.Green

        Else
            tn.ForeColor = Color.Red
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.BackColor = Color.White
    End Sub
    Private Sub department_SelectedIndexChanged(sender As Object, e As EventArgs) Handles department.SelectedIndexChanged
        department.BackColor = Color.White
        position.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_jobpos WHERE department = '" + department.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                position.Items.Add(acsdr("job_pos"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub position_SelectedIndexChanged(sender As Object, e As EventArgs) Handles position.SelectedIndexChanged
        position.BackColor = Color.White

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_jobpos"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                Bpay.Text = acsdr("basic_pay")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog2.Filter = "image file (*.jpg, *.bmp, *.png) | *.jpg; *.bmp; *.png| all files (*.*) | *.* "
        If OpenFileDialog2.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            PictureBox2.Image = Image.FromFile(OpenFileDialog2.FileName)
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked

        If Panel1.Visible = False Then
            Panel1.Visible = True
            GroupBox1.Visible = False
            GroupBox2.Visible = False
            GroupBox3.Visible = False
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = False
            Label1.Visible = False
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If Panel1.Visible = True Then
            Panel1.Visible = False
            GroupBox1.Visible = True
            GroupBox2.Visible = True
            GroupBox3.Visible = True
            Button1.Visible = True
            Button2.Visible = True
            Button3.Visible = True
            Label1.Visible = True
            Label32.Visible = False
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If (TextBox1.Text = "") Then


            Label32.Text = "Please complete all fields!"
            Label32.Visible = True
        Else
            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO tbl_docs ([emp_id],[category],[description]) VALUES ('" + ei.Text.Trim + "','" + ComboBox2.Text.Trim + "','" + TextBox1.Text.Trim + "')"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()

                'TextBox2.Text = SqlQuery

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            PictureBox2.Image.Save(Application.StartupPath & "\docs\" + ei.Text + TextBox1.Text + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            TextBox1.Text = ""
            PictureBox2.Image = Nothing
            MsgBox("Saved Successfully!", MsgBoxStyle.Information)

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ComboBox2.SelectedIndex = 0
        PictureBox2.Image = Nothing
        TextBox1.Text = ""


    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class