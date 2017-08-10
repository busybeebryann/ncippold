Imports System.Data.SqlClient

Public Class update_profile

    Dim age As Integer
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles bd.ValueChanged
        age = Date.Now.Year - bd.Value.Year
        If Date.Now.Month < bd.Value.Month Then
            age = age - 1
            lblage.Text = age
        Else
            lblage.Text = age
        End If
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim gen As String
    Dim emp_id As String
    Private Sub update_profile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user JOIN tbl_logindetails ON tbl_user.emp_id=tbl_logindetails.emp_id WHERE user_name='" + Login.TextBox1.Text.Trim + "'"


            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                Label1.Text = "Update profile of " + acsdr("first_name").trim
                emp_id = acsdr("emp_id").trim
                fn.Text = acsdr("first_name").trim
                ln.Text = acsdr("last_name").trim
                mi.Text = acsdr("middle_name").trim
                gen = acsdr("gender").trim
                bd.Text = acsdr("b_date").trim
                lblage.Text = acsdr("age").trim
                ms.Text = acsdr("marital_status").trim
                dp.Text = acsdr("dependents").trim
                cn.Text = acsdr("c_no").trim
                em.Text = acsdr("email").trim
                add.Text = acsdr("address").trim
                tn.Text = acsdr("TAX").trim
                ss.Text = acsdr("SSS").trim
                ph.Text = acsdr("philhealth").trim
                pi.Text = acsdr("pagibig").trim
                pw.Text = acsdr("pass_word").trim
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        If gen = "Male" Then
            RadioButton1.Select()
        Else
            RadioButton2.Select()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "UPDATE tbl_user SET first_name='" + fn.Text + "',last_name='" + ln.Text + "',middle_name='" + mi.Text + "',gender='" + gen + "',b_date='" + bd.Text + "',age='" + lblage.Text + "',marital_status='" + ms.Text + "',dependents='" + dp.Text + "',c_no='" + cn.Text + "',email='" + em.Text + "',address='" + add.Text + "',TAX='" + tn.Text + "',SSS='" + ss.Text + "',philhealth='" + ph.Text + "',pagibig='" + pi.Text + "' WHERE emp_id='" + emp_id + "'"

            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
            'TextBox1.Text = SqlQuery'

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "UPDATE tbl_logindetails SET pass_word='" + pw.Text + "' WHERE emp_id='" + emp_id + "'"

            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
            'TextBox1.Text = SqlQuery'

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        MessageBox.Show("Personal Information Updated Successful!", "Sucess", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        gen = "Male"
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        gen = "Female"
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

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

End Class