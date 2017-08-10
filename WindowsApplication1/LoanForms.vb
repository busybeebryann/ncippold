Imports System.Data.SqlClient

Public Class LoanForms
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim count As Integer
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ListBox1.Items.Clear()
        If TextBox1.Text = "" Then
            ListBox1.Visible = False
        Else
            ListBox1.Visible = True
            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM tbl_user WHERE emp_id LIKE '%" + TextBox1.Text + "%'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())
                    ListBox1.Items.Add(acsdr("emp_id"))
                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()
                count = ListBox1.Items.Count
                ListBox1.Height = count * 25
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM tbl_user WHERE emp_id = '" + TextBox1.Text + "'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())
                    Label4.Text = acsdr("first_name").trim + " " + acsdr("middle_name").trim + " " + acsdr("last_name").trim
                    Label6.Text = acsdr("date_hired").trim
                    Label8.Text = acsdr("marital_status").trim
                    Label10.Text = acsdr("job_pos").trim
                    Label12.Text = acsdr("dept").trim
                    Label14.Text = acsdr("basic_pay").trim
                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()
                count = ListBox1.Items.Count
                ListBox1.Height = count * 25
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

    End Sub
    Dim bpay As Double

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        TextBox1.Text = ListBox1.Text
        ListBox1.Visible = False
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_user WHERE emp_id = '" + TextBox1.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                Label4.Text = acsdr("first_name").trim + " " + acsdr("middle_name").trim + " " + acsdr("last_name").trim
                Label6.Text = acsdr("date_hired").trim
                Label8.Text = acsdr("marital_status").trim
                Label10.Text = acsdr("job_pos").trim
                Label12.Text = acsdr("dept").trim
                bpay = acsdr("basic_pay").trim
                Label14.Text = "Php " + bpay.ToString("0,00.00")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
            count = ListBox1.Items.Count
            ListBox1.Height = count * 25
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Dim type As String

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        ComboBox3.Visible = False
        type = "Pag-IBIG Housing"
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton1.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        ComboBox3.Visible = False
        type = "NCIPAE Loan"
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton1.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        ComboBox3.Visible = False
        type = "Pag-IBIG Loan"
        RadioButton2.Checked = False
        RadioButton1.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        ComboBox3.Visible = True
        type = "Conso Loan"
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton1.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox3.Text = 6
        ComboBox3.Visible = False
        type = "Emergency Loan"
        RadioButton1.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        ComboBox3.Visible = False
        TextBox3.Text = 8
        type = "Policy Loan"
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
    End Sub
    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        ComboBox3.Visible = False
        TextBox3.Text = 0
        type = "LBP Loan"
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton1.Checked = False
    End Sub
    Dim interest As Double
    Dim finalAm As Double
    Dim months As Integer
    Dim finalMP As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox3.Text = "" Then
                TextBox3.Text = 0
            End If
            Label24.Text = TextBox2.Text
            Label25.Text = TextBox3.Text + " %"

            interest = Val(Label25.Text) / 100
            interest = Val(interest) * Label24.Text
            finalAm = Val(interest) + Val(Label24.Text)
            Label24.Text = finalAm.ToString("0,00.00") + " Php"
            If ComboBox1.Text = "" Then
                months = Val(ComboBox2.Text) * 12
                Label26.Text = Math.Round(Val(finalAm) / months, 2).ToString("0,00.00") + " Php"
                finalMP = Math.Round(Val(finalAm) / months, 2)
            Else
                months = Val(ComboBox1.Text)
                Label26.Text = Math.Round(Val(finalAm) / ComboBox1.Text, 2).ToString("0,00.00") + " Php"
                finalMP = Math.Round(Val(finalAm) / ComboBox1.Text, 2)
            End If
        Catch ex As Exception
            MsgBox("Wrong Amount Inputted Please Double check")
        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox2.Text = Val(ComboBox3.Text) * Val(bpay)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim lf As New LoanForms
        lf.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "INSERT INTO tbl_loans(emp_id,loan_type,interest,months_to_pay,payment_amount,interest_rate)VALUES('" + TextBox1.Text + "','" + type.ToString + "','" + TextBox3.Text + "','" + months.ToString + "','" + finalMP.ToString + "','" + ComboBox4.Text + "')"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
            Name = ""
            MessageBox.Show("Adding of job loan for Employee: " + Label4.Text + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub LoanForms_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class