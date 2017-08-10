Imports System.Data.SqlClient

Public Class attendance_overide
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String

    Private Sub attendance_overide_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ComboBox1.Items.Add(acsdr("emp_id"))

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        admin()
    End Sub
    Dim adminpass As String
    Public Sub admin()
        Try

            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user JOIN tbl_logindetails ON tbl_user.user_id=tbl_logindetails.acc_id WHERE job_pos='Administrator' AND dept='Administrator'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                adminpass = acsdr("pass_word")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Dim d1 As DateTime
    Dim d2 As DateTime
    Dim NH As String
    Dim OT As String
    Dim ND As String
    Dim realval As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(DateTimePicker3.Text + " - " + DateTimePicker4.Text)
        d1 = DateTimePicker3.Text
        d2 = DateTimePicker4.Text
        realval = DateDiff("s", d1, d2) / 60 / 60
        If realval >= 9 Then
            NH = "8"
            OT = Val(realval) - 9
            ND = "0"
        ElseIf d1 >= #10:00:00 PM# Then
            realval = DateDiff("s", d2, d1)

            realval = Val(realval) / 60 / 60 / 2

            If realval >= 9 Then
                ND = "8"
                OT = Val(realval) - 9
                NH = "0"
            ElseIf realval < 9 Then
                ND = realval
                OT = "0"
                NH = "0"
            End If
        Else
            NH = realval
            OT = "0"
            ND = "0"
        End If
        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "INSERT INTO [tbl_AttPayroll] ([emp_id],[time_in],[time_out],[reg_hr],[OT],[ND],[att_date],[status]) VALUES ('" + ComboBox1.Text + "','" + d1 + "','" + d2 + "','" + NH + "','" + OT + "','" + ND + "','" + DateTimePicker1.Text + "','0')"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
        MessageBox.Show("Attendace for employee number " + ComboBox1.Text.Trim + " is added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


End Class