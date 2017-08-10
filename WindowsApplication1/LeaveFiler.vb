Imports System.Data.SqlClient

Public Class LeaveFiler
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
                ListBox1.Height = count * 15
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
                    Label6.Text = acsdr("first_name").trim + " " + acsdr("middle_name").trim + " " + acsdr("last_name").trim
                    Label7.Text = acsdr("marital_status").trim
                    Label8.Text = acsdr("job_pos").trim
                    Label9.Text = acsdr("dept").trim
                    bpay = acsdr("basic_pay").trim
                    Label10.Text = "Php " + bpay.ToString("0,00.00")
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
                Label6.Text = acsdr("first_name").trim + " " + acsdr("middle_name").trim + " " + acsdr("last_name").trim
                Label7.Text = acsdr("marital_status").trim
                Label8.Text = acsdr("job_pos").trim
                Label9.Text = acsdr("dept").trim
                bpay = acsdr("basic_pay").trim
                Label10.Text = "Php " + bpay.ToString("0,00.00")
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label16.Text = Label6.Text
        Label17.Text = TextBox1.Text
        Label18.Text = Label8.Text
        Label19.Text = Label9.Text
        Label20.Text = Label7.Text
        Label37.Text = DateTimePicker1.Text
        Label38.Text = Label34.Text

        Label21.Text = ComboBox1.Text
        Label22.Text = ComboBox2.Text
        Label23.Text = ComboBox3.Text
    End Sub

    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click

    End Sub
    Dim datenow As DateTime
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        datenow = DateTimePicker1.Text
        Label34.Text = datenow.AddDays(Val(ComboBox2.Text) - 1)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim lf As New LeaveFiler
        lf.Show()
        Me.Close()
    End Sub
    Dim totL As Integer
    Dim countl As Integer
    Dim hw As Integer
    Dim dateL As DateTime
    Dim hww As String


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            totL = Val(ComboBox2.Text)
            countl = Val(ComboBox2.Text)
            dateL = DateTimePicker1.Text
            hw = Val(ComboBox3.Text)
            For i As Integer = 0 To totL - 1
                If hw >= 1 Then
                    hww = "8"
                ElseIf hw = 0 Then
                    hww = "0"
                End If
                If totL = countl Then
                    Try
                        Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                        Dim con As New SqlConnection
                        con.ConnectionString = provider
                        If con.State = ConnectionState.Closed Then
                            con.Open()
                        End If
                        Dim SqlCommand As New SqlCommand
                        Dim SqlQuery As String = "INSERT INTO [tbl_AttPayroll] ([emp_id],[time_in],[time_out],[reg_hr],[OT],[ND],[att_date],[status]) VALUES ('" + TextBox1.Text + "','Leave','Leave','" + hww + "','0','0','" + dateL + "','0')"
                        With SqlCommand
                            .CommandText = SqlQuery
                            .Connection = con
                            .ExecuteNonQuery()
                        End With
                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                Else
                    dateL = dateL.AddDays(1)
                    Try
                        Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                        Dim con As New SqlConnection
                        con.ConnectionString = provider
                        If con.State = ConnectionState.Closed Then
                            con.Open()
                        End If
                        Dim SqlCommand As New SqlCommand
                        Dim SqlQuery As String = "INSERT INTO [tbl_AttPayroll] ([emp_id],[time_in],[time_out],[reg_hr],[OT],[ND],[att_date],[status]) VALUES ('" + TextBox1.Text + "','Leave','Leave','" + hww + "','0','0','" + dateL + "','0')"
                        With SqlCommand
                            .CommandText = SqlQuery
                            .Connection = con
                            .ExecuteNonQuery()
                        End With
                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
                countl = countl - 1
                hw = hw - 1

            Next
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
            Dim SqlQuery As String = "INSERT INTO [tbl_leave] ([emp_id],[leave_type],[no_days],[no_paid],[s_d],[e_d]) VALUES ('" + TextBox1.Text + "','" + ComboBox1.Text + "','" + ComboBox2.Text + "','" + ComboBox3.Text + "','" + DateTimePicker1.Text + "','" + Label34.Text + "')"

            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        MessageBox.Show("Leave successfully compiled!", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub LeaveFiler_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class