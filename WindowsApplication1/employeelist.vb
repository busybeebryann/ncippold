Imports System.Data.SqlClient
Imports System.Data
Public Class employeelist
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""

    Private Sub employeelist_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
        pos()
        dept()
    End Sub
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT emp_id AS Employee_id, first_name,last_name,job_pos,dept FROM tbl_user WHERE emp_id LIKE '%" + TextBox1.Text + "%' OR first_name LIKE '%" + TextBox1.Text + "%' OR last_name LIKE '%" + TextBox1.Text + "%' OR job_pos LIKE '%" + TextBox1.Text + "%' OR dept LIKE '%" + TextBox1.Text + "%'"
            conn.Open()
            da = New SqlDataAdapter(SQL, conn)
            da.Fill(dt)
            conn.Close()
            DataGridView1.DataSource = dt
            DataGridView1.CurrentCell = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim name As String
    Dim combo As String
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user JOIN tbl_logindetails ON tbl_user.emp_id = tbl_logindetails.emp_id WHERE tbl_user.emp_id = '" + DataGridView1.CurrentCell.Value.ToString + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                position.Text = acsdr("job_pos")
                department.Text = acsdr("dept")
                dh.Text = acsdr("date_hired")
                ei.Text = acsdr("emp_id")
                tn.Text = acsdr("TAX")
                ss.Text = acsdr("SSS")
                ph.Text = acsdr("Philhealth")
                pi.Text = acsdr("Pagibig")
                name = (acsdr("first_name")).trim + " " + (acsdr("last_name")).trim + " - " + acsdr("emp_id")
                Label4.Text = name.Trim
                pass.Text = (acsdr("pass_word").trim)
             
                ComboBox1.Text = acsdr("educ_loan")

            End While

           

            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub pos()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
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

    Public Sub dept()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        dtgv()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim addemp As New employeelist
        addemp.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "UPDATE tbl_user SET job_pos='" + position.Text + "',dept='" + department.Text + "',date_hired='" + dh.Text + "',emp_id='" + ei.Text + "',TAX='" + tn.Text + "',SSS='" + ss.Text + "',philhealth='" + ph.Text + "',pagibig='" + pi.Text + "',educ_loan='" + ComboBox1.Text + "'  WHERE emp_id='" + DataGridView1.CurrentCell.Value + "'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
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
            Dim SqlQuery As String = "UPDATE tbl_logindetails SET user_name='" + ei.Text + "' WHERE user_name='" + DataGridView1.CurrentCell.Value + "'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
            MessageBox.Show("Employee Profile Update Successful!", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        mainform.Show()
        Me.Close()

    End Sub


End Class