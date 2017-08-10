Imports System.Data.SqlClient
Imports System.Data
Public Class dept_add
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub dept_add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
    End Sub
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT dept_name AS Department_Name FROM tbl_department"
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
    Dim posid As String
    Dim department As String = ""
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        department = DataGridView1.CurrentCell.Value.ToString
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_department WHERE dept_name = '" + DataGridView1.CurrentCell.Value.ToString + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                TextBox1.Text = acsdr("dept_name")
                department = acsdr("dept_name")
                posid = acsdr("dept_id")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If department <> "" Then
            Try
                Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_department SET dept_name='" + TextBox1.Text + "' WHERE dept_id='" + posid + "'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
                Name = ""
                MessageBox.Show("Update of department name for  " + department.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Else
            Try
                Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO tbl_department ([dept_name]) VALUES ('" + TextBox1.Text + "')"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                MessageBox.Show("Adding of department  " + department.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
        Dim addemp As New dept_add
        addemp.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "DELETE FROM tbl_department WHERE dept_id='" + posid + "'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            MessageBox.Show("Deletion of department " + department.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Dim addemp As New dept_add
        addemp.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim addemp As New dept_add
        addemp.Show()
        Me.Close()
    End Sub
End Class