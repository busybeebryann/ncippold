Imports System.Data.SqlClient
Imports System.Data
Public Class job_position
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub job_position_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
        dept()
    End Sub
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT job_pos AS Position,basic_pay AS BasicPay FROM tbl_jobpos"
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
    Dim position As String = ""
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        position = DataGridView1.CurrentCell.Value.ToString
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_jobpos WHERE job_pos = '" + DataGridView1.CurrentCell.Value.ToString + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                TextBox1.Text = acsdr("job_pos").trim
                position = acsdr("job_pos")
                TextBox2.Text = acsdr("basic_pay").trim
                posid = acsdr("pos_id")
                ComboBox1.Text = acsdr("department").trim
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
                ComboBox1.Items.Add(acsdr("dept_name").trim)
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If position <> "" Or TextBox1.Text <> "" Then
            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM tbl_jobpos WHERE job_pos LIKE '%" + TextBox1.Text + "%'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())
                    Label5.Text = acsdr("pos_id")
                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


            If Label5.Text <> "Label5" Then
                Try
                    Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                    Dim con As New SqlConnection
                    con.ConnectionString = provider
                    If con.State = ConnectionState.Closed Then
                        con.Open()
                    End If
                    Dim SqlCommand As New SqlCommand
                    Dim SqlQuery As String = "UPDATE tbl_jobpos SET job_pos='" + TextBox1.Text + "',basic_pay='" + TextBox2.Text + "',department='" + ComboBox1.Text + "' WHERE pos_id='" + posid + "' OR pos_id='" + Label5.Text + "'"
                    With SqlCommand
                        .CommandText = SqlQuery
                        .Connection = con
                        .ExecuteNonQuery()
                    End With
                    con.Close()
                    Name = ""
                    MessageBox.Show("Update of job position " + position.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
                    Dim SqlQuery As String = "INSERT INTO tbl_jobpos(job_pos,basic_pay,department)VALUES('" + TextBox1.Text + "','" + TextBox2.Text + "','" + ComboBox1.Text + "')"
                    With SqlCommand
                        .CommandText = SqlQuery
                        .Connection = con
                        .ExecuteNonQuery()
                    End With
                    con.Close()
                    Name = ""
                    MessageBox.Show("Adding of job position " + position.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If
        Else
            MsgBox("Fill out the form to update or add another position")
        End If

        Dim addemp As New job_position
        addemp.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim addemp As New job_position
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
            Dim SqlQuery As String = "DELETE FROM tbl_jobpos WHERE pos_id='" + posid + "'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            MessageBox.Show("Deletion of job position  " + position.Trim + " is successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Dim addemp As New job_position
        addemp.Show()
        Me.Close()
    End Sub


End Class