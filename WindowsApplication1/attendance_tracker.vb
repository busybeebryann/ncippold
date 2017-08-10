Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Public Class attendance_tracker
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub attendance_tracker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
    End Sub
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try

            If Login.TextBox1.Text = "admin" Then

                SQL = "SELECT emp_id,time_in,time_out,att_date FROM tbl_AttPayroll WHERE att_date BETWEEN '" + DateTimePicker1.Text + "' AND '" + DateTimePicker2.Text + "'"

            Else
                SQL = "SELECT emp_id,time_in,time_out,att_date FROM tbl_AttPayroll WHERE emp_id='" + Login.TextBox1.Text + "' AND att_date BETWEEN '" + DateTimePicker1.Text + "' AND '" + DateTimePicker2.Text + "'"

            End If

            Label2.Text = SQL
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dtgv()
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = Label2.Text
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox1.Items.Add(acsdr("emp_id"))
                ListBox2.Items.Add(acsdr("time_in"))
                ListBox3.Items.Add(acsdr("time_out"))
                ListBox4.Items.Add(acsdr("att_date"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        For t_index As Integer = 0 To ListBox1.Items.Count - 1
            Dim ei As String = CStr(ListBox1.Items(t_index)).Trim
            Dim ti As String = CStr(ListBox2.Items(t_index)).Trim
            Dim tou As String = CStr(ListBox3.Items(t_index)).Trim
            Dim ad As String = CStr(ListBox4.Items(t_index)).Trim

            Try
                Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
                Dim con As New OleDbConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If

                Dim SqlCommand As New OleDbCommand
                Dim SqlQuery As String = "INSERT INTO [tbl_att] ([ei],[ti],[tou],[ad]) VALUES ('" + ei + "','" + ti + "','" + tou + "','" + ad + "')"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With

                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
        Next
        att_printer.Show()
    End Sub
End Class