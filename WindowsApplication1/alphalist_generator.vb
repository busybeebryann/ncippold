Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Public Class alphalist_generator
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub alphalist_generator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
    End Sub
    Dim mon As String
    Public Sub dtgv()
        If ComboBox1.Text = "January" Then
            mon = "1"
        ElseIf ComboBox1.Text = "February" Then
            mon = "2"
        ElseIf ComboBox1.Text = "March" Then
            mon = "3"
        ElseIf ComboBox1.Text = "April" Then
            mon = "4"
        ElseIf ComboBox1.Text = "May" Then
            mon = "5"
        ElseIf ComboBox1.Text = "June" Then
            mon = "6"
        ElseIf ComboBox1.Text = "July" Then
            mon = "7"
        ElseIf ComboBox1.Text = "August" Then
            mon = "8"
        ElseIf ComboBox1.Text = "September" Then
            mon = "9"
        ElseIf ComboBox1.Text = "October" Then
            mon = "10"
        ElseIf ComboBox1.Text = "November" Then
            mon = "11"
        ElseIf ComboBox1.Text = "December" Then
            mon = "12"
        End If
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT emp_id AS Employee_ID,emp_name AS Employee_Name FROM tbl_payroll WHERE month_computed='" + mon + "' AND TAX!='0.00' ORDER BY emp_name"
            Label3.Text = SQL
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        dtgv()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = Label3.Text
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox1.Items.Add(acsdr("Employee_ID").trim)
                ListBox2.Items.Add(acsdr("Employee_Name").trim)
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        For t_index As Integer = 0 To ListBox1.Items.Count - 1
            Dim emp_id As String = CStr(ListBox1.Items(t_index)).Trim
            Dim emp_name As String = CStr(ListBox2.Items(t_index)).Trim

            Try
                Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
                Dim con As New OleDbConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If

                Dim SqlCommand As New OleDbCommand
                Dim SqlQuery As String = "INSERT INTO [tbl_alphalist] ([emp_id],[emp_name],[rep_month]) VALUES ('" + emp_id + "','" + emp_name + "','" + ComboBox1.Text + "')"
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
        alpha.Show()
    End Sub
End Class