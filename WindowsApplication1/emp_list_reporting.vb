Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb


Public Class employee_list_reporting
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub employee_list_reporting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
        job()
        deptname()
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Public Sub job()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_jobpos"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ComboBox1.Items.Add(acsdr("job_pos"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub deptname()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_department"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ComboBox2.Items.Add(acsdr("dept_name"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT * FROM tbl_user WHERE [dept] LIKE '%" + ComboBox2.Text + "%' AND job_pos LIKE '%" + ComboBox1.Text + "%' AND first_name LIKE '%" + TextBox1.Text + "%' OR [dept] LIKE '%" + ComboBox2.Text + "%' AND job_pos LIKE '%" + ComboBox1.Text + "%' AND last_name LIKE '%" + TextBox1.Text + "%'"
            Label6.Text = SQL
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        dtgv()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        dtgv()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        dtgv()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()
        ListBox9.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = Label6.Text
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox1.Items.Add(acsdr("emp_id"))
                ListBox2.Items.Add(acsdr("first_name"))
                ListBox3.Items.Add(acsdr("last_name"))
                ListBox4.Items.Add(acsdr("middle_name"))
                ListBox5.Items.Add(acsdr("gender"))
                ListBox6.Items.Add(acsdr("job_pos"))
                ListBox7.Items.Add(acsdr("dept"))
                ListBox8.Items.Add(acsdr("date_hired"))
                ListBox9.Items.Add(acsdr("DIN"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        For t_index As Integer = 0 To ListBox1.Items.Count - 1
            Dim emp_id As String = CStr(ListBox1.Items(t_index)).Trim
            Dim fn As String = CStr(ListBox2.Items(t_index)).Trim
            Dim ln As String = CStr(ListBox3.Items(t_index)).Trim
            Dim mi As String = CStr(ListBox4.Items(t_index)).Trim
            Dim gr As String = CStr(ListBox5.Items(t_index)).Trim
            Dim jp As String = CStr(ListBox6.Items(t_index)).Trim
            Dim dp As String = CStr(ListBox7.Items(t_index)).Trim
            Dim dh As String = CStr(ListBox8.Items(t_index)).Trim
            Dim dn As String = CStr(ListBox9.Items(t_index)).Trim
            Try
                Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
                Dim con As New OleDbConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If

                Dim SqlCommand As New OleDbCommand
                Dim SqlQuery As String = "INSERT INTO [tbl_user] ([emp_id],[fn],[ln],[mi],[gr],[jp],[dp],[dh],[dn]) VALUES ('" + emp_id + "','" + fn + "','" + ln + "','" + mi + "','" + gr + "','" + jp + "','" + dp + "','" + dh + "','" + dn + "')"
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
        Form1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        mainform.Show()
        Me.Close()

    End Sub
End Class