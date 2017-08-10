Imports System.Data.SqlClient
Imports System.Globalization


Public Class payroll_daterange
    Dim gendata As String
    Dim vMonth As String
    Dim month As String
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Private Sub payroll_daterange_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gendata = "SELECT emp_id,time_in,time_out,reg_hr,OT,ND,att_date FROM [tbl_AttPayroll] WHERE status = '0'"
        dtgv()
        Label6.Text = DateTimePicker1.Text + " TO " + DateTimePicker2.Text

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user"
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

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""

    Dim da2 As SqlDataAdapter = Nothing
    Dim dt2 As New DataTable
    Dim SQL2 As String = ""
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = gendata
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

    Public Sub dtgv2()
        dt2.Clear()
        DataGridView2.DataSource = Nothing
        Try
            SQL2 = "SELECT * FROM tbl_PayrollData WHERE daterange='" + DateTimePicker1.Text + " - " + DateTimePicker2.text + "'"
            conn.Open()
            da2 = New SqlDataAdapter(SQL2, conn)
            da2.Fill(dt2)
            conn.Close()
            DataGridView2.DataSource = dt2
            DataGridView2.CurrentCell = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        gendata = "SELECT emp_id,time_in,time_out,reg_hr,OT,ND,att_date FROM [tbl_AttPayroll] WHERE status = '0' AND att_date BETWEEN '" + DateTimePicker1.Text + "' AND '" + DateTimePicker2.Text + "'"

        dtgv()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Label6.Text = DateTimePicker1.Text + " TO " + DateTimePicker2.Text
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Label6.Text = DateTimePicker1.Text + " TO " + DateTimePicker2.Text
    End Sub
    Dim days As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Weekdays()

        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim totHr As String = ""
            Dim HrH As String = ""
            Dim totOT As String = ""
            Dim OTH As String = ""
            Dim totND As String = ""
            Dim NDH As String = ""
            Dim id As String = CStr(ListBox1.Items(i))
            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM tbl_AttPayroll WHERE status='0' AND emp_id='" + id + "' AND att_date BETWEEN '" + DateTimePicker1.Text + "' AND '" + DateTimePicker2.Text + "'"

                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())

                    HrH = acsdr("reg_hr")
                    totHr = Val(totHr) + Val(HrH)

                    If totHr = "" Then
                        totHr = 0
                    End If

                    OTH = acsdr("OT")
                    totOT = Val(totOT) + Val(OTH)

                    If totOT = "" Then
                        totOT = 0
                    End If

                    NDH = acsdr("ND")
                    totND = Val(totND) + Val(NDH)

                    If totND = "" Then
                        totND = 0
                    End If

                    ListBox2.Items.Add(acsdr("reg_hr"))
                    ListBox3.Items.Add(acsdr("OT"))
                    ListBox4.Items.Add(acsdr("ND"))
                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


            ListBox6.Items.Add(totHr)
            ListBox7.Items.Add(totOT)
            ListBox8.Items.Add(totND)

            ListBox2.Items.Clear()
            ListBox3.Items.Clear()
            ListBox4.Items.Clear()

            Try
                Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO tbl_PayrollData (emp_id,tot_RH,tot_OT,tot_ND,daterange,drange_dur,status) VALUES ('" + id + "','" + totHr + "','" + totOT + "','" + totND + "','" + DateTimePicker1.Text + " - " + DateTimePicker2.Text + "','" + Label7.Text + "','0')"
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
                Dim SqlQuery As String = "UPDATE tbl_AttPayroll SET status='1' WHERE emp_id='" + id + "' AND att_date BETWEEN '" + DateTimePicker1.Text + "' AND '" + DateTimePicker2.Text + "'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Next
        MessageBox.Show("Payroll for daterange " + DateTimePicker1.Text + " - " + DateTimePicker2.Text + " generated successful! You can now compute for the employee salary.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        dtgv2()
    End Sub
    Public Sub Weekdays()
        Dim startDate As DateTime = DateTimePicker1.Text
        Dim endDate As DateTime = DateTimePicker2.Text
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        numWeekdays = 0
        WeekendDays = 0

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1

        For i As Integer = 1 To totalDays

            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next

        numWeekdays = totalDays - WeekendDays
        Label7.Text = numWeekdays

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub
End Class