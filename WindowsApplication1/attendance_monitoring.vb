Imports System.Data
Imports System.Data.SqlClient
Public Class attendance_monitoring
    Dim realval As String
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim name As String
    Dim d1 As DateTime
    Dim d2 As DateTime
    Dim OT As String
    Dim RH As String
    Dim ND As String
    Dim d As String
    Dim DIN As String
    Dim ver As String

    Private Sub attendance_monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM [tbl_user] WHERE [DIN] IS NOT NULL"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                name = acsdr("first_name").trim + " " + acsdr("last_name").trim
                ListBox1.Items.Add(name)
                ListBox2.Items.Add(acsdr("emp_id").trim)

                ListBox8.Items.Add(acsdr("DIN"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        For l_index As Integer = 0 To ListBox8.Items.Count - 1
            Dim id As String = CStr(ListBox8.Items(l_index))
            Try
                acsconn.ConnectionString = "Data Source=MASAGCAFAMILY\SQLEXPRESS;Initial Catalog=RIMS_Attendance;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM [ras_AttRecord] WHERE [DIN] ='" + id + "'"
                Dim acscmdr2 As New SqlCommand
                acscmdr2.CommandText = strsql
                acscmdr2.Connection = acsconn
                acsdr = acscmdr2.ExecuteReader
                While (acsdr.Read())

                    ver = acsdr("VerifyMode")
                    If ver = "1" Then
                        ListBox11.Items.Add(acsdr("DIN"))
                        Dim mystr As String = acsdr("clock")
                        Dim cut_at As String = " "
                        Dim x As Integer = InStr(mystr, cut_at)

                        Dim string_after As String = mystr.Substring(x + cut_at.Length - 1)
                        ListBox3.Items.Add(string_after)

                        Try
                            d = acsdr("clock")
                            Dim mystr2 As String = d
                            Dim cut_at2 As String = " "
                            Dim x2 As Integer = InStr(mystr2, cut_at2)
                            Dim string_before As String = mystr2.Substring(0, x2 - 1)
                            ListBox9.Items.Add(string_before)
                        Catch ex As Exception

                        End Try

                    Else
                        d = acsdr("clock")
                        Dim mystr2 As String = d
                        Dim cut_at2 As String = " "
                        Dim x2 As Integer = InStr(mystr2, cut_at2)
                        Dim string_after As String = mystr2.Substring(x2 + cut_at2.Length - 1)
                        ListBox4.Items.Add(string_after)
                    End If
                End While
                acscmdr2.Dispose()
                acsdr.Close()
                acsconn.Close()



            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


        Next
        For d_index As Integer = 0 To ListBox11.Items.Count - 1
            Dim DIN As String = CStr(ListBox11.Items(d_index))
            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM [tbl_user] WHERE [DIN] ='" + DIN + "'"
                Dim acscmdr2 As New SqlCommand
                acscmdr2.CommandText = strsql
                acscmdr2.Connection = acsconn
                acsdr = acscmdr2.ExecuteReader
                While (acsdr.Read())
                    ListBox10.Items.Add(acsdr("first_name").trim + " " + acsdr("last_name").trim)
                    ListBox12.Items.Add(acsdr("emp_id"))
                End While
                acscmdr2.Dispose()
                acsdr.Close()
                acsconn.Close()



            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Next

        For t_index As Integer = 0 To ListBox11.Items.Count - 1
            Dim d1 As String = CStr(ListBox3.Items(t_index))
            Dim d2 As String = CStr(ListBox4.Items(t_index))
            If d1 >= #10:00:00 PM# Then
                realval = DateDiff("s", d2, d1)

                realval = Val(realval) / 60 / 60 / 2

                If realval >= 9 Then
                    ND = "8"
                    OT = Val(realval) - 9
                    RH = "0"
                ElseIf realval < 9 Then
                    ND = realval
                    OT = "0"
                    RH = "0"
                End If
                ListBox5.Items.Add(RH)
                ListBox6.Items.Add(OT)
                ListBox7.Items.Add(ND)

            Else
                realval = DateDiff("s", d1, d2)

                realval = Val(realval) / 60 / 60

                If realval >= 9 Then
                    RH = "8"
                    OT = Val(realval) - 9
                ElseIf realval < 9 Then
                    RH = realval
                    OT = "0"
                End If

                ListBox5.Items.Add(RH)
                ListBox6.Items.Add(OT)
                ListBox7.Items.Add("0")
            End If



        Next



    End Sub
    Dim conn As New SqlConnection("Data Source=MASAGCAFAMILY\SQLEXPRESS;Initial Catalog=RIMS_Attendance;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT ID as Action_No,DIN as Employee_Biometrics_ID,Clock as Date_Time,VerifyMode FROM [ras_AttRecord]"
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
        For t_index As Integer = 0 To ListBox10.Items.Count - 1
            Dim ID As String = CStr(ListBox12.Items(t_index)).Trim
            Dim d1r As String = CStr(ListBox3.Items(t_index))
            Dim d2r As String = CStr(ListBox4.Items(t_index))
            Dim HRr As String = CStr(ListBox5.Items(t_index))
            Dim OTr As String = CStr(ListBox6.Items(t_index))
            Dim NDr As String = CStr(ListBox7.Items(t_index))
            Dim att_date As String = CStr(ListBox9.Items(t_index)).Trim
            Dim sInDate As String = att_date
            Dim sOutDate As String = Convert.ToDateTime(sInDate).ToString("yyyy-MM-dd")

            Try
                Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "INSERT INTO [tbl_AttPayroll] ([emp_id],[time_in],[time_out],[reg_hr],[OT],[ND],[att_date],[status]) VALUES ('" + ID + "','" + d1r + "','" + d2r + "','" + HRr + "','" + OTr + "','" + NDr + "','" + sOutDate + "','0')"
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
        dtgv2()
    End Sub

    Dim conn2 As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da2 As SqlDataAdapter = Nothing
    Dim dt2 As New DataTable
    Dim SQL2 As String = ""
    Public Sub dtgv2()
        dt2.Clear()
        DataGridView2.DataSource = Nothing
        Try
            SQL2 = "SELECT * FROM [tbl_AttPayroll] WHERE status='0'"
            conn2.Open()
            da2 = New SqlDataAdapter(SQL2, conn2)
            da2.Fill(dt2)
            conn.Close()
            DataGridView2.DataSource = dt2
            DataGridView2.CurrentCell = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
End Class