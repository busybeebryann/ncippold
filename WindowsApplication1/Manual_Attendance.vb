Imports System.Data.SqlClient
Public Class Manual_Attendance
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Private Sub Manual_Attendance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        Label6.Text = Date.Now.ToString("yyyy-M-dd")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label3.Text = TimeOfDay
    End Sub
    Dim nameemp As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" And TextBox2.Text = "" Then
            MsgBox("Please fill in all the fields.", MsgBoxStyle.Exclamation, "NCIP-CAR")

        Else

            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass;MultipleActiveResultSets=True"
            Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM tbl_logindetails WHERE emp_id='" + TextBox1.Text + "' AND pass_word='" + TextBox2.Text + "'", acsconn)
            acsconn.Open()
            Dim sdr As SqlDataReader = cmd.ExecuteReader()
            cmd.Connection = acsconn
            If (sdr.Read() = True) Then
                If acsconn.State = ConnectionState.Closed Then
                    acsconn.Open()
                End If

                strsql = "SELECT * from tbl_user JOIN tbl_logindetails ON tbl_user.emp_id = tbl_logindetails.emp_id WHERE tbl_user.emp_id = '" & TextBox1.Text & "'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader

                While (acsdr.Read())
                    nameemp = acsdr("first_name").trim + " " + acsdr("middle_name").trim + ". " + acsdr("last_name").trim
                End While

                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()


                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass;MultipleActiveResultSets=True"
                Dim cmdLs As SqlCommand = New SqlCommand("SELECT * FROM tbl_AttPayroll WHERE emp_id = '" + TextBox1.Text + "' AND att_date = '" & Label6.Text & "'", acsconn)
                acsconn.Open()
                Dim sdrLs As SqlDataReader = cmdLs.ExecuteReader()
                If (sdrLs.Read() = True) Then
                    MsgBox("Already Timed In! If you want to Time out press the Time Out Button.")

                ElseIf (sdrLs.Read() = False) Then

                    Dim cmds As SqlCommand = New SqlCommand("SELECT * FROM tbl_logindetails WHERE user_name='" + TextBox1.Text + "' AND pass_word='" + TextBox2.Text + "'", acsconn)

                    If acsconn.State = ConnectionState.Closed Then
                        acsconn.Open()
                    End If

                    Try
                        Dim SqlCommand As New SqlCommand
                        Dim SqlQuery As String = "INSERT INTO tbl_AttPayroll ([emp_id],[att_date],[time_in],[time_out],[reg_hr],[OT],[ND],[status]) VALUES ('" + TextBox1.Text + "','" + Label6.Text + "','" + Label3.Text + "','" + "none" + "','" + "0" + "','" + "0" + "','" + "0" + "','" + "0" + "')"
                        With SqlCommand
                            .CommandText = SqlQuery
                            .Connection = acsconn
                            .ExecuteNonQuery()
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                        End With
                        MessageBox.Show("Time In Successful! Welcome " + nameemp + "! Have a nice day.")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                    acsconn.Close()
                Else
                    MsgBox("Wrong Employee ID")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                End If
            End If
            End If
    End Sub
    Dim NH As String
    Dim OT As String
    Dim ND As String
    Dim realval As Integer
    Dim d1 As DateTime
    Dim d2 As DateTime
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text = "" And TextBox2.Text = "" Then
            MsgBox("Please fill in all the fields.", MsgBoxStyle.Exclamation, "NCIP-CAR")

        Else

            Try
                acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass;"
                acsconn.Open()
                strsql = "SELECT * from tbl_AttPayroll WHERE emp_id = '" + TextBox1.Text + "'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())
                    DateTimePicker1.Text = acsdr("time_in")
                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()
            Catch ex As Exception
                'MsgBox(ex.ToString)
            End Try

            d1 = DateTimePicker1.Text
            d2 = DateTimePicker2.Text

            realval = DateDiff("s", d1, d2) / 60 / 60

            If realval < 0.1 Then
                realval = 0
            End If

            If realval >= 9 Then
                NH = "8"
                OT = Val(realval) - 9
                ND = "0"
            ElseIf DateTimePicker1.Text >= #10:00:00 PM# Then
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
                Dim provider As String = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass;"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                'Dim SqlQuery As String = "INSERT INTO [tbl_AttPayroll] ([emp_id],[time_in],[time_out],[reg_hr],[OT],[ND],[att_date],[status]) VALUES ('" + TextBox1.Text + "','" + Label3.Text + "','" + DateTimePicker1.Text + "','" + NH + "','" + OT + "','" + ND + "','" + Date.Now.ToString("M/dd/yyyy") + "','0')"
                Dim SqlQuery As String = "UPDATE tbl_AttPayroll SET time_out='" + DateTimePicker2.Text + "',reg_hr='" + NH + "', OT='" + OT + "',ND='" + ND + "' WHERE emp_id='" + TextBox1.Text + "'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try
            MessageBox.Show("Successfully timed out for employee number " + TextBox1.Text + "!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
    End Sub
End Class