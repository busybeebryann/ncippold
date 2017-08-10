Imports System.Data.SqlClient
Imports System.Data
Public Class E201_list
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        dtgv2()
    End Sub

    Private Sub E201_list_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()

    End Sub

    Public Sub dtgv()
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT * FROM tbl_user WHERE emp_id LIKE '%" + TextBox1.Text + "%' OR first_name LIKE '%" + TextBox1.Text + "%' OR last_name LIKE '%" + TextBox1.Text + "%' OR job_pos LIKE '%" + TextBox1.Text + "%' OR dept LIKE '%" + TextBox1.Text + "%'"
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
        dt.Clear()
        DataGridView1.DataSource = Nothing
        Try
            SQL = "SELECT * FROM tbl_user WHERE emp_id LIKE '%" + TextBox1.Text + "%'"
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

    Dim id As String
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        E201_form.Show()

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_user WHERE tbl_user.emp_id = '" + DataGridView1.CurrentCell.Value.ToString + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                E201_form.Label11.Text = acsdr("c_no")
                E201_form.Label12.Text = acsdr("email")
                E201_form.Label13.Text = acsdr("b_date")
                E201_form.Label14.Text = acsdr("age")
                E201_form.Label15.Text = acsdr("gender")
                E201_form.Label16.Text = acsdr("address")
                E201_form.Label17.Text = acsdr("marital_status")
                E201_form.Label37.Text = acsdr("dependents")
                E201_form.Label28.Text = acsdr("TAX")
                E201_form.Label27.Text = acsdr("sss")
                E201_form.Label34.Text = acsdr("pagibig")
                E201_form.Label35.Text = acsdr("philhealth")

                E201_form.Label33.Text = acsdr("emp_id")
                E201_form.Label32.Text = acsdr("dept")
                E201_form.Label31.Text = acsdr("job_pos")
                E201_form.Label30.Text = acsdr("date_hired")
                E201_form.Label29.Text = acsdr("hr_pr_day")

                E201_form.Label39.Text = acsdr("first_name").trim + " " + acsdr("middle_name").trim
                E201_form.Label1.Text = acsdr("last_name").trim

                id = acsdr("emp_id")
                id = id.Trim

            End While

            Try
                E201_form.PictureBox1.Image = Image.FromFile(Application.StartupPath & "\images\" + id + ".jpg")
            Catch ex As Exception
                E201_form.PictureBox1.Image = Image.FromFile(Application.StartupPath & "\images\default.png")
            End Try

            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        mainform.Show()

    End Sub
End Class