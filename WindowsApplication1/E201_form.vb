Imports Microsoft.VisualBasic.PowerPacks
Imports System.Data
Imports System.Data.SqlClient
Public Class E201_form

    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim pf As New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        pf.Form = Me
        pf.Print()


    End Sub

    Private Sub FileSystemWatcher1_Changed(sender As Object, e As IO.FileSystemEventArgs) Handles FileSystemWatcher1.Changed

    End Sub

    Private Sub E201_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Dim id As String
    Dim desc As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        uploaded_Docs.Show()

        ' Panel1.Visible = False
        Dim con As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM tbl_logindetails WHERE user_name = '" & Label33.Text & "'", con)
        con.Open()
        Dim sdr As SqlDataReader = cmd.ExecuteReader()

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()

            strsql = "SELECT * from tbl_docs ORDER BY description"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())

                uploaded_Docs.ComboBox1.Items.Add(acsdr("description"))
                uploaded_Docs.Label2.Text = acsdr("emp_id").trim
                id = acsdr("emp_id").trim

                Try
                    uploaded_Docs.PictureBox1.Image = Image.FromFile(Application.StartupPath & "\docs\" + id + desc + ".jpg")
                    

                Catch ex As Exception
                    uploaded_Docs.PictureBox1.Image = Image.FromFile(Application.StartupPath & "\images\NCIP.png")
                  
                End Try


            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class