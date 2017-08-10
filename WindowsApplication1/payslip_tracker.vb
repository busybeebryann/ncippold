Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Public Class payslip_tracker
    Dim conn As New SqlConnection("Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass")
    Dim da As SqlDataAdapter = Nothing
    Dim dt As New DataTable
    Dim SQL As String = ""
    Private Sub payslip_filter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtgv()
        drange()
    End Sub
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Public Sub drange()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-ARVI-PC\SQLEXPRESS;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * from tbl_payroll WHERE emp_id='" + Login.TextBox1.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ComboBox1.Items.Add(acsdr("daterange"))
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
            SQL = "SELECT * FROM tbl_payroll WHERE daterange LIKE '%" + ComboBox1.Text.Trim + "%' AND emp_id LIKE '%" + Login.TextBox1.Text + "%'"
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



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
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
        ListBox10.Items.Clear()
        ListBox11.Items.Clear()
        ListBox12.Items.Clear()
        ListBox13.Items.Clear()
        ListBox14.Items.Clear()
        ListBox15.Items.Clear()
        ListBox16.Items.Clear()
        ListBox17.Items.Clear()
        ListBox18.Items.Clear()
        ListBox19.Items.Clear()
        ListBox20.Items.Clear()
        ListBox21.Items.Clear()
        ListBox22.Items.Clear()
        ListBox23.Items.Clear()
        ListBox24.Items.Clear()
        ListBox25.Items.Clear()
        ListBox26.Items.Clear()
        ListBox27.Items.Clear()
        ListBox28.Items.Clear()
        ListBox29.Items.Clear()
        ListBox30.Items.Clear()
        ListBox31.Items.Clear()
        ListBox32.Items.Clear()
        ListBox33.Items.Clear()
        ListBox34.Items.Clear()
        ListBox35.Items.Clear()
        ListBox36.Items.Clear()
        ListBox37.Items.Clear()

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
                ListBox2.Items.Add(acsdr("emp_name"))
                ListBox3.Items.Add(acsdr("position"))
                ListBox4.Items.Add(acsdr("department"))
                ListBox5.Items.Add(acsdr("daterange"))
                ListBox6.Items.Add(acsdr("civil_stat"))
                ListBox7.Items.Add(acsdr("dependents"))
                ListBox8.Items.Add(acsdr("rate_day"))
                ListBox9.Items.Add(acsdr("rate_hour"))
                ListBox10.Items.Add(acsdr("NH"))
                ListBox11.Items.Add(acsdr("OT"))
                ListBox12.Items.Add(acsdr("ND"))
                ListBox13.Items.Add(acsdr("monthly_pay"))
                ListBox14.Items.Add(acsdr("PERA"))
                ListBox15.Items.Add(acsdr("Grosspay"))
                ListBox16.Items.Add(acsdr("policy_loan"))
                ListBox17.Items.Add(acsdr("e_card"))
                ListBox18.Items.Add(acsdr("emergency_loan"))
                ListBox19.Items.Add(acsdr("conso_loan"))
                ListBox20.Items.Add(acsdr("pagibig_loan"))
                ListBox21.Items.Add(acsdr("pagibig_housing"))
                ListBox22.Items.Add(acsdr("ncipae_loan"))
                ListBox23.Items.Add(acsdr("lbp_loan"))
                ListBox24.Items.Add(acsdr("educ_loan"))
                ListBox25.Items.Add(acsdr("uoli_cont"))
                ListBox26.Items.Add(acsdr("gsis_social_cont"))
                ListBox27.Items.Add(acsdr("ncipae_man_fee"))
                ListBox28.Items.Add(acsdr("ncipae_mem_fee"))
                ListBox29.Items.Add(acsdr("Philhealth"))
                ListBox30.Items.Add(acsdr("Pagibig"))
                ListBox31.Items.Add(acsdr("TAX"))
                ListBox32.Items.Add(acsdr("net_pay"))
                ListBox33.Items.Add(acsdr("additional_deductions"))
                ListBox34.Items.Add(acsdr("status"))
                ListBox35.Items.Add(acsdr("date_computed"))
                ListBox36.Items.Add(acsdr("month_computed"))
                ListBox37.Items.Add(acsdr("time_computed"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        For t_index As Integer = 0 To ListBox1.Items.Count - 1
            Dim ei As String = CStr(ListBox1.Items(t_index)).Trim
            Dim en As String = CStr(ListBox2.Items(t_index)).Trim
            Dim jp As String = CStr(ListBox3.Items(t_index)).Trim
            Dim dp As String = CStr(ListBox4.Items(t_index)).Trim
            Dim dr As String = CStr(ListBox5.Items(t_index)).Trim
            Dim cs As String = CStr(ListBox6.Items(t_index)).Trim
            Dim dps As String = CStr(ListBox7.Items(t_index)).Trim
            Dim rd As String = CStr(ListBox8.Items(t_index)).Trim
            Dim rh As String = CStr(ListBox9.Items(t_index)).Trim
            Dim nh As String = CStr(ListBox10.Items(t_index)).Trim
            Dim ot As String = CStr(ListBox11.Items(t_index)).Trim
            Dim nd As String = CStr(ListBox12.Items(t_index)).Trim
            Dim mp As String = CStr(ListBox13.Items(t_index)).Trim
            Dim pr As String = CStr(ListBox14.Items(t_index)).Trim
            Dim gp As String = CStr(ListBox15.Items(t_index)).Trim
            Dim pl As String = CStr(ListBox16.Items(t_index)).Trim
            Dim ec As String = CStr(ListBox17.Items(t_index)).Trim
            Dim el As String = CStr(ListBox18.Items(t_index)).Trim
            Dim cl As String = CStr(ListBox19.Items(t_index)).Trim
            Dim pgl As String = CStr(ListBox20.Items(t_index)).Trim
            Dim pgh As String = CStr(ListBox21.Items(t_index)).Trim
            Dim ncl As String = CStr(ListBox22.Items(t_index)).Trim
            Dim ll As String = CStr(ListBox23.Items(t_index)).Trim
            Dim edl As String = CStr(ListBox24.Items(t_index)).Trim
            Dim uc As String = CStr(ListBox25.Items(t_index)).Trim
            Dim gsc As String = CStr(ListBox26.Items(t_index)).Trim
            Dim nmf As String = CStr(ListBox27.Items(t_index)).Trim
            Dim nmef As String = CStr(ListBox28.Items(t_index)).Trim
            Dim ph As String = CStr(ListBox29.Items(t_index)).Trim
            Dim pa As String = CStr(ListBox30.Items(t_index)).Trim
            Dim tx As String = CStr(ListBox31.Items(t_index)).Trim
            Dim np As String = CStr(ListBox32.Items(t_index)).Trim
            Dim ad As String = CStr(ListBox33.Items(t_index)).Trim
            Dim sta As String = CStr(ListBox34.Items(t_index)).Trim
            Dim dc As String = CStr(ListBox35.Items(t_index)).Trim
            Dim mc As String = CStr(ListBox36.Items(t_index)).Trim
            Dim tc As String = CStr(ListBox37.Items(t_index)).Trim

            Try
                Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
                Dim con As New OleDbConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If

                Dim SqlCommand As New OleDbCommand
                Dim SqlQuery As String = "INSERT INTO tbl_payroll (ei,en,jp,dp,dr,cs,dps,rd,rh,nh,ot,nd,mp,pr,gp,pl,ec,el,cl,pgl,pgh,ncl,ll,edl,uc,gsc,nmf,nmef,ph,pa,tx,np,ad,sta,dc,mc,tc) VALUES ('" + ei + "','" + en + "','" + jp + "','" + dp + "','" + dr + "','" + cs + "','" + dps + "','" + rd + "','" + rh + "','" + nh + "','" + ot + "','" + nd + "','" + mp + "','" + pr + "','" + gp + "','" + pl + "','" + ec + "','" + el + "','" + cl + "','" + pgl + "','" + pgh + "','" + ncl + "','" + ll + "','" + edl + "','" + uc + "','" + gsc + "','" + nmf + "','" + nmef + "','" + ph + "','" + pa + "','" + tx + "','" + np + "','" + ad + "','" + sta + "','" + dc + "','" + mc + "','" + tc + "')"
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
        payslip_printer.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim pa As New payslip_tracker
        pa.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class