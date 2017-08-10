Imports System.Data.SqlClient
Imports System.Data
Imports System.Text.RegularExpressions

Public Class Payroll_computation
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim drange As String
    Dim cdrange As String = "none"
    Private Sub Payroll_computation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_PayrollData WHERE status = '0' ORDER BY daterange"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                drange = acsdr("daterange").trim
                If drange <> cdrange Then
                    ComboBox1.Items.Add(drange)
                    cdrange = drange
                Else
                    cdrange = drange
                End If

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Dim drangedur As String
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox3.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_PayrollData WHERE status = '0' AND daterange='" + ComboBox1.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox3.Items.Add(acsdr("emp_id"))
                drangedur = acsdr("drange_dur")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Dim educloan As String
    Dim pro As String
    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Button2.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = True
        ListBox4.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_user WHERE emp_id='" + ListBox3.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                Label15.Text = acsdr("last_name").trim + ", " + acsdr("first_name").trim + " " + acsdr("middle_name").trim
                Label29.Text = ListBox3.Text
                Label9.Text = acsdr("job_pos").trim
                Label17.Text = ComboBox1.Text
                Label23.Text = acsdr("dept").trim
                Label44.Text = acsdr("marital_status").trim
                Label45.Text = acsdr("dependents").trim
                ratepd = Math.Round(Val(acsdr("basic_pay").trim), 2)
                Label31.Text = Math.Round((Val(ratepd)) / Val(drangedur), 2)
                Label33.Text = Math.Round((Val(Label31.Text) / 8), 2)
                educloan = acsdr("educ_loan")
                pro = acsdr("province")
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        rate()
        workhours()
    End Sub
    Dim ratepd As String
    Public Sub rate()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_jobpos WHERE job_pos='" + Label9.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub workhours()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_PayrollData WHERE emp_id='" + ListBox3.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                T1.Text = acsdr("tot_RH").trim
                T2.Text = acsdr("TOT_OT").trim
                T3.Text = acsdr("TOT_ND").trim

                If T1.Text = "" Then
                    T1.Text = 0
                Else
                    T1.Text = Val(T1.Text) / 8
                End If
                If T2.Text = "" Then
                    T2.Text = 0
                Else
                    T2.Text = Val(T2.Text) / 8
                End If
                If T3.Text = "" Then
                    T3.Text = 0
                Else
                    T3.Text = Val(T3.Text) / 8
                End If
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Dim NH As Double
    Dim OT As Double
    Dim ND As Double
    Dim gross As Double
    Dim OTval As Double
    Dim NDval As Double
    Dim SSStot As Double
    Dim pagibigtot As Double
    Dim philhealthtot As Double
    Dim TAXtot As Double
    Dim totdeduction As Double
    Dim taxable As Double
    Dim loansDed As Double
    Dim gsisded As Double
    Dim g As Double
    Dim n As Double
    Dim t1t As String
    Dim t2t As String
    Dim t3t As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        t1t = Val(T1.Text) * 8
        t2t = Val(T2.Text) * 8
        t3t = Val(T3.Text) * 8
        If educloan = "Yes" Then
            educloan = 20
        Else
            educloan = 0.0
        End If
        GSISLOAN()
        NCIPEA()
        pagibigloan()
        NCIPEALoan()
        LPB()


        OTval = (Val(Label33.Text) * 0.25) + Val(Label33.Text)
        NDval = (Val(Label33.Text) * 0.1) + Val(Label33.Text)
        NH = Val(t1t) * Val(Label33.Text)
        OT = Val(t2t) * OTval
        ND = Val(t3t) * NDval

        gross = NH + OT + ND
        Label98.Text = gross.ToString("0,00.00")
        gross = gross + 2000
        gsisded = gross * 0.09
        Label10.Text = gross.ToString("0,00.00")
        Label84.Text = gsisded.ToString("0,00.00")


        gross = Val(gross)
        loansDed = Val(emL) + Val(cnL) + Val(plL) + Val(educloan) + Val(phL) + Val(pgL) + Val(ncipAmnt) + Val(lpbamnt)

        gross = gross

        PagibigDed()
        philhealth()

        totdeduction = Val(Label13.Text) + Val(Label12.Text) + Val(ncip) + Val(ncip2) + gsisded

        taxable = gross - totdeduction

        taxfunc()

        taxable = taxable - (tax)
        taxable = taxable - loansDed
        Label14.Text = taxable.ToString("0,00.00")

        Label27.Text = T1.Text
        Label36.Text = T2.Text
        Label39.Text = T3.Text
        Button2.Enabled = True
        Button6.Enabled = True
        If bigbonus = "Yes" Then
            extramonth()
        End If
        g = Regex.Replace(Label10.Text, "[^A-Za-z0-9\-/]", "")
        g = g * 0.1
        g = g * 0.1
        n = Regex.Replace(Label14.Text, "[^A-Za-z0-9\-/]", "")
        n = n * 0.1
        n = n * 0.1
        tDed = g - n

        MsgBox(tDed)
    End Sub

    Dim lastpayroll As String

    Dim pagibig As Double
    Sub PagibigDed()
        If Val(gross) <= 1500 Then
            pagibig = Val(gross) * 0.01
            Label13.Text = pagibig.ToString("0,00.00")
        ElseIf Val(gross) >= 1501 And Val(gross) <= 4999 Then
            pagibig = Val(gross) * 0.02
            Label13.Text = pagibig.ToString("0,00.00")
        Else
            pagibig = 100
            Label13.Text = pagibig.ToString("0,00.00")
        End If
    End Sub
    Sub philhealth()
        If Val(gross) > 200 And Val(gross) <= 8999.99 Then
            Label12.Text = 100
        ElseIf Val(gross) >= 9000 And Val(gross) < 9999.99 Then
            Label12.Text = 112.5
        ElseIf Label10.Text >= 10000 And Val(gross) <= 10999.99 Then
            Label12.Text = 125.0
        ElseIf Val(gross) >= 11000 And Val(gross) <= 11999.99 Then
            Label12.Text = 137.5
        ElseIf Val(gross) >= 12000 And Val(gross) <= 12999.99 Then
            Label12.Text = 150
        ElseIf Val(gross) >= 13000 And Val(gross) <= 13999.99 Then
            Label12.Text = 162.5
        ElseIf Val(gross) >= 14000 And Val(gross) <= 14999.99 Then
            Label12.Text = 175
        ElseIf Val(gross) >= 15000 And Val(gross) <= 15999.99 Then
            Label12.Text = 187.5
        ElseIf Val(gross) >= 16000 And Val(gross) <= 16999.99 Then
            Label12.Text = 200
        ElseIf Val(gross) >= 17000 And Val(gross) <= 17999.99 Then
            Label12.Text = 212.5
        ElseIf Val(gross) >= 18000 And Val(gross) <= 18999.99 Then
            Label12.Text = 225
        ElseIf Val(gross) >= 19000 And Val(gross) <= 19999.99 Then
            Label12.Text = 237.5
        ElseIf Val(gross) >= 20000 And Val(gross) <= 20999.99 Then
            Label12.Text = 250
        ElseIf Val(gross) >= 21000 And Val(gross) <= 21999.99 Then
            Label12.Text = 262.5
        ElseIf Val(gross) >= 22000 And Val(gross) <= 22999.99 Then
            Label12.Text = 275
        ElseIf Val(gross) >= 23000 And Val(gross) <= 23999.99 Then
            Label12.Text = 287.5
        ElseIf Val(gross) >= 24000 And Val(gross) <= 24999.99 Then
            Label12.Text = 300
        ElseIf Val(gross) >= 25000 And Val(gross) <= 25999.99 Then
            Label12.Text = 312.5
        ElseIf Val(gross) >= 26000 And Val(gross) <= 26999.99 Then
            Label12.Text = 325
        ElseIf Val(gross) >= 27000 And Val(gross) <= 27999.99 Then
            Label12.Text = 337.5
        ElseIf Val(gross) >= 28000 And Val(gross) <= 28999.99 Then
            Label12.Text = 350
        ElseIf Val(gross) >= 29000 And Val(gross) <= 29999.99 Then
            Label12.Text = 362.5
        ElseIf Val(gross) >= 30000 And Val(gross) <= 30999.99 Then
            Label12.Text = 375
        ElseIf Val(gross) >= 31000 And Val(gross) <= 31999.99 Then
            Label12.Text = 387.5
        ElseIf Val(gross) >= 32000 And Val(gross) <= 32999.99 Then
            Label12.Text = 400
        ElseIf Val(gross) >= 33000 And Val(gross) <= 33999.99 Then
            Label12.Text = 412.5
        ElseIf Val(gross) >= 34000 And Val(gross) <= 34999.99 Then
            Label12.Text = 425
        ElseIf Val(gross) >= 35000 Then
            Label12.Text = 437.5
        End If
    End Sub
    Dim dependents As String
    Dim taxableincome As Double
    Dim tax As Double

    Public Sub taxfunc()

        dependents = Label45.Text
        taxableincome = taxable
        If dependents = "0" Then
            If taxable >= 50 And taxable <= 4166 Then
                Label14.Text = taxable.ToString("0,00.00")
                Label41.Text = 0
            ElseIf taxable >= 4167 And taxable <= 4999 Then
                tax = 0 + (taxableincome - 4167) * 0.05
                Label14.Text = ((taxable - tax).ToString("0,00.00"))
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 5000 And taxable <= 6666 Then
                tax = 41.67 + (taxableincome - 5000) * 0.1
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 6667 And taxable <= 9999 Then
                tax = 208.33 + (taxableincome - 6667) * 0.15
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 10000 And taxable <= 15832 Then
                tax = 708.33 + (taxableincome - 10000) * 0.2
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 15833 And taxable <= 24999 Then
                tax = 1875 + (taxableincome - 15833) * 0.25
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 25000 And taxable <= 45832 Then
                tax = 4166.67 + (taxableincome - 25000) * 0.3
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 45833 Then
                tax = 10416.67 + (taxableincome - 45833) * 0.32
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            End If
        ElseIf dependents = "1" Then
            If taxable >= 75 And taxable <= 6249 Then
                Label14.Text = taxable.ToString("0,00.00")
                Label41.Text = 0
            ElseIf taxable >= 6250 And taxable <= 7082 Then
                tax = 0 + (taxableincome - 6250) * 0.05
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 7083 And taxable <= 8749 Then
                tax = 41.67 + (taxableincome - 7083) * 0.1
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 8750 And taxable <= 12082 Then
                tax = 208.33 + (taxableincome - 8750) * 0.15
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 12082 And taxable <= 17916 Then
                tax = 708.33 + (taxableincome - 12082) * 0.2
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 17917 And taxable <= 27082 Then
                tax = 1875 + (taxableincome - 17917) * 0.25
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 27083 And taxable <= 47916 Then
                tax = 4166.67 + (taxableincome - 27083) * 0.3
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 47917 Then
                tax = 10416.67 + (taxableincome - 47917) * 0.32
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            End If
        ElseIf dependents = "2" Then
            If taxable >= 100 And taxable <= 8332 Then
                Label14.Text = taxable.ToString("0,00.00")
                Label41.Text = 0
            ElseIf taxable >= 8333 And taxable <= 9166 Then
                tax = 0 + (taxableincome - 9166) * 0.05
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 9167 And taxable <= 10832 Then
                tax = 41.67 + (taxableincome - 9167) * 0.1
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 10833 And taxable <= 14166 Then
                tax = 208.33 + (taxableincome - 10833) * 0.15
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 14167 And taxable <= 19999 Then
                tax = 708.33 + (taxableincome - 14167) * 0.2
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 20000 And taxable <= 29166 Then
                tax = 1875 + (taxableincome - 20000) * 0.25
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 29167 And taxable <= 49000 Then
                tax = 4166.67 + (taxableincome - 29167) * 0.3
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 50000 Then
                tax = 10416.67 + (taxableincome - 50000) * 0.32
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax
            End If
        ElseIf dependents = "3" Then
            If taxable >= 125 And taxable <= 10416 Then
                Label14.Text = taxable.ToString("0,00.00")
                Label41.Text = 0
            ElseIf taxable >= 10417 And taxable <= 11249 Then
                tax = 0 + (taxableincome - 10417) * 0.05
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 11250 And taxable <= 12916 Then
                tax = 41.67 + (taxableincome - 11250) * 0.1
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 12917 And taxable <= 16249 Then
                tax = 208.33 + (taxableincome - 12917) * 0.15
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 16250 And taxable <= 22082 Then
                tax = 708.33 + (taxableincome - 16250) * 0.2
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 22083 And taxable <= 31249 Then
                tax = 1875 + (taxableincome - 22083) * 0.25
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 31250 And taxable <= 52082 Then
                tax = 4166.67 + (taxableincome - 31250) * 0.3
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 52083 Then
                tax = 10416.67 + (taxableincome - 52083) * 0.32
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            End If
        ElseIf dependents = "4" Then
            If taxable >= 150 And taxable <= 12499 Then
                Label14.Text = taxable.ToString("0,00.00")
                Label41.Text = 0
            ElseIf taxable >= 12500 And taxable <= 13332 Then
                tax = 0 + (taxableincome - 12500) * 0.05
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 13333 And taxable <= 14999 Then
                tax = 41.67 + (taxableincome - 13333) * 0.1
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 15000 And taxable <= 18332 Then
                tax = 208.33 + (taxableincome - 15000) * 0.15
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 18333 And taxable <= 24166 Then
                tax = 708.33 + (taxableincome - 18333) * 0.2
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 24167 And taxable <= 33332 Then
                tax = 1875 + (taxableincome - 24167) * 0.25
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 33333 And taxable <= 54166 Then
                tax = 4166.67 + (taxableincome - 33333) * 0.3
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            ElseIf taxable >= 54167 Then
                tax = 10416.67 + (taxableincome - 54167) * 0.32
                Label14.Text = (taxable - tax).ToString("0,00.00")
                Label41.Text = tax.ToString("0,00.00")
            End If
        End If
    End Sub
    Dim ncip As Integer
    Dim ncip2 As Integer
    Public Sub NCIPEA()
        ncip = 25
        Label78.Text = "25.00"
        ncip2 = 100
        Label81.Text = "100.00"
    End Sub
    Dim ncipAmnt As String
    Dim ncipAmntMP As String
    Public Sub NCIPEALoan()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='NCIPAE Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ncipAmnt = acsdr("payment_amount").trim
                ncipAmntMP = Val(acsdr("months_to_pay").trim) - 1
                If ncipAmnt = "" Then
                    ncipAmnt = 0
                    Label65.Text = "0.00"
                Else
                    Label65.Text = Val(ncipAmnt).ToString("0,00.00")
                End If
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try
    End Sub

    Public Sub upNCIP()
        If Val(ncipAmnt) >= 0 Then

            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + ncipAmnt + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='NCIPAE Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "DELETE FROM tbl_loans WHERE months_to_pay='0'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Dim lpbamnt As String
    Dim lpbMP As String
    Public Sub LPB()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='LBP Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                lpbamnt = acsdr("payment_amount").trim
                lpbMP = Val(acsdr("months_to_pay").trim) - 1
                If lpbamnt = "" Then
                    lpbamnt = 0
                    Label74.Text = "0.00"
                Else
                    Label74.Text = Val(lpbamnt).ToString("0,00.00")
                End If
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try
    End Sub

    Public Sub upLPB()
        If Val(lpbamnt) >= 0 Then

            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + lpbamnt + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='LBP Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "DELETE FROM tbl_loans WHERE months_to_pay='0'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Dim datenow As String
    Dim addID As String
    Dim mnow As String
    Dim tDed As String
    Public Sub savedata()

        datenow = Date.Now()
        mnow = Date.Now.Month.ToString
        If (mnow = "1") Then
            mnow = "January"
        ElseIf (mnow = "2") Then
            mnow = "February"
        ElseIf (mnow = "3") Then
            mnow = "March"
        ElseIf (mnow = "4") Then
            mnow = "April"
        ElseIf (mnow = "5") Then
            mnow = "May"
        ElseIf (mnow = "6") Then
            mnow = "June"
        ElseIf (mnow = "7") Then
            mnow = "July"
        ElseIf (mnow = "8") Then
            mnow = "August"
        ElseIf (mnow = "9") Then
            mnow = "September"
        ElseIf (mnow = "10") Then
            mnow = "October"
        ElseIf (mnow = "11") Then
            mnow = "November"
        ElseIf (mnow = "12") Then
            mnow = "December"
        End If
        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "INSERT INTO tbl_payroll ([emp_id],[emp_name],[position],[department],[daterange],[civil_stat],[dependents],[rate_day],[rate_hour],[NH],[OT],[ND],[monthly_pay],[PERA],[Grosspay],[policy_loan],[e_card],[emergency_loan],[conso_loan],[pagibig_loan],[pagibig_housing],[ncipae_loan],[lbp_loan],[educ_loan],[uoli_cont],[gsis_social_cont],[ncipae_man_fee],[ncipae_mem_fee],[Philhealth],[Pagibig],[TAX],[net_pay],[status],[date_computed],[month_computed],[time_computed],[tDed],[pro]) VALUES ('" + Label29.Text + "','" + Label15.Text + "','" + Label9.Text + "','" + Label23.Text + "','" + Label17.Text + "','" + Label44.Text + "','" + Label45.Text + "','" + Label31.Text + "','" + Label33.Text + "','" + T1.Text + "','" + T2.Text + "','" + T3.Text + "','" + Label98.Text + "','" + Label93.Text + "','" + Label10.Text + "','" + Label54.Text + "','" + Label60.Text + "','" + Label61.Text + "','" + Label64.Text + "','" + Label63.Text + "','" + Label62.Text + "','" + Label65.Text + "','" + Label74.Text + "','" + Label90.Text + "','" + Label87.Text + "','" + Label84.Text + "','" + Label81.Text + "','" + Label78.Text + "','" + Label12.Text + "','" + Label13.Text + "','" + Label41.Text + "','" + Label14.Text + "','0','" + datenow + "','" + mnow + "','" + TimeOfDay + "','" + tDed + "','" + pro + "')"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        For i As Integer = 0 To ListBox4.Items.Count - 1
            Dim additionalDI As String = CStr(ListBox4.Items(i))
            addID += additionalDI + Environment.NewLine
        Next
        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "UPDATE tbl_payroll SET additional_deductions='" + addID + "' WHERE emp_id='" + Label29.Text + "' AND daterange='" + Label17.Text + "' AND status='0'"
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
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "UPDATE tbl_PayrollData SET status='1' WHERE emp_id='" + Label29.Text.Trim + "' AND daterange='" + Label17.Text.Trim + "' AND status='0'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ListBox4.Items.Clear()
        MessageBox.Show("Salary Computation of " + Label15.Text.Trim + " for the daterange of " + Label17.Text.Trim + " is computed and saved successfully on the system and is now ready for payslip printing.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Dim valueA As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        valueA = Label14.Text + Val(TextBox3.Text)
        Label14.Text = valueA
        ListBox4.Items.Add(TextBox2.Text + " - " + TextBox3.Text + " Php")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        valueA = Label14.Text - Val(TextBox3.Text)
        Label14.Text = valueA
        ListBox4.Items.Add(TextBox2.Text + " - " + TextBox3.Text + " Php")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim paycomp As New Payroll_computation
        paycomp.Show()
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        savedata()
        Try
            If acsconn.State = ConnectionState.Open Then
                acsconn.Close()
            End If

            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_PayrollData WHERE status = '0' AND daterange='" + ComboBox1.Text + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox3.Items.Clear()
                ListBox3.Items.Add(acsdr("emp_id"))
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        upGSIS()
        upPagIBIG()
        upLPB()
        upNCIP()
        Label10.Text = "0.00"
        Label11.Text = "0.00"
        Label12.Text = "0.00"
        Label13.Text = "0.00"
        Label14.Text = "0.00"
        Label15.Text = "0.00"
        Label29.Text = "0.00"
        Label17.Text = "0.00"
        Label9.Text = "0.00"
        Label23.Text = "0.00"
        Label44.Text = "0.00"
        Label45.Text = "0.00"
        Label31.Text = "0.00"
        Label33.Text = "0.00"
        Label41.Text = "0.00"
        Label27.Text = ""
        Label36.Text = ""
        Label39.Text = ""
        Label98.Text = "0.00"
        Label93.Text = "2,000.00"
        Label54.Text = "0.00"
        Label60.Text = "0.00"
        Label61.Text = "0.00"
        Label62.Text = "0.00"
        Label63.Text = "0.00"
        Label64.Text = "0.00"
        Label65.Text = "0.00"
        Label74.Text = "0.00"
        Label90.Text = "0.00"
        Label87.Text = "0.00"
        Label84.Text = "0.00"
        Label81.Text = "0.00"
        Label78.Text = "0.00"
        T1.Text = ""
        T2.Text = ""
        T3.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ListBox4.Items.Clear()
        Button1.Enabled = False
        Button2.Enabled = False
        Button6.Enabled = False


    End Sub
    Dim bigbonus
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            bigbonus = "Yes"
        Else
            bigbonus = "No"
        End If
    End Sub
    Dim amnt As String
    Public Sub extramonth()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_payroll WHERE emp_id = '" + Label29.Text.Trim + "'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                ListBox5.Items.Add(acsdr("rate_hour").trim)
                ListBox6.Items.Add(acsdr("NH").trim)
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try
        For sum As Integer = 0 To ListBox5.Items.Count - 1
            Dim nh As String = CStr(ListBox5.Items(sum)).Trim
            Dim hr As String = CStr(ListBox6.Items(sum)).Trim
            Dim amount = Val(nh) * Val(hr)
            ListBox7.Items.Add(amount)
        Next
        For total As Integer = 0 To ListBox7.Items.Count - 1
            amnt = CStr(ListBox7.Items(total))
            amnt = Val(amnt) + Val(amnt)
        Next
        amnt = Val(amnt) / 12
        ListBox4.Items.Add("13th Month Pay - " + amnt)
        Label14.Text = Val(Label14.Text) + Val(amnt)
    End Sub
    Dim emL As String
    Dim cnL As String
    Dim plL As String
    Dim emLRM As String
    Dim cnLRM As String
    Dim plLRM As String

    Public Sub GSISLOAN()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='Emergency Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                emL = acsdr("payment_amount").trim
                emLRM = Val(acsdr("months_to_pay").trim) - 1
                If emL = "" Then
                    emL = 0
                    Label61.Text = "0.00"
                Else
                    Label61.Text = Val(emL).ToString("0,00.00")
                End If
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='Conso Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                cnL = acsdr("payment_amount").trim
                cnLRM = Val(acsdr("months_to_pay").trim) - 1
                If cnL = "" Then
                    cnL = 0
                    Label64.Text = "0.00"
                Else
                    Label64.Text = Val(cnL).ToString("0,00.00")
                End If
            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='Policy Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                plL = acsdr("payment_amount").trim
                plLRM = Val(acsdr("months_to_pay").trim) - 1
                If plL = "" Then
                    plL = 0
                    Label54.Text = "0.00"
                Else
                    Label54.Text = Val(plL).ToString("0,00.00")
                End If

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try
    End Sub

    Public Sub upGSIS()
        If Val(emLRM) >= 0 Then

            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + emLRM + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='Emergency Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

        If Val(cnLRM) >= 0 Then


            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + cnLRM + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type = 'Conso Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


        End If

        If Val(plLRM) >= 0 Then

            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + plLRM + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='Policy Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If


        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "DELETE FROM tbl_loans WHERE months_to_pay='0'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Dim pgL As String
    Dim pglRM As String
    Dim phL As String
    Dim phLRM As String
    Public Sub pagibigloan()
        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='Pag-IBIG Loan'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                pgL = acsdr("payment_amount").trim
                pglRM = Val(acsdr("months_to_pay").trim) - 1
                If pgL = "" Then
                    pgL = 0
                    Label63.Text = "0.00"
                Else
                    Label63.Text = Val(pgL).ToString("0,00.00")
                End If

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try

        Try
            acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            acsconn.Open()
            strsql = "SELECT * FROM tbl_loans WHERE emp_id = '" + Label29.Text.Trim + "' AND loan_type='Pag-IBIG Housing'"
            Dim acscmdr As New SqlCommand
            acscmdr.CommandText = strsql
            acscmdr.Connection = acsconn
            acsdr = acscmdr.ExecuteReader
            While (acsdr.Read())
                phL = acsdr("payment_amount").trim
                phLRM = Val(acsdr("months_to_pay").trim) - 1
                If phL = "" Then
                    phL = 0
                    Label62.Text = "0.00"
                Else
                    Label62.Text = Val(phL).ToString("0,00.00")
                End If

            End While
            acscmdr.Dispose()
            acsdr.Close()
            acsconn.Close()

        Catch ex As Exception

            acsdr.Close()
            acsconn.Close()
        End Try
    End Sub

    Public Sub upPagIBIG()



        If Val(pglRM) >= 0 Then
            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + pglRM + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='Pag-IBIG Loan'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        If Val(phLRM) >= 0 Then
            Try
                Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                Dim con As New SqlConnection
                con.ConnectionString = provider
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                Dim SqlCommand As New SqlCommand
                Dim SqlQuery As String = "UPDATE tbl_loans SET months_to_pay='" + phLRM + "' WHERE emp_id='" + Label29.Text.Trim + "' AND loan_type='Pag-IBIG Housing'"
                With SqlCommand
                    .CommandText = SqlQuery
                    .Connection = con
                    .ExecuteNonQuery()
                End With
                con.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        Try
            Dim provider As String = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
            Dim con As New SqlConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim SqlCommand As New SqlCommand
            Dim SqlQuery As String = "DELETE FROM tbl_loans WHERE months_to_pay='0'"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Private Sub Label74_Click(sender As Object, e As EventArgs) Handles Label74.Click

    End Sub
End Class