Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Data.OleDb

Public Class full_payroll
    Dim acsconn As New SqlConnection
    Dim acsdr As SqlDataReader
    Dim strsql As String
    Dim province As String
    Dim pc As Integer = 1
    Dim mp, pr, ts, gs, go, pg, ph, tx, cl, el, ml, pl, ex, gl, hl, lp, nl, nm, na, td, np, s5, s0, ne As String
    Dim mp2, pr2, ts2, gs2, go2, pg2, ph2, tx2, cl2, el2, ml2, pl2, ex2, gl2, hl2, lp2, nl2, nm2, na2, td2, np2, s52, s02 As String
    Dim v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19, v20, v21, v22, v23, v24 As String

    Private Sub full_payroll_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
    Dim tblval
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pc = 0
        Do While pc <= 8
            If pc = 1 Then
                province = "Region"
                tblval = "tbl_region"
            ElseIf pc = 2 Then
                province = "Abra"
                tblval = "tbl_abra"
            ElseIf pc = 3 Then
                province = "Apayao"
                tblval = "tbl_apayao"
            ElseIf pc = 4 Then
                province = "Baguio City"
                tblval = "tbl_baguio"
            ElseIf pc = 5 Then
                province = "Benguet"
                tblval = "tbl_benguet"
            ElseIf pc = 6 Then
                province = "Mt. Province"
                tblval = "tbl_mt_province"
            ElseIf pc = 7 Then
                province = "Ifugao"
                tblval = "tbl_ifugao"
            ElseIf pc = 8 Then
                province = "Kalinga"
                tblval = "tbl_kalinga"
            End If

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

            mp2 = "0"
            pr2 = "0"
            ts2 = "0"
            gs2 = "0"
            go2 = "0"
            pg2 = "0"
            ph2 = "0"
            tx2 = "0"
            cl2 = "0"
            el2 = "0"
            ml2 = "0"
            pl2 = "0"
            ex2 = "0"
            gl2 = "0"
            hl2 = "0"
            lp2 = "0"
            nl2 = "0"
            nm2 = "0"
            na2 = "0"
            td2 = "0"
            np2 = "0"
            s52 = "0"
            s02 = "0"

            

            Try
                If acsconn.State = ConnectionState.Open Then
                    acsconn.Close()
                End If

                acsconn.ConnectionString = "Data Source=BUSYBEE-CHAD-PC\SQLEXPRESS2;Initial Catalog=payroll_db;User ID=sa;Password=mainpass"
                acsconn.Open()
                strsql = "SELECT * FROM tbl_payroll WHERE province='" + province + "' AND month_computed = '" + ComboBox1.Text + "'"
                Dim acscmdr As New SqlCommand
                acscmdr.CommandText = strsql
                acscmdr.Connection = acsconn
                acsdr = acscmdr.ExecuteReader
                While (acsdr.Read())
                    mp = ((Val(Regex.Replace(acsdr("monthly_pay"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox1.Items.Add(mp)

                    ListBox25.Items.Add(mp)

                    pr = ((Val(Regex.Replace(acsdr("PERA"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox2.Items.Add(pr)
                    ts = ((Val(Regex.Replace(acsdr("Grosspay"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox3.Items.Add(ts)

                    gs = ((Val(Regex.Replace(acsdr("gsis_social_cont"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox4.Items.Add(gs)
                    go = ((Val(Regex.Replace(acsdr("uoli_cont"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox5.Items.Add(go)
                    pg = ((Val(Regex.Replace(acsdr("Pagibig"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox6.Items.Add(pg)
                    ph = ((Val(Regex.Replace(acsdr("Philhealth"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox7.Items.Add(ph)
                    tx = ((Val(Regex.Replace(acsdr("TAX"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox8.Items.Add(tx)

                    cl = ((Val(Regex.Replace(acsdr("conso_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox9.Items.Add(cl)
                    el = ((Val(Regex.Replace(acsdr("educ_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox10.Items.Add(el)
                    ml = ((Val(Regex.Replace(acsdr("emergency_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox11.Items.Add(mp)
                    pl = ((Val(Regex.Replace(acsdr("policy_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox12.Items.Add(pl)
                    ex = ((Val(Regex.Replace(acsdr("e_card"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox13.Items.Add(ex)
                    gl = ((Val(Regex.Replace(acsdr("pagibig_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox14.Items.Add(gl)
                    hl = ((Val(Regex.Replace(acsdr("pagibig_housing"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox15.Items.Add(hl)
                    lp = ((Val(Regex.Replace(acsdr("lbp_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox16.Items.Add(lp)
                    nl = ((Val(Regex.Replace(acsdr("ncipae_loan"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox17.Items.Add(nl)
                    nm = ((Val(Regex.Replace(acsdr("ncipae_mem_fee"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox18.Items.Add(nm)
                    na = ((Val(Regex.Replace(acsdr("ncipae_man_fee"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox19.Items.Add(na)
                    td = ((Val(Regex.Replace(acsdr("tDed"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox20.Items.Add(td)
                    np = ((Val(Regex.Replace(acsdr("net_pay"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1)
                    ListBox21.Items.Add(np)

                    s5 = (((Val(Regex.Replace(acsdr("net_pay"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1) / 2)
                    ListBox22.Items.Add(s5)
                    s0 = (((Val(Regex.Replace(acsdr("net_pay"), "[^A-Za-z0-9\-/]", "")) * 0.1) * 0.1) / 2)
                    ListBox23.Items.Add(s0)


                    

                    For x As Integer = 0 To ListBox1.Items.Count - 1
                        mp2 = mp2 + Val(ListBox1.Items.Item(x).ToString)
                    Next
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add(mp2)


                    For x As Integer = 0 To ListBox2.Items.Count - 1
                        pr2 = pr2 + Val(ListBox2.Items.Item(x).ToString)
                    Next
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add(pr2)


                    For x As Integer = 0 To ListBox3.Items.Count - 1
                        ts2 = ts2 + Val(ListBox3.Items.Item(x).ToString)
                    Next
                    ListBox3.Items.Clear()
                    ListBox3.Items.Add(ts2)


                    For x As Integer = 0 To ListBox4.Items.Count - 1
                        gs2 = gs2 + Val(ListBox4.Items.Item(x).ToString)
                    Next
                    ListBox4.Items.Clear()
                    ListBox4.Items.Add(gs2)


                    For x As Integer = 0 To ListBox5.Items.Count - 1
                        go2 = go2 + Val(ListBox5.Items.Item(x).ToString)
                    Next
                    ListBox5.Items.Clear()
                    ListBox5.Items.Add(go2)


                    For x As Integer = 0 To ListBox6.Items.Count - 1
                        pg2 = pg2 + Val(ListBox6.Items.Item(x).ToString)
                    Next
                    ListBox6.Items.Clear()
                    ListBox6.Items.Add(pg2)


                    For x As Integer = 0 To ListBox7.Items.Count - 1
                        ph2 = ph2 + Val(ListBox7.Items.Item(x).ToString)
                    Next
                    ListBox7.Items.Clear()
                    ListBox7.Items.Add(ph2)


                    For x As Integer = 0 To ListBox8.Items.Count - 1
                        tx2 = tx2 + Val(ListBox8.Items.Item(x).ToString)
                    Next
                    ListBox8.Items.Clear()
                    ListBox8.Items.Add(tx2)


                    For x As Integer = 0 To ListBox9.Items.Count - 1
                        cl2 = cl2 + Val(ListBox9.Items.Item(x).ToString)
                    Next
                    ListBox9.Items.Clear()
                    ListBox9.Items.Add(cl2)


                    For x As Integer = 0 To ListBox10.Items.Count - 1
                        el2 = el2 + Val(ListBox10.Items.Item(x).ToString)
                    Next
                    ListBox10.Items.Clear()
                    ListBox10.Items.Add(el2)


                    For x As Integer = 0 To ListBox11.Items.Count - 1
                        ml2 = ml2 + Val(ListBox11.Items.Item(x).ToString)
                    Next
                    ListBox11.Items.Clear()
                    ListBox11.Items.Add(ml2)


                    For x As Integer = 0 To ListBox12.Items.Count - 1
                        pl2 = pl2 + Val(ListBox12.Items.Item(x).ToString)
                    Next
                    ListBox12.Items.Clear()
                    ListBox12.Items.Add(pl2)


                    For x As Integer = 0 To ListBox13.Items.Count - 1
                        ex2 = ex2 + Val(ListBox13.Items.Item(x).ToString)
                    Next
                    ListBox13.Items.Clear()
                    ListBox13.Items.Add(ex2)


                    For x As Integer = 0 To ListBox14.Items.Count - 1
                        gl2 = gl2 + Val(ListBox14.Items.Item(x).ToString)
                    Next
                    ListBox14.Items.Clear()
                    ListBox14.Items.Add(gl2)


                    For x As Integer = 0 To ListBox15.Items.Count - 1
                        hl2 = hl2 + Val(ListBox15.Items.Item(x).ToString)
                    Next
                    ListBox15.Items.Clear()
                    ListBox15.Items.Add(hl2)


                    For x As Integer = 0 To ListBox16.Items.Count - 1
                        lp2 = lp2 + Val(ListBox16.Items.Item(x).ToString)
                    Next
                    ListBox16.Items.Clear()
                    ListBox16.Items.Add(lp2)

                    
                    For x As Integer = 0 To ListBox17.Items.Count - 1
                        nl2 = nl2 + Val(ListBox17.Items.Item(x).ToString)
                    Next
                    ListBox17.Items.Clear()
                    ListBox17.Items.Add(nl2)


                    For x As Integer = 0 To ListBox18.Items.Count - 1
                        nm2 = nm2 + Val(ListBox18.Items.Item(x).ToString)
                    Next
                    ListBox18.Items.Clear()
                    ListBox18.Items.Add(nm2)


                    For x As Integer = 0 To ListBox19.Items.Count - 1
                        na2 = na2 + Val(ListBox19.Items.Item(x).ToString)
                    Next
                    ListBox19.Items.Clear()
                    ListBox19.Items.Add(na2)


                    For x As Integer = 0 To ListBox20.Items.Count - 1
                        td2 = td2 + Val(ListBox20.Items.Item(x).ToString)
                    Next
                    ListBox20.Items.Clear()
                    ListBox20.Items.Add(td2)


                    For x As Integer = 0 To ListBox21.Items.Count - 1
                        np2 = np2 + Val(ListBox21.Items.Item(x).ToString)
                    Next
                    ListBox21.Items.Clear()
                    ListBox21.Items.Add(np2)


                    For x As Integer = 0 To ListBox22.Items.Count - 1
                        s52 = s52 + Val(ListBox22.Items.Item(x).ToString)
                    Next
                    ListBox22.Items.Clear()
                    ListBox22.Items.Add(s52)


                    For x As Integer = 0 To ListBox23.Items.Count - 1
                        s02 = s02 + Val(ListBox23.Items.Item(x).ToString)
                    Next
                    ListBox23.Items.Clear()
                    ListBox23.Items.Add(s02)


                End While
                acscmdr.Dispose()
                acsdr.Close()
                acsconn.Close()
                ne = ListBox25.Items.Count
                ListBox24.Items.Add(ne)

                v1 = ""
                v2 = ""
                v3 = ""
                v4 = ""
                v5 = ""
                v6 = ""
                v7 = ""
                v8 = ""
                v9 = ""
                v10 = ""
                v11 = ""
                v12 = ""
                v13 = ""
                v14 = ""
                v15 = ""
                v16 = ""
                v17 = ""
                v18 = ""
                v19 = ""
                v20 = ""
                v21 = ""
                v22 = ""
                v23 = ""
                v24 = ""

                v1 = mp2
                v2 = pr2
                v3 = ts2
                v4 = gs2
                v5 = go2
                v6 = pg2
                v7 = ph2
                v8 = tx2
                v9 = cl2
                v10 = el2
                v11 = ml2
                v12 = pl2
                v13 = ex2
                v14 = gl2
                v15 = hl2
                v16 = lp2
                v17 = nl2
                v18 = nm2
                v19 = na2
                v20 = td2
                v21 = np2
                v22 = s52
                v23 = s02
                v24 = ne

                If v1 = "" Then
                    v1 = "0.00"
                End If
                If v2 = "" Then
                    v2 = "0.00"
                End If
                If v3 = "" Then
                    v3 = "0.00"
                End If
                If v4 = "" Then
                    v4 = "0.00"
                End If
                If v5 = "" Then
                    v5 = "0.00"
                End If
                If v6 = "" Then
                    v6 = "0.00"
                End If
                If v7 = "" Then
                    v7 = "0.00"
                End If
                If v8 = "" Then
                    v8 = "0.00"
                End If
                If v9 = "" Then
                    v9 = "0.00"
                End If
                If v10 = "" Then
                    v10 = "0.00"
                End If
                If v11 = "" Then
                    v11 = "0.00"
                End If
                If v12 = "" Then
                    v12 = "0.00"
                End If
                If v13 = "" Then
                    v13 = "0.00"
                End If
                If v14 = "" Then
                    v14 = "0.00"
                End If
                If v15 = "" Then
                    v15 = "0.00"
                End If
                If v16 = "" Then
                    v16 = "0.00"
                End If
                If v17 = "" Then
                    v17 = "0.00"
                End If
                If v18 = "" Then
                    v18 = "0.00"
                End If
                If v19 = "" Then
                    v19 = "0.00"
                End If
                If v20 = "" Then
                    v20 = "0.00"
                End If
                If v21 = "" Then
                    v21 = "0.00"
                End If
                If v22 = "" Then
                    v22 = "0.00"
                End If
                If v23 = "" Then
                    v23 = "0.00"
                End If
                If v24 = "" Then
                    v24 = "0.00"
                End If
                If (tblval = "") Then
                    'donothing
                Else
                    Try
                        Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
                        Dim con As New OleDbConnection
                        con.ConnectionString = provider
                        If con.State = ConnectionState.Closed Then
                            con.Open()
                        End If

                        Dim SqlCommand As New OleDbCommand
                        Dim SqlQuery As String = "INSERT INTO " + tblval + " ([mp], [pr], [ts], [gs], [go], [pg], [ph], [tx], [cl], [el], [ml], [pl], [ex], [gl], [hl], [lp], [nl], [nm], [na], [td], [np], [s5], [s0], [ne], [pro]) VALUES ('" + v1 + "','" + v2 + "','" + v3 + "','" + v4 + "','" + v5 + "','" + v6 + "','" + v7 + "','" + v8 + "','" + v9 + "','" + v10 + "','" + v11 + "','" + v12 + "','" + v13 + "','" + v14 + "','" + v15 + "','" + v16 + "','" + v17 + "','" + v18 + "','" + v19 + "','" + v20 + "','" + v21 + "','" + v22 + "','" + v23 + "','" + v24 + "','" + ComboBox1.Text + "')"

                        With SqlCommand
                            .CommandText = SqlQuery
                            .Connection = con
                            .ExecuteNonQuery()
                        End With

                        con.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try

                End If
               

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            pc = pc + 1
        Loop

        Form4.Show()
    End Sub
End Class