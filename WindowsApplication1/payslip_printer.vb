Imports System.Data
Imports System.Data.OleDb

Public Class payslip_printer
    Dim dsOrder As New DataSet
    Dim strsql As String
    Private Sub payslip_printer_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        PrintPreview()
        del()
    End Sub

    Function GetOrderList()
        Dim ds As New DataSet
        Dim conn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim da As New OleDbDataAdapter
        Dim sConnstring As String
        Try
            sConnstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
            conn = New OleDbConnection(sConnstring)
            cmd.Connection = conn
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM tbl_payroll"
            da.SelectCommand = cmd
            da.Fill(ds)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return ds
    End Function

    Public Sub PrintPreview()
        Dim rptDocument As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If rptDocument Is Nothing Then
                rptDocument = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            End If

            rptDocument.Load(Application.StartupPath & "\Payslip.rpt")
            dsOrder = GetOrderList()
            rptDocument.SetDataSource(dsOrder.Tables(0))

            CrystalReportViewer1.ReportSource = rptDocument


        Catch ex As Exception
            MessageBox.Show("mod_generic - PrintPreview" & ControlChars.NewLine & ex.Message.ToString())
            Exit Sub
        Finally

        End Try
    End Sub
   
    Public Sub del()
        Try
            Dim provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\report.accdb"
            Dim con As New OleDbConnection
            con.ConnectionString = provider
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            Dim SqlCommand As New OleDbCommand
            Dim SqlQuery As String = "DELETE FROM tbl_payroll"
            With SqlCommand
                .CommandText = SqlQuery
                .Connection = con
                .ExecuteNonQuery()
            End With

            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
End Class
