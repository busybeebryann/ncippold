Public Class mainform

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Add_Employee.Show()
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        employeelist.Show()
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        job_position.Show()
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        dept_add.Show()
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)

        Login.Show()
        Login.TextBox1.Text = ""
        Login.TextBox2.Text = ""
        Me.Close()

    End Sub

    Private Sub LinkLabel18_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Login.Show()
        Login.TextBox1.Text = ""
        Login.TextBox2.Text = ""
        Me.Close()
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        attendance_tracker.Show()
    End Sub

    Private Sub LinkLabel10_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        payroll_daterange.Show()
    End Sub

    Private Sub LinkLabel14_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Payroll_computation.Show()
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs)
        Developers_page.Show()
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        attendance_overide.Show()
    End Sub

    Private Sub LinkLabel11_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        employee_list_reporting.Show()
    End Sub
    Private Sub LinkLabel15_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        payslip_filter.Show()
    End Sub

    Private Sub LinkLabel20_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        payslip_tracker.Show()
    End Sub

    Private Sub LinkLabel19_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        attendance_tracker.Show()
    End Sub

    Private Sub LinkLabel22_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        update_profile.Show()
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        update_profile.Show()
    End Sub

    Private Sub LinkLabel13_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        alphalist_generator.Show()
    End Sub
    Dim pictureSlide As Integer = 0
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        pictureSlide = pictureSlide + 1

        If pictureSlide = 1 Then
            PictureBox11.Image = My.Resources._1

        ElseIf pictureSlide = 2 Then
            PictureBox11.Image = My.Resources._2

        ElseIf pictureSlide = 3 Then
            PictureBox11.Image = My.Resources._3

        ElseIf pictureSlide = 4 Then
            PictureBox11.Image = My.Resources._6
        Else
            pictureSlide = 0
        End If
    End Sub

    Private Sub mainform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()


    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        employeelist.Show()
    End Sub

    Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer1.Tick
        pictureSlide = pictureSlide + 1

        If pictureSlide = 1 Then
            PictureBox11.Image = My.Resources._1

        ElseIf pictureSlide = 2 Then
            PictureBox11.Image = My.Resources._2

        ElseIf pictureSlide = 3 Then
            PictureBox11.Image = My.Resources._3

        ElseIf pictureSlide = 4 Then
            PictureBox11.Image = My.Resources._6
        Else
            pictureSlide = 0
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        job_position.Show()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        dept_add.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        attendance_tracker.Show()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        attendance_overide.Show()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        payroll_daterange.Show()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Payroll_computation.Show()

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        payslip_filter.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        update_profile.Show()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Login.Show()
        Login.TextBox1.Text = ""
        Login.TextBox2.Text = ""
        Me.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        If Panel2.Visible = False Then
            Panel2.Visible = True
            Panel3.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False


        Else
            Panel2.Visible = False
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If Panel3.Visible = False Then
            Panel3.Visible = True
            Panel2.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            Panel6.Visible = False



        Else
            Panel3.Visible = False

        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Panel4.Visible = False Then
            Panel4.Visible = True
            Panel3.Visible = False
            Panel5.Visible = False
            Panel2.Visible = False
            Panel6.Visible = False

        Else
            Panel4.Visible = False
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If Panel5.Visible = False Then
            Panel5.Visible = True
            Panel4.Visible = False
            Panel3.Visible = False
            Panel2.Visible = False
            Panel6.Visible = False

        Else
            Panel5.Visible = False
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If Panel6.Visible = False Then
            Panel6.Visible = True
            Panel5.Visible = False
            Panel4.Visible = False
            Panel3.Visible = False
            Panel2.Visible = False

        Else
            Panel6.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Add_Employee.Show()

    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        Login.Show()
        Login.TextBox1.Text = ""
        Login.TextBox2.Text = ""
        Me.Close()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        attendance_tracker.Show()
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs)
        update_profile.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

        If Panel11.Visible = False Then
            Panel11.Visible = True
            Panel10.Visible = False
            Panel12.Visible = False

        Else
            Panel11.Visible = False
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        If Panel10.Visible = False Then
            Panel10.Visible = True
            Panel11.Visible = False
            Panel12.Visible = False

        Else
            Panel10.Visible = False
        End If

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If Panel12.Visible = False Then
            Panel12.Visible = True
            Panel10.Visible = False
            Panel11.Visible = False

        Else
            Panel12.Visible = False
        End If
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        payslip_tracker.Show()

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Manual_Attendance.Show()

    End Sub

    Private Sub Button31_Click_1(sender As Object, e As EventArgs) Handles Button31.Click
        update_profile.Show()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        E201_list.Show()

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        LoanForms.Show()
    End Sub


    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        LeaveFiler.Show()
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class