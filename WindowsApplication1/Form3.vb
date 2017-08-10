Public Class Form3


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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        job_position.Show()
    End Sub
End Class