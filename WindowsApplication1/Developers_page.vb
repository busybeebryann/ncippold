Public Class Developers_page



    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub Developers_page_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Created by: Timothy Martin Masagca.")
    End Sub
End Class