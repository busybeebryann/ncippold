
Public Class uploaded_Docs
    Private Sub uploaded_Docs_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        PictureBox1.Image = Image.FromFile(Application.StartupPath & "\docs\" + Label2.Text + (ComboBox1.Text).Trim + ".jpg")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        PictureBox1.Image = Nothing
    End Sub
End Class