Public Class Form1


    Public Sub New()
        InitializeComponent()
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Program.Main()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Program.Update()
    End Sub
End Class
