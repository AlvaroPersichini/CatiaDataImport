Public Class Form1


    Public Sub New()

        InitializeComponent()

        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        StartPosition = FormStartPosition.CenterScreen
        Size = New Size(300, 200)
        Text = "CATIA-Excel Sync"
        ToolStripStatusLabel1.Text = "Ready"

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Program.Main()

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Program.Update()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub



End Class
