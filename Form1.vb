Public Class Form1


    Public Sub New()
        InitializeComponent()

        ' Bloquea el redimensionado arrastrando bordes
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        ' Quita el botón de maximizar de la esquina superior derecha
        Me.MaximizeBox = False

        ' Opcional: Centra el formulario al abrirse ya que no se puede mover/agrandar
        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Program.Main()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Program.Update()

    End Sub
End Class
