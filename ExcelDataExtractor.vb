Option Explicit On

Public Class ExcelDataExtractor

    Function ExtractData(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Dictionary(Of String, ExcelData)

        ' 1. Obtener última fila de forma segura
        Dim lastRow As Integer = GetLastRow(oSheet)

        ' Si la hoja no tiene datos después del encabezado, salimos temprano
        Dim oDic As New Dictionary(Of String, ExcelData)
        If lastRow < 3 Then Return oDic

        For i As Integer = 3 To lastRow
            ' Clave única (Columna D)
            Dim key As String = oSheet.Cells(i, 4).Text.ToString()

            ' Validamos que la clave no esté vacía antes de agregar
            If Not String.IsNullOrWhiteSpace(key) AndAlso Not oDic.ContainsKey(key) Then
                Dim oExcelData As New ExcelData With {
                    .NewPartNumber = oSheet.Cells(i, 5).Text.ToString(),  ' E
                    .DescriptionRef = oSheet.Cells(i, 6).Text.ToString(), ' F
                    .Quantity = oSheet.Cells(i, 7).Text.ToString(),       ' G
                    .Source = oSheet.Cells(i, 8).Text.ToString(),         ' H
                    .Nomenclature = oSheet.Cells(i, 10).Text.ToString(),  ' J
                    .Definition = oSheet.Cells(i, 11).Text.ToString()     ' K
                }
                oDic.Add(key, oExcelData)
            End If
        Next

        Return oDic
    End Function

    Private Function GetLastRow(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Try
            Dim lastCell = oSheet.Cells.Find("*", , , ,
                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious)

            If lastCell Is Nothing Then Return 0
            Return lastCell.Row
        Catch
            Return 0
        End Try
    End Function
End Class
