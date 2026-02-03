Public Class ExcelDataUpdater



    Sub UpdateData(oSheetListView As Microsoft.Office.Interop.Excel.Worksheet, oCatiaData As Dictionary(Of String, PwrProduct))

        Dim lastRow As Integer = GetLastRow(oSheetListView)

        If lastRow < 3 Then Return
        oSheetListView.Unprotect()
        oSheetListView.Cells.Locked = False
        Dim i As Integer = 3
        For Each kvp As KeyValuePair(Of String, PwrProduct) In oCatiaData
            Dim oDoc As INFITF.Document = CType(kvp.Value.Product.ReferenceProduct.Parent, INFITF.Document) ' Para el nombre del archivo (Parent es un Document)
            With oSheetListView
                ' Asignación de valores con CType para cumplir con Option Strict On
                CType(.Cells(i, "A"), Microsoft.Office.Interop.Excel.Range).Value2 = i - 2
                CType(.Cells(i, "B"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.FullPath
                CType(.Cells(i, "C"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.FileName
                CType(.Cells(i, "D"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.PartNumber
                .Range(.Cells(i, 1), .Cells(i, 4)).Locked = True
                CType(.Cells(i, "E"), Microsoft.Office.Interop.Excel.Range).Value2 = ""
                CType(.Cells(i, "F"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.DescriptionRef
                CType(.Cells(i, "G"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Quantity
                CType(.Cells(i, "H"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Source
                CType(.Cells(i, "I"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Level
                CType(.Cells(i, "J"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.Nomenclature
                CType(.Cells(i, "K"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.Definition
            End With
            i += 1
        Next
        oSheetListView.Protect(Contents:=True, UserInterfaceOnly:=True)
    End Sub




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


