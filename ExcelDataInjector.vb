Option Explicit On
Option Strict On

Public Class ExcelDataInjector

    Sub InjectData(oSheetListView As Microsoft.Office.Interop.Excel.Worksheet, oDiccType3 As Dictionary(Of String, PwrProduct))

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Step 2/3: Filling Excel with extracted data...")

        Dim i As Integer = 3
        Dim oShape As Microsoft.Office.Interop.Excel.Shape

        ' Formateo masivo de celdas como texto antes de empezar
        Dim ultimoFila As Integer = oDiccType3.Count + 2
        oSheetListView.Range("A3:L" & ultimoFila).NumberFormat = "@"


        For Each kvp As KeyValuePair(Of String, PwrProduct) In oDiccType3

            Dim sImgPath As String = kvp.Value.ImageFilePath
            Dim oDoc As INFITF.Document = CType(kvp.Value.Product.ReferenceProduct.Parent, INFITF.Document) ' Para el nombre del archivo (Parent es un Document)

            With oSheetListView
                ' Asignación de valores con CType para cumplir con Option Strict On
                CType(.Cells(i, "A"), Microsoft.Office.Interop.Excel.Range).Value2 = i - 2
                CType(.Cells(i, "B"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.FullPath
                CType(.Cells(i, "C"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.FileName
                CType(.Cells(i, "D"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.PartNumber
                CType(.Cells(i, "E"), Microsoft.Office.Interop.Excel.Range).Value2 = ""
                CType(.Cells(i, "F"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.DescriptionRef
                CType(.Cells(i, "G"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Quantity
                CType(.Cells(i, "H"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Source
                CType(.Cells(i, "I"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Level
                CType(.Cells(i, "J"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.Nomenclature
                CType(.Cells(i, "K"), Microsoft.Office.Interop.Excel.Range).Value2 = kvp.Value.Product.Definition
                .Range(.Cells(i, 1), .Cells(i, 4)).Locked = True

                ' Inserción de imagen con coordenadas Single (CSng)
                If IO.File.Exists(sImgPath) Then
                    Dim cl As Microsoft.Office.Interop.Excel.Range = CType(.Cells(i, "L"), Microsoft.Office.Interop.Excel.Range)
                    oShape = .Shapes.AddPicture(sImgPath,
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        CSng(CDbl(cl.Left) + 5.5),
                                        CSng(CDbl(cl.Top) + 5.0),
                                        90, 90)
                End If
            End With
            i += 1
        Next
        oSheetListView.Protect(Contents:=True, UserInterfaceOnly:=True)
        oSheetListView.Application.ActiveWindow.DisplayVerticalScrollBar = True
        oSheetListView.Application.ActiveWindow.DisplayHorizontalScrollBar = True

    End Sub

End Class