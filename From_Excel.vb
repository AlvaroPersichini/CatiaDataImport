
Module From_Excel


    'Sub FromExcelToCATIA(oProduct As ProductStructureTypeLib.Product)

    '    Dim myExcel As Microsoft.Office.Interop.Excel.Application = GetObject(, "Excel.Application")
    '    Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.ActiveWorkbook
    '    Dim oWorkSheetGrupos As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.Worksheets.Item(1) '1: Grupos
    '    Dim oWorkSheetNCU As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.Worksheets.Item(2) '1: NCU

    '    Dim PartNumCol As String = Procedimientos.EncuentraColumna("PN-Interno", oWorkSheetNCU)
    '    Dim DescriptionCol As String = Procedimientos.EncuentraColumna("Description", oWorkSheetNCU)
    '    Dim Vendor_Code_IDCol As String = Procedimientos.EncuentraColumna("Vendor_Code_ID", oWorkSheetNCU)
    '    Dim VendorCol As String = Procedimientos.EncuentraColumna("Vendor", oWorkSheetNCU)

    '    Dim MaterialenBrutoCol As String = Procedimientos.EncuentraColumna("Material en Bruto", oWorkSheetGrupos)


    '    Dim PowerDiccionary As Dictionary(Of String, PowerProduct) = DiccT2_PN_PwrProduct(oProduct)

    '    For Each kvp As KeyValuePair(Of String, PowerProduct) In PowerDiccionary

    '        '    For i = 3 To 131

    '        '        If kvp.Value.oProduct.PartNumber = oWorkSheetNCU.Cells.Range(PartNumCol & i).Text Then

    '        '            kvp.Value.oProduct.DescriptionRef = oWorkSheetNCU.Cells.Range(DescriptionCol & i).Text
    '        '            kvp.Value.oProduct.Source = 2
    '        '            kvp.Value.oProduct.Definition = oWorkSheetNCU.Cells.Range(Vendor_Code_IDCol & i).Text

    '        '        End If

    '        '    Next

    '        'Next

    '        'For Each kvp As KeyValuePair(Of String, PowerProduct) In PowerDiccionary

    '        '    For i = 3 To 244



    '        '    Next

    '    Next





    'For Each p As ProductStructureTypeLib.Product In valColl
    '    For i = 3 To 95
    '        If p.PartNumber = oWorkSheet.Cells.Range(PartNumCol & i).Text Then
    '            p.DescriptionRef = oWorkSheet.Cells.Range(DescriptionCol & i).Text
    '            p.Source = 2
    '            p.Definition = oWorkSheet.Cells.Range(Vendor_Code_IDCol & i).Text
    '        End If
    '    Next
    'Next

    'For Each p As ProductStructureTypeLib.Product In valColl
    '    If Left(p.PartNumber, 3) <> "NCU" Then
    '        p.Source = 1
    '    End If
    'Next

    'hay que poner la línea para que actualice el árbol

    '  End Sub





End Module
