Option Explicit On
Module Procedimientos

    'CheckSavedDocs
    'Input: "oAppCATIA As INFITF.Application"
    'Return: "blnHasChanged" (1 si hay documentos sin salvar, 0 si todos los documentos estan salvado)
    'Verifica que todos los documentos esten salvados antes de avanzar en la ejecución.
    'NOTA: No solo revisa los del currentDocument, sino todos los documentos en session
    Public Function CheckSavedDocs(oAppCATIA As INFITF.Application) As Boolean
        Dim colCATIADocuments As INFITF.Documents = oAppCATIA.Documents
        Dim blnHasChanged As Boolean
        For Each doc As INFITF.Document In colCATIADocuments
            blnHasChanged = Not doc.Saved
            If blnHasChanged = True Then
                Exit For
            End If
        Next
        Return blnHasChanged
    End Function


    'Revisa si ya existen archivos en el directorio destino
    'Input:  strDir: Directorio destino
    'Output:  ContenedorNombres As Dictionary(Of String, NameContainer)
    'Return: intFilesInExistance (cantidad de coincidencias)
    Public Function CountFilesInExistance(strDir As String, ContenedorNombres As Dictionary(Of String, NameContainer)) As Integer
        Dim intFilesInExistance As Integer = 0
        For Each kvp As KeyValuePair(Of String, NameContainer) In ContenedorNombres
            If Dir(strDir & "\" & kvp.Value.sNewNameWithExt) <> "" Then
                intFilesInExistance += 1
            End If
        Next
        Return intFilesInExistance
    End Function


    ' Se ingresa con el string de la columna y sale con la letra de esa columna
    ' Letra que representa la columna: A - B - C - D...
    ' Los nombres de los parámetros son buscados desde la columna "A" a las "Z"
    ' NOTA: Si hay parametros mas allá de la Z no seran encontrados
    ' Asigna la letra de la columna encontrada a una variable de tipo string
    Public Function EncuentraColumna(strTituloColumna As String, oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet) As String
        Dim myExcel As Microsoft.Office.Interop.Excel.Application = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.ActiveWorkbook
        Dim oRangeContainingCell As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1:Z1").Find(strTituloColumna, , , Microsoft.Office.Interop.Excel.XlLookAt.xlWhole)

        If oRangeContainingCell Is Nothing Then
            MsgBox("no se ha encontrado la columna")
        End If

        Dim strResultColum As String = Left(oRangeContainingCell.Address(RowAbsolute:=False, ColumnAbsolute:=False), 1)
        Return strResultColum
    End Function


    ' Escribir el prospecto de este procedimiento
    Sub FormatoGrupos(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet)

        oWorkSheet.Activate()
        Dim i As Integer = FindLastRowWithData(oWorkSheet)
        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = oWorkSheet.Parent
        Dim viewGrupos As Microsoft.Office.Interop.Excel.WorksheetView = oWorkBook.Windows.Item(1).SheetViews.Item(1)  'Está asignando el item 1 del total de todas las ventanas.
        Dim oRangeCabezera As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1:U2")
        Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A3", "U3")
        Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
        Dim strColumnLetter As String
        Dim a As String
        Dim b As String

        'Bordes del encabezado
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangeCabezera.Cells
            With c
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            End With
        Next


        'Bordes desde linea 3 hasta fin de los datos
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpo
            strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
            b = strColumnLetter & FindLastRowWithData(oWorkSheet)
            oCurrentRange = oWorkSheet.Range(a, b)
            With oCurrentRange
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            End With
        Next

        With oWorkSheet
            .Activate()
            .Range("A1", "U1").Orientation = 90
            .Range("A1", "U1").Font.Bold = True
            .Range("A1", "C1").Interior.Color = RGB(204, 255, 255)
            .Range("D1", "U1").Interior.ColorIndex = 15
            .Range("A2", "U2").Interior.ColorIndex = 15
            .Range("A3", "U" & i).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            .Range("A3", "U" & i).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        End With

        With oWorkSheet.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        'oWorkSheet.Range("A3", "U" & i).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
        'oWorkSheet.Range("A3", "U" & i).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        For Each C As Microsoft.Office.Interop.Excel.Range In oRangeCabezera
            C.EntireColumn.AutoFit()
        Next

        viewGrupos.DisplayGridlines = False

    End Sub

    '//************************************************************************************************************
    '// Da formato a una hoja de excel para ser completada con la información del procedimiento "CompletaLIstView2"
    '// Falta mejorar las opciones de las imagenes
    '//************************************************************************************************************
    Sub FormatoListView2(oWorkSheetListView As Microsoft.Office.Interop.Excel.Worksheet)

        oWorkSheetListView.Activate() : oWorkSheetListView.Name = "ListView"
        Dim i As Integer = FindLastRowWithData(oWorkSheetListView)
        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = oWorkSheetListView.Parent

        'Está asignando el item 1 del total de todas las ventanas.
        Dim viewListView As Microsoft.Office.Interop.Excel.WorksheetView = oWorkBook.Windows.Item(1).SheetViews.Item(1) : viewListView.DisplayGridlines = False
        Dim oRangoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A1", "U2")
        Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A3", "U3")
        Dim strColumnLetter As String
        Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
        Dim a As String
        Dim b As String


        ' // Arma el diccionario con los textos del encabezado. Si a futuro se requieren otras columnas
        ' // hay que modificar esto. Se pueden armar diccionarios o listas aparte y luego pasarlas como argumentos
        Dim oDicListViewColumnText As New Dictionary(Of String, String) From {
            {"A1", "Grupo"},
            {"B1", "Prefix"},
            {"C1", "Part Number"},
            {"D1", "Description"},
            {"E1", "Cantidad"},
            {"F1", "Conjunto - Parte"},
            {"G1", "Made or Bought"},
            {"H1", "-Libre-"},
            {"I1", "Vendor_Code_ID"},
            {"J1", "-libre-"},
            {"K1", "Material en Bruto"},
            {"L1", "Material"},
            {"M1", "Terminacion Superf"},
            {"N1", "Tratamiento Termico"},
            {"O1", "Peso"},
            {"P1", "Costo Unitario Estimado"},
            {"Q1", "Supplier/Vendor"},
            {"R1", "Lead Time [Week]"},
            {"S1", "Documento"},
            {"T1", "Obs."},
            {"U1", "Image"}
        }
        For Each kvp As KeyValuePair(Of String, String) In oDicListViewColumnText
            oWorkSheetListView.Range(kvp.Key).Value = kvp.Value
        Next

        ' // Bordes del encabezado
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado.Cells
            With c
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            End With
        Next

        ' // Fuente, tamaño y alineado de todo el documento
        With oWorkSheetListView.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        ' // Fuente, tamaño y alineado del encabezado
        With oWorkSheetListView
            .Range("A1", "U1").Orientation = 90
            .Range("A1", "U1").Font.Bold = True
            .Range("A1", "C1").Interior.Color = RGB(204, 255, 255)
            .Range("D1", "U1").Interior.ColorIndex = 15
            .Range("A2", "U2").Interior.ColorIndex = 15
        End With


        ' Hace AutoFit pero a la columna de imagenes no.
        ' Aca hay que incluir la opcion de que si la planilla va a tener imagenes entonces que no haga AutoFit,
        ' pero si son incluidas las imagenes, no debería hacer autofit.
        For Each C As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado
            C.EntireColumn.AutoFit()
        Next


        ' // Formato aplicado a todo el cuerpo
        With oWorkSheetListView
            .Range("A3", "U" & i).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            .Range("A3", "U" & i).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .Range("U3", "U" & i).RowHeight = 100
            .Range("U3", "U" & i).ColumnWidth = 18
        End With


        ' Para aplicar los bordes a cada columna hasta la última fila de datos,
        ' hay que hacer estos pasos para armar el rango
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpo
            strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
            b = strColumnLetter & FindLastRowWithData(oWorkSheetListView)
            oCurrentRange = oWorkSheetListView.Range(a, b)
            With oCurrentRange
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            End With
        Next

    End Sub


    Sub FormatoNCU(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet)

        oWorkSheet.Activate()
        Dim i As Integer = FindLastRowWithData(oWorkSheet)
        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = oWorkSheet.Parent
        Dim viewNCU As Microsoft.Office.Interop.Excel.WorksheetView = oWorkBook.Windows.Item(1).SheetViews.Item(1) 'Está asignando el item 1 del total de todas las ventanas.
        Dim oRangeEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1:F1")
        Dim oRangeToSort As Microsoft.Office.Interop.Excel.Range
        Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A2", "F2")
        Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
        Dim strColumnLetter As String
        Dim a As String
        Dim b As String

        ' // Completa el encabezado
        Dim oDicNCUColumnText As New Dictionary(Of String, String) From {
         {"A1", "-"},
         {"B1", "Descripcion"},
         {"C1", "PN-Interno"},
         {"D1", "Vendor Code"},
         {"E1", "Cantidad"},
         {"F1", "Obs."}
         }
        For Each kvp As KeyValuePair(Of String, String) In oDicNCUColumnText
            oWorkSheet.Range(kvp.Key).Value = kvp.Value
        Next

        ' // Fuente, tamaño y alineado
        With oWorkSheet.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        ' // Estas propiedades solo las aplica al encabezado
        With oRangeEncabezado
            .Interior.Color = 13693658
            .Font.Bold = True
            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            .EntireColumn.AutoFit()
        End With

        ' // Oculta todas las gridlines
        viewNCU.DisplayGridlines = False

        ' // Estas propiedades las aplica a todo el cuerpo con datos
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangeEncabezado
            strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
            b = strColumnLetter & i
            oCurrentRange = oWorkSheet.Range(a, b)
            With oCurrentRange
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                .EntireRow.AutoFit()
            End With
        Next

        ' // Sort order Ascending
        oRangeToSort = oWorkSheet.Range("A:F")
        oRangeToSort.Sort("C3:C" & i, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, , , , , ,
                          Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, , ,
                          Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns, ,
                          Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortTextAsNumbers)
    End Sub


    ' Escribir el prospecto de este procedimiento
    Sub CompletaColumnasGrupos(oWorkSheetGrupos As Microsoft.Office.Interop.Excel.Worksheet)
        oWorkSheetGrupos.Activate() : oWorkSheetGrupos.Name = "Grupos"
        Dim oDicGruposColumnText As New Dictionary(Of String, String) From {
        {"A1", "Grupo"},
        {"B1", "Prefix"},
        {"C1", "Part Number"},
        {"D1", "Description"},
        {"E1", "Cantidad"},
        {"F1", "Conjunto - Parte"},
        {"G1", "Made or Bought"},
        {"H1", "Assembly Level"},
        {"I1", "Vendor_Code_ID"},
        {"J1", "-Libre-"},
        {"K1", "Material en Bruto"},
        {"L1", "Material"},
        {"M1", "Terminacion Superf"},
        {"N1", "Tratamiento Termico"},
        {"O1", "Peso"},
        {"P1", "Costo Unitario Estimado"},
        {"Q1", "Supplier/Vendor"},
        {"R1", "Lead Time [Week]"},
        {"S1", "Documento"},
        {"T1", "Obs."},
        {"U1", "Note"}
    }
        SetColumnsText(oDicGruposColumnText, oWorkSheetGrupos)
    End Sub



    ' Escribir el prospecto de ésta función.-
    Private Function FindLastRowWithData(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Dim oRange As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1", "U1") ' solo trabaja desde columna A hasta U
        Dim i As Integer
        Dim j As Integer = 3
        Dim strColName As String
        For Each c As Microsoft.Office.Interop.Excel.Range In oRange.Columns
            strColName = Left(c.Columns.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            i = oWorkSheet.Range(strColName & oWorkSheet.Rows.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If i > j Then
                j = i
            End If
        Next
        Return j
    End Function

    ' Escribir el prospecto de ésta función.
    Private Function FindLastColumnWithData(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Dim i As Integer
        Dim oCells As Microsoft.Office.Interop.Excel.Range
        oCells = oWorkSheet.Cells(1, oWorkSheet.Columns.Count)
        i = oCells.End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column
        Return i
    End Function


    ' Escribe los textos en cada columna
    Private Sub SetColumnsText(oDicColumnText As Dictionary(Of String, String), oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet)
        For Each kvp As KeyValuePair(Of String, String) In oDicColumnText
            oWorkSheet.Range(kvp.Key).Value = kvp.Value
        Next
    End Sub

End Module







'Sub CompletaColumnasListView(oWorkSheetListView As Microsoft.Office.Interop.Excel.Worksheet)

'    oWorkSheetListView.Activate()
'    oWorkSheetListView.Name = "ListView"
'    Dim oDicListViewColumnText As New Dictionary(Of String, String) From {
'        {"A1", "Grupo"},
'        {"B1", "Prefix"},
'        {"C1", "Part Number"},
'        {"D1", "Description"},
'        {"E1", "Cantidad"},
'        {"F1", "Conjunto - Parte"},
'        {"G1", "Made or Bought"},
'        {"H1", "-Libre-"},
'        {"I1", "Vendor_Code_ID"},
'        {"J1", "-libre-"},
'        {"K1", "Material en Bruto"},
'        {"L1", "Material"},
'        {"M1", "Terminacion Superf"},
'        {"N1", "Tratamiento Termico"},
'        {"O1", "Peso"},
'        {"P1", "Costo Unitario Estimado"},
'        {"Q1", "Supplier/Vendor"},
'        {"R1", "Lead Time [Week]"},
'        {"S1", "Documento"},
'        {"T1", "Obs."},
'        {"U1", "Image"}
'    }
'    SetColumnsText(oDicListViewColumnText, oWorkSheetListView)
'End Sub




'' Escribir el prospecto de este procedimiento
'Sub FormatoListView(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet)

'    oWorkSheet.Activate()
'    Dim i As Integer = FindLastRowWithData(oWorkSheet)
'    Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = oWorkSheet.Parent
'    Dim viewListView As Microsoft.Office.Interop.Excel.WorksheetView = oWorkBook.Windows.Item(1).SheetViews.Item(1) 'Está asignando el item 1 del total de todas las ventanas.
'    Dim oRangeCabezera As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1:T1")
'    Dim oRangoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1", "U2")
'    Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A3", "U3")
'    Dim strColumnLetter As String
'    Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
'    Dim a As String
'    Dim b As String

'    'Bordes del encabezado
'    For Each c As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado.Cells
'        With c
'            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
'            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
'            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
'        End With
'    Next

'    '
'    With oWorkSheet.Cells
'        .Font.Name = "Tahoma"
'        .Font.Size = 8
'        .HorizontalAlignment = -4108
'        .VerticalAlignment = -4107
'    End With

'    With oWorkSheet
'        .Range("A1", "U1").Orientation = 90
'        .Range("A1", "U1").Font.Bold = True
'        .Range("A1", "C1").Interior.Color = RGB(204, 255, 255)
'        .Range("D1", "U1").Interior.ColorIndex = 15
'        .Range("A2", "U2").Interior.ColorIndex = 15
'    End With

'    oWorkSheet.Range("A3", "U" & i).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
'    oWorkSheet.Range("A3", "U" & i).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
'    oWorkSheet.Range("U3", "U" & i).RowHeight = 100
'    oWorkSheet.Range("U3", "U" & i).ColumnWidth = 18

'    ' Hace AutoFit pero a la columna de imagenes no.
'    ' Aca hay que incluir la opcion de que si la planilla va a tener imagenes entonces que no haga AutoFit,
'    ' pero si son incluidas las imagenes, no debería hacer autofit.
'    For Each C As Microsoft.Office.Interop.Excel.Range In oRangeCabezera
'        C.EntireColumn.AutoFit()
'    Next

'    For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpo
'        strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
'        a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
'        b = strColumnLetter & FindLastRowWithData(oWorkSheet)
'        oCurrentRange = oWorkSheet.Range(a, b)
'        With oCurrentRange
'            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
'            .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
'        End With
'    Next

'    viewListView.DisplayGridlines = False

'End Sub