Option Explicit On



Module UserProperties

    ' Property UserRefProperties() As Parameters (Read Only)  (KnowledgewareTypeLib.Parameters)
    ' Returns the collection Object containing the product properties.
    ' All the user defined properties that are created In the reference product might be accessed through that collection. 
    ' Only available On reference products. 

    'Hacer sobrecargas de este metodo. Cada sobrecarga para cada tipo de valor de la propiedad:
    ' 1) String, 2) Boolean, 3) Integer, etc.
    Sub AddUserProperty(oProduct As ProductStructureTypeLib.Product, strPropName As String)

        Dim oRefProduct As ProductStructureTypeLib.Product = oProduct.ReferenceProduct
        Dim Parametros As KnowledgewareTypeLib.Parameters = oRefProduct.UserRefProperties

        'No se porque, pero hay que renombrarlo. Se ve que al crearlo con su nombre, en realidad lo que hace es agregar el nombre
        'a una ruta donde está todo el camino hasta el nombre del parametro.
        'Al renombrarlo, elimina el nombre que tiene de ruta completa,
        'y solo deja el nombre que se le da con el metodo "rename"

        If Parametros.Count = 0 Then
            Parametros.CreateString(strPropName, "")
            Parametros.Item(oProduct.PartNumber & "\" & "Properties" & "\" & strPropName).Rename(strPropName)
        Else
            For Each Parametro As KnowledgewareTypeLib.Parameter In Parametros
                If Parametro.Name = strPropName Then
                    Exit Sub
                End If
            Next
            Parametros.CreateString(strPropName, "")
            Parametros.Item(oProduct.PartNumber & "\" & "Properties" & "\" & strPropName).Rename(strPropName)
        End If

    End Sub

    ' Hacer sobrecargas de este metodo. Cada sobrecarga para cada tipo de valor de la propiedad:
    ' 1) String, 2) Boolean, 3) Integer, etc.
    Public Sub AddUserPropValue(oProduct As ProductStructureTypeLib.Product,
                                         strPropName As String,
                                         strValue As String)
        Dim oRefProduct As ProductStructureTypeLib.Product = oProduct.ReferenceProduct
        Dim Parametros As KnowledgewareTypeLib.Parameters = oRefProduct.UserRefProperties
        If Parametros.Count = 0 Then
            Exit Sub
        Else
            For Each Parametro As KnowledgewareTypeLib.Parameter In Parametros
                If Parametro.Name = strPropName Then
                    oRefProduct.UserRefProperties.Item(strPropName).ValuateFromString(strValue)
                End If
            Next
        End If
    End Sub



    'Esto funciona, pero falta terminar
    Public Sub FillUserPropFromExcel(oProduct As ProductStructureTypeLib.Product)

        Dim oDic As Dictionary(Of String, ProductStructureTypeLib.Product) = Diccionarios.DiccT1_PN_oProduct(oProduct)
        Dim myExcel As Microsoft.Office.Interop.Excel.Application = GetObject(, "Excel.Application")
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.ActiveWorkbook
        ' Dim oWorkSheetGrupos As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.Worksheets.Item(1)
        Dim oSheetListView As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.Worksheets.Item(2)

        Dim oRange As Microsoft.Office.Interop.Excel.Range = oSheetListView.Range("C3", "C229")

        For Each C As Microsoft.Office.Interop.Excel.Range In oRange

            If oDic.ContainsKey(C.Value) Then

                If C.Offset(, 8).Value <> "" Then
                    AddUserProperty(oDic.Item(C.Value), "Material en Bruto")
                    AddUserPropValue(oDic.Item(C.Value), "Material en Bruto", C.Offset(, 8).Value)
                End If

                If C.Offset(, 9).Value <> "" Then
                    AddUserProperty(oDic.Item(C.Value), "Material")
                    AddUserPropValue(oDic.Item(C.Value), "Material", C.Offset(, 9).Value)
                End If

            End If

        Next

    End Sub




    Sub ClearUserProperties()

    End Sub



    Sub RemoveAllUserProperties(oCurrentProduct As ProductStructureTypeLib.Product)
        Dim oDic1 As Dictionary(Of String, ProductStructureTypeLib.Product) = DiccT1_PN_oProduct(oCurrentProduct)
        Dim Parametros As KnowledgewareTypeLib.Parameters
        Dim strParam As String

        For Each kvp As KeyValuePair(Of String, ProductStructureTypeLib.Product) In oDic1
            Parametros = kvp.Value.ReferenceProduct.UserRefProperties
            If Not kvp.Value.ReferenceProduct.UserRefProperties.Count = 0 Then
                For Each Parametro As KnowledgewareTypeLib.Parameter In Parametros
                    strParam = Parametro.Name
                    Parametros.Remove(strParam)
                Next
            Else
                GoTo SiguienteP
            End If
SiguienteP:
        Next

    End Sub


End Module











'    Sub CheckIfPropExist(oSelection As INFITF.Selection)
'        For i = 1 To oSelection.Count2
'            Dim oProduct As ProductStructureTypeLib.Product = oSelection.Item2(i).Value
'            'Ver si por error se ha seleccionado un "Product" tal que: Product.Products > 0
'            If oProduct.Products.Count > 0 Then
'                MsgBox("Ha seleccionado un product, vuelva a seleccionar solo Parts ")
'                Exit Sub
'            End If
'            ' Previene la selecciona erronea de partes del tipo: "NCU"
'            If Left(oProduct.PartNumber, 3) = "NCU" Then
'                MsgBox("Ha seleccionado una NCU, vuelva a seleccionar Parts de tipo: MADE")
'                Exit Sub
'            End If
'            Dim refProduct As ProductStructureTypeLib.Product = oProduct.ReferenceProduct
'            If refProduct.UserRefProperties.Count = 0 Then 'Si no tiene propiedades
'                ' AddUserProperty(refProduct)  ' Crea la propiedad
'                AddMaterial(refProduct)  ' Asigna el material a la propiedad
'                GoTo SiguienteProduct
'            Else ' Si tiene propiedades, buscar si tiene una que se llama "Material"
'                For j = 1 To refProduct.UserRefProperties.Count
'                    If refProduct.UserRefProperties.Item(j).Name = "Material" Then
'                        AddMaterial(refProduct)  ' Asigna el material a la propiedad
'                        GoTo SiguienteProduct
'                    Else
'                        Continue For
'                    End If
'                Next
'                ' AddUserProperty(refProduct)  ' Crea la propiedad
'                AddMaterial(refProduct)  ' Asigna el material a la propiedad
'                GoTo SiguienteProduct
'            End If
'SiguienteProduct:
'        Next
'    End Sub




'Sub AddUserParameters(oDic As Dictionary(Of String, ProductStructureTypeLib.Product))
'    'Con el diccionario que ser le pasa como argumento, se referencia la colección de valores (los Products)
'    Dim valColl As Collections.Generic.Dictionary(Of String, ProductStructureTypeLib.Product).ValueCollection = oDic.Values
'    For Each p As ProductStructureTypeLib.Product In valColl
'        If p.Source = 1 Then
'            Dim Parametros As KnowledgewareTypeLib.Parameters = p.ReferenceProduct.UserRefProperties
'            Parametros.CreateString("Raw Material", "")
'        End If
'        ' Parametros.CreateString("Tratamiento Térmico", "")
'        ' Parametros.CreateString("Pintura", "")
'    Next
'End Sub