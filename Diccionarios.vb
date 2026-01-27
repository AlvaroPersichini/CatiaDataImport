Option Explicit On


Module Diccionarios





    ' Diccionario tipo 1 (String, String)
    ' Key: oDocument.Name (con extensión)
    ' Value: Product.PartNumber
    Public Function DiccT1_DocName_PN(objRootProduct As ProductStructureTypeLib.Product) As Specialized.StringDictionary
        Static Dim strDicc As New Specialized.StringDictionary()
        Dim objProductDocument As ProductStructureTypeLib.ProductDocument = objRootProduct.ReferenceProduct.Parent
        Dim oDocument As INFITF.Document
        ' Esta primera instruccion If es para agregar al diccionario el documento raiz.
        If strDicc.ContainsKey(objProductDocument.Name) = False Then
            strDicc.Add(objProductDocument.Name, objRootProduct.PartNumber)
        End If
        ' Agrega al diccionario todos los componentes que haya aguas abajo
        For Each Product As ProductStructureTypeLib.Product In objRootProduct.Products
            oDocument = Product.ReferenceProduct.Parent
            If strDicc.ContainsKey(oDocument.Name) Then
                GoTo Finish
            Else
                strDicc.Add(oDocument.Name, Product.PartNumber)
                If Product.Products.Count > 0 Then
                    DiccT1_DocName_PN(Product)
                End If
            End If
Finish:
        Next
        Return strDicc
    End Function








    ' Diccionario tipo 1 (String, ProductStructureTypeLib.Product)
    ' Key: oProduct.PartNumber
    ' Value: Product
    Public Function DiccT1_PN_oProduct(oProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, ProductStructureTypeLib.Product)
        Static Dim oDic As New Dictionary(Of String, ProductStructureTypeLib.Product)

        '¡¡¡¡ Esta primera instruccion If es para agregar al diccionario el documento raiz. !!!!
        'Cuidado acá porque al usar el procedimiento "OpenInNewWindow" cuando se está computando el product raiz,
        'al momento que saca la captura de pantalla, el procedimiento lo cierra, dejando así a Catia sin ningun documento activo, o peor aún
        'quizá queda otro documento activo.
        'Lo que hay que hacer al momento de computar el producto raiz es que saque la captura de pantalla,
        'pero que no cierre el raiz, y que continue con el siguiente

        'If oDic.ContainsKey(oProduct.PartNumber) = False Then
        '    oDic.Add(oProduct.PartNumber, oProduct)
        'End If

        For Each Product As ProductStructureTypeLib.Product In oProduct.Products
            If oDic.ContainsKey(Product.PartNumber) Then
                GoTo Finish
            Else
                oDic.Add(Product.PartNumber, Product)
                If Product.Products.Count > 0 Then
                    DiccT1_PN_oProduct(Product)
                End If
            End If
Finish:
        Next
        Return oDic
    End Function




    ' Diccionario "tipo 2".
    ' Los diccionarios de tipo 2 devuelven una pareja de valores,
    ' tal que si un product ya se ha computado, si vuelve a aparecer bajo diferente padre, sí se computa.
    ' Si el product "A" aparece varias veces bajo diferentes padres se tiene en cuenta.
    Public Function DiccT2_PN_PwrProduct(oRootProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, PowerProduct)

        Dim strRootProdPN As String = oRootProduct.PartNumber
        Static Dim oDictionary As New Dictionary(Of String, PowerProduct)
        Static Dim lvl As Integer = 1
        Static Dim intCont As Integer = 1

        ' Para incluir el rootProduct lo resolví de esta manera, utilizando un contador.
        If intCont = 1 Then
            Dim PPRoot As New PowerProduct With {
                .oProduct = oRootProduct
               }
            oDictionary.Add(oRootProduct.PartNumber, PPRoot)
        End If
        intCont += 1


        For Each Product As ProductStructureTypeLib.Product In oRootProduct.Products
            Dim PP As New PowerProduct With {
                .oProduct = Product
            }
            Dim strPartNumber As String = Product.PartNumber ' PartNumber del product
            Dim strRootProdName As String = oRootProduct.ReferenceProduct.PartNumber  ' PartNumber del product padre
            If oDictionary.ContainsKey(strRootProdName & "_" & strPartNumber) Then
                oDictionary.Item(strRootProdName & "_" & strPartNumber).intQuantity = oDictionary.Item(strRootProdName & "_" & strPartNumber).intQuantity + 1
                GoTo Finish
            Else
                PP.intLevel = lvl
                PP.intQuantity = 1
                oDictionary.Add(strRootProdName & "_" & strPartNumber, PP) 'Key: "strRootProdName & "_" & strPartNumber" - item: "PP"
                If PP.oProduct.Products.Count > 0 Then
                    lvl += 1
                    DiccT2_PN_PwrProduct(PP.oProduct)
                    lvl -= 1
                End If
            End If
Finish:
        Next
        Return oDictionary
    End Function



    ' Diccionario del tipo 3
    ' Sepuede armar "PartNumber" vs. "Cantidad"
    ' Key: PowerProduct.oProduct.PartNumber
    ' Value: PowerProduct
    Public Function DiccT3_PN_PwrProduct(oRootProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, PowerProduct)
        Static Dim oDictionary As New Dictionary(Of String, PowerProduct)
        Static Dim intCont As Integer = 1

        ' Para incluir el rootProduct lo resolví de esta manera, utilizando un contador.
        'If intCont = 1 Then
        '    Dim PPRoot As New PowerProduct With {
        '        .oProduct = oRootProduct
        '       }
        '    oDictionary.Add(oRootProduct.PartNumber, PPRoot)
        'End If
        'intCont += 1

        For Each Product As ProductStructureTypeLib.Product In oRootProduct.Products
            Dim PP As New PowerProduct With {
                .oProduct = Product
            }
            If oDictionary.ContainsKey(PP.oProduct.PartNumber) Then
                oDictionary.Item(PP.oProduct.PartNumber).intQuantity = oDictionary.Item(PP.oProduct.PartNumber).intQuantity + 1
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_PN_PwrProduct(PP.oProduct)
                End If
                GoTo Finish
            Else
                PP.intQuantity = 1
                oDictionary.Add(PP.oProduct.PartNumber, PP) 'Key: "strPartNumber" - Value: PP (PowerProduct)
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_PN_PwrProduct(PP.oProduct)
                End If
            End If
Finish:
        Next
        Return oDictionary
    End Function



    ' ************************************************************************************************************
    ' Revision 02 del diccionario de tipo 3:
    ' Actualizaciones:
    '       1) computa el product ráíz y le asigna las propiedades.
    '       2) cambios en la forma en que almacena el Source
    '       3) Utiliza la clase "PwrProduct" que tiene varias propiedades para ser usadas en otros procedimientos.
    ' ************************************************************************************************************
    Public Function DiccT3_Rev2(oRootProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, PwrProduct)
        Static Dim oDictionary As New Dictionary(Of String, PwrProduct)
        Static Dim intCont As Integer = 1
        Dim strProductType As String

        ' Para incluir el rootProduct lo resolví de esta manera, utilizando un contador.
        If intCont = 1 Then
            Dim PPRoot As New PwrProduct
            With PPRoot
                .Product = oRootProduct
                .Quantity = 1
                .ProductType = Replace(TypeName(oRootProduct.ReferenceProduct.Parent), "ProductDocument", "C")
                .Source = Replace([Enum].GetName(GetType(ProductStructureTypeLib.CatProductSource), oRootProduct.Source), "catProduct", "")
            End With
            oDictionary.Add(oRootProduct.PartNumber, PPRoot)
        End If
        intCont += 1
        For Each Product As ProductStructureTypeLib.Product In oRootProduct.Products
            Dim PP As New PwrProduct
            With PP
                strProductType = Replace(TypeName(Product.ReferenceProduct.Parent), "ProductDocument", "C")
                strProductType = Replace(strProductType, "PartDocument", "P")
                .Product = Product
                .ProductType = strProductType
                .Source = Replace([Enum].GetName(GetType(ProductStructureTypeLib.CatProductSource), Product.Source), "catProduct", "")
            End With
            If oDictionary.ContainsKey(PP.Product.PartNumber) Then
                oDictionary.Item(PP.Product.PartNumber).Quantity = oDictionary.Item(PP.Product.PartNumber).Quantity + 1
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_Rev2(PP.Product)
                End If
                GoTo Finish
            Else
                PP.Quantity = 1
                oDictionary.Add(PP.Product.PartNumber, PP)
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_Rev2(PP.Product)
                End If
            End If
Finish:
        Next
        Return oDictionary
    End Function

    Public Function DicNombres(oWillBeCopied As Object, oDic1 As Specialized.StringDictionary) As Dictionary(Of String, NameContainer)
        Dim strDocumentExtension As String
        Dim intLastSlashPosition As Integer
        Dim strDocNameWithExt As String
        Dim strNewNameWithExt As String
        Dim strDocNameWithOutExt As String
        Dim strNewNameWithOutExt As String
        Dim intLastPointPosition As Integer
        Dim intLastPointDocName As Integer

        Dim DiccionarioNombres As New Dictionary(Of String, NameContainer)

        For Each strDoc As String In oWillBeCopied
            Dim contenedorNombres As New NameContainer
            intLastPointPosition = strDoc.LastIndexOf(".")
            intLastSlashPosition = strDoc.LastIndexOf("\")
            strDocumentExtension = strDoc.Substring(intLastPointPosition + 1)
            strDocNameWithExt = strDoc.Substring(intLastSlashPosition + 1)
            intLastPointDocName = strDocNameWithExt.LastIndexOf(".")
            strDocNameWithOutExt = Left(strDocNameWithExt, intLastPointDocName)
            strNewNameWithOutExt = oDic1.Item(strDocNameWithExt)
            strNewNameWithExt = strNewNameWithOutExt & "." & strDocumentExtension

            With contenedorNombres
                .sDocNameWithExt = strDocNameWithExt
                .sDocNameWithOutExt = strDocNameWithOutExt
                .sNewNameWithExt = strNewNameWithExt
                .sNewNameWithOutExt = strNewNameWithOutExt
                .sDocumentExtension = strDocumentExtension
            End With

            DiccionarioNombres.Add(strDoc, contenedorNombres)

        Next
        Return DiccionarioNombres
    End Function

End Module










'Lista del tipo 1
'Los diccionarios o listas de tipo 1, devuelven una colección de "Product",
'en donde si un Product ya se ha computado, no se vuelve a computar aparezca o no bajo otros padres.
'Es decir que no computa las instancias. Solo contempla la primer instancia.
'Si luego aparece otra instancia del mismo partNumber no la tiene en cuenta. 

'    Public Function ListTipo1(objRootProduct As ProductStructureTypeLib.Product) As List(Of ProductStructureTypeLib.Product)

'        Static Dim strDicc As New Specialized.StringDictionary()
'        Static Dim ListOfProducts As New List(Of ProductStructureTypeLib.Product)
'        Dim strDocumentName As String
'        Dim objProductDocument As ProductStructureTypeLib.ProductDocument
'        Dim objPartDocument As MECMOD.PartDocument
'        Dim Product As ProductStructureTypeLib.Product

'        'Esta primera instruccion If es para agregar al diccionario el documento raiz.

'        'If strDicc.ContainsKey(objProductDocument.Name) = False Then
'        'strDicc.Add(objProductDocument.Name, objRootProduct.PartNumber)
'        'End If

'        'Agrega al diccionario todos los componentes que haya aguas abajo
'        For Each Product In objRootProduct.Products
'            Dim v As String = TypeName(Product.ReferenceProduct.Parent)
'            Select Case v
'                Case "ProductDocument"
'                    objProductDocument = Product.ReferenceProduct.Parent
'                    strDocumentName = objProductDocument.Name
'                Case "PartDocument"
'                    objPartDocument = Product.ReferenceProduct.Parent
'                    strDocumentName = objPartDocument.Name
'                Case Else
'                    MsgBox("El diccionario solo trabaja con Parts o Products")
'            End Select

'            If ListOfProducts.Contains(Product) Then
'                ' If strDicc.ContainsKey(strDocumentName) Then
'                GoTo Finish
'            Else
'                ListOfProducts.Add(Product)
'                '  strDicc.Add(strDocumentName, Product.PartNumber)
'                If Product.Products.Count > 0 Then
'                    ListTipo1(Product)
'                End If
'            End If
'Finish:
'        Next
'        Return ListOfProducts
'    End Function



'  Dim objProductDocument As ProductStructureTypeLib.ProductDocument
'  Dim objPartDocument As MECMOD.PartDocument

'Esta primera instruccion If es para agregar al diccionario el documento raiz.
'If oDic.ContainsKey(objProductDocument.Name) = False Then
'    oDic.Add(objProductDocument.Name, objRootProduct.PartNumber)
'End If

'Agrega al diccionario todos los componentes que haya aguas abajo






'Este diccionario es de tipo 1, este deberia ser el oficial. EL otro que hay 
'de tipo 1, trabaja con argumentos que no es bueno. Mejor es siguir desarrollando éste y migrar
'los procedimientos para usar éste diccionario.
'Utiliza un "oProduct" como "value", no un "PowerProduct".
'Hay que migrarlo para que use powerproduct.
'    Public Function DicTipo1PN(oProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, ProductStructureTypeLib.Product)
'        Static Dim oDic As New Dictionary(Of String, ProductStructureTypeLib.Product)
'        For Each Product As ProductStructureTypeLib.Product In oProduct.Products
'            If oDic.ContainsKey(Product.PartNumber) Then
'                GoTo Finish
'            Else
'                oDic.Add(Product.PartNumber, Product)
'                If Product.Products.Count > 0 Then
'                    DicTipo1PN(Product)
'                End If
'            End If
'Finish:
'        Next
'        Return oDic
'    End Function