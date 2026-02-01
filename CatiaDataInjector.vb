Option Explicit On
Option Strict On

Public Class CatiaDataInjector

    ''' <summary>
    ''' Inyecta los datos desde la lista de ExcelData hacia el árbol de CATIA.
    ''' </summary>
    Public Sub InjectData(oRootProduct As ProductStructureTypeLib.Product,
                          dataToInject As Dictionary(Of String, ExcelData))

        ' Obtenemos el documento raíz para la lógica de componentes
        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)
        Dim processedFiles As New HashSet(Of String)

        ' Procesar el Root
        ApplyToDocument(oRootProduct, rootDoc, dataToInject, processedFiles)

        ' Iniciar recorrido recursivo para los hijos
        ProcesarHijosRecursivo(oRootProduct, dataToInject, processedFiles, rootDoc)

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Injection complete.")
    End Sub

    Private Sub ProcesarHijosRecursivo(oParent As ProductStructureTypeLib.Product,
                                      dataToInject As Dictionary(Of String, ExcelData),
                                      ByRef processedFiles As HashSet(Of String),
                                      oParentDoc As INFITF.Document)

        For Each oChild As ProductStructureTypeLib.Product In oParent.Products
            Dim oChildDoc As INFITF.Document = CType(oChild.ReferenceProduct.Parent, INFITF.Document)

            ' Si es un COMPONENT (mismo archivo que el padre)
            If oChildDoc.FullName = oParentDoc.FullName Then
                ProcesarHijosRecursivo(oChild, dataToInject, processedFiles, oParentDoc)
            Else
                ' Si es un ARCHIVO REAL (Part o Product) que no procesamos todavía
                If Not processedFiles.Contains(oChildDoc.FullName) Then
                    ApplyToDocument(oChild, oChildDoc, dataToInject, processedFiles)
                End If

                ' Si tiene estructura interna, seguimos bajando
                If TypeOf oChildDoc Is ProductStructureTypeLib.ProductDocument Then
                    ProcesarHijosRecursivo(oChild, dataToInject, processedFiles, oChildDoc)
                End If
            End If
        Next
    End Sub

    Private Sub ApplyToDocument(oInstancia As ProductStructureTypeLib.Product,
                            oDoc As INFITF.Document,
                            data As Dictionary(Of String, ExcelData),
                            ByRef processed As HashSet(Of String))

        Dim currentPN As String = oInstancia.PartNumber

        If data.ContainsKey(currentPN) Then
            Dim info As ExcelData = data(currentPN)
            Dim oRefProd As ProductStructureTypeLib.Product = oInstancia.ReferenceProduct

            ' 1. PartNumber
            If Not String.IsNullOrEmpty(info.NewPartNumber) AndAlso oRefProd.PartNumber <> info.NewPartNumber Then
                oRefProd.PartNumber = info.NewPartNumber
            End If

            ' 2. Definition
            If oRefProd.Definition <> info.Definition Then
                oRefProd.Definition = info.Definition
            End If

            ' 3. Nomenclature
            If oRefProd.Nomenclature <> info.Nomenclature Then
                oRefProd.Nomenclature = info.Nomenclature
            End If

            ' 4. DescriptionRef
            If oRefProd.DescriptionRef <> info.DescriptionRef Then
                oRefProd.DescriptionRef = info.DescriptionRef
            End If

            ' 5. Source
            ' Convertimos a Integer o el tipo base para comparar antes de asignar
            Dim newSource As ProductStructureTypeLib.CatProductSource = CType(info.Source, ProductStructureTypeLib.CatProductSource)
            If oRefProd.Source <> newSource Then
                oRefProd.Source = newSource
            End If

        End If

        ' Marcamos como procesado
        If Not String.IsNullOrEmpty(oDoc.FullName) Then
            processed.Add(oDoc.FullName)
        End If
    End Sub

End Class
