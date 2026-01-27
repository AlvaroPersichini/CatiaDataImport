Module TextReplace
    'Module TextReplace
    '    ''' <summary>
    '    ''' Reemplaza un string en el arbol - Solo reemplaza la primer ocurrencia
    '    ''' El procedimiento por default es "case-sensitive"
    '    ''' </summary>
    '    ''' <param name="objCurrentProduct">Item Part or Product to be renamed</param>
    '    ''' <param name="strToSearch">Word to be serched</param>
    '    ''' <param name="strReplacement">Replacement string to apply</param>
    '    Sub TextReplace(ByRef objCurrentProduct As ProductStructureTypeLib.Product, strToSearch As String, strReplacement As String)

    '        Dim i As Integer
    '        Dim strOldPartNumber As String
    '        Dim newText As String
    '        Static Dim objDictionary As Scripting.Dictionary = CreateObject("Scripting.Dictionary")
    '        Static Dim objDictionaryAux As Scripting.Dictionary = CreateObject("Scripting.Dictionary")
    '        Dim match As Text.RegularExpressions.Match
    '        Static Dim intReplaced As Integer
    '        Static Dim intNonReplaced As Integer

    '        objCurrentProduct = objCurrentProduct.ReferenceProduct
    '        For i = 1 To objCurrentProduct.Products.Count
    '            strOldPartNumber = objCurrentProduct.Products.Item(i).PartNumber
    '            match = Text.RegularExpressions.Regex.Match(strOldPartNumber, strToSearch, Text.RegularExpressions.RegexOptions.None)
    '            If match.Success Then
    '                newText = strOldPartNumber.Remove(match.Index, match.Length)
    '                objCurrentProduct.Products.Item(i).PartNumber = newText.Insert(match.Index, strReplacement)
    '                If strOldPartNumber = objCurrentProduct.Products.Item(i).PartNumber Then
    '                    If objDictionaryAux.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
    '                        GoTo Continuar
    '                    Else
    '                        objDictionaryAux.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
    '                        intNonReplaced += 1
    '                    End If
    '                Else
    '                    intReplaced += 1
    '                End If
    '            End If
    'Continuar:
    '        Next
    '        For i = 1 To objCurrentProduct.Products.Count
    '            If objDictionary.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
    '                GoTo Finish
    '            Else
    '                objDictionary.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
    '                TextReplace(objCurrentProduct.Products.Item(i), strToSearch, strReplacement)
    '            End If
    'Finish:
    '        Next
    '    End Sub

End Module
