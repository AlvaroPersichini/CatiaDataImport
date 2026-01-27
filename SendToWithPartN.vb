Option Explicit On

Module SendToWithPartN

    ' IMPORTANTE:
    ' "oWillBeCopied" es una lista que puede o no contener todos los archivos que se van a copiar en funcion de su tamaño (lenght)
    ' Si "oWillBeCopied" se declara de un tamaño menor a la cantidad de archivos que se van a copiar,
    ' no da error, pero el renombrado solo se lleva a cabo con la cantidad que esa lista contiene. Si "oWillBeCopied" es mas grande, da error.

    ' Cauando se utilizan referencias externas (por ejemplo utilizando el módulo "Structure Design" hay archivos de extension .CATMaterial o también
    ' Parts que estan en memoria pero no cargados (unloaded), esto no es tenido en cuenta por el diccionario de PartNumbers.
    ' Es por esto que, si el product raíz "arrastra archivos que son usados como referencia externas como es el caso de los croquiz de los perfiles,
    ' el oDic no los computa y el oDic.Count va a dar diferente a la cantidad "oWillBeCopied.lenght")

    ' También, en el procedimiento de verificar si los archivos ya existen en el directorio destino,
    ' no se tienen en cuenta las referencias externas (.CATMaterial, croquiza de CATParts, etc,)
    ' entonces, al querer pisar nuevamente todos los arhivos, el número de "n archivos ya existen" puede diferir de lo que contiene "oWillBeCopied"

    ' (*) Me fijo si el dicionario ya contiene un nombre de los nuevos,
    ' porque lo que estaría pasando es que quiera asignar un nombre nuevo que es identico a uno que ya existe
    ' Es decir quiere dar el nombre "A" a una pieza, pero ese nombre "A" ya es el nombre de otro archivo de mas abajo.
    ' Si ese es el caso, entonces no puedo renombrar en este momento.
    ' Lo que hace es, guarda ese par en el diccionario "oDicNoRenamed" y lo procesa luego cuando la pieza de mas abajo, ya no es mas "A"
    ' Utilizar un Segundo cilco de renombrado: NO FUNCIONA SIEMPRE!

    ' Conclusión:
    ' Es preferible utilizar el servicio "SendTo" sin referencias externas, es decir, los product que forma el Structure Design,
    ' cambiarlos a "allCatPart" o eliminar las referencias externas, para que solo queden archivos del tipo "CATProdcut" y "CATPart".


    ''' <summary>
    ''' Realiza un "SendTo" con renombrado de archivos con los ParNumbers del árbol
    ''' </summary>
    ''' <param name="oProductDocument">Product Raíz</param>
    ''' <param name="strDir">Directorio Destino</param>
    Public Sub SendTOWPN(oProductDocument As ProductStructureTypeLib.ProductDocument, oDic1 As Specialized.StringDictionary, strDir As String)



        Dim intFilesInExistance As Integer = 0 'Cantidad de archivos que ya existen en el Dir destino
        Dim oDicNoRenamed As New Specialized.StringDictionary ' Contiene los pares que no se pudieron renombrar en la primera vuelta
        Dim oDicRenamed As New Specialized.StringDictionary   ' Contiene los pares que si se pueden renombrar en la primera vuelt
        Dim objAppCATIA As INFITF.Application = oProductDocument.Application
        ' Dim oDic1 As Specialized.StringDictionary = DiccT1_DocName_PN(oProductDocument.Product)
        Dim SendTo As INFITF.SendToService
        Dim oWillBeCopied(oDic1.Count - 1) As Object
        Dim ContenedorNombres As Dictionary(Of String, NameContainer)




        'Arma el objeto SendTo, Referencia lista y Initial File
        SendTo = objAppCATIA.CreateSendTo()
        SendTo.SetInitialFile(oProductDocument.FullName)
        SendTo.GetListOfToBeCopiedFiles(oWillBeCopied)
        SendTo.SetDirectoryFile(strDir)


        'Arma el contenedor de nombres y Realiza el renombrado, pero no ejecuta
        'El contenedor de nombres se arma con una funcion especifica para el procedimiento "SendTo"
        ContenedorNombres = Diccionarios.DicNombres(oWillBeCopied, oDic1)
        Renombrado(oDicRenamed, oDicNoRenamed, SendTo, ContenedorNombres, oDic1)


        'Revisar si ya existen documentos con el mismo nombre en el directorio destino
        intFilesInExistance = Procedimientos.CountFilesInExistance(strDir, ContenedorNombres)


        'Finalmente executa "SendTo.Run() e informa
        Finalizacion(intFilesInExistance, SendTo, oWillBeCopied, strDir)



        If Err.Number = 0 Then
            Exit Sub
        End If



    End Sub

    Sub Renombrado(oDicRenamed As Specialized.StringDictionary,
                   oDicNoRenamed As Specialized.StringDictionary,
                   SendTo As INFITF.SendToService,
                   ContenedorNombres As Dictionary(Of String, NameContainer),
                   oDic1 As Specialized.StringDictionary)


        ' Arma los pares a ser renombrados en uno o dos pases
        ' Dim ContenedorNombres = Diccionarios.DicNombres(oWillBeCopied, oDic1)
        For Each kvp As KeyValuePair(Of String, NameContainer) In ContenedorNombres
            If kvp.Value.sNewNameWithOutExt <> kvp.Value.sDocNameWithOutExt Then
                If oDic1.ContainsKey(kvp.Value.sNewNameWithExt) Then '(*) 
                    oDicNoRenamed.Add(kvp.Value.sDocNameWithExt, kvp.Value.sNewNameWithOutExt)  'Arma el diccionario de los que no pudo renombrar en primera vuelta
                Else
                    oDicRenamed.Add(kvp.Value.sDocNameWithExt, kvp.Value.sNewNameWithOutExt)  'Arma el diccionario de lo que se pueden renombrar en primera vuelta
                End If
            End If
        Next


        'Primera ciclo de renombrado
        'El diccionario oDicRenamed almacena los pares que pueden ser renombrados en primera vuelta
        For Each de As DictionaryEntry In oDicRenamed
            SendTo.SetRenameFile(de.Key, de.Value)
        Next


        'Segundo ciclo de renombrado
        'Si quedaron pares sin renombrar se renombran con este ciclo
        'El diccionario oDicNoRenamed almacena los pares que no fueron renombrados en la primera vuelta
        'hay casos en que utilizar un segundo renombrado no funciona - NO ESTA RESUELTO 
        'If oDicNoRenamed.Count <> 0 Then
        '    For Each de As DictionaryEntry In oDicNoRenamed
        '        '  SendTo.SetRenameFile(de.Key, de.Value)
        '    Next
        'End If

    End Sub

    Sub Finalizacion(intFilesInExistance As Integer, SendTo As INFITF.SendToService, oWillBeCopied As Object, strDir As String)

        Dim intResponse As Integer
        Dim strMsg As String

        If intFilesInExistance > 0 Then
            strMsg = intFilesInExistance & " Archivo(s) ya existe(n) en el directorio de destino" & vbCrLf & vbCrLf
            strMsg = strMsg & " ¿Desea reemplazar el/los archivo(s)?" & vbCrLf & vbCrLf
            strMsg &= " Click OK to proceed or Cancel"
            intResponse = MsgBox(strMsg, 65, "Replace Files")
            If intResponse = 1 Then
                SendTo.Run()
                'Informa lo que se ha copiado en un mensaje y exit
                MsgBox("Done" & vbCrLf _
                       & oWillBeCopied.Length & " Files has been exported to " & strDir, 64, "Succeed")
                Exit Sub
            Else
                Exit Sub
            End If
        Else 'Si no hay archivos existentes procede con la ejecución
            SendTo.Run()
        End If
        'Informa lo que se ha copiado en un mensaje y exit
        MsgBox(oWillBeCopied.Length & " Files has been exported to " & strDir, 64, "Succeed")
        Exit Sub  'Colocando esta línea se evita que el procedimiento continue con CheckError
    End Sub














End Module