'Option Explicit On

'Module SendToWithPartN

'    ' IMPORTANTE:
'    ' "oWillBeCopied" es una lista que puede o no contener todos los archivos que se van a copiar en funcion de su tamaño (lenght)
'    ' Si "oWillBeCopied" se declara de un tamaño menor a la cantidad de archivos que se van a copiar,
'    ' no da error, pero el renombrado solo se lleva a cabo con la cantidad que esa lista contiene. Si "oWillBeCopied" es mas grande, da error.

'    ' Cauando se utilizan referencias externas (por ejemplo utilizando el módulo "Structure Design" hay archivos de extension .CATMaterial o también
'    ' Parts que estan en memoria pero no cargados (unloaded), esto no es tenido en cuenta por el diccionario de PartNumbers.
'    ' Es por esto que, si el product raíz "arrastra archivos que son usados como referencia externas como es el caso de los croquiz de los perfiles,
'    ' el oDic no los computa y el oDic.Count va a dar diferente a la cantidad "oWillBeCopied.lenght")

'    ' También, en el procedimiento de verificar si los archivos ya existen en el directorio destino,
'    ' no se tienen en cuenta las referencias externas (.CATMaterial, croquiza de CATParts, etc,)
'    ' entonces, al querer pisar nuevamente todos los arhivos, el número de "n archivos ya existen" puede diferir de lo que contiene "oWillBeCopied"

'    ' (*) Me fijo si el dicionario ya contiene un nombre de los nuevos,
'    ' porque lo que estaría pasando es que quiera asignar un nombre nuevo que es identico a uno que ya existe
'    ' Es decir quiere dar el nombre "A" a una pieza, pero ese nombre "A" ya es el nombre de otro archivo de mas abajo.
'    ' Si ese es el caso, entonces no puedo renombrar en este momento.
'    ' Lo que hace es, guarda ese par en el diccionario "oDicNoRenamed" y lo procesa luego cuando la pieza de mas abajo, ya no es mas "A"
'    ' Utilizar un Segundo cilco de renombrado: NO FUNCIONA SIEMPRE!

'    ' Conclusión:
'    ' Es preferible utilizar el servicio "SendTo" sin referencias externas, es decir, los product que forma el Structure Design,
'    ' cambiarlos a "allCatPart" o eliminar las referencias externas, para que solo queden archivos del tipo "CATProdcut" y "CATPart".

'    ' Al realizar el SendTo el comando no tiene en cuanta si el product tiene propiedades como ser "Description", "Source", "Definition", etc.
'    ' Entonces al hacer el SendTo, esas propiedades no se copian al nuevo archivo. Hay que realizar un proceso aparte para copiar esas propiedades.


'    ' Una advertencia sobre el tamaño del Array
'    ' Como estás manteniendo la línea: Dim oWillBeCopied(oDic1.Count - 1) As Object
'    ' Si el producto raíz tiene referencias externas (como un .CATMaterial que no está en tu diccionario), el SendTo querrá meterlo en el array.
'    ' Como tu array tiene el tamaño exacto de tu diccionario, y el raíz ya ocupa un lugar,
'    ' si hay elementos "extra" que CATIA detecta, la línea GetListOfToBeCopiedFiles podría darte el Error de rango esperado que vimos antes.
'    ' El único riesgo técnico sigue siendo que CATIA encuentre más archivos de los que tu diccionario tiene contabilizados




'    Public Sub SendTOWPN(oProductDocument As ProductStructureTypeLib.ProductDocument, oDic1 As Specialized.StringDictionary, strDir As String)

'        ' IMPORTANTE: Si el documento no está guardado, SendTo no verá los links
'        ' oProductDocument.Save() ' Descomenta esta línea si oWillBeCopied sigue dando 1

'        Dim objAppCATIA As INFITF.Application = oProductDocument.Application
'        Dim SendTo As INFITF.SendToService = objAppCATIA.CreateSendTo()

'        ' 1. Seteamos el archivo raíz
'        SendTo.SetInitialFile(oProductDocument.FullName)

'        ' 2. Dimensionamos según el diccionario (133)
'        ' Usamos una variable intermedia para ver qué detecta CATIA realmente
'        Dim oWillBeCopied(oDic1.Count - 1) As Object
'        SendTo.GetListOfToBeCopiedFiles(oWillBeCopied)
'        SendTo.SetDirectoryFile(strDir)

'        ' --- CARTEL DE CONTROL ---
'        MsgBox("ESTADO DE CARGA:" & vbCrLf &
'               "oWillBeCopied.Length: " & oWillBeCopied.Length & vbCrLf &
'               "oDic1.Count: " & oDic1.Count)



'        ' --- CICLO DE RENOMBRADO ---
'        Dim i As Integer
'        For i = 0 To UBound(oWillBeCopied)

'            If oWillBeCopied(i) Is Nothing Then Continue For

'            ' Extracción manual para evitar caracteres ilegales
'            Dim strFullPath As String = oWillBeCopied(i).ToString()
'            Dim lastSlash As Integer = strFullPath.LastIndexOf("\")
'            Dim strFileName As String = If(lastSlash > -1, strFullPath.Substring(lastSlash + 1), strFullPath)

'            ' Si el nombre del archivo está en nuestro diccionario
'            If oDic1.ContainsKey(strFileName) Then
'                Dim strNewName As String = oDic1(strFileName)

'                ' --- VALIDACIÓN PARA EVITAR FAIL POR DUPLICADOS ---
'                ' 1. Verificamos que el nombre nuevo no sea igual al actual
'                Dim dotIdx As Integer = strFileName.LastIndexOf(".")
'                Dim currentNameNoExt As String = If(dotIdx > 0, strFileName.Substring(0, dotIdx), strFileName)

'                If strNewName <> currentNameNoExt Then

'                    MsgBox(strNewName & "--" & currentNameNoExt)

'                    ' 2. Verificamos si el nombre "NCU-2517" ya existe en la lista de CATIA
'                    ' Para esto, comparamos el strNewName contra todos los archivos que CATIA va a copiar
'                    Dim yaExisteEnConjunto As Boolean = False
'                    For Each objPath In oWillBeCopied
'                        If objPath IsNot Nothing AndAlso objPath.ToString().Contains(strNewName & ".") Then
'                            yaExisteEnConjunto = True
'                            Exit For
'                        End If
'                    Next

'                    ' Solo renombramos si el nombre no existe todavía en el conjunto
'                    If Not yaExisteEnConjunto Then
'                        SendTo.SetRenameFile(strFileName, strNewName)
'                    Else
'                        ' Si ya existe, imprimimos en consola para saber cuál saltamos
'                        System.Console.WriteLine("Saltado por duplicado: " & strFileName & " -> " & strNewName)
'                    End If
'                End If
'            End If
'        Next

'        ' 4. Ejecución
'        SendTo.Run()

'        MsgBox("SendTo finalizado con éxito.")

'    End Sub



'End Module

Option Explicit On

Module SendToWithPartN

    Public Sub SendTOWPN(oProductDocument As ProductStructureTypeLib.ProductDocument, oDic1 As Specialized.StringDictionary, strDir As String)

        Dim objAppCATIA As INFITF.Application = oProductDocument.Application
        Dim SendTo As INFITF.SendToService = objAppCATIA.CreateSendTo()

        ' 1. Seteamos el archivo raíz
        SendTo.SetInitialFile(oProductDocument.FullName)

        ' 2. Dimensionamos según el diccionario (133)
        Dim oWillBeCopied(oDic1.Count - 1) As Object
        SendTo.GetListOfToBeCopiedFiles(oWillBeCopied)
        SendTo.SetDirectoryFile(strDir)

        ' --- CARTEL DE CONTROL ---
        MsgBox("ESTADO DE CARGA:" & vbCrLf &
               "oWillBeCopied.Length: " & oWillBeCopied.Length & vbCrLf &
               "oDic1.Count: " & oDic1.Count)

        ' --- DICCIONARIO PARA SEGUNDA PASADA ---
        Dim oDicPendientes As New Specialized.StringDictionary()

        ' --- CICLO DE RENOMBRADO (PRIMERA PASADA) ---
        Dim i As Integer
        For i = 0 To UBound(oWillBeCopied)

            If oWillBeCopied(i) Is Nothing Then Continue For

            Dim strFullPath As String = oWillBeCopied(i).ToString()
            Dim lastSlash As Integer = strFullPath.LastIndexOf("\")
            Dim strFileName As String = If(lastSlash > -1, strFullPath.Substring(lastSlash + 1), strFullPath)

            If oDic1.ContainsKey(strFileName) Then
                Dim strNewName As String = oDic1(strFileName)

                Dim dotIdx As Integer = strFileName.LastIndexOf(".")
                Dim currentNameNoExt As String = If(dotIdx > 0, strFileName.Substring(0, dotIdx), strFileName)

                If strNewName <> currentNameNoExt Then
                    ' Verificamos si el nombre ya existe en el conjunto de CATIA
                    Dim yaExisteEnConjunto As Boolean = False
                    For Each objPath In oWillBeCopied
                        If objPath IsNot Nothing AndAlso objPath.ToString().Contains(strNewName & ".") Then
                            yaExisteEnConjunto = True
                            Exit For
                        End If
                    Next

                    ' Si no existe, renombramos ahora. Si existe, lo mandamos a la lista de pendientes.
                    If Not yaExisteEnConjunto Then
                        SendTo.SetRenameFile(strFileName, strNewName)
                    Else
                        If Not oDicPendientes.ContainsKey(strFileName) Then
                            oDicPendientes.Add(strFileName, strNewName)
                        End If
                    End If
                End If
            End If
        Next

        ' --- SEGUNDO CICLO DE RENOMBRADO (PARA LOS PENDIENTES) ---
        If oDicPendientes.Count > 0 Then
            ' COPIA DE LLAVES: Creamos una lista estática para evitar el error de "Collection was modified"
            Dim llavesPendientes(oDicPendientes.Count - 1) As String
            oDicPendientes.Keys.CopyTo(llavesPendientes, 0)

            For Each strFileKey As String In llavesPendientes
                ' Verificamos que no sea nulo (por seguridad al copiar el array)
                If strFileKey Is Nothing Then Continue For

                Try
                    SendTo.SetRenameFile(strFileKey, oDicPendientes(strFileKey))
                    System.Console.WriteLine("Resuelto en segunda pasada: " & strFileKey)
                Catch
                    System.Console.WriteLine("No se pudo renombrar en 2da pasada (sigue duplicado): " & strFileKey)
                End Try
            Next
        End If

        ' 4. Ejecución
        SendTo.Run()

        MsgBox("SendTo finalizado con éxito." & vbCrLf & "Pendientes intentados: " & oDicPendientes.Count)

    End Sub

End Module











