Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine(New String("-"c, 60))



        ' Conexión con CATIA mediante y acceso al Product activo
        Dim session As New CatiaSession()
        If session.Status <> CatiaSession.CatiaSessionStatus.ProductDocument Then
            MsgBox("Error: Se requiere un Product activo." & Environment.NewLine &
                   "Estado actual: " & session.Description, MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim oAppCatia As INFITF.Application = session.Application
        oAppCatia.DisplayFileAlerts = False
        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = CType(oAppCatia.ActiveDocument, ProductStructureTypeLib.ProductDocument)



        ' Chequeo de estado de guardado
        If Not CheckSaveStatus(oProductDocument) Then
            MsgBox("El documento actual no ha sido guardado. Guárdelo antes de continuar.", Microsoft.VisualBasic.MsgBoxStyle.Exclamation, "Aviso")
            Exit Sub
        End If



        ' --- GESTIÓN DE DIRECTORIOS ---
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = System.IO.Path.Combine(baseDir, "Export_" & timestamp)

        ' Verificamos si la carpeta existe, y si no, la creamos
        If Not IO.Directory.Exists(folderPath) Then
            ' CreateDirectory crea toda la ruta necesaria (incluyendo carpetas padre si no existen)
            IO.Directory.CreateDirectory(folderPath)
            Console.WriteLine("Carpeta creada: " & folderPath)
        Else
            Console.WriteLine("La carpeta ya existe: " & folderPath)
        End If






        ' Generación del Diccionario
        Dim oDic1 As Specialized.StringDictionary = GetMap(oProductDocument.Product)





        ' Ejecución del SendTo con renombrado
        SendToWithPartN.SendTOWPN(oProductDocument, oDic1, folderPath)




    End Sub










    ' --- LÓGICA DE MAPEO ---
    Public Function GetMap(objRoot As ProductStructureTypeLib.Product) As System.Collections.Specialized.StringDictionary
        Dim dicc As New System.Collections.Specialized.StringDictionary()
        FillMap(objRoot, dicc)
        Return dicc
    End Function








    Private Sub FillMap(current As ProductStructureTypeLib.Product, ByRef dicc As System.Collections.Specialized.StringDictionary)
        ' Si el link está roto, el error saltará aquí (Depuración directa)
        Dim docName As String = current.ReferenceProduct.Parent.Name
        Dim pn As String = current.PartNumber

        If Not dicc.ContainsKey(docName) Then
            dicc.Add(docName, pn)
        End If

        ' Navegación por la colección de productos
        For Each child As ProductStructureTypeLib.Product In current.Products
            FillMap(child, dicc)
        Next
    End Sub






    ' --- VERIFICACIÓN DE ESTADO DE GUARDADO ---
    Private Function CheckSaveStatus(oDoc As INFITF.Document) As Boolean
        If System.String.IsNullOrEmpty(oDoc.Path) Then Return False
        If Not oDoc.Saved Then Return False
        Return True
    End Function





End Module