
Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine(New String("-"c, 60))



        ' Catia
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            Console.WriteLine(">>> [ABORT] CATIA Error: " & session.Description)
            Return
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False



        ' Excel
        Dim xlSession As New ExcelSession()
        If Not xlSession.IsReady Then
            Console.WriteLine(xlSession.ErrorMessage)
            Return
        End If



        ' Extraccion
        Console.WriteLine(">>> Extracting data from Excel...")
        Dim oExcelDataExtractor As New ExcelDataExtractor()
        ' Usamos la hoja de la sesión
        Dim oDic As Dictionary(Of String, ExcelData) = oExcelDataExtractor.ExtractData(xlSession.ActiveSheet)



        ' Aplicacion
        Console.WriteLine(">>> Injecting data into CATIA tree...")
        Dim oCatiaDataInjector As New CatiaDataInjector()
        oCatiaDataInjector.InjectData(oProduct, oDic)



        ' Limpieza
        session.Application.DisplayFileAlerts = True
        Dim cleaner As New CatiaDataCOMCleaner()
        cleaner.Release(xlSession.ActiveSheet,
                        xlSession.Workbook,
                        xlSession.Application,
                        oProduct,
                        session.Application)


        Console.WriteLine(New String("-"c, 60))
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))

    End Sub

End Module