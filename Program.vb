Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine(New String("-"c, 60))


        ' --- 1. CONEXIÓN Y VALIDACIÓN ---
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            MsgBox(session.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False


        ' Conexión con Excel y acceso a oActiveSheet
        Dim oExcelApp As Microsoft.Office.Interop.Excel.Application = GetObject(, "Excel.Application")
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = oExcelApp.ActiveWorkbook
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.ActiveSheet

        ' -------------------
        ' ExcelDataExtractor
        ' -------------------
        Dim oExcelDataExtractor As New ExcelDataExtractor
        Dim oDic As Dictionary(Of String, ExcelData) = oExcelDataExtractor.ExtractData(oSheet)


        ' -------------------
        ' CatiaDataApplier
        ' -------------------
        Dim oCatiaDataInjector As New CatiaDataInjector
        oCatiaDataInjector.InjectData(oProduct, oDic)


        ' --- 4. LIMPIEZA FINAL (SIMPLE) ---
        ' Restauramos alertas
        session.Application.DisplayFileAlerts = True

        ' Usamos el limpiador para soltar todo lo que tocamos
        Dim cleaner As New CatiaDataCOMCleaner
        cleaner.Release(oSheet, oWorkbook, oExcelApp, oProduct, session.Application)

        Console.WriteLine(">>> Finalizado con éxito.")


    End Sub


End Module