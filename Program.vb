Option Explicit On
Option Strict On
Imports CatiaExcelClassLibrary


Module Program

    Sub Main()


        ' Catia
        Dim session As New CatiaSession
        If Not session.IsReady Then
            Form1.ToolStripStatusLabel1.Text = session.Description
            Return
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False


        ' Excel
        Dim xlSession As New ExcelSession()
        If Not xlSession.IsReady Then
            Form1.ToolStripStatusLabel1.Text = xlSession.ErrorMessage
            Return
        End If


        ' Extraccion de datos de Excel:
        ' Se extraen los datos de Excel a un diccionario.
        ' Ese diccionario luego 
        Form1.ToolStripStatusLabel1.Text = "> Extracting data from Excel..."
        Dim oExcelDataExtractor As New ExcelDataExtractor()
        Dim oDic As Dictionary(Of String, ExcelData) = oExcelDataExtractor.ExtractData(xlSession.ActiveSheet)


        ' Aplicacion
        ' Acá se puede cambiar a que sea una function no un sub. y que la function devuelva el diccionario de los que se modificaron
        Form1.ToolStripStatusLabel1.Text = "> Exporting data to CATIA..."
        Dim oCatiaDataInjector As New CatiaDataInjector()
        oCatiaDataInjector.InjectData(oProduct, oDic)


        ' Extract Catia Data
        Form1.ToolStripStatusLabel1.Text = "> Extracting data from CATIA..."
        Dim oCatiaDataExtractor As New CatiaDataExtractor()
        Dim oCatiaData As Dictionary(Of String, PwrProduct) = oCatiaDataExtractor.ExtractData(oProduct, "", False)

        ' Inject into Excel
        Form1.ToolStripStatusLabel1.Text = "> Injecting data into Excel..."
        Dim oExcelDataInjector As New ExcelDataInjector()
        oExcelDataInjector.InjectData(xlSession.ActiveSheet, oCatiaData)


        ' Limpieza
        Form1.ToolStripStatusLabel1.Text = ">>> Cleaning up resources..."
        session.Application.DisplayFileAlerts = True
        Dim cleaner As New CatiaDataCOMCleaner()
        cleaner.Release(xlSession.ActiveSheet, xlSession.Workbook, xlSession.Application, oProduct, session.Application)

        ' Fin
        Form1.ToolStripStatusLabel1.Text = "Done!"



    End Sub



    Sub Update()

        ' Catia
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            Form1.ToolStripStatusLabel1.Text = session.Description
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

        ' Extract Catia Data
        Console.WriteLine(">>> Updating Excel with results...")
        Dim oCatiaDataExtractor As New CatiaDataExtractor()
        Dim oCatiaData As Dictionary(Of String, PwrProduct) = oCatiaDataExtractor.ExtractData(oProduct, "", False)


        ' Inject into Excel
        Console.WriteLine(">>> Injecting data into Excel...")
        Dim oExcelDataInjector As New ExcelDataInjector()
        oExcelDataInjector.InjectData(xlSession.ActiveSheet, oCatiaData)


        ' Limpieza
        Console.WriteLine(">>> Cleaning up resources...")
        session.Application.DisplayFileAlerts = True
        Dim cleaner As New CatiaDataCOMCleaner()
        cleaner.Release(xlSession.ActiveSheet, xlSession.Workbook, xlSession.Application, oProduct, session.Application)

        ' Fin
        Console.WriteLine(New String("-"c, 60))
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))

    End Sub

End Module