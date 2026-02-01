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


    End Sub


End Module