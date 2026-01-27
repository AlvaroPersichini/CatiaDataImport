
Option Explicit On
Option Strict On
Public Class CatiaSession

    Private ReadOnly _app As INFITF.Application
    Private ReadOnly _status As CatiaSessionStatus

    Public Sub New()
        _app = Connect()
        _status = EvaluateStatus(_app)
    End Sub

    Private Function Connect() As INFITF.Application
        Try
            Return CType(GetObject(, "CATIA.Application"), INFITF.Application)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function EvaluateStatus(app As INFITF.Application) As CatiaSessionStatus
        If app Is Nothing Then Return CatiaSessionStatus.NotRunning
        If app.Windows.Count = 0 Then Return CatiaSessionStatus.NoWindowsOpen
        Select Case TypeName(app.ActiveDocument)
            Case "ProductDocument" : Return CatiaSessionStatus.ProductDocument
            Case "PartDocument" : Return CatiaSessionStatus.PartDocument
            Case "DrawingDocument" : Return CatiaSessionStatus.DrawingDocument
            Case "CatalogDocument" : Return CatiaSessionStatus.CatalogDocument
            Case "AnalysisDocument" : Return CatiaSessionStatus.AnalysisDocument
            Case "CATProcessDocument" : Return CatiaSessionStatus.ProcessDocument
            Case "CATScriptDocument" : Return CatiaSessionStatus.ScriptDocument
            Case Else : Return CatiaSessionStatus.Unknown
        End Select
    End Function

    Public ReadOnly Property Application As INFITF.Application
        Get
            Return _app
        End Get
    End Property

    Public ReadOnly Property Status As CatiaSessionStatus
        Get
            Return _status
        End Get
    End Property

    Public ReadOnly Property IsReady As Boolean
        Get
            Return Status = CatiaSessionStatus.ProductDocument _
                OrElse Status = CatiaSessionStatus.PartDocument _
                OrElse Status = CatiaSessionStatus.DrawingDocument _
                OrElse Status = CatiaSessionStatus.CatalogDocument
        End Get
    End Property

    Public ReadOnly Property Description As String
        Get
            Select Case Status
                Case CatiaSessionStatus.NotRunning : Return "CATIA is not running."
                Case CatiaSessionStatus.NoWindowsOpen : Return "CATIA is running but has no document open."
                Case CatiaSessionStatus.ProductDocument : Return "Product document is active."
                Case CatiaSessionStatus.PartDocument : Return "Part document is active."
                Case CatiaSessionStatus.DrawingDocument : Return "Drawing document is active."
                Case CatiaSessionStatus.CatalogDocument : Return "Catalog document is active."
                Case CatiaSessionStatus.AnalysisDocument : Return "Analysis document is active."
                Case CatiaSessionStatus.ProcessDocument : Return "Process document is active."
                Case CatiaSessionStatus.ScriptDocument : Return "Script document is active."
                Case Else : Return "Unknown CATIA state."
            End Select
        End Get
    End Property


    Public Enum CatiaSessionStatus
        NotRunning = 0
        NoWindowsOpen = 1
        ProductDocument = 2
        PartDocument = 3
        DrawingDocument = 4
        CatalogDocument = 5
        AnalysisDocument = 6
        ProcessDocument = 7
        ScriptDocument = 8
        Unknown = -1
    End Enum


End Class
