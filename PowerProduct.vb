Option Explicit On
Public Class PowerProduct

    Public oProduct As ProductStructureTypeLib.Product
    Public intLevel As Integer
    Public intQuantity As Integer
    Public oProductFather As ProductStructureTypeLib.Product


    Public Sub New()

    End Sub


    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub



    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub


End Class


Public Class PwrProduct

    Private oProduct As ProductStructureTypeLib.Product
    Private intQuantity As Integer
    Private intLevel As Integer
    Private strProductType As String
    Private strSource As String



    Public Property Product As ProductStructureTypeLib.Product
        Get
            Return oProduct
        End Get
        Set(value As ProductStructureTypeLib.Product)
            oProduct = value
        End Set
    End Property

    Public Property Quantity As Integer
        Get
            Return intQuantity
        End Get
        Set(value As Integer)
            intQuantity = value
        End Set
    End Property

    Public Property AssemblyLevel As Integer
        Get
            Return intLevel
        End Get
        Set(value As Integer)
            intLevel = value
        End Set
    End Property

    Public Property ProductType As String
        Get
            Return strProductType
        End Get
        Set(value As String)
            strProductType = value
        End Set
    End Property

    Public Property Source As String
        Get
            Return strSource
        End Get
        Set(value As String)
            strSource = value
        End Set
    End Property

End Class