Imports System.Runtime.InteropServices

Public Class ExcelSession
    ' Declaración de la API de Windows para obtener el PID
    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    Public Property Application As Microsoft.Office.Interop.Excel.Application
    Public Property Workbooks As Microsoft.Office.Interop.Excel.Workbooks
    Public Property Workbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property IsReady As Boolean = False
    Public Property ErrorMessage As String = ""

    Public Sub New()
        Try
            ' 1. Conexión con Excel
            Me.Application = CType(Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)

            ' 2. Obtener el PID (Process ID) real
            ' Usamos una llamada a la API de Windows para convertir el Hwnd de la ventana en un PID
            Dim excelHwnd As New IntPtr(Me.Application.Hwnd)
            Dim excelPid As Integer
            GetWindowThreadProcessId(excelHwnd, excelPid)

            ' 3. Ahora sí, usamos AppActivate con el PID real
            AppActivate(excelPid)

            ' 4. Enviamos el ESC
            SendKeys.SendWait("{ESC}")


            If Me.Application.ActiveWorkbook Is Nothing Then
                Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
                Return
            End If


            Me.Workbooks = Me.Application.Workbooks
            Me.Workbook = Me.Application.ActiveWorkbook
            Me.ActiveSheet = CType(Me.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            Me.IsReady = True

        Catch ex As COMException
            Me.ErrorMessage = ">>> [ERROR] No se detectó ninguna instancia de Excel abierta."
        Catch ex As System.Exception
            Me.ErrorMessage = ">>> [ERROR] " & ex.Message
        End Try
    End Sub




    Sub CreateNewWorkbook()
        Me.Workbook = Me.Workbooks.Add()
        Me.ActiveSheet = CType(Me.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
    End Sub


    Sub GetActiveWorkbook()
        Try
            ' 1. Conexión con Excel
            Me.Application = CType(Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)




            ' 2. Obtener el PID (Process ID) real
            ' Usamos una llamada a la API de Windows para convertir el Hwnd de la ventana en un PID
            Dim excelHwnd As New IntPtr(Me.Application.Hwnd)
            Dim excelPid As Integer
            GetWindowThreadProcessId(excelHwnd, excelPid)

            ' 3. Ahora sí, usamos AppActivate con el PID real
            AppActivate(excelPid)

            ' 4. Enviamos el ESC
            SendKeys.SendWait("{ESC}")


            If Me.Application.ActiveWorkbook Is Nothing Then
                Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
                Return
            End If


            Me.Workbooks = Me.Application.Workbooks
            Me.Workbook = Me.Application.ActiveWorkbook
            Me.ActiveSheet = CType(Me.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            Me.IsReady = True

        Catch ex As COMException
            Me.ErrorMessage = ">>> [ERROR] No se detectó ninguna instancia de Excel abierta."
        Catch ex As System.Exception
            Me.ErrorMessage = ">>> [ERROR] " & ex.Message
        End Try

    End Sub







End Class