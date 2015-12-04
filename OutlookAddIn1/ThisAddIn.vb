Public Class ThisAddIn

    Friend lang As String
    Public Work_phone_undefined

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As  _
    Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

End Class
