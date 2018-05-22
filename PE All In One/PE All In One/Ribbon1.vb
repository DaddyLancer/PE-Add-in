Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Dim frmCombine As New frmCombineSheets
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCombineSheets_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCombineSheets.Click
        frmCombine.ShowDialog()

    End Sub
End Class
