Imports Microsoft.Office.Interop.Excel
Public Class frmCombineSheets

    Dim wsCount As Integer

    Private Sub rbnNew_CheckedChanged(sender As Object, e As EventArgs) Handles rbnNew.CheckedChanged
        If rbnNew.Checked = True Then
            lblWarning1.Visible = False
            txtSheetName.Enabled = True
            cmbTarget.Enabled = False
        Else
            lblWarning1.Visible = True
            txtSheetName.Enabled = False
            cmbTarget.Enabled = True
        End If
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Dim frmAboutCombine As New AboutCombine
        frmAboutCombine.ShowDialog()
    End Sub

    Private Sub frmCombineSheets_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim curSheet As Excel.Worksheet = app.ActiveSheet
        Dim curWorkbook As Excel.Workbook = app.ActiveWorkbook
        Dim curSelection As Excel.Range = CType(app.Selection, Excel.Range)
        Dim wrkSheet As Worksheet
        wsCount = curWorkbook.Worksheets.Count
        lstSelectedSheets.Items.Clear()
        cmbTarget.Items.Clear()
        Try
            For Each wrkSheet In curWorkbook.Worksheets
                lstSelectedSheets.Items.Add(wrkSheet.Name)
                cmbTarget.Items.Add(wrkSheet.Name)
            Next
            cmbTarget.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnAll_Click(sender As Object, e As EventArgs) Handles btnAll.Click
        For i = 0 To lstSelectedSheets.Items.Count - 1
            lstSelectedSheets.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub btnNone_Click(sender As Object, e As EventArgs) Handles btnNone.Click
        For i = 0 To lstSelectedSheets.Items.Count - 1
            lstSelectedSheets.SetItemChecked(i, False)
        Next
    End Sub

    Private Sub btnInvert_Click(sender As Object, e As EventArgs) Handles btnInvert.Click
        For i = 0 To lstSelectedSheets.Items.Count - 1
            lstSelectedSheets.SetItemChecked(i, Not lstSelectedSheets.GetItemChecked(i))
        Next
    End Sub

    Private Sub rbtnExisting_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnExisting.CheckedChanged
        If rbtnExisting.Checked = True Then
            lstSelectedSheets.SetItemChecked(cmbTarget.SelectedIndex, False)
        End If
    End Sub

    Private Sub cmbTarget_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTarget.SelectedIndexChanged
        If rbtnExisting.Checked = True Then
            lstSelectedSheets.SetItemChecked(cmbTarget.SelectedIndex, False)
        End If
    End Sub
End Class