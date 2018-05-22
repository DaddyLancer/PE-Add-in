Public Class frmCombineSheets
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

    Private Sub frmCombineSheets_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class