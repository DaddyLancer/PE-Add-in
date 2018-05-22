<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCombineSheets
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblWarning1 = New System.Windows.Forms.Label()
        Me.rbnNew = New System.Windows.Forms.RadioButton()
        Me.rbtnExisting = New System.Windows.Forms.RadioButton()
        Me.lstSelectedSheets = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnAll = New System.Windows.Forms.Button()
        Me.btnInvert = New System.Windows.Forms.Button()
        Me.btnNone = New System.Windows.Forms.Button()
        Me.cmbTarget = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSheetName = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnCombine = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblWarning1
        '
        Me.lblWarning1.AutoSize = True
        Me.lblWarning1.ForeColor = System.Drawing.Color.Maroon
        Me.lblWarning1.Location = New System.Drawing.Point(6, 171)
        Me.lblWarning1.Name = "lblWarning1"
        Me.lblWarning1.Size = New System.Drawing.Size(94, 26)
        Me.lblWarning1.TabIndex = 0
        Me.lblWarning1.Text = "Any existing data" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "will be overwritten."
        Me.lblWarning1.Visible = False
        '
        'rbnNew
        '
        Me.rbnNew.AutoSize = True
        Me.rbnNew.Checked = True
        Me.rbnNew.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.rbnNew.Location = New System.Drawing.Point(9, 60)
        Me.rbnNew.Name = "rbnNew"
        Me.rbnNew.Size = New System.Drawing.Size(76, 17)
        Me.rbnNew.TabIndex = 1
        Me.rbnNew.TabStop = True
        Me.rbnNew.Text = "New sheet"
        Me.rbnNew.UseVisualStyleBackColor = True
        '
        'rbtnExisting
        '
        Me.rbtnExisting.AutoSize = True
        Me.rbtnExisting.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.rbtnExisting.Location = New System.Drawing.Point(9, 124)
        Me.rbtnExisting.Name = "rbtnExisting"
        Me.rbtnExisting.Size = New System.Drawing.Size(90, 17)
        Me.rbtnExisting.TabIndex = 2
        Me.rbtnExisting.Text = "Existing sheet"
        Me.rbtnExisting.UseVisualStyleBackColor = True
        '
        'lstSelectedSheets
        '
        Me.lstSelectedSheets.CheckOnClick = True
        Me.lstSelectedSheets.FormattingEnabled = True
        Me.lstSelectedSheets.Location = New System.Drawing.Point(6, 33)
        Me.lstSelectedSheets.Name = "lstSelectedSheets"
        Me.lstSelectedSheets.Size = New System.Drawing.Size(148, 229)
        Me.lstSelectedSheets.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.Label1.Location = New System.Drawing.Point(6, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Select sheets to combine"
        '
        'btnAll
        '
        Me.btnAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAll.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnAll.Location = New System.Drawing.Point(6, 268)
        Me.btnAll.Name = "btnAll"
        Me.btnAll.Size = New System.Drawing.Size(65, 23)
        Me.btnAll.TabIndex = 8
        Me.btnAll.Text = "All"
        Me.btnAll.UseVisualStyleBackColor = True
        '
        'btnInvert
        '
        Me.btnInvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnInvert.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnInvert.Location = New System.Drawing.Point(6, 297)
        Me.btnInvert.Name = "btnInvert"
        Me.btnInvert.Size = New System.Drawing.Size(148, 23)
        Me.btnInvert.TabIndex = 9
        Me.btnInvert.Text = "Invert selection"
        Me.btnInvert.UseVisualStyleBackColor = True
        '
        'btnNone
        '
        Me.btnNone.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNone.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnNone.Location = New System.Drawing.Point(89, 268)
        Me.btnNone.Name = "btnNone"
        Me.btnNone.Size = New System.Drawing.Size(65, 23)
        Me.btnNone.TabIndex = 10
        Me.btnNone.Text = "None"
        Me.btnNone.UseVisualStyleBackColor = True
        '
        'cmbTarget
        '
        Me.cmbTarget.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTarget.Enabled = False
        Me.cmbTarget.FormattingEnabled = True
        Me.cmbTarget.Location = New System.Drawing.Point(9, 147)
        Me.cmbTarget.Name = "cmbTarget"
        Me.cmbTarget.Size = New System.Drawing.Size(121, 21)
        Me.cmbTarget.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.Label2.Location = New System.Drawing.Point(3, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(140, 26)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Where should the combined" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "data be displayed?"
        '
        'txtSheetName
        '
        Me.txtSheetName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSheetName.Location = New System.Drawing.Point(9, 83)
        Me.txtSheetName.Name = "txtSheetName"
        Me.txtSheetName.Size = New System.Drawing.Size(121, 20)
        Me.txtSheetName.TabIndex = 13
        Me.txtSheetName.Text = "Combined"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnCancel.Location = New System.Drawing.Point(6, 297)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(137, 23)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnCombine
        '
        Me.btnCombine.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnCombine.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCombine.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnCombine.Location = New System.Drawing.Point(6, 268)
        Me.btnCombine.Name = "btnCombine"
        Me.btnCombine.Size = New System.Drawing.Size(137, 23)
        Me.btnCombine.TabIndex = 16
        Me.btnCombine.Text = "Combine"
        Me.btnCombine.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.lstSelectedSheets)
        Me.GroupBox1.Controls.Add(Me.btnAll)
        Me.GroupBox1.Controls.Add(Me.btnInvert)
        Me.GroupBox1.Controls.Add(Me.btnNone)
        Me.GroupBox1.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.GroupBox1.Location = New System.Drawing.Point(12, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(160, 328)
        Me.GroupBox1.TabIndex = 17
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Source"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnHelp)
        Me.GroupBox2.Controls.Add(Me.btnCombine)
        Me.GroupBox2.Controls.Add(Me.btnCancel)
        Me.GroupBox2.Controls.Add(Me.cmbTarget)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.rbnNew)
        Me.GroupBox2.Controls.Add(Me.lblWarning1)
        Me.GroupBox2.Controls.Add(Me.txtSheetName)
        Me.GroupBox2.Controls.Add(Me.rbtnExisting)
        Me.GroupBox2.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.GroupBox2.Location = New System.Drawing.Point(178, 13)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(149, 328)
        Me.GroupBox2.TabIndex = 18
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Target"
        '
        'btnHelp
        '
        Me.btnHelp.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnHelp.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHelp.ForeColor = System.Drawing.Color.LightSkyBlue
        Me.btnHelp.Location = New System.Drawing.Point(113, 232)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(30, 30)
        Me.btnHelp.TabIndex = 17
        Me.btnHelp.Text = "?"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'frmCombineSheets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(336, 349)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmCombineSheets"
        Me.Text = "Combine multiple sheets into one"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblWarning1 As Windows.Forms.Label
    Friend WithEvents rbnNew As Windows.Forms.RadioButton
    Friend WithEvents rbtnExisting As Windows.Forms.RadioButton
    Friend WithEvents lstSelectedSheets As Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnAll As Windows.Forms.Button
    Friend WithEvents btnInvert As Windows.Forms.Button
    Friend WithEvents btnNone As Windows.Forms.Button
    Friend WithEvents cmbTarget As Windows.Forms.ComboBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents txtSheetName As Windows.Forms.TextBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnCombine As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents btnHelp As Windows.Forms.Button
End Class
