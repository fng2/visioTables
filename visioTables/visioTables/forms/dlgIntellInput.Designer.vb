<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgIntellInput
    Inherits System.Windows.Forms.Form

    'The form overrides dispose to clear the list of components.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required for Windows Forms Form Designer
    Private components As System.ComponentModel.IContainer

    'Note: The following procedure is required for Windows Forms Designer
    'To change it, use the Windows Form Designer. 
    'Do not change it in the source code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmbText = New System.Windows.Forms.ComboBox()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmbText
        '
        Me.cmbText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbText.FormattingEnabled = True
        Me.cmbText.Location = New System.Drawing.Point(12, 12)
        Me.cmbText.Name = "cmbText"
        Me.cmbText.Size = New System.Drawing.Size(343, 21)
        Me.cmbText.TabIndex = 0
        '
        'btnHelp
        '
        Me.btnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHelp.AutoSize = True
        Me.btnHelp.Location = New System.Drawing.Point(369, 11)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(23, 23)
        Me.btnHelp.TabIndex = 4
        Me.btnHelp.Text = "?"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'dlgIntellInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(404, 47)
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.cmbText)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(600, 85)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(420, 85)
        Me.Name = "dlgIntellInput"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Entering text by moving through cells"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbText As System.Windows.Forms.ComboBox
    Friend WithEvents btnHelp As System.Windows.Forms.Button
End Class
