﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgLinkData
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
        Me.lblCountRow = New System.Windows.Forms.Label()
        Me.lblSourseData = New System.Windows.Forms.Label()
        Me.cmbSourseData = New System.Windows.Forms.ComboBox()
        Me.txtNameTable = New System.Windows.Forms.TextBox()
        Me.ckbInsertName = New System.Windows.Forms.CheckBox()
        Me.ckbTitleColumns = New System.Windows.Forms.CheckBox()
        Me.ckbFontBold = New System.Windows.Forms.CheckBox()
        Me.ckbInvisibleZero = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.cmdRefreshAll = New System.Windows.Forms.Button()
        Me.cmb_DataID = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblCountRow
        '
        Me.lblCountRow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblCountRow.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.lblCountRow.Location = New System.Drawing.Point(12, 9)
        Me.lblCountRow.Name = "lblCountRow"
        Me.lblCountRow.Size = New System.Drawing.Size(513, 23)
        Me.lblCountRow.TabIndex = 0
        Me.lblCountRow.Text = "Source"
        '
        'lblSourseData
        '
        Me.lblSourseData.AutoSize = True
        Me.lblSourseData.Location = New System.Drawing.Point(12, 31)
        Me.lblSourseData.Name = "lblSourseData"
        Me.lblSourseData.Size = New System.Drawing.Size(135, 13)
        Me.lblSourseData.TabIndex = 1
        Me.lblSourseData.Text = "Selecting a data source"
        '
        'cmbSourseData
        '
        Me.cmbSourseData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSourseData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSourseData.FormattingEnabled = True
        Me.cmbSourseData.Location = New System.Drawing.Point(12, 47)
        Me.cmbSourseData.Name = "cmbSourseData"
        Me.cmbSourseData.Size = New System.Drawing.Size(513, 21)
        Me.cmbSourseData.TabIndex = 2
        '
        'txtNameTable
        '
        Me.txtNameTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNameTable.Location = New System.Drawing.Point(12, 107)
        Me.txtNameTable.Name = "txtNameTable"
        Me.txtNameTable.Size = New System.Drawing.Size(513, 20)
        Me.txtNameTable.TabIndex = 3
        '
        'ckbInsertName
        '
        Me.ckbInsertName.AutoSize = True
        Me.ckbInsertName.Checked = True
        Me.ckbInsertName.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbInsertName.Location = New System.Drawing.Point(12, 84)
        Me.ckbInsertName.Name = "ckbInsertName"
        Me.ckbInsertName.Size = New System.Drawing.Size(259, 17)
        Me.ckbInsertName.TabIndex = 4
        Me.ckbInsertName.Text = "Insert table name in first row"
        Me.ckbInsertName.UseVisualStyleBackColor = True
        '
        'ckbTitleColumns
        '
        Me.ckbTitleColumns.AutoSize = True
        Me.ckbTitleColumns.Checked = True
        Me.ckbTitleColumns.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbTitleColumns.Location = New System.Drawing.Point(300, 84)
        Me.ckbTitleColumns.Name = "ckbTitleColumns"
        Me.ckbTitleColumns.Size = New System.Drawing.Size(130, 17)
        Me.ckbTitleColumns.TabIndex = 5
        Me.ckbTitleColumns.Text = "Column Headings"
        Me.ckbTitleColumns.UseVisualStyleBackColor = True
        '
        'ckbFontBold
        '
        Me.ckbFontBold.AutoSize = True
        Me.ckbFontBold.Location = New System.Drawing.Point(12, 142)
        Me.ckbFontBold.Name = "ckbFontBold"
        Me.ckbFontBold.Size = New System.Drawing.Size(227, 17)
        Me.ckbFontBold.TabIndex = 6
        Me.ckbFontBold.Text = "Make headings bold"
        Me.ckbFontBold.UseVisualStyleBackColor = True
        '
        'ckbInvisibleZero
        '
        Me.ckbInvisibleZero.AutoSize = True
        Me.ckbInvisibleZero.Location = New System.Drawing.Point(245, 142)
        Me.ckbInvisibleZero.Name = "ckbInvisibleZero"
        Me.ckbInvisibleZero.Size = New System.Drawing.Size(280, 17)
        Me.ckbInvisibleZero.TabIndex = 7
        Me.ckbInvisibleZero.Text = "Do not show zero and empty cell values"
        Me.ckbInvisibleZero.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(379, 178)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 18
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 12
        Me.OK_Button.Text = "ОК"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 13
        Me.Cancel_Button.Text = "Cancel"
        '
        'cmdRefreshAll
        '
        Me.cmdRefreshAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdRefreshAll.AutoSize = True
        Me.cmdRefreshAll.Location = New System.Drawing.Point(12, 181)
        Me.cmdRefreshAll.Name = "cmdRefreshAll"
        Me.cmdRefreshAll.Size = New System.Drawing.Size(182, 23)
        Me.cmdRefreshAll.TabIndex = 19
        Me.cmdRefreshAll.Text = "Refresh all data sources"
        Me.cmdRefreshAll.UseVisualStyleBackColor = True
        '
        'cmb_DataID
        '
        Me.cmb_DataID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmb_DataID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_DataID.FormattingEnabled = True
        Me.cmb_DataID.Location = New System.Drawing.Point(278, 183)
        Me.cmb_DataID.Name = "cmb_DataID"
        Me.cmb_DataID.Size = New System.Drawing.Size(83, 21)
        Me.cmb_DataID.TabIndex = 20
        Me.cmb_DataID.Visible = False
        '
        'dlgLinkData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 213)
        Me.Controls.Add(Me.cmb_DataID)
        Me.Controls.Add(Me.cmdRefreshAll)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.ckbInvisibleZero)
        Me.Controls.Add(Me.ckbFontBold)
        Me.Controls.Add(Me.ckbTitleColumns)
        Me.Controls.Add(Me.ckbInsertName)
        Me.Controls.Add(Me.txtNameTable)
        Me.Controls.Add(Me.cmbSourseData)
        Me.Controls.Add(Me.lblSourseData)
        Me.Controls.Add(Me.lblCountRow)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(700, 252)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(551, 252)
        Me.Name = "dlgLinkData"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Communication with external data"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblCountRow As System.Windows.Forms.Label
    Friend WithEvents lblSourseData As System.Windows.Forms.Label
    Friend WithEvents cmbSourseData As System.Windows.Forms.ComboBox
    Friend WithEvents txtNameTable As System.Windows.Forms.TextBox
    Friend WithEvents ckbInsertName As System.Windows.Forms.CheckBox
    Friend WithEvents ckbTitleColumns As System.Windows.Forms.CheckBox
    Friend WithEvents ckbFontBold As System.Windows.Forms.CheckBox
    Friend WithEvents ckbInvisibleZero As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents cmdRefreshAll As System.Windows.Forms.Button
    Friend WithEvents cmb_DataID As System.Windows.Forms.ComboBox
End Class
