Imports Visio = Microsoft.Office.Interop.Visio

Public Class dlgLinkData

    Dim lngRowIDs() As Integer

    Private Sub dlgLinkData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim vsoDataRecordset As Visio.DataRecordset, arrDataRecordset() As String
        Dim i As Byte

        For i = 1 To vsoApp.ActiveDocument.DataRecordsets.Count
            vsoDataRecordset = vsoApp.ActiveDocument.DataRecordsets.Item(i)
            arrDataRecordset = Split(vsoDataRecordset.DataConnection.ConnectionString, ";")
            cmbSourseData.Items.Add(Strings.Right(arrDataRecordset(2), Len(arrDataRecordset(2)) - 12) & " - " & vsoDataRecordset.Name)
            cmb_DataID.Items.Add(vsoDataRecordset.ID)
        Next

        cmbSourseData.SelectedIndex = 0
        txtNameTable.Enabled = ckbInsertName.Checked
        ckbInsertName.Checked = False
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Call LinkToDataInShapes(Val(cmb_DataID.Text), ckbInsertName.Checked, txtNameTable.Text,
        ckbTitleColumns.Checked, ckbInvisibleZero.Checked,
        UBound(lngRowIDs), vsoApp.ActiveDocument.DataRecordsets.ItemFromID(Val(cmb_DataID.Text)).DataColumns.Count, ckbFontBold.Checked)

        Me.Close()

    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub cmdRefreshAll_Click(sender As Object, e As EventArgs) Handles cmdRefreshAll.Click
        Call RefreshDataInShapes()
    End Sub

    Private Sub cmbSourseData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSourseData.SelectedIndexChanged
        cmb_DataID.SelectedIndex = cmbSourseData.SelectedIndex
        txtNameTable.Text = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).Name
        lngRowIDs = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).GetDataRowIDs("")
        lblCountRow.Text = "Source contains " &
        vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).DataColumns.Count & " columns " _
        & UBound(lngRowIDs) + 1 & " rows of data"
    End Sub

    Private Sub ckbInsertName_CheckedChanged(sender As Object, e As EventArgs) Handles ckbInsertName.CheckedChanged
        txtNameTable.Enabled = ckbInsertName.Checked
    End Sub

    Sub RefreshDataInShapes() ' Updating External Data Sources

        If vsoApp.ActiveDocument.DataRecordsets.Count = 0 Then
            Exit Sub
        End If

        Dim vsoDataRecordset As Visio.DataRecordset
        Dim i As Byte

        Call RecUndo("Refresh all data")

        On Error GoTo ErrorHandler

        For i = 1 To vsoApp.ActiveDocument.DataRecordsets.Count
            vsoDataRecordset = vsoApp.ActiveDocument.DataRecordsets.Item(i)
            vsoDataRecordset.Refresh()
        Next
        Call RecUndo("0")

        MsgBox("Updated external data sources - " & vsoApp.ActiveDocument.DataRecordsets.Count, vbInformation, "Data Update")
        Exit Sub

ErrorHandler:
        MsgBox("External source:" & vbCrLf & vsoDataRecordset.Name & vbCrLf & Err.Description, vbExclamation, "Data Update")
    End Sub

End Class