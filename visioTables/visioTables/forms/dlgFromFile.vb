Imports Visio = Microsoft.Office.Interop.Visio

Public Class dlgFromFile

    Private fso As Object, folder As String, fileObj As Object
    Private intlistStart As Integer

    'vsoObj.Count is the number of cells in the table that are selected
    'user MUST select one or more cells before clicking OK
    'user can select cells in a row, in a column, diagonally or ANY order
    'if user selects enough cells then all of the text in the file will be inserted, otherwise it will be truncated to the number of cells selected

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Dim vsoObj As Visio.Selection = winObj.Selection
        If ListBox2.Items.Count = 0 Then Exit Sub

        Dim iC As Integer = vsoObj(1).Cells(UTC).Result("")
        Dim iR As Integer = intlistStart

        Call RecUndo("Paste from file")

        On Error Resume Next

        For iC = 1 To vsoObj.Count
            vsoObj(iC).Characters.Text = ListBox2.Items(iR)
            If iR = ListBox2.Items.Count - 1 Then
                iR = 0
            Else : iR = iR + 1
            End If
        Next

        Call RecUndo("0")
        Me.Close()
    End Sub

    Private Sub ReadFolder()
        fso = CreateObject("Scripting.FileSystemObject")
        folder = vsoApp.MyShapesPath & "\" & "Backfill files"

        If fso.FolderExists(folder) Then
            lblTxt.Text = Strings.Right(folder, 75)
            For Each Me.fileObj In fso.GetFolder(folder).Files
                ListBox1.Items.Add(fileObj.Name)
            Next
            ToolTip1.SetToolTip(lblTxt, folder)
        Else
            MsgBox("Folder does not exist: " & vbCrLf & folder, vbExclamation)
            Me.Close()
        End If

        If ListBox1.Items.Count = 0 Then
            MsgBox("There are no files in the folder: " & vbCrLf & folder, vbInformation)
            Me.Close()
        End If

        fso = Nothing
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim TextStream

        fso = CreateObject("Scripting.FileSystemObject")
        fileObj = fso.GetFile(folder & "\" & ListBox1.SelectedItem)

        TextStream = fileObj.OpenAsTextStream(1)
        ListBox2.Items.Clear()
        While Not TextStream.AtEndOfStream
            ListBox2.Items.Add(TextStream.ReadLine())
        End While
        TextStream.Close()
        intlistStart = 0
        fso = Nothing

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        intlistStart = sender.SelectedIndex
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub dlgFromFile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ReadFolder()
    End Sub

End Class