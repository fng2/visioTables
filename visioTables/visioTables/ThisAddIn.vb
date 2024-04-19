Imports System.Drawing
Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports Visio = Microsoft.Office.Interop.Visio
Partial Public Class ThisAddIn
    Public Property AddinUI As AddinUI = New AddinUI()

    Private accessVBA As ClassVBA 'new


    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addinUI
    End Function

    'new
    Protected Overrides Function RequestComAddInAutomationService() As Object
        If accessVBA Is Nothing Then
            accessVBA = New ClassVBA()
        End If
        Return accessVBA
    End Function

    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String, commandTag As String)

        'If Application Is Nothing OrElse Application.ActiveWindow Is Nothing Then Exit Sub

        ' Cleanse?          'new
        winObj = vsoApp.ActiveWindow
        docObj = vsoApp.ActiveDocument
        pagObj = vsoApp.ActivePage
        shpsObj = pagObj.Shapes
        ' Cleanse?

        'new
        Select Case commandId
            Case "btn_newtable" : CreatingTable.Load_dlgNewTable() : Return 'label="New table"
            Case "btn_lockpicture" : LoadDlg(5) : Return 'New dlgPictures, label="Lock shapes"
            Case "btn_help" : CallHelp() : Return
        End Select

        If Not CheckSelCells() Then Exit Sub

        Select Case commandId
            Case "btn_newcolumnbefore", "btn_newcolumnafter" : AddColumns(commandTag) 'label="Column left, label="Column right
            Case "btn_newrowbefore", "btn_newrowafter" : AddRows(commandTag) 'label="Row above, label="Row below
            Case "btn_onwidth" : AllAlignOnText(True, False, 0, 0, True, True) 'label="By text width
            Case "btn_onheight" : AllAlignOnText(False, True, 0, 0, True, True) 'label="By text height
            Case "btn_onwidthheight" : AllAlignOnText(True, True, 0, 0, False, False) 'label="Automatic selection by text
            Case "btn_seltable", "btn_selrange", "btn_selcolumn", "btn_selrow" : SelCell(commandTag) 'label="Range
            Case "btn_seltxt", "btn_selnum", "btn_selnotnum", "btn_seldate", "btn_selempty", "btn_selnotempty", "btn_selinvert" : SelInContent(commandTag) 'label="Select values, label="Select Not values, label="Select dates, label="Select empty, label="Select Not empty
            Case "btn_text", "btn_date", "btn_time", "btn_comment", "btn_numcol", "btn_numrow" : InsertText(commandTag) 'label="Date", label="Time", label="A comment", label="Column number", label="Line number"
            Case "btn_intdeint" : IntDeIntCells() 'label="Merge/Disconnect"
            Case "btn_gut" : GutT() 'label="Cut"
            Case "btn_copy" : CopyT() 'label="Copy"
            Case "btn_paste" : PasteT() 'label="Insert"
            Case "btn_delcolumn" : DelColRows(0) 'label="Column"
            Case "btn_delrow" : DelColRows(1) 'label="Line"
            Case "btn_deltable" : DelTab(True) 'label="Delete table"
            Case "btn_intellinput" : LoadDlg(4) 'New dlgIntellInput, label="Intelligence set"
            Case "btn_sizeonwidth", "btn_sizeonheight" : AlignOnSize(commandTag) 'label="Align column widths"
            Case "btn_size" : LoadDlg(0) 'New dlgTableSize, label="Dimensions"
            Case "btn_autosize" : LoadDlg(1) 'New dlgTableSize, label="Autodimensions" 
            Case "btn_sorttabledata" : LoadDlg(7) 'New dlgSortTable, label="Sort data"
            Case "btn_fromfile" : LoadDlg(2) 'New dlgFromFile, label="Insert from file"
            Case "btn_dropdownlist" : LoadDlg(6) 'New dlgSelectFromList, label="Insert from the list"
            Case "btn_altlinesrow", "btn_altlinescol" : AlternatLines(commandTag) 'label="Alternate lines"
            Case "btn_extdata" : LoadDlg(3) ' New dlgLinkData, label="External data"
            Case "btn_rotatetext" : AllRotateText() 'label="Rotate text"
            Case "btn_convert1Shape" : ConvertInto1Shape() 'label="Convert to 1 shape"
        End Select
    End Sub


    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command shoudl be enabled in the user interface.
    ''' By default, all commands are enabled.
    ''' 
    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "Command1"
                ' make command1 always enabled
                Return True

            Case "Command2"
                ' make command2 enabled only if a window is opened
                'Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing AndAlso Application.ActiveWindow.Selection.Count > 0
                Return True
            Case Else
                Return True
        End Select
    End Function

    'new
    Public Function IsCommandAltEnabled(commandId As String) As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    'new
    Sub Application_ShapeAdded(ByVal Sh As Microsoft.Office.Interop.Visio.Shape)
        Dim nC As Integer = Val(Strings.Left(Matrica, Strings.InStr(1, Matrica, "x", 1) - 1))
        Dim nR As Integer = Val(Strings.Right(Matrica, Strings.Len(Matrica) - Strings.InStr(1, Matrica, "x", 1)))
        strNameTable = "TbL"
        RemoveHandler Application.ShapeAdded, AddressOf Application_ShapeAdded
        Call CreatTable(strNameTable, 4, nC, nR, 0, 0, 0, 0, True, False)
        Application.DoCmd(1907)
    End Sub

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
        Return False
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Returns a label associated with given command.
    ''' We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
    ''' 
    Public Function GetCommandLabel(command As String) As String
        Return My.Resources.ResourceManager.GetString(command & "_Label")
    End Function

    ''' 
    ''' Returns a icon associated with given command.
    ''' We assume for simplicity that icon ids are named after command commandId.
    ''' 
    Public Function GetCommandBitmap(command As String) As Bitmap
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Bitmap)
    End Function


    Sub UpdateUI()
        AddinUI.UpdateRibbon()

    End Sub

    Public Sub Application_SelectionChanged(window As Visio.Window)
        UpdateUI()
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
