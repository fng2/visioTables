Option Explicit On

Imports System.Data
Imports visio = Microsoft.Office.Interop.Visio

Module SpecialTables
    Private _PartsDataTable As New DataTable()
    Private vsoSel As visio.Selection

    Private Sub makeBomTable()

        Dim dlg_DataRow As DataRow

        Try
            MsgBox("makeBomTable2")
            _PartsDataTable.Clear()
            If Not _PartsDataTable.Columns.Contains("Item") Then
                _PartsDataTable.Columns.Add("Item", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Qty") Then
                _PartsDataTable.Columns.Add("Qty", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Part") Then
                _PartsDataTable.Columns.Add("Part", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("AltPart") Then
                _PartsDataTable.Columns.Add("AltPart", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Kit") Then
                _PartsDataTable.Columns.Add("Kit", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Description") Then
                _PartsDataTable.Columns.Add("Description", Type.GetType("System.String"))
            End If

            dlg_DataRow = _PartsDataTable.NewRow()
            dlg_DataRow("Item") = "1"
            dlg_DataRow("Qty") = "2"
            dlg_DataRow("Part") = "widget"
            dlg_DataRow("AltPart") = "hisWidget"
            dlg_DataRow("Kit") = "widgetKit"
            dlg_DataRow("Description") = "this is a widget"
            _PartsDataTable.Rows.Add(dlg_DataRow)

            dlg_DataRow = _PartsDataTable.NewRow()
            dlg_DataRow("Item") = "2"
            dlg_DataRow("Qty") = "3"
            dlg_DataRow("Part") = "megawidget"
            dlg_DataRow("AltPart") = "hismegaWidget"
            dlg_DataRow("Kit") = "megawidgetKit"
            dlg_DataRow("Description") = "this is a mega widget"
            _PartsDataTable.Rows.Add(dlg_DataRow)
            MsgBox("done with makeBomTable2")
        Catch ex As Exception
            MsgBox("error making data table")
        End Try

    End Sub

    Private Function GetLineByComment(ID) As Integer
        'MsgBox("GetLineByComment")
        Dim line As String = String.Empty
        Dim Parts() As String
        'need some error checking!!!
        Try
            If Not ID.contains("Line") Then
                Return 0
            End If

            'Parts = Split(ID, "&") 'dont work
            'Parts = Split(ID, "\&")  'dont work
            'Parts = Split(ID, "%26") 'dont work
            Parts = Split(ID, " ") 'should give us 3 parts, i.e. "Column", "1" & vblf & "line" and "1"

            If Parts.Length = 3 Then
                line = Convert.ToInt16(Parts(2))
            End If
            Return line
        Catch ex As Exception
            MsgBox("error in GetLineByComment")
            Return 0
        End Try

    End Function

    Public Sub BOM()
        'CreatTable(strNameTable, bytInsertType, nudColumns.Value, nudRows.Value, w, h, wT, hT, ckbDelShape.Checked, True)
        'User.TableCol and User.TableRow are in each "working" cell and give the cell's location in the table
        'TableName is "BOM", insert type is 1 (default), 6 columns, 4 rows, width of cell, height of cell, width of table, height of table, delete shape checkbox and progress bar


        'makeBomTable()

        Dim numberOfColumns As Integer = 6
        Dim numberOfRows As Integer = 4 'remember the top two are headers!
        Dim ID As String = String.Empty
        Dim rowNum As Integer = 0
        Dim NewTable As New VisioTable
        Dim chars As visio.Characters
        Dim dlg_DataRow As DataRow

        Try
            MsgBox("in makeBomTablex")
            _PartsDataTable.Clear()
            If Not _PartsDataTable.Columns.Contains("Item") Then
                _PartsDataTable.Columns.Add("Item", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Qty") Then
                _PartsDataTable.Columns.Add("Qty", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Part") Then
                _PartsDataTable.Columns.Add("Part", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("AltPart") Then
                _PartsDataTable.Columns.Add("AltPart", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Kit") Then
                _PartsDataTable.Columns.Add("Kit", Type.GetType("System.String"))
            End If
            If Not _PartsDataTable.Columns.Contains("Description") Then
                _PartsDataTable.Columns.Add("Description", Type.GetType("System.String"))
            End If

            dlg_DataRow = _PartsDataTable.NewRow()
            dlg_DataRow("Item") = "1"
            dlg_DataRow("Qty") = "2"
            dlg_DataRow("Part") = "widget"
            dlg_DataRow("AltPart") = "hisWidget"
            dlg_DataRow("Kit") = "widgetKit"
            dlg_DataRow("Description") = "this is a widget"
            _PartsDataTable.Rows.Add(dlg_DataRow)

            dlg_DataRow = _PartsDataTable.NewRow()
            dlg_DataRow("Item") = "2"
            dlg_DataRow("Qty") = "3"
            dlg_DataRow("Part") = "megawidget"
            dlg_DataRow("AltPart") = "hismegaWidget"
            dlg_DataRow("Kit") = "megawidgetKit"
            dlg_DataRow("Description") = "this is a mega widget"
            _PartsDataTable.Rows.Add(dlg_DataRow)
            MsgBox("done with makeBomTablex")
            'Catch ex As Exception
            '    MsgBox("error making data table")
            'End Try

            'Try
            '    MsgBox("debug in tables")
            'if there is already a BOM table bail!

            NewTable.CreatTable("BOM", 1, numberOfColumns, numberOfRows, 1, 0.5, 1, 1, True, True)


            Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
            vsoSel = Globals.ThisAddIn.Application.ActiveWindow.Selection


            For Each shape As visio.Shape In Globals.ThisAddIn.Application.ActivePage.Shapes
                'does User section exist? if not, bail
                If shape.SectionExists(242, 1) Then
                    'there could be other tables on the page, make sure we only get cells from the BOM table!
                    If shape.CellsSRC(visio.VisSectionIndices.visSectionUser, visio.VisRowIndices.visRowUser, visio.VisCellIndices.visUserValue).ResultStr("") = "BOM" Then

                        'if it's the control cell, then turn off headers?
                        If shape.Name = "BOM" Then
                            shape.CellsSRC(visio.VisSectionIndices.visSectionAction, visio.VisRowIndices.visRowAction, visio.VisCellIndices.visActionChecked).FormulaForceU = 0
                        End If

                        ID = shape.CellsSRC(visio.VisSectionIndices.visSectionObject, visio.VisRowIndices.visRowMisc, visio.VisCellIndices.visComment).ResultStr("")
                        'if the cell belongs to line 1 add it to a collection to be merged later
                        If ID.Contains("Line 1") Then
                            vsoSel.Select(shape, visio.VisSelectArgs.visSelect)
                        End If

                        'if it's the first cell, then add our title
                        'make it BOLD
                        If shape.Name = "ClW" Then
                            'ClW should be the row1/column1 cell
                            'add the title
                            shape.Text = "Bill Of Materials"
                            chars = shape.Characters
                            chars.Begin = 0
                            chars.End = 17
                            chars.CharProps(visio.VisCellIndices.visCharacterSize) = 16 'this works fine
                            'chars.CharProps(visio.VisCellIndices.visCharacterStyle) = visbold 'visbold is not declared
                            chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                        End If

                        'if we are on line 2, add the column descriptions
                        'make these BOLD
                        If ID.Contains("Line 2") Then
                            Select Case True
                                Case ID.Contains("Column 1")
                                    shape.Text = "Item" & vbLf & "Number"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 4
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                                Case ID.Contains("Column 2")
                                    shape.Text = "Qty"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 3
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                                Case ID.Contains("Column 3")
                                    shape.Text = "Part" & vbLf & "Number"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 11
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                                Case ID.Contains("Column 4")
                                    shape.Text = "Alt part" & vbLf & "Number"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 14
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                                Case ID.Contains("Column 5")
                                    shape.Text = "Kit" & vbLf & "Number"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 9
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                                Case ID.Contains("Column 6")
                                    shape.Text = "Description"
                                    chars = shape.Characters
                                    chars.Begin = 0
                                    chars.End = 11
                                    chars.CharProps(visio.VisCellIndices.visCharacterStyle) = 17.0# 'BOLD
                            End Select
                        End If

                        rowNum = GetLineByComment(ID)

                        'if we have exhausted the rows in the datatable but still have rows in the BOM, then DON'T do this next code fragment
                        'that will not happen if we call createtable with the number of rows in the datatable (+2)
                        If Not ID.Contains("Line 1") And Not ID.Contains("Line 2") Then
                            Select Case True
                                Case ID.Contains("Column 1")
                                    shape.Text = _PartsDataTable.Rows(rowNum - 3).Item("Item") 'subtracting 3 from rownum to account for the first two header rows
                                Case ID.Contains("Column 2")
                                    shape.Text = _PartsDataTable.Rows.Item(rowNum - 3).Item("Qty")
                                Case ID.Contains("Column 3")
                                    shape.Text = _PartsDataTable.Rows.Item(rowNum - 3).Item("Part")
                                Case ID.Contains("Column 4")
                                    shape.Text = _PartsDataTable.Rows.Item(rowNum - 3).Item("AltPart")
                                Case ID.Contains("Column 5")
                                    shape.Text = _PartsDataTable.Rows.Item(rowNum - 3).Item("Kit")
                                Case ID.Contains("Column 6")
                                    shape.Text = _PartsDataTable.Rows.Item(rowNum - 3).Item("Description")
                            End Select
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("error in BOM " & ex.Message)
        End Try

        Try
            'merge the cells of line 1
            If Not vsoSel Is Nothing Then
                If vsoSel.Count > 2 Then
                    'Dim junk1 As Integer = Globals.ThisAddIn.Application.ActiveWindow.Selection.Count
                    Globals.ThisAddIn.Application.ActiveWindow.Selection = vsoSel
                    'junk1 = Globals.ThisAddIn.Application.ActiveWindow.Selection.Count
                    IntDeIntCells()
                End If
                Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox("error merging cells")
        End Try


        NewTable = Nothing
        _PartsDataTable = Nothing

    End Sub

    Public Sub noBOM()
        MsgBox("this works")
    End Sub

End Module
