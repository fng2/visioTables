Option Explicit On

Imports System.Windows.Forms
Imports visio = Microsoft.Office.Interop.Visio


Imports System.Data
Imports System.Data.OleDb



Module CreatingTable

#Region "LIST OF VARIABLES AND CONSTANTS"
    Friend Matrica As String = ""
    Friend CheckArrID As String = ""
    Friend vsoApp As Visio.Application = Globals.ThisAddIn.Application
    Friend winObj As Visio.Window
    Friend docObj As Visio.Document
    Friend pagObj As Visio.Page
    Friend shpsObj As Visio.Shapes
    Friend frmNewTable As System.Windows.Forms.Form
    Friend strNameTable As String = ""
    Friend FlagPage As Byte = 0
    Friend ArrShapeID(,) As Integer
    Friend NT As String = ""

    Friend Const UTN = "User.TableName"
    Friend Const UTR = "User.TableRow"
    Friend Const UTC = "User.TableCol"

    Dim selObj As Visio.Selection
    Dim vsoSelection As Visio.Selection
    Dim MemSel As Visio.Shape
    Dim NoDupes As New Collection
    Dim UndoScopeID As Long = 0
    Dim LayerVisible As String = ""
    Dim LayerLock As String = ""


#End Region



#Region "Load Sub"

    Sub CreatTable(a, b, c, d, e, f, g, h, i, j)
        Dim NewTable As New VisioTable
        NewTable.CreatTable(a, b, c, d, e, f, g, h, i, j)
        NewTable = Nothing
    End Sub

    Sub Load_dlgNewTable()
        If frmNewTable Is Nothing Then
            frmNewTable = New dlgNewTable
            frmNewTable.Show()
        End If
    End Sub

    Sub LoadDlg(arg)
        Dim dlgNew As System.Windows.Forms.Form = Nothing

        Try
            Select Case arg
                Case 0, 1 : FlagPage = arg : dlgNew = New dlgTableSize
                Case 2 : dlgNew = New dlgFromFile
                Case 3
                    If vsoApp.ActiveDocument.DataRecordsets.Count = 0 Then '<-bombs here "This operation Is Not supported in Microsoft Visio Standard 2010."
                        MsgBox("There are no external data connections available in the active document.", vbCritical, "Error")
                        'GoTo err
                    End If
                    dlgNew = New dlgLinkData
                Case 4 : dlgNew = New dlgIntellInput
                Case 5 : dlgNew = New dlgPictures
                Case 6 : dlgNew = New dlgSelectFromList
                Case 7 : dlgNew = New dlgSortTable
            End Select

            dlgNew.ShowDialog()
            dlgNew = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        'err:
        'dlgNew = Nothing
    End Sub

#End Region

#Region "Friend Sub"

    ' Insert a new column. Basic procedure
    Sub AddColumns(arg As Byte)
        'arg = 0 inserts a column before the selected one, arg = 1 inserts a column after the selected one

        Call ClearControlCells(UTC)
        If winObj.Selection.Count = 0 Then Exit Sub

        shpsObj = winObj.Page.Shapes

        Dim shObj As Visio.Shape = winObj.Selection(1)
        Dim strF As String

        Call InitArrShapeID(NT)
        winObj.DeselectAll()

        Call PropLayers(1)

        Call SelectCls(shObj.Cells(UTC).Result(""), 0, shObj.Cells(UTC).Result(""), UBound(ArrShapeID, 2))

        Dim vs As Visio.Selection = winObj.Selection
        Dim iAll As Integer = shpsObj.Item(NT).Cells(UTC).Result("")
        Dim nCol As Integer = shObj.Cells(UTC).Result("")

        Call RecUndo("Add Column")
        vs.Duplicate() : Dim vsoDups As Visio.Selection = winObj.Selection

        For i = 2 To vsoDups.Count
            With vsoDups(i)
                If Not .Characters.IsField Then .Characters.Text = ""
                .Cells("Comment").FormulaForceU = "=Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                If InStr(1, .Cells(WI).FormulaU, "SUM") <> 0 Then
                    .Cells(LD).FormulaForceU = 0
                    .Delete()
                End If
            End With
        Next

        Dim NTNew As String = vsoDups(1).Name

        If arg = 0 Then   ' Insert a column before the selection
            With vs(1)
                .Cells(PX).FormulaForceU = GU & NTNew & "!PinX+(" & NTNew & "!Width/2)+(Width/2))"
            End With
            For i = nCol To UBound(ArrShapeID, 1) ' Renumbering control cells
                shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).Result("") + 1 & ")"
            Next

        ElseIf arg = 1 Then  ' Insert a column after the selected one
            With vsoDups(1)
                .Cells(PX).FormulaForceU = GU & vs(1).Name & "!PinX+(" & vs(1).Name & "!Width/2)+(Width/2))"
                .Cells(UTC).FormulaForceU = GU & .Cells(UTC).Result("") + 1 & ")"
            End With
            If nCol <> shpsObj.Item(NT).Cells(UTC).Result("") Then
                shpsObj.ItemFromID(ArrShapeID(vs(1).Cells(UTC).Result("") + 1, 0)).Cells(PX).FormulaForceU = GU & vsoDups(1).Name & "!PinX+(" & vsoDups(1).Name & "!Width/2)+(Width/2))"
                For i = nCol + 1 To UBound(ArrShapeID, 1) ' Renumbering control cells
                    shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).Result("") + 1 & ")"
                Next
            End If
        End If

        With shpsObj
            .Item(NT).Cells(UTC).FormulaForceU = "GUARD(" & iAll + 1 & ")"
            .Item(NTNew).Cells("Controls.ControlWidth").FormulaForceU = "Width*1"
            .Item(NTNew).SendToBack()

            If vsoDups.Count <> iAll + 1 Then ' Defining merged cells and processing them
                For j = 1 To UBound(ArrShapeID, 1)
                    For i = 1 To UBound(ArrShapeID, 2)
                        If ArrShapeID(j, i) <> 0 Then
                            With .ItemFromID(ArrShapeID(j, i))
                                If InStr(1, .Cells(WI).FormulaU, "SUM", 1) <> 0 Then
                                    If InStr(1, .Cells(WI).FormulaU, vs(1).Name & "!", 1) <> 0 Then
                                        If arg = 0 Then
                                            .Cells(WI).FormulaForceU = Replace$(.Cells(WI).FormulaU, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            strF = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!PinX-", NTNew & "!PinX-", 1)
                                            strF = Replace$(strF, vs(1).Name & "!Width/2", NTNew & "!Width/2", 1)
                                            .Cells(PX).FormulaForceU = Replace$(strF, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            .Cells(UTC).FormulaForceU = Replace$(.Cells(UTC).FormulaU, vs(1).Name & "!", NTNew & "!", 1)
                                        End If
                                        If arg = 1 Then
                                            .Cells(WI).FormulaForceU = Replace$(.Cells(WI).FormulaU, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            .Cells(PX).FormulaForceU = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!Width,", vs(1).Name & "!Width" & "," & NTNew & "!Width,", 1)
                                            .Cells(PX).FormulaForceU = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!Width)", vs(1).Name & "!Width" & "," & NTNew & "!Width)", 1)
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    Next
                Next
            End If
        End With

        Call RecUndo("0")
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
        On Error Resume Next
        winObj.Selection = vsoDups

        Call PropLayers(0)

    End Sub

    ' Insert a new line. Basic procedure
    Sub AddRows(arg As Byte)
        'arg = 0 inserts a line before the selected one, arg = 1 inserts a line after the selected one

        Call ClearControlCells(UTR)
        If winObj.Selection.Count = 0 Then Exit Sub

        shpsObj = winObj.Page.Shapes

        Dim shObj As Visio.Shape = winObj.Selection(1)
        Dim strF As String

        Call InitArrShapeID(NT) ': winObj.Select(winObj.Page.Shapes.item(1), 256)
        winObj.DeselectAll()

        Call PropLayers(1)

        Call SelectCls(0, shObj.Cells(UTR).Result(""), UBound(ArrShapeID, 1), shObj.Cells(UTR).Result(""))

        Dim vs As Visio.Selection = winObj.Selection
        Dim iAll As Integer = shpsObj.Item(NT).Cells(UTR).Result("")
        Dim nRow As Integer = shObj.Cells(UTR).Result("")

        Call RecUndo("Add line")
        vs.Duplicate() : Dim vsoDups As Visio.Selection = winObj.Selection

        For i = 2 To vsoDups.Count
            With vsoDups(i)
                If Not .Characters.IsField Then .Characters.Text = ""
                .Cells("Comment").FormulaForceU = "=Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                If InStr(1, .Cells(HE).FormulaU, "SUM") <> 0 Then
                    .Cells(LD).FormulaForceU = 0
                    .Delete()
                End If
            End With
        Next

        Dim NTNew As String = vsoDups(1).Name

        If arg = 0 Then ' Insert a line before the highlighted one
            With vs(1)
                .Cells(PY).FormulaForceU = GU & NTNew & "!PinY-(" & NTNew & "!Height/2)-(Height/2))"
            End With
            For i = nRow To UBound(ArrShapeID, 2) ' Renumbering control cells
                shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).Result("") + 1 & ")"
            Next

        ElseIf arg = 1 Then ' Insert a line after the highlighted one
            With vsoDups(1)
                .Cells(PY).FormulaForceU = GU & vs(1).Name & "!PinY-(" & vs(1).Name & "!Height/2)-(Height/2))"
                .Cells(UTR).FormulaForceU = GU & .Cells(UTR).Result("") + 1 & ")"
            End With
            If nRow <> shpsObj.Item(NT).Cells(UTR).Result("") Then
                shpsObj.ItemFromID(ArrShapeID(0, vs(1).Cells(UTR).Result("") + 1)).Cells(PY).FormulaForceU = GU & vsoDups(1).Name & "!PinY-(" & vsoDups(1).Name & "!Height/2)-(Height/2))"
                For i = nRow + 1 To UBound(ArrShapeID, 2) ' Renumbering control cells
                    shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).Result("") + 1 & ")"
                Next
            End If
        End If

        With shpsObj
            .Item(NT).Cells(UTR).FormulaForceU = "GUARD(" & iAll + 1 & ")"
            .Item(NTNew).Cells("Controls.ControlHeight").FormulaForceU = "Guard(Height*0)"
            .Item(NTNew).SendToBack()

            If vsoDups.Count <> iAll + 1 Then ' Defining merged cells and processing them
                For j = 1 To UBound(ArrShapeID, 1)
                    For i = 1 To UBound(ArrShapeID, 2)
                        If ArrShapeID(j, i) <> 0 Then
                            With .ItemFromID(ArrShapeID(j, i))
                                If InStr(1, .Cells(HE).FormulaU, "SUM", 1) <> 0 Then
                                    If InStr(1, .Cells(HE).FormulaU, vs(1).Name & "!", 1) <> 0 Then
                                        If arg = 0 Then
                                            .Cells(HE).FormulaForceU = Replace$(.Cells(HE).FormulaU, vs(1).Name & "!Height", NTNew & "!Height" & "," & vs(1).Name & "!Height", 1)
                                            strF = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!PinY+", NTNew & "!PinY+", 1)
                                            strF = Replace$(strF, vs(1).Name & "!Height/2", NTNew & "!Height/2", 1)
                                            .Cells(PY).FormulaForceU = Replace$(strF, vs(1).Name & "!Height", NTNew & "!Height" & "," & vs(1).Name & "!Height", 1)
                                            .Cells(UTR).FormulaForceU = Replace$(.Cells(UTR).FormulaU, vs(1).Name & "!", NTNew & "!", 1)
                                        End If
                                        If arg = 1 Then
                                            .Cells(HE).FormulaForceU = Replace$(.Cells(HE).FormulaU, vs(1).Name & "!Height", vs(1).Name & "!Height" & "," & NTNew & "!Height", 1)
                                            .Cells(PY).FormulaForceU = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!Height,", vs(1).Name & "!Height" & "," & NTNew & "!Height,", 1)
                                            .Cells(PY).FormulaForceU = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!Height)", vs(1).Name & "!Height" & "," & NTNew & "!Height)", 1)
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    Next
                Next
            End If
        End With

        Call RecUndo("0")
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
        On Error Resume Next
        winObj.Selection = vsoDups

        Call PropLayers(0)

    End Sub

    ' Align/auto-align table cells to text width/height. Preliminary procedure
    Sub AllAlignOnText(booOnWidth As Boolean, booOnHeight As Boolean, bytNothingOrAutoOrLockColumns As Byte, _
        bytNothingOrAutoOrLockRows As Byte, booOnlySelectedColumns As Boolean, booOnlySelectedRow As Boolean)


        Dim vsoSel As Visio.Selection
        Dim ShNum As Integer, iCount As Integer, bytColumnOrRow As Byte

        shpsObj = winObj.Page.Shapes : vsoSel = winObj.Selection
        Call InitArrShapeID(NT)

        vsoApp.ShowChanges = False

        With shpsObj

            If booOnWidth And Not booOnlySelectedColumns Then ' Auto-align all columns
                bytColumnOrRow = 4
                Call RecUndo("Align everything to text width")
                For iCount = 1 To UBound(ArrShapeID, 1)
                    ShNum = .ItemFromID(ArrShapeID(iCount, 0)).Cells(UTC).Result("") : Call AlignOnText(ShNum, bytColumnOrRow, bytNothingOrAutoOrLockColumns)
                Next
                Call RecUndo("0")
            End If

            If booOnHeight And Not booOnlySelectedRow Then ' Auto-align all lines
                bytColumnOrRow = 5
                Call RecUndo("Align everything to text height")
                For iCount = 1 To UBound(ArrShapeID, 2)
                    ShNum = .ItemFromID(ArrShapeID(0, iCount)).Cells(UTR).Result("") : Call AlignOnText(ShNum, bytColumnOrRow, bytNothingOrAutoOrLockRows)
                Next
                Call RecUndo("0")
            End If

            If booOnWidth And booOnlySelectedColumns Then ' Auto-align only selected columns
                bytColumnOrRow = 4 : NotDub(vsoSel, UTC)
                Call RecUndo("Align to text width")
                For iCount = 1 To NoDupes.Count
                    If NoDupes(iCount) <> 0 Then Call AlignOnText(NoDupes(iCount), bytColumnOrRow, bytNothingOrAutoOrLockColumns)
                Next
                Call RecUndo("0")
                NoDupes.Clear()
            End If

            If booOnHeight And booOnlySelectedRow Then ' Auto-align only selected lines
                bytColumnOrRow = 5 : NotDub(vsoSel, UTR)
                Call RecUndo("Align to text height")
                For iCount = 1 To NoDupes.Count
                    If NoDupes(iCount) <> 0 Then Call AlignOnText(NoDupes(iCount), bytColumnOrRow, bytNothingOrAutoOrLockRows)
                Next
                Call RecUndo("0")
                NoDupes.Clear()
            End If

        End With

        vsoApp.ShowChanges = True
        winObj.Selection = vsoSel

    End Sub


    Sub AlignOnSize(arg As Byte)

        ' MsgBox("alignOnSize in create table")

        Call InitArrShapeID(NT)

        Select Case arg
            Case 4 : ClearControlCells(UTC)
            Case 5 : ClearControlCells(UTR)
        End Select

        If winObj.Selection.Count = 0 Then GoTo err

        Dim strCellWH As String = "", dblResult As Double
        Dim vsoSel As visio.Selection = winObj.Selection

        If vsoSel.Count = 1 Then
            Select Case arg
                Case 4 : Call SelectCls(1, 0, UBound(ArrShapeID, 1), 0)
                Case 5 : Call SelectCls(0, 1, 0, UBound(ArrShapeID, 2))
            End Select
        End If

        vsoSelection = winObj.Selection

        Select Case arg
            Case 4 : NotDub(vsoSelection, UTC) : strCellWH = WI : Call RecUndo("Align column widths")
            Case 5 : NotDub(vsoSelection, UTR) : strCellWH = HE : Call RecUndo("Align row heights")
        End Select

        With winObj.Page.Shapes
            For i = 1 To NoDupes.Count
                Select Case arg
                    Case 4 : dblResult = dblResult + .ItemFromID(ArrShapeID(NoDupes(i), 0)).Cells(strCellWH).Result(64)
                    Case 5 : dblResult = dblResult + .ItemFromID(ArrShapeID(0, NoDupes(i))).Cells(strCellWH).Result(64)
                End Select
            Next
            dblResult = dblResult / NoDupes.Count
            For i = 1 To NoDupes.Count
                Select Case arg
                    Case 4 : .ItemFromID(ArrShapeID(NoDupes(i), 0)).Cells(strCellWH).Result(64) = dblResult
                    Case 5 : .ItemFromID(ArrShapeID(0, NoDupes(i))).Cells(strCellWH).Result(64) = dblResult
                End Select
            Next
        End With

        NoDupes.Clear()
        Call RecUndo("0")
        winObj.Selection = vsoSel

err:
    End Sub

    ' Rotate text in selected cells according to a specified angle
    Sub AllRotateText(Optional arg As Boolean = False, Optional ang As Double = 0)
        Dim iAng As Double
        vsoSelection = winObj.Selection

        If arg Then
            iAng = ang
        Else
            iAng = Val(InputBox("Specify the angle in degrees.", "Text rotation", "90"))
        End If

        Call RecUndo("Rotate text")

        For i = 1 To vsoSelection.Count
            With vsoSelection(i)
                If .CellExistsU(UTN, 0) Then
                    If InStr(1, .Name, "ClW", 1) <> 0 Then
                        .Cells("TxtPinX").FormulaU = "Width*0.5"
                        .Cells("TxtPinY").FormulaU = "Height*0.5"
                        .Cells("TxtLocPinX").FormulaU = "TxtWidth*0.5"
                        .Cells("TxtLocPinY").FormulaU = "TxtHeight*0.5"
                        If iAng = 0 Or iAng = 180 Then
                            .Cells("TxtWidth").FormulaU = "Width*1"
                            .Cells("TxtHeight").FormulaU = "Height*1"
                        Else
                            .Cells("TxtWidth").FormulaU = "Height*1"
                            .Cells("TxtHeight").FormulaU = "Width*1"
                        End If
                        .Cells("TxtAngle").FormulaU = Str(iAng) & " deg"
                    End If
                End If
            End With
        Next

        Call RecUndo("0")

    End Sub

    ' Creating a "striped" table across columns or rows
    Sub AlternatLines(iAlt As Byte)
        Dim i As Integer, j As Integer, strCellWH As String = UTC
        shpsObj = winObj.Page.Shapes : MemSel = winObj.Selection(1)

        Call InitArrShapeID(NT)

        If iAlt = 5 Then strCellWH = UTR

        Call RecUndo("Row/Column Alternation")

        With shpsObj
            For i = 1 To UBound(ArrShapeID, 1)
                For j = 1 To UBound(ArrShapeID, 2)
                    If ArrShapeID(i, j) <> 0 Then
                        If .ItemFromID(ArrShapeID(i, j)).Cells(strCellWH).Result("") Mod 2 = 0 Then .ItemFromID(ArrShapeID(i, j)).Cells("FillForegnd").FormulaU = "THEMEGUARD(MSOTINT(RGB(255,255,255),-10))"
                    End If
                Next
            Next
        End With

        Call RecUndo("0")

        'winObj.Select(winObj.Page.Shapes.item(1), 256)
        winObj.Select(MemSel, 2)

    End Sub

    'Calling Help File - Tables in Visio.chm
    Sub CallHelp()
        Dim RetVal, strPath As String
        strPath = "C:\Windows\hh.exe " & vsoApp.MyShapesPath & "\" & "Tables in Visio.chm"
        RetVal = Shell(strPath, 1)
    End Sub

    ' Convert a table into one grouped shape
    Sub ConvertInto1Shape()
        Dim visWorkCells As Visio.Selection, i As Integer
        Dim dblTop As Double, dblBottom As Double, dblLeft As Double, dblRight As Double

        Call InitArrShapeID(winObj.Selection(1).Cells(UTN).ResultStr(""))
        winObj.Page.Shapes.ItemU(NT).BoundingBox(1, dblLeft, dblBottom, dblRight, dblTop)

        Call SelCell(2, False) : visWorkCells = winObj.Selection
        winObj.DeselectAll()

        Call SelectCls(0, 0, UBound(ArrShapeID, 1), 0)
        Call SelectCls(0, 1, 0, UBound(ArrShapeID, 2))

        Call RecUndo("Convert to 1 shape")
        On Error GoTo err
        With winObj
            .Selection.Group()
            .Selection.DeleteEx(0)
            .Selection = visWorkCells
            .Selection.Group()
        End With

        For i = 1 To visWorkCells.Count
            With visWorkCells(i)
                .DeleteSection(242)
                .DeleteSection(240)
                .DeleteSection(243)
                .CellsSRC(1, 17, 16).FormulaU = ""
                .CellsSRC(1, 15, 5).FormulaForceU = "0"
                .CellsSRC(1, 15, 8).FormulaForceU = "0"
                .CellsSRC(1, 2, 3).FormulaForceU = ""
                .CellsSRC(1, 6, 0).FormulaU = ""
            End With
        Next

        With winObj.Selection(1)
            .Cells(PX).Result("") = dblLeft + (.Cells(PX).Result(""))
            .Cells(PY).Result("") = dblBottom - (.Cells(PY).Result(""))
            .Name = NT
        End With

err:
        Call RecUndo("0")
    End Sub

    ' Cutting content from selected table cells
    Sub GutT()
        Dim txt As String = ""
        My.Computer.Clipboard.SetText(fArrT(txt))

        Call RecUndo("Cut text from cells")

        Dim vsoSelection As Visio.Selection = winObj.Selection
        For Each shpObj In vsoSelection
            shpObj.Characters.Text = ""
        Next

        Call RecUndo("0")
    End Sub

    ' Copying the contents of selected table cells
    Sub CopyT()
        Dim txt As String = ""
        My.Computer.Clipboard.SetText(fArrT(txt))
    End Sub

    ' Paste clipboard contents into table cells
    Sub PasteT()
        shpsObj = winObj.Page.Shapes

        Call InitArrShapeID(NT)

        Dim ShapeObj As Visio.Shape
        Dim arrId(,) As String, arrTMP() As String, arrTMP1() As String, txt As String
        Dim i As Integer, j As Integer

        On Error GoTo err

        txt = My.Computer.Clipboard.GetText
        arrTMP = Split(txt, vbCrLf)
        arrTMP1 = Split(arrTMP(0), vbTab)

        ReDim arrId(UBound(arrTMP, 1) - 1, UBound(arrTMP1, 1))
        For i = LBound(arrId, 1) To UBound(arrId, 1)
            arrTMP1 = Split(arrTMP(i), vbTab)
            For j = LBound(arrTMP1, 1) To UBound(arrTMP1, 1)
                arrId(i, j) = arrTMP1(j)
            Next
        Next

        ShapeObj = winObj.Selection(1) : shpsObj = winObj.Page.Shapes

        On Error Resume Next

        Call RecUndo("Insert text into cells")

        For i = LBound(arrId, 1) To UBound(arrId, 1)
            For j = LBound(arrId, 2) To UBound(arrId, 2)
                shpsObj.ItemFromID(ArrShapeID(j + ShapeObj.Cells(UTC).Result(""), i + ShapeObj.Cells(UTR).Result(""))).Characters.Text = arrId(i, j)
            Next
        Next

err:
        Call RecUndo("0")
        Erase arrId : Erase arrTMP : Erase arrTMP1
    End Sub

    'Removing columns/rows from the active table from the shape. Preliminary procedure
    Sub DelColRows(bytColsOrRows As Byte)
        shpsObj = winObj.Page.Shapes
        Dim vsoSel As Visio.Selection = winObj.Selection, shObj As Visio.Shape

        For Each shObj In vsoSel
            If shObj.CellExists(UTN, 0) = -1 Then
                Select Case bytColsOrRows
                    Case 0
                        If shpsObj.Item(NT).Cells(UTC).Result("") = 1 Then GoTo err
                        Call InitArrShapeID(NT) : Call DeleteColumn(shObj)
                    Case 1
                        If shpsObj.Item(NT).Cells(UTR).Result("") = 1 Then GoTo err
                        Call InitArrShapeID(NT) : Call DeleteRow(shObj)
                End Select
            End If
        Next
err:
        vsoSel = Nothing
    End Sub

    ' Deleting the active table. Basic procedure
    Sub DelTab(arg As Boolean)
        On Error GoTo errD
        Dim Response As Byte = 0
        ' 6 - Yes, 7 - no, 2 - cancellation
        'If Not CheckSelCells() Then Exit Sub

        'If Response = 0 Then
        If arg Then
            Response = MsgBox("Are you sure you want to delete this table?", 67, "Removal!")
        Else
            Response = 6
        End If
        'End If

        If Response = 6 Then
            winObj = vsoApp.ActiveWindow
            shpsObj = winObj.Page.Shapes
            NT = winObj.Selection(1).Cells(UTN).ResultStr("")
            Call RecUndo("Delete table")

            Dim frm As New dlgWait
            frm.Label1.Text = " " & vbCrLf & "Deleting a table..."
            frm.Show() : frm.Refresh()

            'winObj.Select(winObj.Page.Shapes.item(1), 256)
            vsoApp.ShowChanges = False

            Dim dblW As Integer, iCount As Integer
            dblW = shpsObj.Count

            For iCount = shpsObj.Count To 1 Step -1
                frm.lblProgressBar.Width = (300 / dblW) * iCount : frm.lblProgressBar.Refresh() : Application.DoEvents()
                With shpsObj.Item(iCount)
                    If .CellExistsU(UTN, 0) Then
                        If StrComp(.Cells(UTN).ResultStr(""), NT) = 0 Then
                            .Cells(LD).FormulaForceU = 0
                            .Delete()
                        End If
                    End If
                End With
            Next

            frm.Close()
            vsoApp.ShowChanges = True
            Call RecUndo(0)
        End If
        Exit Sub
errD:
        MessageBox.Show("DelTab" & vbNewLine & Err.Description)
    End Sub

    ' Merging/Disconnecting cells from a shape. Preliminary procedure
    Sub IntDeIntCells()
        Call ClearControlCells(UTC)
        Call ClearControlCells(UTR)
        If Not CheckSelCells() Then Exit Sub
        Dim shObj As visio.Shape, vsoSel As visio.Selection = winObj.Selection
        shpsObj = winObj.Page.Shapes
        shObj = vsoSel(1)
        Call InitArrShapeID(NT)
        If InStr(1, shObj.Cells("Width").FormulaU, "SUM", 1) <> 0 Or InStr(1, shObj.Cells("Height").FormulaU, "Sum", 1) <> 0 Then
            Call RecUndo("Disconnect cells")
            Call DeIntegrateCells(shObj)
        Else
            Call RecUndo("Merge cells")
            Call IntegrateCells()
        End If
        Call RecUndo("0")
    End Sub

    ' Filling the array of ID shapes of the active table
    Sub InitArrShapeID(NSh)
        'NSh - string variable, cell value "User.TableName" any shape from the active table
        Dim Ui = winObj.Page.Shapes(NT).UniqueID(0)
        If Len(Ui) = 0 Then Ui = winObj.Page.Shapes(NT).UniqueID(1)
        If CheckArrID = Ui & "0" Then Exit Sub

        Dim shObj As Visio.Shape
        Dim cMax As Integer = shpsObj.Item(NSh).Cells(UTC).Result("")
        Dim rMax As Integer = shpsObj.Item(NSh).Cells(UTR).Result("")

        ReDim ArrShapeID(cMax, rMax)

        For Each shObj In shpsObj
            With shObj
                If .CellExistsU(UTN, 0) AndAlso StrComp(.Cells(UTN).ResultStr(""), NSh) = 0 Then ArrShapeID(.Cells(UTC).Result(""), .Cells(UTR).Result("")) = .ID
                'If Left$(.NameU, 5) = "Sheet" Then MsgBox("The table may be damaged!") 
            End With
        Next

        ArrShapeID(0, 0) = shpsObj.Item(NSh).ID
        If ArrShapeID(cMax, rMax) = ArrShapeID(0, 0) Then ArrShapeID(cMax, rMax) = 0
        CheckArrID = Ui & "0"

    End Sub

    'Insert text, date, time, comment, column number, row number into cells
    Sub InsertText(arg)
        Dim title As String, msgComm As String, txt As String = "", i As Integer
        Dim vsoSel As Visio.Selection = winObj.Selection, arrArg() As String

        title = "Paste into cells"
        Call RecUndo("Paste into cells")

        Dim TextInsert = Sub()
                             If txt <> "" Then
                                 For i = 1 To vsoSel.Count
                                     vsoSel(i).Characters.Text = txt
                                 Next
                             End If
                         End Sub

        Dim NumInsert = Sub()
                            For i = 1 To vsoSel.Count
                                vsoSel(i).Characters.AddCustomFieldU(txt, 0)
                            Next
                        End Sub

        Select Case arg ' We need to fix RecUndo
            Case 0 : txt = InputBox("Insert text", title, "Text...") : TextInsert()
            Case 1 : txt = InputBox("Insert date", title, Today) : TextInsert()
            Case 2 : txt = InputBox("Insert time", title, TimeString) : TextInsert()
            Case 3
                msgComm = "0 - Restore to default" & vbCrLf & "1 - Cell text to comment" & vbCrLf & "2 - Comment text in cell" & vbCrLf
                txt = InputBox("Comments:" & vbCrLf & msgComm, title, "A comment...")
                Select Case txt
                    Case "" ' Cancel
                        GoTo err
                    Case 0 ' Restore to default
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = "Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                        Next
                    Case 1 ' Cell text to comment
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = Chr(34) & vsoSel(i).Characters.Text & Chr(34)
                        Next
                    Case 2 ' Comment text in cell
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Characters.Text = vsoSel(i).Cells("Comment").ResultStr("")
                        Next
                    Case Else ' User comment
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = Chr(34) & txt & Chr(34)
                        Next
                End Select
            Case 4
                msgComm = "Format:" & vbCrLf & "Prefix, Offset, Postfix"
                txt = InputBox("Insert column number" & vbCrLf & msgComm & vbCrLf, title, ",0,")
                If txt = "" Then GoTo err
                arrArg = Split(txt, ",")
                If UBound(arrArg) <> 2 Then GoTo err
                txt = """" & arrArg(0) & """" & "&" & "User.TableCol" & "+" & Int(Val(arrArg(1))) & "&" & """" & arrArg(2) & """" : NumInsert()
            Case 5
                msgComm = "Format:" & vbCrLf & "Prefix, Offset, Postfix"
                txt = InputBox("Insert column number" & vbCrLf & msgComm & vbCrLf, title, ",0,")
                If txt = "" Then GoTo err
                arrArg = Split(txt, ",")
                If UBound(arrArg) <> 2 Then GoTo err
                txt = """" & arrArg(0) & """" & "&" & "User.TableRow" & "+" & Int(Val(arrArg(1))) & "&" & """" & arrArg(2) & """" : NumInsert()
        End Select

err:
        Call RecUndo("0")
    End Sub

    ' Linking tables to external data sources
    Sub LinkToDataInShapes(ID, booInsertTableName, TblName, booTitleColumns, _
         booInvisibleZero, intCountRowSourse, intCountColSourse, booFontBold)
        ' Need to check merged cells


        Dim vsoDataRecordset As Visio.DataRecordset
        Dim shpObj As Visio.Shape
        Dim intEndCol As Integer, intEndRow As Integer
        Dim intRowStart As Integer, intCountCur As Integer, intRS As Integer

        Call InitArrShapeID(NT)

        shpsObj = winObj.Page.Shapes
        vsoDataRecordset = vsoApp.ActiveDocument.DataRecordsets.ItemFromID(ID)

        'Call InitArrShapeID(NT)
        intRowStart = 0

        ' Defining the cell range for communication
        If booInsertTableName Then intRS = intRS + 1
        If booTitleColumns Then intRS = intRS + 1
        intEndCol = shpsObj.Item(NT).Cells(UTC).Result("")
        intEndRow = shpsObj.Item(NT).Cells(UTR).Result("")
        If shpsObj.Item(NT).Cells(UTC).Result("") > intCountColSourse Then intEndCol = intCountColSourse
        If shpsObj.Item(NT).Cells(UTR).Result("") - intRS > intCountRowSourse Then intEndRow = intCountRowSourse + intRS + 1

        Dim frm As New dlgWait
        frm.Label1.Text = " " & vbCrLf & " " & vbCrLf & "Wait..."
        frm.Show() : frm.Refresh()

        Call RecUndo("Link data to shapes")

        On Error GoTo ExitLine

        ' Insert table title         ' (Check 1 line for concatenation!!!!!)
        If booInsertTableName Then
            With winObj
                .DeselectAll()
                .Select(shpsObj.ItemFromID(GetShapeId(1, 1)), 2)
                .Select(shpsObj.ItemFromID(GetShapeId(UBound(ArrShapeID, 1), 1)), 2)

                Call IntegrateCells()
                shpObj = shpsObj.ItemFromID(GetShapeId(1, 1))

                shpObj.Characters.Text = TblName
                If booFontBold Then shpObj.CellsSRC(3, 0, 2).FormulaU = 1
            End With
            intRowStart = intRowStart + 1
            'Call InitArrShapeID(NT)
            winObj.Select(winObj.Page.Shapes.item(1), 256)
        End If

        ' Insert Column Headers
        If booTitleColumns Then
            For i = 1 To UBound(ArrShapeID, 1)
                With shpsObj.ItemFromID(GetShapeId(i, 1 + intRowStart))
                    .DeleteSection(243)
                    If .Cells(UTC).Result("") <= intCountColSourse Then
                        .LinkToData(vsoDataRecordset.ID, 1, False)
                        .Characters.AddCustomFieldU("=" & .CellsSRC(243, i - 1, 2).Name, 0)
                        If booFontBold Then .CellsSRC(3, 0, 2).FormulaU = 1
                    End If
                End With
            Next
            intRowStart = intRowStart + 1
        End If

        ' Link table cells to external data
        For c = 1 To intEndCol
            Application.DoEvents()
            For r = intRowStart + 1 To intEndRow

                With shpsObj.ItemFromID(ArrShapeID(c, r))
                    .DeleteSection(243)
                    intCountCur = .Cells(UTR).Result("") - intRowStart '+ intRowStartSourse - 1
                    .LinkToData(vsoDataRecordset.ID, intCountCur, False)

                    .Characters.AddCustomFieldU("=" & .CellsSRC(243, c - 1, 0).Name, 0)
                    If booInvisibleZero Then
                        .Cells("Fields.Format").FormulaU = "=IF(" & .CellsSRC(243, c - 1, 0).Name & ".Type=2,IF(" & .CellsSRC(243, c - 1, 0).Name & "=0," & """#""" & ",FIELDPICTURE(0)),FIELDPICTURE(0))"
                    Else
                        .Cells("Fields.Format").FormulaU = "=" & "FIELDPICTURE(0)"
                    End If
                End With
            Next
        Next

        Call RecUndo("0")
        winObj.Select(shpsObj.ItemFromID(GetShapeId(1, 1)), 2)
        frm.Close()
        Exit Sub

ExitLine:
        Call RecUndo("0")
        frm.Close()
        winObj.Select(shpsObj.ItemFromID(GetShapeId(1, 1)), 2)
        MsgBox("Communication with external data failed.", 48, "Error!")

        'lngRowIDs = vsoDataRecordset.GetDataRowIDs("") ' Array of all strings
    End Sub

    ' Pin images to table cells
    Sub LockPicture(hAL As Byte, Val As Byte, shN As Boolean, lF As Boolean, msg As Boolean)
        ' hAL -horizontal alignment(1-3), vAL - vertical alignment(1-3)
        ' shN - place names (True, False), lF - block formulas (True, False)
        Dim vsoSel As Visio.Selection = winObj.Selection
        Dim shpObj As Visio.Shape, shpObj1 As Visio.Shape
        Dim strH As String = "", strV As String = "", strL As String = "", strL1 As String = ""
        Dim cnt As Integer, intDot As Integer
        Dim resvar As Byte

        shpsObj = winObj.Page.Shapes

        Select Case hAL
            Case 1
                strH = "Width*0"
            Case 2
                strH = "Width*0.5"
            Case 3
                strH = "Width*1"
        End Select

        Select Case Val
            Case 1
                strV = "Height*1"
            Case 2
                strV = "Height*0.5"
            Case 3
                strV = "Height*0"
        End Select

        Select Case lF
            Case True
                strL = "Guard(" : strL1 = ")"
            Case False
                strL = "" : strL1 = ""
        End Select

        Call RecUndo("Pin images")

        For Each shpObj In vsoSel
            If shpObj.CellExistsU(UTN, 0) Then
                For Each shpObj1 In shpsObj
                    If Not shpObj1.CellExistsU(UTN, 0) Then
                        resvar = shpObj1.SpatialRelation(shpObj, 0, 10)
                        If resvar = 4 Then
                            shpObj1.Cells("LocPinX").FormulaForceU = strH
                            shpObj1.Cells("LocPinY").FormulaForceU = strV

                            Select Case hAL
                                Case 1 'X слева
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX+" & shpObj.Name & "!LeftMargin-" & shpObj.Name & "!Width/2" & strL1
                                Case 2 'X по центру
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX" & strL1
                                Case 3 'X справа
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX-" & shpObj.Name & "!RightMargin+" & shpObj.Name & "!Width*0.5" & strL1
                            End Select

                            Select Case Val
                                Case 1 'Y сверху
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY-" & shpObj.Name & "!TopMargin+" & shpObj.Name & "!Height/2" & strL1
                                Case 2 'Y по центру
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY" & strL1
                                Case 3 'Y снизу
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY+" & shpObj.Name & "!BottomMargin-" & shpObj.Name & "!Height*0.5" & strL1
                            End Select

                            If shN Then
                                intDot = InStr(1, shpObj1.Name, ".")
                                If intDot <> 0 Then
                                    shpObj.Characters.Text = Left$(shpObj1.Name, intDot - 1)
                                Else
                                    shpObj.Characters.Text = shpObj1.Name
                                End If
                            End If
                            cnt = cnt + 1
                            Exit For
                        End If
                    End If
                Next
            End If
        Next

        Call RecUndo("0")

        ' Results of the SpatialRelation (resvar) method
        ' 4 - figure inside a figure
        ' 1 - figures overlap
        ' 8 - the figures touch
        ' 0 - the figures do not have equal points
        If msg Then MsgBox("Ready." & vbCrLf & "Pinned " & cnt & " figures in cells.")

    End Sub

    ' Saving data for Undo, Redo operations
    Sub RecUndo(index)
        If index <> "0" Then
            UndoScopeID = vsoApp.BeginUndoScope(index)
        Else
            vsoApp.EndUndoScope(UndoScopeID, True)
        End If
    End Sub

    ' Resizing table cells. Basic procedure
    Sub ResizeCells(bytCellsOrTable As Byte, booOnlyActiveCells As Boolean, _
        sngWidthCell As Single, sngHeightCell As Single, sngWidthTable As Single, _
        sngHeightTable As Single, booWidth As Boolean, booHeight As Boolean)


        Dim vsoSel As Visio.Selection, i As Integer

        shpsObj = winObj.Page.Shapes : vsoSel = winObj.Selection

        Call InitArrShapeID(NT)

        Call RecUndo("Dimensions")

        With shpsObj
            Select Case bytCellsOrTable
                Case 1
                    If booOnlyActiveCells Then
                        'Dim Shp As Visio.Shape
                        If booWidth Then ' Justified, selected columns
                            NotDub(vsoSel, UTC)
                            For i = 1 To NoDupes.Count
                                If .ItemFromID(ArrShapeID(NoDupes(i), 0)).NameU Like "ThC*" Then .ItemFromID(ArrShapeID(NoDupes(i), 0)).Cells(WI).Result(64) = sngWidthCell
                            Next
                            NoDupes.Clear()
                        End If
                        If booHeight Then ' By height, highlighted lines
                            NotDub(vsoSel, UTR)
                            For i = 1 To NoDupes.Count
                                If .ItemFromID(ArrShapeID(0, NoDupes(i))).NameU Like "TvR*" <> 0 Then .ItemFromID(ArrShapeID(0, NoDupes(i))).Cells(HE).Result(64) = sngHeightCell
                            Next
                            NoDupes.Clear()
                        End If
                    Else '----------------------------------------------------------------------------------
                        If booWidth Then ' Justified, all columns
                            For i = 1 To UBound(ArrShapeID, 1)
                                .ItemFromID(ArrShapeID(i, 0)).Cells(WI).Result(64) = sngWidthCell
                            Next
                        End If
                        If booHeight Then ' By height, all lines
                            For i = 1 To UBound(ArrShapeID, 2)
                                .ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) = sngHeightCell
                            Next
                        End If
                    End If
                Case 2
                    If booWidth Then ' Width, table
                        Dim factorW = sngWidthTable / fSTWH(winObj.Selection(1), 1, False)
                        For i = 1 To UBound(ArrShapeID, 1)
                            .ItemFromID(ArrShapeID(i, 0)).Cells(WI).Result(64) = .ItemFromID(ArrShapeID(i, 0)).Cells(WI).Result(64) * factorW
                        Next
                    End If
                    If booHeight Then ' By height, table
                        Dim factorH = sngHeightTable / fSTWH(winObj.Selection(1), 2, False)
                        For i = 1 To UBound(ArrShapeID, 2)
                            .ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) = .ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) * factorH
                        Next
                    End If
            End Select
        End With

        Call RecUndo("0")

    End Sub

    ' Selecting (miscellaneous) table cells
    Sub SelCell(arg As Byte, Optional booInit As Boolean = True)
        Dim vsoSel As Visio.Selection, intMaxC As Integer, intMaxR As Integer
        Dim iCount As Integer, Shp As visio.Shape

        vsoSel = winObj.Selection
        If booInit Then Call InitArrShapeID(NT)
        winObj.DeselectAll()

        intMaxC = UBound(ArrShapeID, 1) : intMaxR = UBound(ArrShapeID, 2)

        Select Case arg

            Case 1 ' Selecting a table from the user interface
                Call SelectCls(0, 0, intMaxC, intMaxR)

            Case 2 ' Selecting a table without UI
                Call SelectCls(1, 1, intMaxC, intMaxR)

            Case 3 ' Selecting a range of cells
                If vsoSel.Count < 2 Then
                    MsgBox("At least two cells must be selected.", 48, "Error!")
                    GoTo err
                Else
                    Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer
                    Call ClearControlCells(UTC) : Call ClearControlCells(UTR)
                    If Not GetMinMaxRange(vsoSel, cMin, cMax, rMin, rMax) Then GoTo err
                    If cMin = 0 Then cMin = 1
                    If rMin = 0 Then rMin = 1
                    Call SelectCls(cMin, rMin, cMax, rMax)
                End If

            Case 4 ' Column selection
                NotDub(vsoSel, UTC)
                For iCount = 1 To NoDupes.Count
                    Call SelectCls(NoDupes(iCount), 1, NoDupes(iCount), intMaxR)
                Next
                NoDupes.Clear() : Shp = Nothing

            Case 5 ' Selecting a line
                NotDub(vsoSel, UTR)
                For iCount = 1 To NoDupes.Count
                    Call SelectCls(1, NoDupes(iCount), intMaxC, NoDupes(iCount))
                Next
                NoDupes.Clear() : Shp = Nothing

            Case 6 ' Selecting columns
                Call SelectCls(1, 0, intMaxC, 0)

            Case 7 ' Selection of rows
                Call SelectCls(0, 1, 0, intMaxR)

        End Select

        Exit Sub

err:
    End Sub

    'Selection by blocks in the table according to specified values
    Sub SelectCls(StartCol, StartRow, EndCol, EndRow)
        Dim vsoSel As Visio.Selection = vsoApp.ActiveWindow.Selection

        On Error Resume Next

        With winObj.Page.Shapes
            For c = StartCol To EndCol
                For r = StartRow To EndRow
                    vsoSel.Select(.ItemFromID(ArrShapeID(c, r)), 2)
                Next
            Next
        End With

        vsoApp.ActiveWindow.Selection = vsoSel

        On Error GoTo 0
    End Sub

    ' Selecting table cells by criterion (text, date, value, empty/not empty). Basic procedure
    Sub SelInContent(arg)
        Dim vsoSel As visio.Selection = winObj.Selection
        Dim shpObj As visio.Shape = Nothing

        Dim SelShp = Sub() vsoSel.Select(shpObj, 1)

        If arg <> 8 Then
            For Each shpObj In vsoApp.ActiveWindow.Selection
                With shpObj
                    If .Cells(UTC).Result("") = 0 Or .Cells(UTR).Result("") = 0 Then SelShp()
                    Select Case arg
                        Case 1 'Text
                            If IsNumeric(.Characters.Text) Or
                            Trim(.Characters.Text) = "" Or
                            IsDate(.Characters.Text) Then SelShp()
                        Case 2 'Numbers
                            If Not IsNumeric(.Characters.Text) Then SelShp()
                        Case 3 'Dates
                            If Not IsDate(.Characters.Text) Or IsNumeric(.Characters.Text) Then SelShp()
                        Case 5 'Not numbers
                            If IsNumeric(.Characters.Text) Or
                            Trim(.Characters.Text) = "" Then SelShp()
                        Case 6 'Empty
                            If Trim(.Characters.Text) <> "" Then SelShp()
                        Case 7 'Not empty
                            If Trim(.Characters.Text) = "" Then SelShp()
                    End Select
                End With
            Next
            vsoApp.ActiveWindow.Selection = vsoSel
        Else 'Invert relative to table
            Call InitArrShapeID(NT)
            Call SelectCls(1, 1, UBound(ArrShapeID, 1), UBound(ArrShapeID, 2))

            For Each shpObj In vsoSel
                vsoApp.ActiveWindow.Select(shpObj, 1)
            Next
        End If

    End Sub

    ' Set the text of a table cell/cells by column and row number
    Sub SetText(arg As Object, intStartCol As Integer, intStartRow As Integer, intEndCol As Integer, intEndRow As Integer, byColOrRow As Byte)
        Dim bytParam As Short, iCount As Integer, jCount As Integer

        If IsArray(arg) Then
            Select Case arg.Rank
                Case 1 : bytParam = 1 : iCount = LBound(arg)
                Case 2 : bytParam = 2 : iCount = LBound(arg, 1) : jCount = LBound(arg, 2)
                Case Else : Exit Sub
            End Select
        Else
            bytParam = 0
            If Len(arg) = 0 Then Exit Sub
        End If

        Dim intColNum As Integer, intRowNum As Integer
        Call InitArrShapeID(NT)

        If intEndCol < intStartCol Then
            Dim x As Integer = intEndCol : intEndCol = intStartCol : intStartCol = x
        End If
        If intEndRow < intStartRow Then
            Dim y As Integer = intEndRow : intEndRow = intStartRow : intStartRow = y
        End If

        If intStartCol < 1 Then intStartCol = 0
        If intEndCol > UBound(ArrShapeID, 1) Then intEndCol = UBound(ArrShapeID, 1)
        If intStartRow < 1 Then intStartRow = 0
        If intEndRow > UBound(ArrShapeID, 2) Then intEndRow = UBound(ArrShapeID, 2)

        On Error Resume Next

        If byColOrRow = 0 Then

            For intColNum = intStartCol To intEndCol
                For intRowNum = intStartRow To intEndRow
                    If bytParam = 0 Then shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg)
                    If bytParam = 1 Then
                        shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg(iCount))
                        iCount += 1
                    End If
                    If bytParam = 2 Then
                        shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg(iCount, jCount))
                        If iCount = UBound(arg, 1) And jCount = UBound(arg, 2) Then GoTo Line1

                        If jCount = UBound(arg, 2) Then
                            jCount = 0 : iCount += 1
                        Else
                            jCount += 1
                        End If
                    End If
                Next
            Next

        ElseIf byColOrRow = 1 Then

            For intRowNum = intStartRow To intEndRow
                For intColNum = intStartCol To intEndCol
                    If bytParam = 0 Then shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg)
                    If bytParam = 1 Then
                        shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg(iCount))
                        iCount += 1
                    End If
                    If bytParam = 2 Then
                        shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Characters.Text = CStr(arg(iCount, jCount))
                        If iCount = UBound(arg, 1) And jCount = UBound(arg, 2) Then GoTo Line1

                        If jCount = UBound(arg, 2) Then
                            jCount = 0 : iCount += 1
                        Else
                            jCount += 1
                        End If
                    End If
                Next
            Next

        End If
Line1:
    End Sub

    ' Set the formula of a table cell/cells by column and row number
    Sub SetFormula(cell As String, intStartCol As Integer, intStartRow As Integer, intEndCol As Integer, intEndRow As Integer, txt As Object)
        Dim bytParam As Short, iCount As Integer, jCount As Integer

        If IsArray(txt) Then
            Select Case txt.Rank
                Case 1 : bytParam = 1 : iCount = LBound(txt)
                Case Else : Exit Sub
            End Select
        Else
            bytParam = 0
            If Len(txt) = 0 Then Exit Sub
        End If

        Dim intColNum As Integer, intRowNum As Integer
        Call InitArrShapeID(NT)

        On Error Resume Next

        For intColNum = intStartCol To intEndCol
            For intRowNum = intStartRow To intEndRow
                If bytParam = 0 Then shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Cells(cell).FormulaU = CStr(txt)
                If bytParam = 1 Then
                    shpsObj.ItemFromID(ArrShapeID(intColNum, intRowNum)).Cells(cell).FormulaU = CStr(txt(iCount))
                    iCount += 1
                End If
            Next
        Next

Line1:
    End Sub

    ' Sort selected table cells
    Sub SortTableData(NumColumn, DigOrTxt, SortDirection)
        Dim vsoSel As Visio.Selection = winObj.Selection
        If vsoSel.Count < 2 Then
            MsgBox("At least two cells must be selected! Without control cells.", 48, "Error!")
            Exit Sub
        End If
        '------------------------------- START --------------------------------------------------------
        Call RecUndo("Sort data")

        Dim arr, arrTMP
        Dim cS, rS, ci, ri As Integer
        Dim i, j, allN As Integer, TCol As Byte

        SelCell(3, True) ' selecting a range of cells

        cS = SelColRow(1) ' cS = SelColRow(1) ' number of selected columns
        rS = SelColRow(2) ' number of allocated lines
        ci = vsoSel.PrimaryItem.Cells(UTC).Result("") ' index(column) of the first selected cell
        ri = vsoSel.PrimaryItem.Cells(UTR).Result("") ' index(row) of the first selected cell
        Call GetCellsProperties(arr, ci, ri, ci + cS - 1, ri + rS - 1, "Text") ' reading text from selected cells into an array

        ReDim arrTMP(cS - 1, rS - 1) ' array override
        allN = -1

        On Error GoTo err
        ' filling a two-dimensional array with data from a one-dimensional one
        For i = 0 To UBound(arrTMP, 1)
            For j = 0 To UBound(arrTMP, 2)
                allN = allN + 1
                arrTMP(i, j) = arr(allN)
            Next
        Next

        ' getting target column number
        TCol = NumColumn - 1
        If TCol < 0 Or TCol > UBound(arrTMP, 1) Then TCol = 0

        ' bubble sorting
        Dim n As Integer, SortDir
        Dim Temp

        Select Case DigOrTxt ' choice: sort text or numbers
            Case False ' text sorting
                SortDir = Microsoft.VisualBasic.Switch(SortDirection = False, 1, SortDirection = True, -1)
                For i = 0 To UBound(arrTMP, 2)
                    For j = i + 1 To UBound(arrTMP, 2)
                        If StrComp(arrTMP(TCol, i), arrTMP(TCol, j), 1) = SortDir Then
                            ' bubble sort subroutine
                            ' sort target column
                            Temp = arrTMP(TCol, i)
                            arrTMP(TCol, i) = arrTMP(TCol, j)
                            arrTMP(TCol, j) = Temp

                            ' sort other columns according to target column
                            For n = 0 To UBound(arrTMP, 1)
                                If n <> TCol Then
                                    Temp = arrTMP(n, i)
                                    arrTMP(n, i) = arrTMP(n, j)
                                    arrTMP(n, j) = Temp
                                End If
                            Next
                        End If
                    Next
                Next
            Case True ' number sorting
                SortDir = Microsoft.VisualBasic.Switch(SortDirection = False, True, SortDirection = True, False)
                For i = 0 To UBound(arrTMP, 2)
                    For j = i + 1 To UBound(arrTMP, 2)
                        If CDbl(arrTMP(TCol, i)) > CDbl(arrTMP(TCol, j)) = SortDir Then
                            ' bubble sort subroutine
                            ' sort target column
                            Temp = arrTMP(TCol, i)
                            arrTMP(TCol, i) = arrTMP(TCol, j)
                            arrTMP(TCol, j) = Temp

                            ' sort other columns according to target column
                            For n = 0 To UBound(arrTMP, 1)
                                If n <> TCol Then
                                    Temp = arrTMP(n, i)
                                    arrTMP(n, i) = arrTMP(n, j)
                                    arrTMP(n, j) = Temp
                                End If
                            Next
                        End If
                    Next
                Next
        End Select

        Call SetText(arrTMP, ci, ri, ci + cS - 1, ri + rS - 1, 0)

        Call RecUndo("0")
        Exit Sub

err:
        MsgBox("Error number: " & Err.Number & vbNewLine & "Error description: " & Err.Description, 48, "Error!")
        Call RecUndo("0")
    End Sub

    ' Retrieving formula/values ​​of specified cells from the active table
    Sub GetCellsProperties(ByRef arr As Object, c As Integer, r As Integer, c1 As Integer, r1 As Integer, arg As String)
        Call InitArrShapeID(NT)
        Dim Coll As New Collection

        For i = c To c1
            For j = r To r1
                Select Case StrConv(arg, vbLowerCase)
                    Case "id", "0"
                        If ArrShapeID(i, j) <> 0 Then Coll.Add(ArrShapeID(i, j))
                    Case "name", "1"
                        If ArrShapeID(i, j) <> 0 Then Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).NameU)
                    Case "text", "2"
                        If ArrShapeID(i, j) <> 0 Then Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Characters.Text)
                    Case "comment", "3"
                        If ArrShapeID(i, j) <> 0 Then Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Cells("Comment").ResultStr(""))
                End Select
            Next
        Next

        If c = c1 And r = r1 Then
            arr = Coll.Item(1)
        Else
            ReDim arr(Coll.Count - 1)
            For i = 1 To Coll.Count
                arr(i - 1) = Coll.Item(i)
            Next
        End If

    End Sub

    ' Get the formula of a table cell/cells by column and row number
    Sub GetFormula(cell As String, c As Integer, r As Integer, c1 As Integer, r1 As Integer, ByRef arr As Object, res As String)
        Call InitArrShapeID(NT)
        Dim Coll As New Collection

        On Error Resume Next

        For i = c To c1
            For j = r To r1
                Select Case StrConv(res, vbLowerCase)
                    Case "number", "0"
                        Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Cells(cell).Result(""))
                    Case "string", "1"
                        Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Cells(cell).ResultStr(""))
                    Case "formula", "2"
                        Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Cells(cell).FormulaU)
                    Case Else
                        Coll.Add(shpsObj.ItemFromID(ArrShapeID(i, j)).Cells(cell).Result(CInt(res)))
                End Select
            Next
        Next

        If c = c1 And r = r1 Then
            arr = Coll.Item(1)
        Else
            ReDim arr(Coll.Count - 1)
            For i = 1 To Coll.Count
                arr(i - 1) = Coll.Item(i)
            Next
        End If

    End Sub

    ' Align/auto-align table cells to text width/height. Basic procedure
    Sub AlignOnText(ShNum As Integer, bytColumnOrRow As Byte, Optional bytNothingOrAutoOrLock As Byte = 0)
        Dim cellName As String = "", txt As String = "", txt1 As String, txt2 As String, lentxt As Integer
        Dim intCount As Integer, iC As Integer, iR As Integer

        Select Case bytColumnOrRow
            Case 4
                cellName = WI : txt = "MAX(TEXTWIDTH(" : txt1 = "!TheText),TEXTWIDTH(" : lentxt = Strings.Len(txt)
                For intCount = 1 To UBound(ArrShapeID, 2)
                    With shpsObj.ItemFromID(ArrShapeID(ShNum, intCount))
                        If InStr(1, .Cells(cellName).FormulaU, ",", 1) = 0 _
                            And ArrShapeID(ShNum, intCount) <> 0 Then txt = txt & .Name & txt1
                    End With
                Next
                iC = ShNum : iR = 0
            Case 5
                cellName = HE : txt = "MAX(TEXTHEIGHT(" : txt1 = "!TheText," : txt2 = "!Width),TEXTHEIGHT(" : lentxt = Strings.Len(txt)
                For intCount = 1 To UBound(ArrShapeID, 1)
                    With shpsObj.ItemFromID(ArrShapeID(intCount, ShNum))
                        If InStr(1, .Cells(cellName).FormulaU, ",", 1) = 0 _
                            And ArrShapeID(intCount, ShNum) <> 0 Then txt = txt & .Name & txt1 & .Name & txt2
                    End With
                Next
                iC = 0 : iR = ShNum
        End Select

        ' On Error Resume Next
        With shpsObj.ItemFromID(ArrShapeID(iC, iR))
            .Cells(cellName).FormulaForceU = Strings.Left(txt, Strings.Len(txt) - lentxt + 3) & ")"
            If bytNothingOrAutoOrLock = 0 Then .Cells(cellName).Result(64) = .Cells(cellName).Result(64)
            If bytNothingOrAutoOrLock = 2 Then .Cells(cellName).FormulaForceU = GU & Strings.Left(txt, Strings.Len(txt) - lentxt + 3) & "))"
        End With

    End Sub

    ' Finding text in cells
    Sub SearchText(Operand As String, Pattern As String, Action As String)
        Dim selObj As Visio.Selection = winObj.Selection, sh As Visio.Shape, booCheck As Boolean = False


        If StrConv(Action, vbLowerCase) = "select" OrElse StrConv(Action, vbLowerCase) = "0" Then winObj.DeselectAll()

        If StrConv(Action, vbLowerCase) = "clear text" OrElse StrConv(Action, vbLowerCase) = "2" Then Call RecUndo("Очистить текст")

        For Each sh In selObj
            With sh.Characters
                Select Case StrConv(Operand, vbLowerCase)
                    Case "equal", "=", "0" : If .Text = Pattern Then booCheck = True
                    Case "not equal", "<>", "1" : If .Text <> Pattern Then booCheck = True
                    Case "contains", "*", "2" : If .Text Like Pattern Then booCheck = True
                    Case "not contains", "!*", "3" : If Not .Text Like Pattern Then booCheck = True
                End Select
                Select Case StrConv(Action, vbLowerCase)
                    Case "select", "0" : If booCheck Then winObj.Select(sh, 2)
                    Case "deselect", "1" : If booCheck Then winObj.Select(sh, 1)
                    Case "clear text", "2" : If booCheck Then .Text = ""
                End Select
                booCheck = False
            End With
        Next

        If StrConv(Action, vbLowerCase) = "clear text" OrElse StrConv(Action, vbLowerCase) = "2" Then Call RecUndo("0")

    End Sub

    ' Replacing text in cells
    Sub ReplaceTxt(txt As String, txt1 As String, Optional istart As Integer = 1, Optional icount As Integer = -1)
        Dim selObj As Visio.Selection = winObj.Selection, sh As Visio.Shape

        Call RecUndo("Replacing text")

        On Error Resume Next

        For Each sh In selObj
            With sh.Characters
                .Text = Strings.Replace(.Text, txt, txt1, istart, icount)
            End With
        Next

        Call RecUndo("0")

    End Sub

#End Region

#Region "Private Sub"

    ' Deselecting columns or rows
    Public Sub ClearControlCells(arg)
        Dim shObj As visio.Shape

        With winObj
            For Each shObj In .Selection
                If shObj.Cells(arg).Result("") = 0 Then .Select(shObj, 1)
                If shObj.Name = NT Then .Select(shObj, 1)
            Next
        End With

    End Sub

    ' Merges selected cells into one while preserving the contents. Basic procedure
    Private Sub IntegrateCells()
        Dim vsoSel As Visio.Selection = winObj.Selection
        Dim shObj As Visio.Shape, flagCheck As Boolean
        flagCheck = True

        If vsoSel.Count < 2 Then
            MsgBox("Merging cells:" & vbNewLine & "At least two cells must be selected! Without control cells.", 48, "Error!")
            Exit Sub
        End If
        '------------------------------- START --------------------------------------------------------
        Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer, NText As String
        Dim Matr As Integer, i As Integer, fWn As String, fHn As String, fXn As String, fYn As String, shObj1 As Visio.Shape

        If Not GetMinMaxRange(vsoSel, cMin, cMax, rMin, rMax) Then GoTo err
        shObj1 = shpsObj.ItemFromID(GetShapeId(cMin, rMin))
        NText = shObj1.Characters.Text

        winObj.DeselectAll()
        Call SelectCls(cMin, rMin, cMax, rMax)
        vsoSel = winObj.Selection

        ' Start Checking for lice------------------------------------------------
        Matr = (cMax - cMin + 1) * (rMax - rMin + 1)
        If vsoSel.Count <> Matr Then flagCheck = False
        For i = 1 To vsoSel.Count
            If InStr(1, winObj.Selection(i).Cells(WI).FormulaU, "SUM", 1) <> 0 Or InStr(1, winObj.Selection(i).Cells(HE).FormulaU, "SUM", 1) <> 0 Then flagCheck = False
        Next
        If flagCheck = False Then GoTo err
        ' End Checking for lice -------------------------------------------------

        'winObj.Select(winObj.Page.Shapes.item(1), 256)

        'Start Generating and redefining formulas for a merged cell: PinX, PinY, Width, Height
        If cMax - cMin <> 0 Then
            fWn = "=GUARD(SUM("
            fXn = "=GUARD(Sheet." & ArrShapeID(cMin, 0) & "!PinX-(Sheet." & ArrShapeID(cMin, 0) & "!Width/2)+SUM("
            For i = cMin To cMax
                shObj = shpsObj.ItemFromID(ArrShapeID(i, 0))
                fWn = fWn & shObj.Name & "!Width,"
                fXn = fXn & shObj.Name & "!Width,"
            Next
            fWn = Left$(fWn, Len(fWn) - 1) & "))"
            fXn = Left$(fXn, Len(fXn) - 1) & ")/2)"
        Else
            fWn = GU & shpsObj.ItemFromID(ArrShapeID(cMin, 0)).Name & "!Width)"
            fXn = GU & shpsObj.ItemFromID(ArrShapeID(cMin, 0)).Name & "!PinX)"
        End If

        '---------------------------------------------------------------------------
        If rMax - rMin <> 0 Then
            fHn = "=GUARD(SUM("
            fYn = "=GUARD(Sheet." & ArrShapeID(0, rMin) & "!PinY+(Sheet." & ArrShapeID(0, rMin) & "!Height/2)-SUM("
            For i = rMin To rMax
                shObj = shpsObj.ItemFromID(ArrShapeID(0, i))
                fHn = fHn & shObj.Name & "!Height,"
                fYn = fYn & shObj.Name & "!Height,"
            Next
            fHn = Left$(fHn, Len(fHn) - 1) & "))"
            fYn = Left$(fYn, Len(fYn) - 1) & ")/2)"
        Else
            fHn = GU & shpsObj.ItemFromID(ArrShapeID(0, rMin)).Name & "!Height)"
            fYn = GU & shpsObj.ItemFromID(ArrShapeID(0, rMin)).Name & "!PinY)"
        End If

        '---------------------------------------------------------------------------
        With shObj1 ' overriding cell formulas
            .Cells(PX).FormulaForceU = fXn
            .Cells(PY).FormulaForceU = fYn
            .Cells(WI).FormulaForceU = fWn
            .Cells(HE).FormulaForceU = fHn
            .BringToFront()
            .Characters.Text = NText
        End With
        'End formula redefinition =============================================

        For i = 2 To vsoSel.Count ' Removing waste cells
            vsoSel(i).Cells(LD).FormulaForceU = 0
            vsoSel(i).Delete()
        Next

        winObj.Select(shObj1, 2)
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
        Exit Sub

err:
        Dim msg As String
        msg = "Possible causes of the error:" & vbCrLf
        msg = msg & "A cell that has already been merged is highlighted." & vbCrLf
        msg = msg & "Something went wrong." & vbCrLf
        MsgBox(msg, 48, "Error!")
    End Sub

    ' Unlink the selected cell while preserving the contents. Basic procedure
    Private Sub DeIntegrateCells(shObj As Visio.Shape)
        Dim flagCheck As Boolean, flagTxt As Boolean
        flagCheck = True

        If InStr(1, shObj.Cells(WI).FormulaU, "SUM", 1) = 0 And InStr(1, shObj.Cells(HE).FormulaU, "SUM", 1) = 0 Then flagCheck = False
        If Not flagCheck Then GoTo err

        '------------------------------- START --------------------------------------------------------
        Dim vsoDup As Visio.Shape
        Dim fx As String, fy As String, arrX() As String, arrY() As String, NText As String
        Dim j As Integer, i As Integer

        With shObj
            fx = .Cells(PX).FormulaU : fy = .Cells(PY).FormulaU
            NText = .Characters.Text
        End With

        '---------------------------------------------------------------
        If InStr(1, fx, "SUM", 1) <> 0 Then
            fx = Left$(fx, Len(fx) - 4)
            fx = Right$(fx, Len(fx) - InStr(1, fx, "+") - 4)
            fx = Replace$(fx, WI, "", 1)
            arrX = Split(fx, ",")
        Else
            fx = Replace$(fx, "GUARD(", "", 1)
            ReDim arrX(0)
            arrX(0) = Replace$(fx, "PinX)", "", 1)
        End If

        '---------------------------------------------------------------
        If InStr(1, fy, "SUM", 1) <> 0 Then
            fy = Left$(fy, Len(fy) - 4)
            fy = Right$(fy, Len(fy) - InStr(1, fy, "-") - 4)
            fy = Replace$(fy, HE, "", 1)
            arrY = Split(fy, ",")
        Else
            fy = Replace$(fy, "GUARD(", "", 1)
            ReDim arrY(0)
            arrY(0) = Replace$(fy, "PinY)", "", 1)
        End If

        '---------------------------------------------------------------
        flagTxt = True

        shObj.Characters.Text = NText

        For j = 0 To UBound(arrY)
            For i = 0 To UBound(arrX)
                vsoDup = shObj.Duplicate
                With vsoDup ' overriding cell formulas
                    .Cells(PX).FormulaForceU = GU & arrX(i) & "PinX)"
                    .Cells(PY).FormulaForceU = GU & arrY(j) & "PinY)"
                    .Cells(WI).FormulaForceU = GU & arrX(i) & "Width)"
                    .Cells(HE).FormulaForceU = GU & arrY(j) & "Height)"
                    .Cells(UTN).FormulaForceU = shObj.Cells(UTN).FormulaU
                    .Cells(UTC).FormulaForceU = GU & arrX(i) & "User.TableCol)"
                    .Cells(UTR).FormulaForceU = GU & arrY(j) & "User.TableRow)"
                    If j <> 0 Or i <> 0 Then .Cells("Comment").FormulaForceU = "Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"

                    If flagTxt Then
                        .Characters.Text = NText
                        flagTxt = False
                    Else
                        .Characters.Text = ""
                    End If
                End With

            Next
        Next
        shObj.Cells(LD).FormulaForceU = 0
        shObj.Delete()
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
        Exit Sub

err:
        Dim msg As String
        msg = "Possible causes of the error:" & vbCrLf & vbCrLf
        msg = msg & "1. More than one cell is selected" & vbCrLf
        msg = msg & "2. Unmerged cell selected" & vbCrLf
        MsgBox(msg, 48, "Error!")

    End Sub

    ' Delete a column. Basic procedure
    Private Sub DeleteColumn(shObj)
        If shObj.Cells(UTC).Result("") = 0 Or shObj.Cells(UTN).FormulaU = "GUARD(NAME(0))" Then Exit Sub
        Call RecUndo("Удалить столбец")

        Dim iAll As Integer = shpsObj.Item(NT).Cells(UTC).Result("")
        Dim iDel As Integer = shObj.Cells(UTC).Result("")
        Dim NTDel As String = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Name
        Dim i As Integer, j As Integer
        Dim strF As String, tmpName As String = "", PropC(1) As String

        If iDel < iAll Then ' Saving the properties of the deleted control. cells
            PropC(0) = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Cells(PX).FormulaU
            PropC(1) = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Cells(PY).FormulaU
        End If

        Call PropLayers(1)

        If iDel <> iAll Then tmpName = shpsObj.ItemFromID(ArrShapeID(iDel + 1, 0)).Name

        With shpsObj ' Defining merged cells And processing them
            For j = 1 To UBound(ArrShapeID, 1)
                For i = 1 To UBound(ArrShapeID, 2)
                    With .ItemFromID(ArrShapeID(j, i))
                        If InStr(1, .Cells(WI).FormulaU, "SUM") <> 0 Then
                            If InStr(1, .Cells(WI).FormulaU, NTDel) <> 0 Then 'AndAlso InStr(1, .Cells(WI).FormulaU, ",") <> 0 Then
                                strF = Replace$(.Cells(WI).FormulaU, NTDel & "!Width", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(WI).FormulaForceU = Replace$(strF, ",,", ",", 1)
                                '----------------------------------------------------------------------------------------------------------------
                                strF = .Cells(PX).FormulaU
                                If iDel <> iAll Then
                                    strF = Replace$(.Cells(PX).FormulaU, NTDel & "!PinX", tmpName & "!PinX", 1)
                                    strF = Replace$(strF, NTDel & "!Width/2", tmpName & "!Width/2", 1)
                                    .Cells(UTC).FormulaForceU = Replace$(.Cells(UTC).FormulaU, NTDel & "!", tmpName & "!", 1)
                                End If
                                strF = Replace$(strF, NTDel & "!Width", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(PX).FormulaForceU = Replace$(strF, ",,", ",", 1)
                            End If
                            If .Cells(WI).Result(64) = shpsObj.ItemFromID(ArrShapeID(j, 0)).Cells(WI).Result(64) _
                                AndAlso .Cells(UTC).Result("") = iDel + 1 Then
                                .Cells(WI).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j + 1, 0)).Name & "!Width)"
                                .Cells(PX).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j + 1, 0)).Name & "!PinX)"
                            End If
                            If .Cells(WI).Result(64) = shpsObj.ItemFromID(ArrShapeID(j, 0)).Cells(WI).Result(64) _
                                AndAlso .Cells(UTC).Result("") <> iDel + 1 Then
                                .Cells(WI).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j, 0)).Name & "!Width)"
                                .Cells(PX).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j, 0)).Name & "!PinX)"
                            End If
                        End If
                    End With
                Next
            Next

            For i = LBound(ArrShapeID, 2) To UBound(ArrShapeID, 2) 'Removing selected cells based on criteria
                With .ItemFromID(ArrShapeID(iDel, i))
                    If ArrShapeID(iDel, i) <> 0 Then
                        If .Cells(UTC).Result("") = iDel Then
                            If InStr(1, .Cells(WI).FormulaU, "SUM", 1) = 0 Then
                                .Cells(LD).FormulaForceU = 0
                                .Delete()
                                ArrShapeID(iDel, i) = 0
                            End If
                        End If
                    End If
                End With
            Next
        End With

        shpsObj.Item(NT).Cells(UTC).FormulaForceU = "GUARD(" & iAll - 1 & ")"

        If iDel < iAll Then
            With shpsObj.ItemFromID(ArrShapeID(iDel + 1, 0))
                .Cells(PX).FormulaForceU = PropC(0)
                .Cells(PY).FormulaForceU = PropC(1)
            End With

            With shpsObj 'Renumbering Columns
                j = 0
                For i = 1 To UBound(ArrShapeID, 1)
                    If ArrShapeID(i, 0) <> 0 Then
                        j = j + 1
                        .ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & j & ")"
                    End If
                Next
            End With
        End If

        Erase PropC
        Call PropLayers(0)
        winObj.Select(shpsObj.ItemFromID(ArrShapeID(0, 0)), 2)
        Call RecUndo("0")
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
    End Sub

    ' Deleting a line. Basic procedure
    Private Sub DeleteRow(shObj)
        If shObj.Cells(UTR).Result("") = 0 Or shObj.Cells(UTN).FormulaU = "GUARD(NAME(0))" Then Exit Sub
        Call RecUndo("Удалить строку")

        Dim iAll As Integer = shpsObj.Item(NT).Cells(UTR).Result("")
        Dim iDel As Integer = shObj.Cells(UTR).Result("")
        Dim NTDel As String = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Name
        Dim i As Integer, j As Integer
        Dim strF As String, tmpName As String = "", PropC(1) As String

        If iDel < iAll Then ' Saving the properties of the deleted control. cells
            PropC(0) = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Cells(PX).FormulaU
            PropC(1) = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Cells(PY).FormulaU
        End If

        Call PropLayers(1)

        If iDel <> iAll Then tmpName = shpsObj.ItemFromID(ArrShapeID(0, iDel + 1)).Name

        With shpsObj ' Defining merged cells and processing them
            For j = 1 To UBound(ArrShapeID, 1)
                For i = 1 To UBound(ArrShapeID, 2)
                    With .ItemFromID(ArrShapeID(j, i))
                        If InStr(1, .Cells(HE).FormulaU, "SUM") <> 0 Then
                            If InStr(1, .Cells(HE).FormulaU, NTDel) <> 0 Then 'AndAlso InStr(1, .Cells(HE).FormulaU, ",") <> 0 Then
                                strF = Replace$(.Cells(HE).FormulaU, NTDel & "!Height", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(HE).FormulaForceU = Replace$(strF, ",,", ",", 1)
                                '-------------------------------------------------------------------------------------------------------------
                                strF = .Cells(PY).FormulaU
                                If iDel <> iAll Then
                                    strF = Replace$(.Cells(PY).FormulaU, NTDel & "!PinY", tmpName & "!PinY", 1)
                                    strF = Replace$(strF, NTDel & "!Height/2", tmpName & "!Height/2", 1)
                                    .Cells(UTR).FormulaForceU = Replace$(.Cells(UTR).FormulaU, NTDel & "!", tmpName & "!", 1)
                                End If
                                strF = Replace$(strF, NTDel & "!Height", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(PY).FormulaForceU = Replace$(strF, ",,", ",", 1)
                            End If
                            If .Cells(HE).Result(64) = shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) _
                                AndAlso .Cells(UTR).Result("") = iDel + 1 Then
                                .Cells(HE).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i + 1)).Name & "!Height)"
                                .Cells(PY).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i + 1)).Name & "!PinY)"
                            End If
                            If .Cells(HE).Result(64) = shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) _
                                AndAlso .Cells(UTR).Result("") <> iDel + 1 Then
                                .Cells(HE).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Name & "!Height)"
                                .Cells(PY).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Name & "!PinY)"
                            End If
                        End If
                    End With
                Next
            Next

            For i = LBound(ArrShapeID, 1) To UBound(ArrShapeID, 1) 'Removing selected cells based on criteria
                With .ItemFromID(ArrShapeID(i, iDel))
                    If ArrShapeID(i, iDel) <> 0 Then
                        If .Cells(UTR).Result("") = iDel Then
                            If InStr(1, .Cells(HE).FormulaU, "SUM", 1) = 0 Then
                                .Cells(LD).FormulaForceU = 0
                                .Delete()
                                ArrShapeID(i, iDel) = 0
                            End If
                        End If
                    End If
                End With
            Next
        End With

        shpsObj.Item(NT).Cells(UTR).FormulaForceU = "GUARD(" & iAll - 1 & ")"

        If iDel < iAll Then
            With shpsObj.ItemFromID(ArrShapeID(0, iDel + 1))
                .Cells(PX).FormulaForceU = PropC(0)
                .Cells(PY).FormulaForceU = PropC(1)
            End With

            With shpsObj ' Renumbering lines
                j = 0
                For i = 1 To UBound(ArrShapeID, 2)
                    If ArrShapeID(0, i) <> 0 Then
                        j = j + 1
                        .ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & j & ")"
                    End If
                Next
            End With
        End If

        Erase PropC
        Call PropLayers(0)
        winObj.Select(shpsObj.ItemFromID(ArrShapeID(0, 0)), 2)
        Call RecUndo("0")
        CheckArrID = winObj.Page.Shapes(NT).UniqueID(0) & "1"
    End Sub

    ' Populating a collection with values ​​without duplicates
    Private Sub NotDub(vsoSel, UT)
        Dim Shp As Visio.Shape

        On Error Resume Next

        For Each Shp In vsoSel
            NoDupes.Add(Shp.Cells(UT).Result(""), CStr(Shp.Cells(UT).Result("")))
        Next

        On Error GoTo 0

    End Sub

    ' Enabling/disabling visibility and blocking of layers during code execution - Titles_Tables and Cells_Tables
    Public Sub PropLayers(arg As Byte)
        With winObj.Page.Layers
            Select Case arg
                Case 1
                    LayerVisible = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaU
                    '            LayerVisible1 = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0) + 2).CellsC(4).FormulaU
                    LayerLock = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(7).FormulaU
                    '            LayerLock1 = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0) + 2).CellsC(7).FormulaU
                    If LayerVisible <> 1 Then .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaForceU = "1"
                    If LayerLock <> 0 Then .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(7).FormulaForceU = "0"
                Case 0
                    .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaForceU = LayerVisible
                    .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(7).FormulaForceU = LayerLock
            End Select
        End With

    End Sub

#End Region


#Region "Functions"

    ' Data Formatting Features
    Function PtoD(arg)
        Return vsoApp.FormatResult(arg, 50, 64, "####0.####")
    End Function

    Function DtoD(arg)
        Return vsoApp.FormatResult(arg, 64, 64, "####0.####")
    End Function

    Function DtoP(arg)
        Return vsoApp.FormatResult(arg, 64, 50, "####0.####")
    End Function

    Function ItoD(arg)
        Return vsoApp.FormatResult(arg, 65, 64, "####0.####")
    End Function

    ' function for counting the number of allocated rows/columns
    Function SelColRow(arg)
        Dim col As New Collection, num As Integer

        On Error Resume Next

        If arg = 1 Then
            For i = 1 To winObj.Selection.Count
                num = winObj.Selection(i).Cells(UTC).Result("")
                col.Add(num, CStr(num))
            Next
        ElseIf arg = 2 Then
            For i = 1 To winObj.Selection.Count
                num = winObj.Selection(i).Cells(UTR).Result("")
                col.Add(num, CStr(num))
            Next
        End If

        Return col.Count
    End Function

    ' Message about missing/incorrect selection on the sheet
    Function CheckSelCells() As Boolean

        With winObj
            If .Selection.Count = 0 Then GoTo ErrMsg

            Dim shObj As visio.Shape

            For Each shObj In .Selection
                If Not shObj.CellExistsU(UTN, 0) Then .Select(shObj, 1)
            Next
            If .Selection.Count = 0 Then GoTo ErrMsg

            NT = .Selection(1).Cells(UTN).ResultStr("")

            For Each shObj In .Selection
                If StrComp(shObj.Cells(UTN).ResultStr(""), NT) <> 0 Then _
                .Select(shObj, 1)
            Next
        End With

        Return True
ErrMsg:
        MsgBox("There are no selected cells in the table on the active sheet!" &
           vbCrLf & "No further work is possible." & vbCrLf &
           "You need to select a cell in the table and perform the operation again." & vbCrLf, 48, "Attention")
        Return False
    End Function

    ' Filling an array with data from table cells
    Private Function fArrT(txt)
        Dim i As Integer, j As Integer, arrId(,) As String, Response As Boolean
        Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer

        Call ClearControlCells(UTC) : Call ClearControlCells(UTR)

        Dim vsoSelection As Visio.Selection = winObj.Selection

        Response = GetMinMaxRange(vsoSelection, cMin, cMax, rMin, rMax)

        ReDim arrId(rMax + 1, cMax + 1)

        For i = 1 To vsoSelection.Count
            With vsoSelection(i)
                arrId(.Cells(UTR).Result(""), .Cells(UTC).Result("")) = .Characters.Text
            End With
        Next

        For j = rMin To rMax
            For i = cMin To cMax
                txt = IIf(i = cMax, txt & arrId(j, i) & vbCrLf, txt & arrId(j, i) & vbTab)
            Next
        Next

        Erase arrId
        fArrT = txt
    End Function

    'Function for calculating the size of the active table
    Function fSTWH(sh As Visio.Shape, strWidthOrHeight As Byte, booInit As Boolean)
        If booInit Then Call InitArrShapeID(sh.Cells(UTN).ResultStr(""))

        With winObj.Page.Shapes
            Select Case strWidthOrHeight
                Case 1
                    For i = 1 To UBound(ArrShapeID, strWidthOrHeight)
                        fSTWH = fSTWH + .ItemFromID(ArrShapeID(i, 0)).Cells(WI).Result(64)
                    Next
                Case 2
                    For i = 1 To UBound(ArrShapeID, strWidthOrHeight)
                        fSTWH = fSTWH + .ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64)
                    Next
            End Select
        End With

    End Function

    'Function for determining the minimum and maximum number of columns/rows among a selected range of cells
    Function GetMinMaxRange(ByVal vsoSel As Visio.Selection, ByRef cMin As Integer, ByRef cMax As Integer, ByRef rMin As Integer, ByRef rMax As Integer) As Boolean
        Dim i As Integer
        rMin = 1000 : cMin = 1000 : rMax = 0 : cMax = 0

        On Error GoTo err

        For i = 1 To vsoSel.Count
            With vsoSel(i)
                If rMin > .Cells(UTR).Result("") Then rMin = .Cells(UTR).Result("")
                If cMin > .Cells(UTC).Result("") Then cMin = .Cells(UTC).Result("")
                If rMax < .Cells(UTR).Result("") Then rMax = .Cells(UTR).Result("")
                If cMax < .Cells(UTC).Result("") Then cMax = .Cells(UTC).Result("")
            End With
        Next

        GetMinMaxRange = True
        Exit Function

err:
        GetMinMaxRange = False
    End Function

    ' Getting table cell ID by column and row number
    Function GetShapeId(ByVal intColNum As Integer, ByVal intRowNum As Integer) As Integer
        On Error GoTo err

        If ArrShapeID(intColNum, intRowNum) <> 0 Then
            GetShapeId = ArrShapeID(intColNum, intRowNum)
        Else
            Dim i As Integer, j As Integer, cN As String, rN As String
            With winObj.Page.Shapes
                cN = .ItemFromID(ArrShapeID(intColNum, 0)).Name : rN = .ItemFromID(ArrShapeID(0, intRowNum)).Name
                For i = 1 To intColNum
                    For j = 1 To intRowNum
                        If ArrShapeID(i, j) <> 0 Then
                            If InStr(1, .ItemFromID(ArrShapeID(i, j)).Cells(PX).FormulaU, cN) <> 0 And _
                               InStr(1, .ItemFromID(ArrShapeID(i, j)).Cells(PY).FormulaU, rN) <> 0 Then
                                GetShapeId = ArrShapeID(i, j)
                                Exit Function
                            End If
                        End If
                    Next
                Next
            End With
        End If
        Exit Function

err:
        GetShapeId = 0
    End Function

#End Region

End Module