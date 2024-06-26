﻿Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports visio = Microsoft.Office.Interop.Visio

Friend Class VisioTable

    'on a 2x2 table newshape is called 4 times (once for each "cell") and drawcells is called 9 times (once for each "cell" and once for each "header" item)

#Region "List Of Fields"
    Private strNameTable As String
    Private bytInsertType As Byte
    Private intColumnsCount As Integer
    Private intRowsCount As Integer
    Private sngWidthCells As Single
    Private sngHeightCells As Single
    Private sngWidthTable As Single
    Private sngHeightTable As Single
    Private booDeleteTargetShape As Boolean
    Private booVisibleProgressBar As Boolean
#End Region

#Region "List Of Variables"
    Private vsoApp As visio.Application = Globals.ThisAddIn.Application
    Private winObj As visio.Window = vsoApp.ActiveWindow
    Private pagObj As visio.Page = vsoApp.ActivePage
    Private shpObj As visio.Shape

    Private shape_TbL As visio.Shape
    Private shape_ThC As visio.Shape
    Private shape_TvR As visio.Shape
    Private shape_ClW As visio.Shape

    Private sngTW As Double = 0
    Private sngTH As Double = 0
    Private sngTX As Double = 0
    Private sngTY As Double = 0

    Private arrNewID() As Integer
    Private CountID As Integer = 0

    'Dim Vr As String = Strings.Left(Strings.Replace(vsoApp.Version, ".", ",", 1), InStr(vsoApp.Version, ",") - 1) 'error length must be zero or greater
    Dim Vr As String = Left(vsoApp.Version, InStr(vsoApp.Version, ".") - 1)
    Dim arrL = {"0.1 pt", "0", "1", "0 mm", "0", "0", "0", "0", "0", "0%"}
    Dim arrF = {"1", "0", "1", "0", "1", "0", "0%", "0%", "0%", "0%", "0", "0 mm", "0 mm", "0 deg", "100%"}
    Dim VarCell As Byte = 0

    'these next 6 are in constants, get rid of them?
    Private Const PX = "PinX"
    Private Const PY = "PinY"
    Private Const WI = "Width"
    Private Const HE = "Height"
    Private Const LD = "LockDelete"
    Private Const GU = "=GUARD("

    Private Const CA = "Angle"
    Private Const strATC = "!Actions.Titles.Checked=1,"
    Private Const strACC = "!Actions.Comments.Checked=1,"
    Private Const strThGu000 = "GUARD(MSOTINT(RGB(0,0,0),50))"
    Private Const strThGu255 = "GUARD(RGB(255,255,255))"
    Private Const strThGu191 = "GUARD(RGB(191,191,191))"
    Private Const GS = "=GUARD(Sheet."
    Private Const GI = "=Guard(IF("
    Private Const sh = "Sheet."
    Private Const GU5 = "=GUARD(10 mm)" ' Remake on DrawUn
    Private Const P50 = "50%"
    Private Const GT = "GUARD(TRUE)"
    Private Const G1 = "Guard(1)"
#End Region

    Public Sub CreatTable(ByVal a As String, ByVal b As Byte, ByVal c As Integer, ByVal d As Integer, ByVal e As Single,
                   ByVal f As Single, ByVal g As Single, ByVal h As Single, ByVal i As Boolean, ByVal j As Boolean) 'Implements IVisioTable.CreatTable
        On Error GoTo errD
        Dim frm As New dlgWait
        Dim vsoLayerTitles As visio.Layer, vsoLayerCells As visio.Layer, MemSHID As Integer
        Dim TypeCell As String
        Dim jGT As Integer = 0
        Dim iGT As Integer = 0

        strNameTable = IIf(Trim(a) = "", "TbL", a)
        bytInsertType = IIf(b < 1 Or b > 4, 1, b)
        intColumnsCount = IIf(c < 1 Or c > 1000, 5, c)
        intRowsCount = IIf(d < 1 Or d > 1000, 5, d)
        sngWidthCells = IIf(e = 0 Or e < 0, PtoD(56.69291339), e)
        sngHeightCells = IIf(f = 0 Or f < 0, PtoD(28.34645669), f)
        sngWidthTable = IIf(g = 0 Or g < 0, PtoD(283.4646), g)
        sngHeightTable = IIf(h = 0 Or h < 0, PtoD(283.4646), h)
        booDeleteTargetShape = i
        booVisibleProgressBar = j

        winObj = vsoApp.ActiveWindow
        pagObj = vsoApp.ActivePage

        If bytInsertType = 2 Then
            Dim sngPW As Single, sngPH As Single, sngPLM As Single, sngPRM As Single, sngPTM As Single, sngPBM As Single
            With pagObj
                sngPW = .PageSheet.Cells("PageWidth").Result(64)
                sngPH = .PageSheet.Cells("PageHeight").Result(64)
                sngPLM = .PageSheet.Cells("PageLeftMargin").Result(64)
                sngPRM = .PageSheet.Cells("PageRightMargin").Result(64)
                sngPTM = .PageSheet.Cells("PageTopMargin").Result(64)
                sngPBM = .PageSheet.Cells("PageBottomMargin").Result(64)
                sngTW = (sngPW - sngPLM - sngPRM) / intColumnsCount
                sngTH = (sngPH - sngPTM - sngPBM) / intRowsCount
            End With
        End If

        If bytInsertType = 4 Then
            If winObj.Selection.Count = 0 Then
                MsgBox("You must choose one figure")
                Exit Sub
            Else
                With winObj.Selection(1)
                    .BoundingBox(3, sngTX, sngTY, sngTW, sngTH) ' L, B, R, T
                    MemSHID = .ID
                    sngTX = ItoD(sngTX)
                    sngTY = ItoD(sngTH)
                    sngTW = .Cells(WI).Result(64) / intColumnsCount
                    sngTH = .Cells(HE).Result(64) / intRowsCount
                End With
            End If
        End If

        ReDim arrNewID((intRowsCount * intColumnsCount) + (intRowsCount + intColumnsCount))
        CountID = -1

        If booVisibleProgressBar Then
            frm.Label1.Text = " " & vbCrLf & "Creating a new table"
            frm.Show() : frm.Refresh()
        End If

        Call RecUndo("Creating a table...")

        ' Adding and changing layer properties
        vsoLayerTitles = pagObj.Layers.Add("Titles_Tables")
        vsoLayerCells = pagObj.Layers.Add("Cells_Tables")
        If vsoLayerTitles.CellsC(4).Result("") = 0 Then vsoLayerTitles.CellsC(4).FormulaForceU = 1 ' Make the layer visible if necessary
        If vsoLayerTitles.CellsC(7).Result("") = 1 Then vsoLayerTitles.CellsC(7).FormulaForceU = 0 ' Unlock a layer if necessary
        If vsoLayerCells.CellsC(4).Result("") = 0 Then vsoLayerCells.CellsC(4).FormulaForceU = 1 ' Make the layer visible if necessary
        If vsoLayerCells.CellsC(7).Result("") = 1 Then vsoLayerCells.CellsC(7).FormulaForceU = 0 ' Unlock a layer if necessary
        vsoLayerTitles.CellsC(5).FormulaU = "GUARD(0)" ' Layer is always not printed

        vsoApp.ShowChanges = False

        TypeCell = strNameTable : VarCell = 3 'Inserting 1 cell
        NewShape(TypeCell)
        shpObj = shape_TbL
        DrawOfCells(iGT, jGT)
        vsoLayerTitles.Add(shpObj, 1)

        TypeCell = "ThC" : VarCell = 2 'Inserting 1 row of a table
        For iGT = 1 To intColumnsCount
            If iGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_ThC
            Else
                shpObj = shape_ThC.Duplicate
            End If
            DrawOfCells(iGT, jGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "TvR" : VarCell = 1 'Insert 1 table column
        For jGT = 1 To intRowsCount
            If jGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_TvR
            Else
                shpObj = shape_TvR.Duplicate
            End If
            DrawOfCells(iGT, jGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "ClW" : VarCell = 0 'Inserting work cells
        For jGT = 1 To intRowsCount
            If booVisibleProgressBar Then
                frm.lblProgressBar.Width = (300 / intRowsCount) * jGT : frm.lblProgressBar.Refresh() : Application.DoEvents()
            End If
            For iGT = 1 To intColumnsCount
                If jGT = 1 And iGT = 1 Then
                    NewShape(TypeCell)
                    shpObj = shape_ClW
                Else
                    shpObj = shape_ClW.Duplicate
                End If
                DrawOfCells(iGT, jGT)
                vsoLayerCells.Add(shpObj, 0)
            Next
        Next

        shpObj = pagObj.Shapes.ItemFromID(arrNewID(0))
        shpObj.Cells(UTC).FormulaU = GU & intColumnsCount & ")"
        shpObj.Cells(UTR).FormulaU = GU & intRowsCount & ")"

        For iGT = 0 To intColumnsCount + intRowsCount
            winObj.Page.Shapes.ItemFromID(arrNewID(iGT)).Cells("LockTextEdit").FormulaU = "Guard(1)"
        Next

        Call RecUndo("0")

        If booVisibleProgressBar Then
            frm.Close() : frm = Nothing
        End If


        If bytInsertType = 4 AndAlso booDeleteTargetShape Then
            If pagObj.Shapes.ItemFromID(MemSHID).Cells(LD).Result("") = 0 Then pagObj.Shapes.ItemFromID(MemSHID).DeleteEx(0)
        End If

        vsoApp.ShowChanges = True
        winObj.Select(shpObj, 258)
        Exit Sub
errD:
        MsgBox("CreatTable-Class" & vbNewLine & Err.Description)
    End Sub

    Private Sub DrawOfCells(ByVal iGT, ByVal jGT)
        On Error GoTo errD
        'User.TableCol and User.TableRow are in each "working" cell and give the cell's location in the table

        With pagObj
            CountID = CountID + 1
            arrNewID(CountID) = shpObj.ID
            With shpObj
                Select Case VarCell

                    Case 0 'Inserting work cells
                        .Cells(PX).FormulaForceU = GS & arrNewID(iGT) & "!PinX)"
                        .Cells(PY).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!PinY)"
                        .Cells(WI).FormulaForceU = GS & arrNewID(iGT) & "!Width)"
                        .Cells(HE).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!Height)"
                        .Cells(UTN).FormulaForceU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GS & arrNewID(iGT) & "!User.TableCol)"
                        .Cells(UTR).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!User.TableRow)"
                        'char(10) is linefeed
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"

                    Case 2 ' control line
                        .Cells(PX).FormulaForceU = GS & arrNewID(CountID - 1) & "!PinX+(Sheet." & arrNewID(CountID - 1) & "!Width/2)+(Width/2))"
                        .Cells(PY).FormulaForceU = GS & arrNewID(0) & "!PinY)"
                        .Cells(HE).FormulaForceU = GS & arrNewID(0) & "!Height)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(WI).Result(64) = sngTW
                        ElseIf bytInsertType = 3 Then
                            .Cells(WI).Result(64) = sngWidthTable / intColumnsCount
                        Else
                            .Cells(WI).Result(64) = sngWidthCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GU & iGT & ")"
                        .Cells(UTR).FormulaForceU = GU & 0 & ")"
                        .Characters.AddCustomFieldU(UTC, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTC & ")"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrNewID(0) & "!Actions.Titles.Checked))"
                        '.Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Управляющая ячейка столбца""" & "," & """""" & "))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Column control cell""" & "," & """""" & "))"

                    Case 1 ' control column
                        .Cells(PX).FormulaU = GS & arrNewID(0) & "!PinX)"
                        If jGT = 1 Then
                            .Cells(PY).FormulaU = GS & arrNewID(CountID - intColumnsCount - 1) & "!PinY-(Sheet." & arrNewID(CountID - intColumnsCount - 1) & "!Height/2)-(Height/2))"
                        Else
                            .Cells(PY).FormulaForceU = GS & arrNewID(CountID - 1) & "!PinY-(Sheet." & arrNewID(CountID - 1) & "!Height/2)-(Height/2))"
                        End If
                        .Cells(WI).FormulaU = GS & arrNewID(0) & "!Width)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(HE).Result(64) = sngTH
                        ElseIf bytInsertType = 3 Then
                            .Cells(HE).Result(64) = sngHeightTable / intRowsCount
                        Else
                            .Cells(HE).Result(64) = sngHeightCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaU = GU & 0 & ")"
                        .Cells(UTR).FormulaForceU = GU & jGT & ")"
                        .Characters.AddCustomFieldU(UTR, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTR & ")"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrNewID(0) & "!Actions.Titles.Checked))"
                        '.Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Управляющая ячейка строки""" & "," & """""" & "))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Row control cell""" & "," & """""" & "))"

                    Case 3 ' 1 GlavUpr ?
                        Const frm = "###0.0###"
                        '.Cells(WI).FormulaForceU = GU & PtoD(28.34645669) & ")"
                        '.Cells(HE).FormulaForceU = GU & PtoD(28.34645669) & ")"
                        .Cells(WI).FormulaForceU = GU & Str(vsoApp.FormatResult(10, "mm", "", frm)) & ")"
                        .Cells(HE).FormulaForceU = GU & Str(vsoApp.FormatResult(10, "mm", "", frm)) & ")"
                        If bytInsertType = 4 Then
                            .Cells(PX).Result(64) = sngTX + (.Cells(WI).Result(64) / 2) - .Cells(WI).Result(64)
                            .Cells(PY).Result(64) = sngTY - (.Cells(HE).Result(64) / 2) + .Cells(HE).Result(64)
                        Else
                            .Cells(PX).FormulaU = "=ThePage!PageLeftMargin-5 mm"
                            .Cells(PY).FormulaU = "=ThePage!PageHeight-ThePage!PageTopMargin+5 mm"
                        End If
                        .UpdateAlignmentBox()
                        .Cells(UTN).FormulaU = "=GUARD(Name(0))"
                End Select

                .Cells(CA).FormulaU = GU & "0 deg)"
            End With
        End With
        Exit Sub
errD:
        MsgBox("DrawOfCells" & vbNewLine & Err.Description)
    End Sub

    Private Sub NewShape(TypeCell)
        On Error GoTo errD
        ' Subprocedure for creating table shapes and setting them up
        Dim vsoShape As visio.Shape
        Dim AddSectionNum As Integer, intArrNum() As Integer, arrRowData
        vsoShape = winObj.Page.DrawRectangle(0, 0, 1, 1)

        With vsoShape
            .Name = TypeCell

            ' add User section for all cells
            AddSectionNum = 242
            'visSectionUser	242	Stores cells created and used by an external solution.
            intArrNum = {0, 1}
            'arrRowData = {{"TableName", "Name(0)", """Таблица"""},
            '              {"TableCol", """""", """Столбец"""},
            '              {"TableRow", """""", """Строка"""}}
            arrRowData = {{"TableName", "Name(0)", """Table"""},
                          {"TableCol", """""", """Column"""},
                          {"TableRow", """""", """Line"""}}
            AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum, True)

            .Cells("LocPinX").FormulaU = "Guard(Width*0.5)"
            .Cells("LocPinY").FormulaU = "Guard(Height*0.5)"
            .Cells("UpdateAlignBox").FormulaForceU = GT
            .Cells("LockDelete").FormulaU = G1
            .Cells("LockRotate").FormulaU = G1
            .CellsSRC(1, 16, 1).FormulaU = "char(169)&char(32)&char(82)&char(79)&char(77)&char(65)&char(78)&char(79)&char(86)&char(32)&char(86)&char(56)&char(46)&char(48)"

            Select Case TypeCell
                Case strNameTable, "ThC", "TvR"
                    ' Setting formats
                    Call FormatLFM(vsoShape, TypeCell)
                    .Cells("LockFormat").FormulaU = G1
                    .Cells("LockFromGroupFormat").FormulaU = G1
                    .Cells("LockThemeColors").FormulaU = G1
                    .Cells("LockThemeEffects").FormulaU = G1
                    ' Setting up Miscellaneous
                    .Cells("NoObjHandles").FormulaForceU = GT
                    .Cells("NonPrinting").FormulaForceU = GT
            End Select

            Select Case TypeCell
                Case "ClW" ' Work cell
                    shape_ClW = vsoShape

                Case "TvR" ' Work cell
                    AddSectionNum = 9 ' Add a Control section
                    'visSectionControls	9	Stores an object's control handles.
                    intArrNum = {0, 1, 2, 3, 6, 8} ' Make no less than zero
                    'arrRowData = {{"ControlHeight", "GUARD(Width*0)", "Height*0", "GUARD(Controls.ControlHeight)", "GUARD(Controls.ControlHeight.Y)", "False", """Изменение высоты ячейки"""}}
                    arrRowData = {{"ControlHeight", "GUARD(Width*0)", "Height*0", "GUARD(Controls.ControlHeight)", "GUARD(Controls.ControlHeight.Y)", "False", """Changing cell height"""}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .CellsSRC(10, 1, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 2, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 5, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    shape_TvR = vsoShape

                Case "ThC" ' Column ID
                    AddSectionNum = 9 ' Добавить Control секцию
                    'visSectionControls	9	Stores an object's control handles.
                    intArrNum = {0, 1, 2, 3, 6, 8} ' Сделать не меньше нуля
                    'arrRowData = {{"ControlWidth", "Width*1", "GUARD(Height)", "GUARD(Controls.ControlWidth)", "GUARD(Controls.ControlWidth.Y)", "False", """Изменение ширины ячейки"""}}
                    arrRowData = {{"ControlWidth", "Width*1", "GUARD(Height)", "GUARD(Controls.ControlWidth)", "GUARD(Controls.ControlWidth.Y)", "False", """Changing cell width"""}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .CellsSRC(10, 2, 0).FormulaU = "GUARD(Controls.ControlWidth)"
                    .CellsSRC(10, 3, 0).FormulaU = "GUARD(Controls.ControlWidth)"
                    shape_ThC = vsoShape

                Case strNameTable ' Home UYA
                    AddSectionNum = 240 'Add an Action section
                    'visSectionAction	240	Stores the actions that appear on the shortcut menu.
                    intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    'arrRowData = {{"Titles", "SETF(GetRef(Actions.Titles.Checked),NOT(Actions.Titles.Checked))", """П&оказывать заголовки""", """""", 5, 1, "FALSE", "TRUE"},
                    '    {"Comments", "SETF(GetRef(Actions.Comments.Checked),NOT(Actions.Comments.Checked))", """Показывать коммента&рии""", """""", 6, 1, "FALSE", "FALSE"},
                    '    {"FixingTable", "SETF(GetRef(Actions.FixingTable.Checked),NOT(Actions.FixingTable.Checked))", """Заф&иксировать таблицу""", """""", 7, 0, "FALSE", "FALSE"}}

                    arrRowData = {{"Titles", "SETF(GetRef(Actions.Titles.Checked),NOT(Actions.Titles.Checked))", """&Render headers""", """""", 5, 1, "FALSE", "TRUE"},
                        {"Comments", "SETF(GetRef(Actions.Comments.Checked),NOT(Actions.Comments.Checked))", """Show comments""", """""", 6, 1, "FALSE", "FALSE"},
                        {"FixingTable", "SETF(GetRef(Actions.FixingTable.Checked),NOT(Actions.FixingTable.Checked))", """Freeze the table""", """""", 7, 0, "FALSE", "FALSE"}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .Cells("LineColor").FormulaForceU = "GUARD(IF(Actions.Titles.Checked=1,RGB(191,191,191),RGB(255,255,255)))"
                    .Cells("FillForegnd").FormulaForceU = "GUARD(IF(Actions.Titles.Checked=1,MSOTINT(RGB(0,0,0),50),RGB(255,255,255)))"
                    .Cells("FillForegndTrans").FormulaForceU = "GUARD(IF(Actions.Titles.Checked=1,0%,50%))"

                    .Cells("LockMoveX").FormulaU = "GUARD(Actions.FixingTable.Checked)"
                    .Cells("LockMoveY").FormulaU = "GUARD(Actions.FixingTable.Checked)"
                    .Cells("Width").FormulaU = GU5
                    .Cells("Height").FormulaU = GU5
                    '.Cells("Comment").FormulaU = "GUARD(IF(Actions.Comments.Checked=1," & "User.TableName&CHAR(10)&" & """Основная управляющая ячейка"", """"))"
                    .Cells("Comment").FormulaU = "GUARD(IF(Actions.Comments.Checked=1," & "User.TableName&CHAR(10)&" & """Main control cell"", """"))"
                    Dim Ui = .UniqueID(1)
                    shape_TbL = vsoShape
            End Select
        End With
        vsoShape = Nothing
        Erase intArrNum : Erase arrRowData
        Exit Sub
errD:
        MsgBox("NewShape" & vbNewLine & Err.Description)
    End Sub

    Private Sub AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum, Optional DelS = False)
        On Error GoTo errD
        ' Sub-process of adding a given Section, the required number of lines in Section and setting these lines
        Dim intI As Byte, intJ As Byte

        With vsoShape
            Select Case Strings.Left(.NameU, 3)
                Case "ClW"
                    If Not .SectionExists(AddSectionNum, 0) Then .AddSection(AddSectionNum)
                Case Else
                    If .SectionExists(AddSectionNum, 0) AndAlso DelS Then .DeleteSection(AddSectionNum)
                    If Not .SectionExists(AddSectionNum, 0) Then .AddSection(AddSectionNum)
            End Select

            Select Case AddSectionNum
                Case 242
                    For intI = 0 To UBound(arrRowData)
                        .AddNamedRow(AddSectionNum, arrRowData(intI, 0), 0)
                        .cells("User." & arrRowData(intI, 0) & ".Value").FormulaU = arrRowData(intI, 1)
                        .cells("User." & arrRowData(intI, 0) & ".Prompt").FormulaU = arrRowData(intI, 2)
                    Next
                Case Else
                    For intI = 0 To UBound(arrRowData)
                        .AddRow(AddSectionNum, 0, 0)
                        .CellsSRC(AddSectionNum, intI, 0).RowNameU = arrRowData(intI, 0)
                        For intJ = 0 To UBound(intArrNum)
                            .CellsSRC(AddSectionNum, intI, intArrNum(intJ)).FormulaU = arrRowData(intI, intJ + 1)
                        Next
                    Next
            End Select
        End With
        Exit Sub
errD:
        MsgBox("AddSections" & vbNewLine & Err.Description)
    End Sub

    Private Sub FormatLFM(Shp, TC)
        'arrL = Array("LineWeight", "LineColor", "LinePattern", "Rounding", "EndArrowSize", "BeginArrow", "EndArrow", "LineCap", "BeginArowSize", "LineColorTrans")
        'arrF = Array("FillForegnd", "FillBkgnd", "FillPattern", "ShdwForegnd", "ShdwBkgnd", "ShdwPattern", "FillForegndTrans", "FillBkgndTrans", "ShdwForegndTrans", "ShdwBkgndTrans", "ShapeShdwType", "ShapeShdwOffsetX", "ShapeShdwOffsetY", "ShapeShdwObliqueAngle", "ShapeShdwScaleFactor")        

        On Error GoTo errD
        With Shp

            For i = 0 To 9 ' Line
                .CellsSRC(1, 2, i).FormulaForceU = GU & arrL(i) & ")"
            Next

            For i = 0 To 14 ' Fill
                .CellsSRC(1, 3, i).FormulaForceU = GU & arrF(i) & ")"
            Next

            For i = 0 To 3 ' Text block margins
                .CellsSRC(1, 11, i).FormulaForceU = GU & 0 & " pt)"
            Next

            ' Setting color, font size, text alignment in cells
            .Cells("Char.Color[1]").FormulaForceU = strThGu255
            .Cells("Char.Font[1]").FormulaForceU = GU & vsoApp.ActiveDocument.Fonts("Courier New").ID & ")"
            .Cells("Char.Size[1]").FormulaForceU = GU & 10 & " pt)"
            .Cells("VerticalAlign").FormulaForceU = G1
            .Cells("Para.HorzAlign[1]").FormulaForceU = G1

            Select Case TC ' Dependencies
                Case "ThC", "TvR"
                    .Cells("FillForegnd").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu000 & "," & strThGu255 & "))"
                    .Cells("FillForegndTrans").FormulaForceU = GI & sh & arrNewID(0) & strATC & "0%" & "," & "50%" & "))"
                    .Cells("LineColor").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu191 & "," & strThGu255 & "))"
            End Select

            If Vr = 15 Then ' Theme Properties section for Visio 2013
                For i = 0 To 7
                    .CellsSRC(1, 31, i).FormulaForceU = GU & 0 & ")"
                Next
            End If

        End With
        Exit Sub
errD:
        MsgBox("FormatLFM" & vbNewLine & Err.Description)
    End Sub

End Class
