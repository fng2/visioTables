Imports System.Windows.Forms
Imports System.Text

Public Class dlgIntellInput
    Dim bytColOrRow As Byte = 0
    Dim KeyArg As Integer
    Dim cT As Integer = winObj.Page.Shapes(NT).Cells(UTC).Result("")
    Dim rT As Integer = winObj.Page.Shapes(NT).Cells(UTR).Result("")
    Dim NoDupes As New Collection

    Private Sub dlgIntellInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call InitArrShapeID(NT)
        Call RecUndo("Manual typing")
    End Sub

    Private Sub cmbText_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbText.KeyDown
        KeyArg = e.KeyCode
        Select Case e.KeyCode
            Case Keys.F1 : bytColOrRow = 0 : InsertText() : JumpToNext()
            Case Keys.F2 : bytColOrRow = 1 : InsertText() : JumpToNext()
            Case Keys.F3, Keys.F5, Keys.F6, Keys.F7, Keys.F8 : JumpToNext()
            Case Keys.Enter : InsertText() : Me.Close()
            Case Keys.Escape : JumpToNext()
            Case Keys.F9 : DeleteText()
        End Select
    End Sub

    Private Sub InsertText()
        With winObj.Selection(1)
            .Characters.Text = cmbText.Text
        End With
        cmbText.SelectionStart = 0
        cmbText.SelectionLength = Len(cmbText.Text)
        Call SaveText()
    End Sub

    Private Sub DeleteText()
        winObj.Selection(1).Characters.Text = ""
    End Sub

    Private Sub JumpToNext()
        Dim c As Integer = winObj.Selection(1).Cells(UTC).Result("")
        Dim r As Integer = winObj.Selection(1).Cells(UTR).Result("")
        winObj.DeselectAll()

        Select Case KeyArg
            Case Keys.F1 : SelC(c + 1, r)
            Case Keys.F2 : SelC(c, r + 1)
            Case Keys.F3
                Select Case bytColOrRow
                    Case 0 : SelC(c + 1, r)
                    Case 1 : SelC(c, r + 1)
                End Select
            Case Keys.F5 : SelC(1, r)
            Case Keys.F6 : SelC(cT, r)
            Case Keys.F7 : SelC(c, 1)
            Case Keys.F8 : SelC(c, rT)
            Case Keys.Escape
                Select Case bytColOrRow
                    Case 0 : SelC(c - 1, r)
                    Case 1 : SelC(c, r - 1)
                End Select
        End Select

    End Sub

    Private Sub SelC(c, r)

        Select Case c + r
            Case Is > cT + rT : c = 1 : r = 1 : SelC(c, r)
            Case Is < 2 : c = cT : r = rT : SelC(c, r)
        End Select

        Select Case c
            Case Is > cT : c = 1 : r += 1 : SelC(c, r)
            Case Is < 1 : c = cT : r -= 1 : SelC(c, r)
        End Select

        Select Case r
            Case Is > rT : c += 1 : r = 1 : SelC(c, r)
            Case Is < 1 : c -= 1 : r = rT : SelC(c, r)
        End Select

        If ArrShapeID(c, r) = 0 Then
            Select Case KeyArg
                Case Keys.F1 : c += 1 : SelC(c, r)
                Case Keys.F2 : r += 1 : SelC(c, r)
                Case Keys.Escape
                    Select Case bytColOrRow
                        Case 0 : c -= 1 : SelC(c, r)
                        Case 1 : r -= 1 : SelC(c, r)
                    End Select
            End Select
        End If

        SelectCls(c, r, c, r)

    End Sub

    Private Sub SaveText()
        On Error GoTo err
        NoDupes.Add(cmbText.Text, CStr(cmbText.Text))
        cmbText.Items.Add(NoDupes(NoDupes.Count))
err:
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Dim msg As String
        msg = "Navigation keys:" & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "F1 - insert text and move RIGHT along the line." & vbCrLf
        msg = msg & "F2 - insert text and move DOWN the column." & vbCrLf
        msg = msg & "F3 - move to the NEXT cell. (skip the current cell)." & vbCrLf
        msg = msg & "F4 - open a list to select saved values." & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "F5 - go to the beginning of the line." & vbCrLf
        msg = msg & "F6 - go to the end of the line." & vbCrLf
        msg = msg & "F7 - go to the beginning of the column." & vbCrLf
        msg = msg & "F8 - go to the end of the column." & vbCrLf
        msg = msg & "F9 - remove text from a cell." & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "ENTER - paste text into the selected cell and close the dialog." & vbCrLf
        msg = msg & "ESC - go back in reverse order."
        MsgBox(msg, 64, "Clue")
    End Sub

    Private Sub dlgIntellInput_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Call RecUndo("0")
    End Sub

End Class