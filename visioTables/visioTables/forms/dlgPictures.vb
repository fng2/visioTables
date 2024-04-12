Public Class dlgPictures

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Dim hAL As Byte = 1, Val As Byte = 2, shN As Boolean = False, lF As Boolean = False

        If optAlignCenterH.Checked Then hAL = 2 ' hAL - horizontal alignment(1-3)
        If optAlignRightH.Checked Then hAL = 3

        If optAlignTopV.Checked Then Val = 1 ' vAL - vertical alignment(1-3)
        If optAlignBottomV.Checked Then Val = 3

        If ckbInsertName.Checked Then shN = True ' shN - put names in a cell(0,1)

        If ckbLockFormulas.Checked Then lF = True ' lF - block formulas(True,False)
        Call LockPicture(hAL, Val, shN, lF, True)

        Me.Close()
    End Sub

    Private Sub dlgPictures_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me.ToolTip1
            .SetToolTip(ckbInsertName, "Inserting a shape name into the same table cell ")
            .SetToolTip(ckbLockFormulas, "Locking shapes (PinX and PinY formulas) in table cells")
        End With
    End Sub
End Class