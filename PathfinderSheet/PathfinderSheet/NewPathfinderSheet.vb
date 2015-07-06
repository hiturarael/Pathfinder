Option Strict On

Public Class NewPathfinderSheet

    Private Sub FileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem.Click

    End Sub

    Private Sub StatsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatsToolStripMenuItem.Click
        DiceRoller.ShowDialog()
    End Sub

    Private Sub PointBuyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PointBuyToolStripMenuItem.Click

    End Sub

    Private Sub btnAddRace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddRace.Click
        AddRace.ShowDialog()
    End Sub

    Private Sub txtPlayerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlayerName.TextChanged
        txtCharacterLevel.Enabled = True
    End Sub

    Private Sub enableFields()

    End Sub

    Private Sub txtCharacterLevel_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCharacterLevel.LostFocus
        Dim result As Integer = MsgBox("Confirm Your Level - Once you click Okay then you will not be able to change it again.",
                                       MsgBoxStyle.OkCancel, "Confirm Level")
        If result = DialogResult.Cancel Then
            txtCharacterLevel.Text = ""

        ElseIf result = DialogResult.OK Then
            txtCharacterLevel.Enabled = False
            ddlExpSpeed.Enabled = True

            btnAddRace.Enabled = True
            btnAddClass.Enabled = True
        End If
    End Sub

    Private Sub ddlExpSpeed_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlExpSpeed.SelectedIndexChanged
        Dim level As String = txtCharacterLevel.Text
        Dim speed As String = CStr(ddlExpSpeed.SelectedItem)

        txtNextLevel.Text = XML.setToNextLevel(level, speed).ToString
        txtExperience.Text = XML.setCurrentExp(level, speed).ToString
    End Sub

    Private Sub CurrentDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CurrentDirectoryToolStripMenuItem.Click
        MessageBox.Show(CurDir())
    End Sub

    Private Sub NewPathfinderSheet_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub NewPathfinderSheet_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        XML.workbook.Close()
        XML.APP.Quit()
    End Sub
End Class
