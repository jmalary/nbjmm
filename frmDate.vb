Imports Comptable_NBJMM.GlobalVariables
Public Class frmDate

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        MsgBox(DateTimePicker1.Value.ToString)

        Me.Close()

        Me.WindowState = FormWindowState.Maximized
        frmMenu.MdiParent = Form_windows

        'frmMenu.Show()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        'MsgBox(DateTimePicker1.Value)


        DateCourant = DateTimePicker1.Value
        changer_date = True

        Me.Close()

        Me.WindowState = FormWindowState.Maximized
        'frmMenu.MdiParent = Form_windows

        'frmMenu.Show()

    End Sub

    Private Sub frmDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub
End Class