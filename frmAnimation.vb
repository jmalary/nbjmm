Public Class frmAnimation
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Label1.Text = Label1.Text
        If Len(Label1.Text) <> 0 Then
            Label1.Text = Microsoft.VisualBasic.Right(Label1.Text, (Len(Label1.Text) - 1))
        Else
            Timer1.Enabled = False
        End If

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub frmAnimation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class