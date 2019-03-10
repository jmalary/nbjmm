Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization


Public Class frmDownload
    Private Sub frmDownload_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 0 To 2000
            label1.Text = i.ToString()
            Application.DoEvents()
        Next

    End Sub
End Class