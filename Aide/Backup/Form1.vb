''' <summary>
''' Generic About Box Dialog demo form
''' </summary>
''' <remarks>
''' Jeff Atwood
''' http://www.codinghorror.com
''' </remarks>
Public Class Form1
    Inherits System.Windows.Forms.Form

    ''' <summary>
    ''' loads the default values from a new instance of the AboutBox form
    ''' not exactly efficient, but hey, it's demo code
    ''' </summary>
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim frmAbout As New AboutBox
        txtTitle.Text = frmAbout.AppTitle
        txtDescription.Text = frmAbout.AppDescription
        txtVersion.Text = frmAbout.AppVersion
        txtCopyright.Text = frmAbout.AppCopyright
        txtMoreInfo.Text = frmAbout.AppMoreInfo

        '-- add some demonstrations of the RichTextBox auto-URLs:
        txtMoreInfo.Text &= Environment.NewLine
        txtMoreInfo.Text &= Environment.NewLine
        txtMoreInfo.Text &= "http://www.codinghorror.com/blog/" & Environment.NewLine
        txtMoreInfo.Text &= "mailto:jatwood@codinghorror.com?subject=your+code+stinks" & Environment.NewLine
    End Sub

    ''' <summary>
    ''' creates a new instance of the AboutBox form, sets all properties, and displays it
    ''' </summary>
    Private Sub ShowAbout()
        With New AboutBox
            .AppTitle = txtTitle.Text
            .AppDescription = txtDescription.Text
            .AppVersion = txtVersion.Text
            .AppCopyright = txtCopyright.Text
            .AppMoreInfo = txtMoreInfo.Text
            .AppDetailsButton = chkDetails.Checked
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        ShowAbout()
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ShowAbout()
    End Sub
End Class