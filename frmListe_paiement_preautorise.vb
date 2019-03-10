Imports Comptable_NBJMM.GlobalVariables

Public Class frmListe_paiement_preautorise

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.test2ConnectionString



    Private Sub frmListe_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_preautorises", myConnection)
        Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

        While myReader.Read
            Dim newListViewItem As New ListViewItem
            newListViewItem.Text = myReader.GetInt32(0)
            newListViewItem.SubItems.Add(myReader.GetString(1))
            newListViewItem.SubItems.Add(myReader.GetDateTime(2))

            newListViewItem.SubItems.Add(myReader.GetDateTime(3))
            'newListViewItem.SubItems.Add(Format$(myReader.GetInt32(4), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(4), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(5), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(6), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(7), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(8), "Currency"))
            ListView1.Items.Add(newListViewItem)
        End While



        proc_A()
        proc_B()
        proc_C()
        proc_D()
        proc_E()


        myConnection.Close()


    End Sub



    Sub proc_A()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")


        End While

    End Sub

    Sub proc_B()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label3.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While

    End Sub

    Sub proc_C()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label4.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_D()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label5.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_E()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label7.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'frmMenu.Show()
        Me.Hide()

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (ListView1.SelectedItems.Count > 0) Then
            For Each i As ListViewItem In ListView1.SelectedItems
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_preautorises WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@numero", i.Text)

                myCommand.ExecuteNonQuery()
                myConnection.Close()

                ListView1.Items.Remove(i)
            Next
        Else
            MessageBox.Show("please select an item to delete")
        End If
    End Sub

    Private Sub frmListe_paiement_preautorise_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        Dim x As Integer
        Dim y As Integer
        Dim r As Rectangle

        If Not Parent Is Nothing Then
            r = Parent.ClientRectangle
            x = r.Width - Me.Width + Parent.Left
            y = r.Height - Me.Height + Parent.Top
        Else
            r = Screen.PrimaryScreen.WorkingArea
            x = r.Width - Me.Width
            y = r.Height - Me.Height
        End If

        x = CInt(x / 2)
        y = CInt(y / 2)

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(x, y)
    End Sub
End Class