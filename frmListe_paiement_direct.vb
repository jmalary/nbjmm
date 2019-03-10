Imports Comptable_NBJMM.GlobalVariables


Public Class frmListe_paiement_direct

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.test2ConnectionString



    Private Sub frmListe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_direct", myConnection)
        Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

        While myReader.Read
            Dim newListViewItem As New ListViewItem
            newListViewItem.Text = myReader.GetInt32(0)

            newListViewItem.SubItems.Add(myReader.GetString(2))
            '********************************************************************************************** '*************************************************************************************

            Dim today = myReader.GetDateTime(3)

            Dim jour As String = ""

            Dim day = today.Day

            Dim month = today.Month

            Dim years = today.Year

            ' MsgBox(Date.Today.DayOfWeek)


            Dim Mois As String = ""

            Select Case Date.Today.DayOfWeek
                Case 0
                    jour = "Dimanche"
                Case 1
                    jour = "Lundi"
                Case 2
                    jour = "Mardi"
                Case 3
                    jour = "Mercredi"
                Case 4
                    jour = "Jeudi"
                Case 5
                    jour = "Vendredi"
                Case 6
                    jour = "Samedi"

            End Select

            Select Case month
                Case 1
                    Mois = "janvier"
                Case 2
                    Mois = "février"
                Case 3
                    Mois = "mars"
                Case 4
                    Mois = "avril"
                Case 5
                    Mois = "mai"
                Case 6
                    Mois = "juin"
                Case 7
                    Mois = "juillet"
                Case 8
                    Mois = "aout"
                Case 9
                    Mois = "septembre"
                Case 10
                    Mois = "octobre"
                Case 11
                    Mois = "novembre"
                Case 12
                    Mois = "decembre"

            End Select

            Dim courant = String.Concat(jour, " ", day, " ", Mois, " ", years)

            newListViewItem.SubItems.Add(courant)

            '*************************************************************************************

            today = myReader.GetDateTime(4)



            day = today.Day

            month = today.Month

            years = today.Year

            ' MsgBox(Date.Today.DayOfWeek)




            Select Case Date.Today.DayOfWeek
                Case 0
                    jour = "Dimanche"
                Case 1
                    jour = "Lundi"
                Case 2
                    jour = "Mardi"
                Case 3
                    jour = "Mercredi"
                Case 4
                    jour = "Jeudi"
                Case 5
                    jour = "Vendredi"
                Case 6
                    jour = "Samedi"

            End Select

            Select Case month
                Case 1
                    Mois = "janvier"
                Case 2
                    Mois = "février"
                Case 3
                    Mois = "mars"
                Case 4
                    Mois = "avril"
                Case 5
                    Mois = "mai"
                Case 6
                    Mois = "juin"
                Case 7
                    Mois = "juillet"
                Case 8
                    Mois = "aout"
                Case 9
                    Mois = "septembre"
                Case 10
                    Mois = "octobre"
                Case 11
                    Mois = "novembre"
                Case 12
                    Mois = "decembre"

            End Select

            courant = String.Concat(jour, " ", day, " ", Mois, " ", years)

            newListViewItem.SubItems.Add(courant)



            '***********************************************************************************************
            'newListViewItem.SubItems.Add(myReader.GetDateTime(2))

            'nwListViewItem.SubItems.Add(myReader.GetDateTime(3))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(6), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(7), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(8), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(9), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(10), "Currency"))

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
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")


        End While

    End Sub

    Sub proc_B()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label3.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While

    End Sub

    Sub proc_C()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label4.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_D()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label5.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_E()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label7.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'frmMenu.Show()
        Me.Hide()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (ListView1.SelectedItems.Count > 0) Then
            For Each i As ListViewItem In ListView1.SelectedItems
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_direct WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@numero", i.Text)

                myCommand.ExecuteNonQuery()
                myConnection.Close()

                ListView1.Items.Remove(i)
            Next
        Else
            MessageBox.Show("please select an item to delete")
        End If
    End Sub

    Private Sub frmListe_paiement_direct_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
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