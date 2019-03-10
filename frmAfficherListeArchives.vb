Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb


Public Class frmAfficherListeArchives



    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Dim myConnection As New OleDb.OleDbConnection(myConnString)
    Dim con As System.Data.OleDb.OleDbConnection
    Dim sqlStr As String
    Dim cmd As System.Data.OleDb.OleDbCommand

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()
        Me.WindowState = FormWindowState.Maximized
        'frmMenu.MdiParent = Form_windows
        'frmMenu.Show()
        Form_windows.ListeDesArchivesToolStripMenuItem.Enabled = True

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub frmAfficherListeArchives_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If backup = True Then
            Button3.Enabled = False
            Button5.Enabled = False
            Button2.Enabled = True
            ListView1.Enabled = False

        Else

            Button3.Enabled = True
            Button5.Enabled = True
            Button2.Enabled = False
            ListView1.Enabled = True

        End If




        ChargerLister()



    End Sub

    Sub ChargerLister()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM bidon", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            Dim value1 As String = "xxxxxxxxx"

            While myReader.Read
                Dim newListViewItem As New ListViewItem
                newListViewItem.Text = myReader.GetInt32(0)
                newListViewItem.SubItems.Add(myReader.GetString(1))





                '*************************************************************************************

                Dim today = myReader.GetDateTime(2)

                Dim jour As String = ""

                Dim day = today.Day

                Dim month = today.Month

                Dim years = today.Year

                ' MsgBox(Date.Today.DayOfWeek)


                Dim Mois As String = ""

                Select Case today.DayOfWeek
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



                '*************************************************************************************




                ListView1.Items.Add(newListViewItem)
            End While


            'proc_A()
            'proc_B()
            'proc_C()
            'proc_D()
            'proc_E()


            myConnection.Close()

            'proc_Vider()

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'Dim I As Integer
        If (ListView1.SelectedItems.Count > 0) Then
            Dim a = MsgBox("Vous etes certain de supprimer quelques entrée ?", MsgBoxStyle.YesNo, "Confirmation")
            If a = vbYes Then

                For Each i As ListViewItem In ListView1.SelectedItems
                    Try
                        Dim myConnection As New OleDb.OleDbConnection(myConnString)
                        myConnection.Open()

                        Dim myCommand As New OleDb.OleDbCommand("DELETE FROM bidon WHERE numero = @numero", myConnection)
                        myCommand.Parameters.AddWithValue("@numero", i.Text)
                        myCommand.ExecuteNonQuery()
                        myConnection.Close()
                        myCommand.Dispose()


                        ListView1.Items.Remove(i)

                    Catch ex As Exception

                    End Try

                Next



            End If

        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à supprimer")
        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (ListView1.SelectedItems.Count > 0) Then

            'Dim numero As String


            'frmUpdate_check.Label2.Text = ListView1.SelectedItems(0).SubItems(0).Text
            backup = True
            num_backup = ListView1.SelectedItems(0).SubItems(0).Text

            changer_date(num_backup)
            MessageBox.Show("backup a été activé")
            frmInterface.Label2.Text = "Backup.."
            frmInterface.Label2.ForeColor = Color.Red
            frmInterface.ToolStripButton8.Enabled = True
            frmInterface.lister_button(False)
            frmInterface.ToolStripButton8.Enabled = True
            frmInterface.ToolStripButton8.BackColor = Color.Firebrick

            Me.Close()
            Form_windows.ListeDesArchivesToolStripMenuItem.Enabled = True
            Form_windows.ListeDesDépôtsToolStripMenuItem.Enabled = False
            Form_windows.ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = False

            'frmMenu.MdiParent = Form_windows
            'frmMenu.Show()

            'frmUpdate_check.MdiParent = Form_windows
            'frmUpdate_check.Show()

        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à mettre à jour")
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        backup = False
        DateCourant = Date.Now
        ListView1.Enabled = True
        MessageBox.Show("Le fichier a été fermé")
        Button3.Enabled = True
        Button5.Enabled = True
        Button2.Enabled = False
        frmInterface.lister_button(True)
        frmInterface.ToolStripButton8.Enabled = False
        frmInterface.ToolStripButton8.BackColor = Color.Transparent
        Form_windows.ListeDesDépôtsToolStripMenuItem.Enabled = True
        Form_windows.ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = True

        frmInterface.Label2.Text = "Initialisation terminée"
        frmInterface.Label2.ForeColor = Color.Black












    End Sub

    Sub changer_date(ByVal numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim datenr As Date


            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM bidon where numero = " & numero & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim nombre As Integer


            While myReader.Read
                Dim newListViewItem As New ListViewItem

                datenr = myReader.GetDateTime(2)


            End While

            DateCourant = datenr


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try

    End Sub

    Private Sub proc_archivage()

    End Sub

End Class