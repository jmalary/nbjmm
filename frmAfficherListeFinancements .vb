Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb
Imports System.Globalization


Public Class frmAfficherListeFinancements

    Public myConnToAccess As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim tables As DataTableCollection

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader

    Dim sqlStr As String
    'Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;"

    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()



        Me.WindowState = FormWindowState.Maximized
        'frmMenu.MdiParent = Form_windows
        'frmMenu.Show()
        Form_windows.ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'frmDetailsFinancement.Show()
        'Me.Close()
        'frmDetailsFinancement.MdiParent = Form_windows
        'frmDetailsFinancement.Show()

        If (ListView1.SelectedItems.Count > 0) Then


            'frmUpdate_Paiement_preautorise.Label7.Text = ListView1.SelectedItems(0).SubItems(0).Text

            num_list_financement = ListView1.SelectedItems(0).SubItems(0).Text
            'Me.Hide()
            Me.Close()
            'frmDetailsFinancement.MdiParent = Form_windows
            'frmDetailsFinancement.Show()

            Dim f As New frmDetailsFinancement()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            frmInterface.Panel1.Controls.Add(f)


        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à Afficher")
        End If




    End Sub

    Private Sub frmAfficherListeFinancements_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        ChargerLister()

    End Sub

    Sub ChargerLister()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                Dim newListViewItem As New ListViewItem


                newListViewItem.Text = myReader.GetInt32(0)

                Dim sNom As String = Recupere_description(myReader.GetString(2))

                newListViewItem.SubItems.Add(sNom)

                Dim payer As Date = myReader.GetDateTime(3)

                newListViewItem.SubItems.Add(
                payer.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA")))

                newListViewItem.SubItems.Add(myReader.GetString(4))
                newListViewItem.SubItems.Add(myReader.GetString(5))

                newListViewItem.SubItems.Add(myReader.GetString(6) + " $")

                Dim payer1 As Date = myReader.GetDateTime(8)

                newListViewItem.SubItems.Add(
                payer1.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA")))



                '*************************************************************************************


                ListView1.Items.Add(newListViewItem)

                'ListView1.Bounds = New Rectangle(New Point(10, 10), New Size(300, 200))

                ' Set the view to show details.
                ListView1.View = View.Details
                ' Allow the user to edit item text.
                ListView1.LabelEdit = True
                ' Allow the user to rearrange columns.
                ListView1.AllowColumnReorder = True
                ' Display check boxes.
                ListView1.CheckBoxes = True
                ' Select the item and subitems when selection is made.
                ListView1.FullRowSelect = True
                ' Display grid lines.
                ListView1.GridLines = True
                ' Sort the items in the list in ascending order.
                ListView1.Sorting = SortOrder.Ascending

                'ListView1.CheckBoxes = True

                'ListView1.Items(5).Selected = True

                'ListView1.CheckBoxes = True



                'ListView1.RightToLeft = Windows.Forms.LeftRightAlignment.Left




                'ListView1.RightToLeftLayout = True
                'ListView1.RightToLeft = Windows.Forms.RightToLeft.Yes
                'ListView1.View = View.Details




            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()


            proc_Vider()


        Catch ex As Exception

        End Try


    End Sub

    Sub proc_Vider()


        Try
            myConnToAccess = New OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter("SELECT numero from check_preautorises", myConnToAccess)
            da.Fill(ds, "employeur")
            myConnToAccess.Dispose()

            myConnToAccess = Nothing



            Dim quantity As Integer = ds.Tables.Count

            myConnToAccess.Close()
            da.Dispose()
            ds.Dispose()


        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim emptyFieldExists As Boolean = False

        'For Each item As ListViewItem In ListView1.Items
        '    If item.SubItems(0) Is Nothing Or item.SubItems(0).Text <> "" Then
        '        emptyFieldExists = True
        '    End If
        'Next

        'If emptyFieldExists Then
        '    MsgBox("viDEZ")
        '    'Do what you need to do!


        'End If
        For Each itm As ListViewItem In ListView1.Items
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()
                If itm.Checked Then
                    'MsgBox("item checked: " & itm.Text)
                    supprimer_lister(itm.Text)
                    Dim myCommand As New OleDb.OleDbCommand("DELETE FROM list_financement WHERE id = @numero", myConnection)

                    myCommand.Parameters.AddWithValue("@numero", itm.Text)
                    myCommand.ExecuteNonQuery()
                    myConnection.Close()

                    itm.Remove()

                    myConnection.Close()
                    myCommand.Dispose()

                End If
            Catch ex As Exception
                MsgBox("FAIL a")
            End Try


        Next
    End Sub

    Private Sub supprimer_lister(numero As String)

        'Dim numero As String = Label2.Text
        Dim num_financement As String = ""
        Dim categorie As String = ""
        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where id = " & numero & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read


                num_financement = myReader.GetString(1)
                categorie = myReader.GetString(4)


                'MsgBox(myReader.GetValue(5))
            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()


        Catch ex As Exception

        End Try


        Select Case categorie
            Case "preautorise"
                supprimer_mensuel(num_financement)
                supprimer_mensuel_check(num_financement)
            Case "revenu"
                supprimer_mensuel_revenu(num_financement)
                supprimer_mensuel_check_revenu(num_financement)
            Case "depense"

                supprimer_mensuel_depense(num_financement)
                supprimer_mensuel_check_depense(num_financement)

        End Select


    End Sub

    Private Sub supprimer_mensuel(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM mensuel WHERE numero = @numero", myConnection)

            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub

    Private Sub supprimer_mensuel_check(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_preautorises where payer like " & numero & "", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub

    Private Sub supprimer_mensuel_revenu(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM revenu WHERE numero = @numero", myConnection)

            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub
    Private Sub supprimer_mensuel_check_revenu(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_revenu where payer like " & numero & "", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub
    Private Sub supprimer_mensuel_depense(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM depense WHERE numero = @numero", myConnection)

            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub

    Private Sub supprimer_mensuel_check_depense(numero As String)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_depenses where payer like " & numero & "", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numero)
            myCommand.ExecuteNonQuery()
            myConnection.Close()



            myConnection.Close()
            myCommand.Dispose()


        Catch ex As Exception
            MsgBox("Fail mensuel")
        End Try


    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class