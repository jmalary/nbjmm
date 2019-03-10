Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports iTextSharp.text.Font

Imports System.IO
Imports System.Data


Public Class frmAfficherListeDepense

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

    Private Sub frmAfficherListeDepense_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        ajouter_bouton()

        ChargerLister()

    End Sub

    Sub ChargerLister()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_depenses", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                Dim newListViewItem As New ListViewItem


                newListViewItem.Text = myReader.GetInt32(0)

                'newListViewItem.SubItems.Add(myReader.GetString(1))


                Dim sNom As String = Recupere_description(myReader.GetString(1))
                newListViewItem.SubItems.Add(sNom)


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

                today = myReader.GetDateTime(3)



                day = today.Day

                month = today.Month

                years = today.Year

                ' MsgBox(Date.Today.DayOfWeek)




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

                courant = String.Concat(jour, " ", day, " ", Mois, " ", years)

                newListViewItem.SubItems.Add(courant)


                '*************************************************************************************


                'newListViewItem.SubItems.Add(myReader.GetInt32(4))

                newListViewItem.SubItems.Add(Format$(myReader.GetDouble(4), "Currency"))
                newListViewItem.SubItems.Add(Format$(myReader.GetDouble(5), "Currency"))
                newListViewItem.SubItems.Add(Format$(myReader.GetDouble(6), "Currency"))
                newListViewItem.SubItems.Add(Format$(myReader.GetDouble(7), "Currency"))
                newListViewItem.SubItems.Add(Format$(myReader.GetDouble(8), "Currency"))

                ListView1.Items.Add(newListViewItem)

                'ListView1.Bounds = New Rectangle(New Point(10, 10), New Size(300, 200))

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
                ListView1.Columns(0).DisplayIndex = ListView1.Columns.Count - 1


            End While

            If ListView1.Items.Count = 0 Then
                proc_zero()
            Else
                proc_A()
                proc_B()
                proc_C()
                proc_D()
                proc_E()

            End If


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()


            proc_Vider()


        Catch ex As Exception

        End Try


    End Sub
    Private Sub ajouter_bouton()

        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width
        'y = Me.Size.Height() - 90
        y = 470

        Label1.Location = New Point(400, y)
        Label2.Location = New Point(520, y)
        Label3.Location = New Point(660, y)
        Label4.Location = New Point(775, y)
        Label5.Location = New Point(905, y)
        Label7.Location = New Point(1025, y)

    End Sub

    Sub proc_zero()

        Label2.Text = "0.00 $"
        Label3.Text = "0.00 $"
        Label4.Text = "0.00 $"
        Label5.Text = "0.00 $"

        Label7.Text = "0.00 $"

    End Sub

    Sub proc_A()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_depenses", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

            End While
            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

    End Sub

    Sub proc_B()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_depenses", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                Label3.Text = Format$(myReader1.Item(0).ToString, "Currency")
            End While

            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

    End Sub

    Sub proc_C()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_depenses", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                Label4.Text = Format$(myReader1.Item(0).ToString, "Currency")
            End While
            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try


    End Sub

    Sub proc_D()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_depenses", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                Label5.Text = Format$(myReader1.Item(0).ToString, "Currency")
            End While
            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()


        Catch ex As Exception

        End Try

    End Sub

    Sub proc_E()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_depenses      ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                Label7.Text = Format$(myReader1.Item(0).ToString, "Currency")
            End While
            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()


        Catch ex As Exception

        End Try

    End Sub

    Sub proc_Vider()


        Try
            myConnToAccess = New OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter("SELECT numero from check_depenses", myConnToAccess)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If check_menu = False Then
            Me.Close()
            'frmMenu.Show()
        Else
            Me.Close()
            Me.WindowState = FormWindowState.Maximized
            'frmMenu.MdiParent = Form_windows
            'frmMenu.Show()
        End If

        FrmInterface.Lister_button(True)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If (ListView1.SelectedItems.Count > 0) Then
            Dim a = MsgBox("Vous etes certain de supprimer quelques entrée ?", MsgBoxStyle.YesNo, "Confirmation")
            If a = vbYes Then

                For Each i As ListViewItem In ListView1.SelectedItems
                    Try
                        Dim numFour As String
                        Dim date1 As Date
                        Dim date2 As Date

                        Dim montant As Double

#Disable Warning BC42030 ' La variable 'numFour' est passée par référence avant qu'une valeur lui ait été attribuée. Cela peut provoquer une exception de référence null au moment de l'exécution.
                        obtenir_numDepense(i.Text, numFour, date1, date2, montant)
#Enable Warning BC42030 ' La variable 'numFour' est passée par référence avant qu'une valeur lui ait été attribuée. Cela peut provoquer une exception de référence null au moment de l'exécution.
                        'obtenir_numPaiementDirect(i.Text, numFour, date1, date2, numPaiem, montant)
                        Dim myConnection As New OleDb.OleDbConnection(myConnString)
                        myConnection.Open()

                        Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_depenses WHERE numero = @numero", myConnection)
                        myCommand.Parameters.AddWithValue("@numero", i.Text)
                        myCommand.ExecuteNonQuery()
                        myConnection.Close()

                        ListView1.Items.Remove(i)
                        enregistrer = True
                        myConnection.Close()
                        myCommand.Dispose()


                        supprimer_depense(i.Text, numFour, date1, date2, montant)

                    Catch ex As Exception

                    End Try

                Next

            End If

            If ListView1.Items.Count = 0 Then
                proc_zero()
            Else
                proc_A()
                proc_B()
                proc_C()
                proc_D()
                proc_E()

            End If



        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à supprimer")
        End If

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
                    Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_depenses WHERE numero = @numero", myConnection)

                    myCommand.Parameters.AddWithValue("@numero", itm.Text)
                    myCommand.ExecuteNonQuery()
                    myConnection.Close()

                    itm.Remove()
                    enregistrer = True
                    myConnection.Close()
                    myCommand.Dispose()

                End If
            Catch ex As Exception

            End Try

        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (ListView1.SelectedItems.Count > 0) Then
            'frmUpdate_Paiement_preautorise.Show()

            'frmUpdate_Paiement_preautorise.Label7.Text = ListView1.SelectedItems(0).SubItems(0).Text
            'Me.Hide()
            num_preautorise = ListView1.SelectedItems(0).SubItems(0).Text
            Me.Close()
            'frmMenu.Close()

            'frmUpdate_Paiement_preautorise.MdiParent = Form_windows
            'frmUpdate_Paiement_preautorise.Show()
            Dim fc As New frmUpdate_Paiement_depense()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True





            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            FrmInterface.Panel1.Controls.Add(fc)



        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à mettre à jour")
        End If
    End Sub
End Class