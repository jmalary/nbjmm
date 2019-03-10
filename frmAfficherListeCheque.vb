Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO





Public Class frmAfficherListeCheque
    Public myConnToAccess As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim tables As DataTableCollection

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader

    Dim sqlStr As String
    'Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;"

    Dim CW As Integer = Me.Width ' Current Width
    Dim CH As Integer = Me.Height ' Current Height
    Dim IW As Integer = Me.Width ' Initial Width
    Dim IH As Integer = Me.Height ' Initial Height



    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Private Sub frmAfficherListeCheque_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        check_menu = False



    End Sub

   




    Private Sub frmAfficherListeCheque_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'IW = Me.Width
        'IH = Me.Height




        Me.WindowState = FormWindowState.Maximized

        ChargerLister()
        Dim y As Integer
        y = FrmInterface.Panel1.Height - 40

        Me.Height = FrmInterface.Panel1.Height

        'Label1.Location = New Point(500, y)
        'Label2.Location = New Point(615, y)
        'Label3.Location = New Point(760, y)
        'Label4.Location = New Point(890, y)
        'Label5.Location = New Point(1020, y)
        'Label7.Location = New Point(1145, y)


        'Label1.Location = New Point(400, y)
        'Label2.Location = New Point(600, y)
        'Label3.Location = New Point(760, y)
        'Label4.Location = New Point(790, y)
        'Label5.Location = New Point(920, y)
        'Label7.Location = New Point(1045, y)



        Me.Height = FrmInterface.Panel1.Height
        Label1.Top = FrmInterface.Panel1.Height - 40
        Label2.Top = FrmInterface.Panel1.Height - 40
        Label3.Top = FrmInterface.Panel1.Height - 40
        Label4.Top = FrmInterface.Panel1.Height - 40
        Label5.Top = FrmInterface.Panel1.Height - 40
        Label7.Top = FrmInterface.Panel1.Height - 40

    End Sub



    Sub ChargerLister()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_emis ORDER BY date_paiement", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            Dim value1 As String = "xxxxxxxxx"

            While myReader.Read


                Dim fin As Double = myReader.GetDouble(12)

                If fin <> 0 Then

                Else

                    Dim newListViewItem As New ListViewItem

                    newListViewItem.Text = myReader.GetInt32(0)
                    newListViewItem.SubItems.Add(myReader.GetString(1))
                    newListViewItem.SubItems.Add(myReader.GetString(2))



                    '*************************************************************************************

                    Dim today = myReader.GetDateTime(3)

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

                    today = myReader.GetDateTime(4)



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


                    newListViewItem.SubItems.Add(myReader.GetInt32(5))

                    newListViewItem.SubItems.Add(Format$(myReader.GetDouble(7), "Currency"))
                    newListViewItem.SubItems.Add(Format$(myReader.GetDouble(8), "Currency"))
                    newListViewItem.SubItems.Add(Format$(myReader.GetDouble(9), "Currency"))
                    newListViewItem.SubItems.Add(Format$(myReader.GetDouble(10), "Currency"))
                    newListViewItem.SubItems.Add(Format$(myReader.GetDouble(11), "Currency"))

                    ListView1.Items.Add(newListViewItem)
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
                    'ListView1.Columns(0).DisplayIndex = ListView1.Columns.Count - 1

                End If


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
            MessageBox.Show(ex.Message)

        End Try



    End Sub


    Sub proc_A()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_emis", myConnection)
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
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_emis", myConnection)
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
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_emis", myConnection)
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
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_emis", myConnection)
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
            Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_emis", myConnection)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'frmMenu.Show()
        'quitter = True
        If check_menu = False Then
            Me.Close()

        Else
            Me.Close()
            'Me.WindowState = FormWindowState.Maximized
            'frmMenu.MdiParent = Form_windows
            'frmMenu.Show()
        End If

        frmInterface.lister_button(True)
    End Sub

    Sub proc_zero()

        Label2.Text = "1.00 $"
        Label3.Text = "2.00 $"
        Label4.Text = "3.00 $"
        Label5.Text = "4.00 $"
        Label7.Text = "5.00 $"


    End Sub


    Sub proc_Vider()

        Try
            myConnToAccess = New OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter("SELECT numero from check_emis", myConnToAccess)
            da.Fill(ds, "employeur")
            myConnToAccess.Dispose()

            'myConnToAccess = Nothing



            Dim quantity As Integer = ds.Tables.Count

            myConnToAccess.Close()
            da.Dispose()
            ds.Dispose()



        Catch ex As Exception

        End Try





    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

        If (ListView1.SelectedItems.Count > 0) Then
            For Each i As ListViewItem In ListView1.SelectedItems
                Try
                    Dim myConnection As New OleDb.OleDbConnection(myConnString)
                    myConnection.Open()

                    Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE numero = @numero", myConnection)
                    myCommand.Parameters.AddWithValue("@numero", i.Text)

                    myCommand.ExecuteNonQuery()
                    myConnection.Close()

                    ListView1.Items.Remove(i)


                    If ListView1.Items.Count = 0 Then
                        proc_zero()
                    Else
                        proc_A()
                        proc_B()
                        proc_C()
                        proc_D()
                        proc_E()
                        MsgBox("Element supprimé")
                        enregistrer = True
                    End If

                    myConnection.Close()
                    myCommand.Dispose()




                Catch ex As Exception

                End Try

            Next
        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à supprimer")
        End If
    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If (ListView1.SelectedItems.Count > 0) Then

            num_check = ListView1.SelectedItems(0).SubItems(0).Text
            Me.Close()
            'frmUpdate_check.Show()
            'frmUpdate_check.Label2.Text = ListView1.SelectedItems(0).SubItems(0).Text

            ''Me.Close()
            'frmMenu.Close()

            'frmUpdate_check.MdiParent = Form_windows
            'frmUpdate_check.Show()

            Dim fc As New frmUpdate_check()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            frmInterface.Panel1.Controls.Add(fc)
        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à mettre à jour")
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim id As Double
        id = Val(InputBox("Entrez le montant"))

        'id  contain the value of ID

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()
            'copier la connection string above
            'ouvrir la connection 
            'MsgBox("connection Ok")
            sqlStr = "Select * from check_emis where un=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            'cmd.ExecuteNonQuery()

            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then
                While dr.Read()
                    'TextBox1.Text = dr(0)
                    'TextBox2.Text = dr(1)
                End While
            End If
            'ExecuteReader fo
            MsgBox("Record Found")
            con.Close()
            cmd.Dispose()
            dr.Close()


        Catch ex As Exception
            MessageBox.Show("could not find record")

        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        'Dim I As Integer
        If (ListView1.SelectedItems.Count > 0) Then
            Dim a = MsgBox("Vous etes certain de supprimer quelques entrée ?", MsgBoxStyle.YesNo, "Confirmation")
            If a = vbYes Then

                For Each i As ListViewItem In ListView1.SelectedItems
                    Try
                        Dim numFour As String
                        Dim date1 As Date
                        Dim date2 As Date
                        Dim numCheque As Integer
                        Dim montant As Double

#Disable Warning BC42030 ' La variable 'numFour' est passée par référence avant qu'une valeur lui ait été attribuée. Cela peut provoquer une exception de référence null au moment de l'exécution.
                        obtenir_numCheque(i.Text, numFour, date1, date2, numCheque, montant)
#Enable Warning BC42030 ' La variable 'numFour' est passée par référence avant qu'une valeur lui ait été attribuée. Cela peut provoquer une exception de référence null au moment de l'exécution.

                        Dim myConnection As New OleDb.OleDbConnection(myConnString)
                        myConnection.Open()

                        Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE numero = @numero", myConnection)
                        myCommand.Parameters.AddWithValue("@numero", i.Text)
                        myCommand.ExecuteNonQuery()
                        myConnection.Close()

                        ListView1.Items.Remove(i)

                        enregistrer = True

                        myConnection.Close()
                        myCommand.Dispose()

                        supprimer_cheque(i.Text, numFour, date1, date2, numCheque, montant)
                        'supprimer_cheque(obtenir_numCheque(i.Text))
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

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
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
                Dim numFour As Integer
                Dim date1 As Date
                Dim date2 As Date
                Dim numCheque As Integer
                Dim montant As Double

                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()
                If itm.Checked Then



                    obtenir_numCheque(itm.Text, numFour, date1, date2, numCheque, montant)
                    'MsgBox("item checked: " & itm.Text)
                    Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE numero = @numero", myConnection)

                    myCommand.Parameters.AddWithValue("@numero", itm.Text)
                    myCommand.ExecuteNonQuery()
                    'myConnection.Close()

                    itm.Remove()
                    enregistrer = True
                    myCommand.Dispose()
                End If
                myConnection.Close()

                'supprimer_cheque(obtenir_numCheque(itm.Text))
                supprimer_cheque(itm.Text, numFour, date1, date2, numCheque, montant)
            Catch ex As Exception

            End Try

        Next

    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ajouter_bouton()

        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width
        y = Me.Size.Height() - 90
        y = 470

        Label1.Location = New Point(450, y)
        Label2.Location = New Point(600, y)
        Label3.Location = New Point(737, y)
        Label4.Location = New Point(875, y)
        Label5.Location = New Point(985, y)
        Label7.Location = New Point(1110, y)







    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub frmAfficherListeCheque_Resize(sender As Object, e As EventArgs) Handles Me.Resize

        'Me.Width = FrmInterface.Panel1.Width
        'Me.Height = FrmInterface.Panel1.Height

        'Dim RW As Double = (Me.Width - CW) / CW ' Ratio change of width
        'Dim RH As Double = (Me.Height - CH) / CH ' Ratio change of height

        'For Each Ctrl As Control In Controls
        '    Ctrl.Width += CInt(Ctrl.Width * RW)
        '    Ctrl.Height += CInt(Ctrl.Height * RH)
        '    Ctrl.Left += CInt(Ctrl.Left * RW)
        '    Ctrl.Top += CInt(Ctrl.Top * RH)
        'Next

        'CW = Me.Width
        'CH = Me.Height
    End Sub


    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        If (ListView1.SelectedItems.Count > 0) Then

            num_check = ListView1.SelectedItems(0).SubItems(0).Text
            Me.Close()
            'frmUpdate_check.Show()
            'frmUpdate_check.Label2.Text = ListView1.SelectedItems(0).SubItems(0).Text

            ''Me.Close()
            'frmMenu.Close()

            'frmUpdate_check.MdiParent = Form_windows
            'frmUpdate_check.Show()

            Dim fc As New frmUpdate_check()
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

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'frmMenu.Show()
        'quitter = True
        If check_menu = False Then
            Me.Close()

        Else
            Me.Close()
            'Me.WindowState = FormWindowState.Maximized
            'frmMenu.MdiParent = Form_windows
            'frmMenu.Show()
        End If

        FrmInterface.Lister_button(True)
    End Sub

    Private Sub frmAfficherListeCheque_RightToLeftChanged(sender As Object, e As EventArgs) Handles Me.RightToLeftChanged

    End Sub
End Class
