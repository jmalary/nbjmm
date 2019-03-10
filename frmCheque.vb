Imports System.Data.OleDb
Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization

Public Class frmCheque

    Public myConnToAccess As OleDbConnection
    Dim mySQLCommand As OleDbCommand
    Dim mySQLStrg As String
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim tables As DataTableCollection
    Dim btrouve As Boolean = False


    'Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\test2.mdb;"
    Dim myConnString As String = GlobalVariables.test2ConnectionString
    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String

    Private Sub frmCheque_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        'Me.Close()

        'frmListe_cheque.Close()



    End Sub


    Private Sub frmCheque_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Try
            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\test2.mdb;")

            myConnToAccess = New OleDbConnection(GlobalVariables.test2ConnectionString)



            'myConnToAccess = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;")
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            da = New OleDbDataAdapter("SELECT * from employeur ORDER BY numero", myConnToAccess)
            da.Fill(ds, "employeur")
            Dim view1 As New DataView(tables(0))


            'ComboBox2.Items.Insert(0, "Copenhagen")


            With ComboBox2
                .DataSource = ds.Tables("employeur")
                .DisplayMember = "description"
                .ValueMember = "numFournisseur"
                .SelectedIndex = 0
                '.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                '.AutoCompleteSource = AutoCompleteSource.ListItems
            End With

            Dim today = DateCourant

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

            'Label8.Text = DateTime.Now.ToString("dd/MM/yyyy")

            Label8.Text = DateCourant.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA"))

            lblerreur1.Text = "Veuillez entrer la numéro de chèque"
            lblerreur.Text = "Veuillez entrer un montant en utilisant " & ControlChars.NewLine & "le format suivant comme exemple : 120,00"


            test()
            myConnToAccess.Close()
            ds.Dispose()
            da.Dispose()

            view1.Dispose()



        Catch ex As Exception

        End Try

        TextBox2.Focus()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Me.WindowState = FormWindowState.Maximized
        'frmMenu.MdiParent = Form_windows
        'frmMenu.Show()

        Me.Close()


        'frmMenu.Show()

        frmInterface.btm_cheque.Enabled = True
        Form_windows.ToolStripMenuItemCheque.Enabled = True




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim courant As String = DateTimePicker2.Value

        Dim dt1 As DateTime = DateCourant



        Dim dt2 As DateTime = Convert.ToDateTime(DateTimePicker2.Text)
        Dim ts As TimeSpan = dt2.Subtract(dt1)

        Dim nbr = Convert.ToInt32(ts.Days) + 1

        If Convert.ToInt32(ts.Days) >= 0 Then
            'MessageBox.Show("Total Days are " & nbr)
            Dim id As Integer
            Dim ids As Integer
            id = Val(TextBox1.Text)
            ids = Val(TextBox2.Text)

            If (id > 0) Then
                If id > 0 And ids > 0 Then
                    'test()

                    If ComboBox2.SelectedValue <> 24 And TextBox2.Text <> " " Then
                        proc_finder()
                        Dim numfour As String = Recupere_description(ComboBox2.SelectedValue.ToString())
                        If btrouve = True Then



                            Select Case MessageBox.Show("Ce numéro de chèque existe, voulez-vous le changer?", "Saisir la numéro chéque ", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)

                                Case MsgBoxResult.No

                                    'MsgBox("YES Clicked.")
                                    Dim myConnection As New OleDb.OleDbConnection(myConnString)
                                    myConnection.Open()

                                    'Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE no_cheque = @numero", myConnection)
                                    'myCommand.Parameters.AddWithValue("@numero", TextBox2.Text)
                                    'myCommand.ExecuteNonQuery()
                                    'myConnection.Close()
                                    'myCommand.Dispose()
                                    proc_A(ComboBox2.SelectedValue.ToString)
                                    btrouve = False

                                    'e.Cancel = True
                                    'frmMenu.Close()
                                    'frmEnregistrer.MdiParent = Me
                                    'frmEnregistrer.Show()
                                Case MsgBoxResult.Yes

                                    TextBox2.Text = ""
                                    TextBox2.Focus()
                                    btrouve = False
                                    'quitter = True
                                    'Application.Exit()

                                Case MsgBoxResult.Cancel

                                    TextBox2.Text = ""
                                    TextBox1.Text = ""
                                    DateTimePicker2.Value = Date.Now
                                    ComboBox2.SelectedIndex = 0
                                    btrouve = False
                            End Select

                        Else

                            proc_A(ComboBox2.SelectedValue.ToString)

                        End If


                    Else
                        MsgBox("Vous devez choisir un fournisseur !!!")
                        ComboBox2.BackColor = Color.IndianRed
                        Panel1.BackColor = Color.IndianRed

                    End If

                Else
                    If ComboBox2.Text = "Choisir un employeur" Then
                        ComboBox2.BackColor = Color.IndianRed
                        Panel1.BackColor = Color.IndianRed

                    End If
                    TextBox2.BackColor = Color.IndianRed

                End If


            Else
                MsgBox("Le montant doit etre inscrit")
                TextBox1.BackColor = Color.IndianRed
                If ComboBox2.Text = "Choisir un employeur" Then
                    ComboBox2.BackColor = Color.IndianRed
                    Panel1.BackColor = Color.IndianRed

                End If

                If ids = 0 Then
                    TextBox2.BackColor = Color.IndianRed

                End If

            End If
        Else
            MessageBox.Show("Erreur, la date trop tot Invalid Input")
        End If



    End Sub

    Sub test()
        'Try
        '    Dim myConnection As New OleDb.OleDbConnection(myConnString)
        '    myConnection.Open()

        '    Dim myCommand As New OleDb.OleDbCommand("Select * FROM nbr where numero = 1", myConnection)
        '    Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

        '    Dim nombre As Integer
        '    While myReader.Read
        '        Dim newListViewItem As New ListViewItem

        '        nombre = myReader.GetInt32(1)
        '        'MsgBox(nombre)

        '        'newListViewItem.SubItems.Add(myReader.GetInt32(4))

        '    End While

        '    TextBox2.Text = nombre

        '    myConnection.Close()
        '    myCommand.Dispose()
        '    myReader.Close()


        'Catch ex As Exception

        'End Try



    End Sub

    Sub proc_finder()

        Dim id As Integer
        id = Val(TextBox2.Text)

        'id  contain the value of ID

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()
            'copier la connection string above
            'ouvrir la connection 
            'MsgBox("connection Ok")
            ' MsgBox("heyhey")
            sqlStr = "Select * from check_emis where no_cheque=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()

            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then
                'While dr.Read()
                '    TextBox1.Text = dr(0)
                '    TextBox2.Text = dr(1)
                '    MsgBox(dr(0))
                '    MsgBox(dr(1))
                'End While
                'MsgBox("Record Found")
                btrouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

    End Sub

    Sub proc_A(ByVal numFour As String)

        Dim dt1 As DateTime = DateCourant


        Dim dt2 As DateTime = Convert.ToDateTime(DateTimePicker2.Text)
        Dim ts As TimeSpan = dt2.Subtract(dt1)

        Dim nbr = Convert.ToInt32(ts.Days) + 1

        'Corrigé 31 aout

        nbr = nbr - 1

        Dim nom As String = Recupere_description(numFour)

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()



            Select Case nbr
                'Case 0 To 6
                Case 0 To 6
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "','" & DateCourant & "' , '" & DateTimePicker2.Value.ToString() & "', '" & TextBox2.Text & "',0 ,'" & TextBox1.Text & "', 0, 0, 0, 0, 0)"
                Case 7 To 13
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "', '" & DateCourant & "', '" & DateTimePicker2.Value.ToString() & "','" & TextBox2.Text & "' ,0, 0, '" & TextBox1.Text & "', 0, 0, 0, 0)"
                Case 14 To 20
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "','" & DateCourant & "' , '" & DateTimePicker2.Value.ToString() & "', '" & TextBox2.Text & "',0, 0,0 , '" & TextBox1.Text & "', 0, 0, 0)"
                Case 21 To 27
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "','" & DateCourant & "' , '" & DateTimePicker2.Value.ToString() & "', '" & TextBox2.Text & "',0, 0, 0, 0,'" & TextBox1.Text & "', 0, 0)"
                Case 28 To 30
                    'Case Else
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "','" & DateCourant & "' , '" & DateTimePicker2.Value.ToString() & "','" & TextBox2.Text & "',0, 0, 0, 0,0 , '" & TextBox1.Text & "', 0)"
                Case Is >= 31
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & ComboBox2.SelectedValue.ToString() & "', '" & nom & "','" & DateCourant & "' , '" & DateTimePicker2.Value.ToString() & "','" & TextBox2.Text & "',0, 0, 0, 0, 0, 0, '" & TextBox1.Text & "')"
            End Select


            cmd = New OleDb.OleDbCommand(sqlStr, con)

            cmd.ExecuteNonQuery()



            MsgBox("Chequé ajoutée")
            enregistrer = True


            DateTimePicker2.Value = DateCourant


            TextBox1.Text = ""
            TextBox2.Text = ""

            ComboBox2.SelectedIndex = 0



            cmd.Dispose()

            con.Close()

        Catch ex As Exception

            MessageBox.Show(ex.Message)


        End Try




    End Sub

    Sub proc_ajouter()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()
            Dim num_cheque As Integer = TextBox2.Text
            num_cheque = num_cheque + 1


            Dim myCommand As New OleDb.OleDbCommand("update nbr set circulation = @numero WHERE numero = 1", myConnection)
            'myCommand.Parameters.AddWithValue("@numero", Val(TextBox2.Text))
            'myCommand.Parameters.AddWithValue("@numero", 1)
            myCommand.Parameters.AddWithValue("@numero", num_cheque)

            myCommand.ExecuteNonQuery()

            myCommand.Dispose()
            oledbAdapter.Dispose()
            myConnection.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text <> "Choisir un employeur" Then
            ComboBox2.BackColor = Color.White
            Panel1.BackColor = Color.Transparent


        End If
    End Sub




    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim allowedChars As String = "0123456789,"
        If allowedChars.IndexOf(e.KeyChar) = -1 AndAlso
            Not e.KeyChar = ChrW(8) Then
            ' Invalid Character
            e.Handled = True
        End If


    End Sub





    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        Dim allowedChars As String = "0123456789"
        If allowedChars.IndexOf(e.KeyChar) = -1 AndAlso
            Not e.KeyChar = ChrW(8) Then
            ' Invalid Character
            e.Handled = True
        End If


    End Sub




    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not IsNumeric(TextBox1.Text) Then
            'TextBox1.Text = ""
            'MessageBox.Show("nombre non valide")
            lblerreur.Text = "Veuillez entrer un montant en utilisant " & ControlChars.NewLine & "le format suivant comme exemple : 120,00"
            TextBox1.BackColor = Color.IndianRed
        End If
    End Sub



    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If Not String.IsNullOrWhiteSpace(TextBox2.Text) AndAlso Not IsNumeric(TextBox2.Text) Then

            lblerreur1.Text = "Veuillez entrer un nombre de chèque "
            TextBox2.BackColor = Color.IndianRed
        End If

    End Sub



    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not IsNumeric(TextBox1.Text) Then
            'TextBox1.Text = ""
            'MessageBox.Show("nombre non valide")
            lblerreur.Text = "Veuillez entrer un montant en utilisant " & ControlChars.NewLine & "le format suivant comme exemple : 120,00"
            TextBox1.BackColor = Color.IndianRed
        End If
    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If Not String.IsNullOrWhiteSpace(TextBox2.Text) AndAlso Not IsNumeric(TextBox2.Text) Then
            'TextBox1.Text = ""
            'MessageBox.Show("nombre non valide")
            lblerreur1.Text = "Veuillez entrer un nombre de chèque "
            TextBox2.BackColor = Color.IndianRed
        End If
    End Sub



    Private Sub TextBox1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not IsNumeric(TextBox1.Text) Then
            ''TextBox1.Text = ""
            ''MessageBox.Show("nombre non valide  ")
            lblerreur.Text = "Veuillez entrer un montant en utilisant " & ControlChars.NewLine & "le format suivant comme exemple : 120,00"
            TextBox1.BackColor = Color.IndianRed
        End If
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        If Not String.IsNullOrWhiteSpace(TextBox2.Text) AndAlso Not IsNumeric(TextBox2.Text) Then
            ''TextBox1.Text = ""
            ''MessageBox.Show("nombre non valide  ")
            lblerreur1.Text = "Veuillez entrer un nombre de chèque "
            TextBox2.BackColor = Color.IndianRed
        End If
    End Sub



    Private Sub frmCheque_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

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
        Me.Location = New Point(x, 10)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> "" Then
            TextBox1.BackColor = Color.White
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text <> "" Then
            TextBox2.BackColor = Color.White
        End If
    End Sub

    Private Sub frmCheque_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub

    Private Function trouver_numFour(ByVal numero As Integer) As String

        Dim nom As String = ""

        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur where numero = " & numero & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim nombre As Integer


            While myReader.Read
                nom = myReader.GetString(1)

            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        Return nom


    End Function

End Class