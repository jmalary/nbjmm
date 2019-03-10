Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization

Public Class Form_windows

    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Dim myConnection As New OleDb.OleDbConnection(myConnString)
    Dim con As System.Data.OleDb.OleDbConnection
    Dim sqlStr As String
    Dim cmd As System.Data.OleDb.OleDbCommand

    Dim test_a As String

    'Structure longStructure
    '    Dim nom As String
    '    Dim adresse As String
    '    Dim ville As String
    '    Dim postal As String

    'End Structure

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim not_back As Boolean


        Me.WindowState = FormWindowState.Maximized


        'Afficher le menu principale
        'frmMenu.MdiParent = Me
        'frmMenu.Show()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM listdate", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            Dim heure As DateTime



            While myReader.Read

                auj = myReader.GetDateTime(1)
                heure = myReader.GetDateTime(3)





            End While
            'afficher la date recule

            If auj <= DateCourant Then
                not_back = False
            Else
                not_back = True
            End If




            If changer_date = True Then
                auj = DateCourant
            End If

            'auj.ToString("dddd dd MMMM yyyy",
            'CultureInfo.CreateSpecificCulture("fr-CA"))



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try




        FrmInterface.MdiParent = Me
        FrmInterface.Dock = DockStyle.Fill
        FrmInterface.Show()


        If not_back = False Then

            Changertemps()


        Else
            FrmInterface.Close()
            My.Application.SplashScreen.Close()

            MsgBox("Votre logiciel NBJMM doit être desinstallée ou la date doit etre après  " + auj.ToString("dddd dd MMMM yyyy", CultureInfo.CreateSpecificCulture("fr-CA")) + " dans le système de votre ordinateur")
            Application.Exit()


        End If




    End Sub


    Private Sub ToolStripLabel1_Click(sender As Object, e As EventArgs)
        ' Me.WindowState = FormWindowState.Maximized
        'frmMenu.MdiParent = Me
        'frmMenu.Show()

    End Sub

    Private Sub ÀProposDeNBJMMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ÀProposDeNBJMMToolStripMenuItem.Click
        frmMenu.Close()
        'frmPropos.Show()
        'frmAbout.Show()

    End Sub

    Private Sub QuitterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click






        Me.Close()

    End Sub

    Private Sub CalculatriceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculatriceToolStripMenuItem.Click
        'frmInterface.Close()
        'FrmCalculatrice.MdiParent = Me
        'FrmCalculatrice.Show()

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                FrmInterface.Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            FrmInterface.Panel1.Controls.Clear()

            FrmInterface.Lister_button(True)

            'Dim fc As New frmAfficherListeFinancements()
            Dim fc As New FrmCalculatrice()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(100, 100)

            'frmInterface.Panel1.Remove(fc)


            FrmInterface.Panel1.Controls.Add(fc)
            'ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = False


        End If



    End Sub

    Private Sub NoteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NoteToolStripMenuItem.Click

        FrmInterface.Close()
        frmNotepad.MdiParent = Me
        frmNotepad.Show()

        'If faire_vider = True Then

        '    Dim result As DialogResult

        '    result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
        '    If result = DialogResult.No Then
        '        faire_vider = False
        '        FrmInterface.Lister_button(True)

        '    ElseIf result = DialogResult.Yes Then

        '    End If


        'End If

        'If faire_vider = False Then

        '    FrmInterface.Panel1.Controls.Clear()

        '    FrmInterface.Lister_button(True)

        '    'Dim fc As New frmAfficherListeFinancements()
        '    Dim fc As New frmNotepad()
        '    'InitializeComponent()
        '    fc.TopLevel = False
        '    'f.WindowState = FormWindowState.Maximized
        '    fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        '    fc.Visible = True
        '    fc.Location = New Point(1, 0)

        '    'frmInterface.Panel1.Remove(fc)


        '    FrmInterface.Panel1.Controls.Add(fc)
        '    'ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = False


        'End If





    End Sub

    Private Sub ContacterCentreDassistanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContacterCentreDassistanceToolStripMenuItem.Click
        frmMenu.Close()
        frmEmail.MdiParent = Me
        frmEmail.Show()
    End Sub

    Private Sub FichierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FichierToolStripMenuItem.Click

        Me.WindowState = FormWindowState.Maximized


        'frmMenu.Show()


    End Sub

    Private Sub MenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuToolStripMenuItem.Click
        Me.WindowState = FormWindowState.Maximized


        'Afficher le menu principale
        'frmMenu.MdiParent = Me
        'frmMenu.Show()



        frmInterface.MdiParent = Me
        frmInterface.Show()


    End Sub

    Private Sub SaisieToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaisieToolStripMenuItem.Click

    End Sub

    Private Sub ChangeLheureToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeLheureToolStripMenuItem.Click



        Dim fc As New frmDate()
        ' InitializeComponent()
        fc.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        fc.Visible = True
        fc.Location = New Point(1, 0)

        'frmInterface.Panel1.Remove(fc)

        frmInterface.Panel1.Controls.Add(fc)




    End Sub

    Private Sub OutilsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OutilsToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub ListeDesChèquesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesChèquesToolStripMenuItem.Click


        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmAfficherListeCheque()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            frmInterface.Panel1.Controls.Add(fc)
            ListeDesChèquesToolStripMenuItem.Enabled = True
            FrmInterface.btm_lister_cheque.Enabled = True



        End If


    End Sub

    Private Sub ListeDesPaiementsPréautoriséToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesPaiementsPréautoriséToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmAfficherListePaiementPreautorise()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            frmInterface.Panel1.Controls.Add(fc)
            ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = True
            FrmInterface.btm_lister_paiement_preautorise.Enabled = True



        End If

    End Sub

    Private Sub ListeDesPaiementsDirectsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesPaiementsDirectsToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmAfficherListePaiementDirect()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            ListeDesPaiementsDirectsToolStripMenuItem.Enabled = True
            FrmInterface.btm_lister_paiement_direct.Enabled = True



        End If



        '********************************************************************

    End Sub

    Private Sub ListeDesArchivesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesArchivesToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmAfficherListeArchives()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            ListeDesArchivesToolStripMenuItem.Enabled = True


        End If





    End Sub

    Private Sub EnregistrerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnregistrerToolStripMenuItem.Click
        'frmMenu.Close()
        'frmEnregistrer.MdiParent = Me
        'frmEnregistrer.Show()

        Dim f As New frmEnregistrer()
        'InitializeComponent()
        f.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.Visible = True
        f.Location = New Point(1, 0)

        frmInterface.Panel1.Controls.Add(f)

    End Sub



    Sub Miseajour()

        Dim cnn As New OleDb.OleDbConnection

        Try
            cnn = New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            Dim command As String
            'command = "INSERT INTO listerdate(description, datenr, compression) VALUES (@description, @datenr, @compression)"
            command = "INSERT INTO listdate(miseajour, description) VALUES(@datenr, @description)"
            cnn.Open()
            Dim cmd As OleDb.OleDbCommand
            cmd = New OleDb.OleDbCommand(command, cnn)
            cmd.Parameters.AddWithValue("@datenr", DateCourant)
            cmd.Parameters.AddWithValue("@description", "aaaa")
            cmd.ExecuteNonQuery()
            'MessageBox.Show("Add réussi")
        Catch exceptionObject As Exception
            MessageBox.Show(exceptionObject.Message)
        Finally
            cnn.Close()
        End Try


    End Sub


    Sub Changertemps()



        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()



            Dim myCommand As New OleDb.OleDbCommand("Update listdate set miseajour = @enrdate WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@enrdate", DateCourant)

            Dim newValueConverted As Integer = Val("2")


            myCommand.Parameters.AddWithValue("@numero", newValueConverted)


            myCommand.ExecuteNonQuery()
            myConnection.Close()
            myCommand.Dispose()

        Catch ex As Exception
            MessageBox.Show(" Error occurred: " & ex.Message)

        End Try





    End Sub


    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing



        If backup = False Then

            'changertemps()
        End If




        If enregistrer = True Then



            Select Case MsgBox("Voulez-vous enregistrer les modifications que vous avez modifié ?", MsgBoxStyle.Critical + MsgBoxStyle.YesNoCancel, "")

                Case MsgBoxResult.Yes

                    e.Cancel = True
                    frmMenu.Close()
                    frmEnregistrer.MdiParent = Me
                    frmEnregistrer.Show()
                Case MsgBoxResult.No
                    quitter = True
                    Application.Exit()

                Case MsgBoxResult.Cancel

                    e.Cancel = True
            End Select
        End If

        '*********************************************************************

        'Dim message As String =
        '    "Are you sure that you would like to close the form?"
        'Dim caption As String = "Form Closing"
        'Dim result = MessageBox.Show(message, caption,
        '                             MessageBoxButtons.YesNoCancel,
        '                             MessageBoxIcon.Question)

        '' If the no button was pressed ...
        'If (result = DialogResult.Cancel) Then
        '    ' cancel the closure of the form.
        '    e.Cancel = True
        'ElseIf (result = DialogResult.No) Then
        '    'Application.Exit()
        'ElseIf (result = DialogResult.Yes) Then
        '    frmMenu.Close()
        '    frmEnregistrer.MdiParent = Me
        '    frmEnregistrer.Show()

        'End If


        '*********************************************

        'Dim result As Integer = MessageBox.Show("message", "caption", MessageBoxButtons.YesNoCancel)
        'If result = DialogResult.Cancel Then
        '    MessageBox.Show("Cancel pressed")
        'ElseIf result = DialogResult.No Then
        '    MessageBox.Show("No pressed")
        'ElseIf result = DialogResult.Yes Then
        '    MessageBox.Show("Yes pressed")
        'End If




    End Sub 'Form1_Closing

    Private Sub ListeDesFinancementsPréautorisésToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesFinancementsPréautorisésToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmAfficherListeFinancements()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            ListeDesFinancementsPréautorisésToolStripMenuItem.Enabled = True


        End If



    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ToolStripMenuItemCheque_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemCheque.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then
            'InitializeComponent()
            FrmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmCheque()

            fc.TopLevel = False
            fc.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            FrmInterface.btm_cheque.Enabled = True
            ToolStripMenuItemCheque.Enabled = True





        End If

    End Sub

    Private Sub ToolStripMenuItemPaiement_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemPaiement.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmPaiement()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            FrmInterface.btm_paiement.Enabled = True
            ToolStripMenuItemPaiement.Enabled = True




        End If

    End Sub

    Private Sub ToolStripMenuItemFournisseur_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemFournisseur.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmEmployeur()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            frmInterface.Panel1.Controls.Add(fc)
            FrmInterface.btm_fournisseur.Enabled = True
            ToolStripMenuItemFournisseur.Enabled = True




        End If

    End Sub

    Private Sub TransfererTousLesDonneesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransfererTousLesDonneesToolStripMenuItem.Click


    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk


        'Me.Hide()
        'FrmInterface.Hide()
        'Form_windows.Hide()



        frmAnimation.Show()
        frmAnimation.TopMost = True
        My.Application.DoEvents()

        'frmAnimation.Label1.Text = " Pourcentage " & SystemInformation.PowerStatus.BatteryLifePercent * 100
        'frmAnimation.Label2.Text = " Temp restant " & SystemInformation.PowerStatus.BatteryLifeRemaining / 60




        Try
            'importation = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            importation = System.IO.File.ReadAllText(OpenFileDialog1.FileName, System.Text.Encoding.Default)

        Catch ex As Exception
        End Try




        'Dim bloc As String() = importation.Split(New Char() {"0"c})

        Dim bloc() As String = Split(importation, "{0_FFOU;")

        For i = 1 To UBound(bloc)
            bloc(i - 1) = bloc(i)
        Next i
        ReDim Preserve bloc(UBound(bloc) - 1)


        Dim word As String
        Dim no_emp As Integer = 0
        Dim nouveau As Integer = 0


        Dim pourcentage As Integer = bloc.Length()

        frmAnimation.Label1.Text = ""

        For Each word In bloc

            Dim chaine As String() = word.Split(New Char() {";"c})



            Dim bString As String = chaine(0)

            bString = bString.Remove(0, 5)

            Dim sValueAsArray = bString.ToCharArray()

            If IsNumeric(sValueAsArray(0)) Then
                'Console.WriteLine("First character is numeric")
                Dim trouve As Boolean = Fournisseur_existant(bString)
                If trouve = False Then
                    Ajouter_fournisseur(bString, chaine(2), chaine(4), chaine(5), chaine(6))
                    no_emp += 1
                    nouveau += 1
                Else
                    enlever_fournisseur(bString)
                    Ajouter_fournisseur(bString, chaine(2), chaine(4), chaine(5), chaine(6))
                    no_emp += 1

                End If

            Else
                bString = bString.Remove(0, 1)
                sValueAsArray = bString.ToCharArray()
                If IsNumeric(sValueAsArray(0)) Then
                    'Console.WriteLine("First character is numeric")

                    Dim trouve As Boolean = Fournisseur_existant(bString)

                    If trouve = False Then
                        Ajouter_fournisseur(bString, chaine(2), chaine(4), chaine(5), chaine(6))
                        no_emp += 1
                        nouveau += 1
                    Else
                        enlever_fournisseur(bString)
                        Ajouter_fournisseur(bString, chaine(2), chaine(4), chaine(5), chaine(6))
                        no_emp += 1
                    End If

                Else

                    'Console.WriteLine("First character is not numeric")
                End If
                Console.WriteLine("First character is not numeric")
            End If



            frmAnimation.Label3.Text = no_emp & " / " & pourcentage
            Application.DoEvents()
            frmAnimation.Label3.Refresh()





        Next

        frmEmployeur.Label7.Text = "Total = " + no_emp.ToString() + " fournisseurs."

        'frmLoading.Close()
        'frmDownload.Close()
        frmAnimation.Close()


        MsgBox("Termine,  " & nouveau.ToString & " nouveaux fournisseurs")

        FrmInterface.Panel1.Controls.Clear()


    End Sub

    Private Sub Ajouter_fournisseur(ByVal numero As String, ByVal nom As String, ByVal adresse As String, ByVal ville As String, ByVal postal As String)

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()



            adresse = adresse.Remove(0, 5)
            postal = postal.Remove(0, 5)



            If nom.Contains("'") Then
                nom = Replace(nom, "'", "''")
            End If

            If adresse.Contains("'") Then
                adresse = Replace(adresse, "'", "''")
            End If

            If ville.Contains("'") Then
                ville = Replace(ville, "'", "''")
            End If

            sqlStr = "INSERT INTO employeur (numFournisseur, description,Adresse,code ,Ville) VALUES('" & numero & "' ,'" & nom & "' , '" & adresse & "', '" & postal & "','" & ville & "')"


            ' sqlStr = "INSERT INTO check_direct (nom,quand,un,deux,trois,quatre) VALUES('" & ComboBox2.SelectedValue.ToString & "', '" & DateTimePicker1.Value.ToString() & "',1, 2, 3, 4)"
            cmd = New OleDb.OleDbCommand(sqlStr, con)

            cmd.ExecuteNonQuery()

            'proc_ajouter()

            enregistrer = True


            cmd.Dispose()

            con.Close()

        Catch ex As Exception
            MessageBox.Show("Erreur Importation = " + numero)

        End Try

    End Sub

    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk



        'frmLoading.Show() 'frmDownload.Show()
        frmAnimation.Show()
        frmAnimation.TopMost = True
        My.Application.DoEvents()

        Dim no_transaction As Integer = 0

        Try
            'importation = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            importation = System.IO.File.ReadAllText(OpenFileDialog2.FileName, System.Text.Encoding.Default)

        Catch ex As Exception

        End Try



        Dim bloc() As String = Split(importation, "{0_FCHQ;")

        nbr_element = bloc.Length




        For i = 1 To UBound(bloc)
            bloc(i - 1) = bloc(i)
        Next i
        ReDim Preserve bloc(UBound(bloc) - 1)


        Dim word As String

        For Each word In bloc


            Dim chaine As String() = word.Split(New Char() {";"c})


            Dim fournisseur As String = chaine(0)
            Dim calendar As String = chaine(1)
            Dim cheque As String = chaine(4)
            Dim sorte As Char = chaine(7)


            Dim somme() As String = Split(word, "{1_FPAI;")

            For i = 1 To UBound(somme)
                somme(i - 1) = somme(i)
            Next i
            ReDim Preserve somme(UBound(somme) - 1)

            Dim total As Decimal = 0



            For Each mot In somme

                Dim sous_chaine As String() = mot.Split(New Char() {";"c})


                Dim aAsDecimal As Decimal = CDec(Val(sous_chaine(1)))

                total += aAsDecimal


            Next

            Dim trouve, trouve1 As Boolean



            Dim str_nombre As String = total.ToString()

            If sorte = "Y" Then
                'Verifier si le numero existe dans la liste des cheque
                trouve = Numero_existant(cheque)
                'Verifier si le numero existe dans la liste des cheques supprime
                trouve1 = Numero_existant_cheque(cheque)

            Else

                trouve = Numero_existant_d(cheque)
                'Verifier si le numero existe dans la liste des cheques supprime
                trouve1 = Numero_existant_direct(cheque)

            End If



            If sorte = "Y" Then


                If trouve = False And trouve1 = False Then
                    fournisseur = fournisseur.Remove(0, 5)
                    Ajouter_transactions_cheque(fournisseur, calendar, cheque, str_nombre)
                ElseIf trouve = True And cheque <> "" Then
                    fournisseur = fournisseur.Remove(0, 5)
                    enlever_cheque(cheque)
                    Ajouter_transactions_cheque(fournisseur, calendar, cheque, str_nombre)
                End If



            Else

                If trouve = False And trouve1 = False Then
                    fournisseur = fournisseur.Remove(0, 5)
                    Ajouter_transactions_direct(fournisseur, calendar, cheque, str_nombre)
                ElseIf trouve = True And cheque <> "" Then
                    fournisseur = fournisseur.Remove(0, 5)
                    enlever_paiement_direct(cheque)
                    Ajouter_transactions_direct(fournisseur, calendar, cheque, str_nombre)
                End If

            End If

            no_transaction += 1

            frmAnimation.Label3.Text = no_transaction & " / " & nbr_element
            Application.DoEvents()
            frmAnimation.Label3.Refresh()


        Next

        'frmLoading.Close()
        frmAnimation.Close()




        MsgBox("Fin de telechargement")





    End Sub



    Private Sub FourniesseurToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FourniesseurToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()



    End Sub

    Private Sub TransactionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransactionToolStripMenuItem.Click
        OpenFileDialog2.ShowDialog()
    End Sub




    Private Sub Ajouter_transactions_cheque(ByVal fournisseur As String, ByVal calendar As String, ByVal cheque As String, ByVal total As String)

        Dim dt1 As DateTime = DateCourant


        Dim annee As String


        Dim c2 As String = calendar
        Dim c3 As String = calendar

        annee = calendar.Substring(0, 4)


        Dim mois1 As Char
        Dim mois2 As Char
        mois1 = calendar.Chars(4) ' myChar = "D"
        mois2 = calendar.Chars(5) ' myChar = "D"

        Dim jour1 As Char
        Dim jour2 As Char
        jour1 = calendar.Chars(6) ' myChar = "D"
        jour2 = calendar.Chars(7) ' myChar = "D"



        Dim value As String = annee + "-" + mois1 + mois2 + "-" + jour1 + jour2
        Dim time As DateTime = DateTime.Parse(value)


        Dim dt2 As DateTime = time
        Dim ts As TimeSpan = dt2.Subtract(dt1)

        Dim nbr = Convert.ToInt32(ts.Days) + 1

        Dim nom_fournisseur As String = Recupere_description(fournisseur)
        'Corrigé 31 aout

        nbr = nbr - 1

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            Select Case nbr
                Case Is < 7
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "',0, '" & total & "', 0, 0, 0, 0, 0)"
                Case 7 To 13
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "', '" & DateCourant & "', '" & time & "','" & cheque & "' ,0, 0, '" & total & "', 0, 0, 0, 0)"
                Case 14 To 20
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "',0, 0, 0, '" & total & "', 0, 0, 0)"
                Case 21 To 27
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "',0, 0, 0, 0, '" & total & "', 0, 0)"
                Case 28 To 30
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "','" & cheque & "',0, 0, 0, 0, 0, '" & total & "', 0)"
                Case Is >= 31
                    sqlStr = "INSERT INTO check_emis (numFour, nom,quand,date_paiement,no_cheque,zero, un,deux,trois,quatre,cinq, six) VALUES('" & fournisseur & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "','" & cheque & "',0, 0, 0, 0, 0, 0,'" & total & "')"
            End Select



            cmd = New OleDb.OleDbCommand(sqlStr, con)

            cmd.ExecuteNonQuery()




            enregistrer = True



            cmd.Dispose()

            con.Close()

        Catch ex As Exception
            MessageBox.Show("Cheque : " + cheque + " Error occurred: " & ex.Message)

        End Try

    End Sub


    Private Sub Ajouter_transactions_direct(ByVal fournisseur As String, ByVal calendar As String, ByVal cheque As String, ByVal total As String)

        Dim dt1 As DateTime = DateCourant


        Dim annee As String



        Dim c2 As String = calendar
        Dim c3 As String = calendar

        annee = calendar.Substring(0, 4)


        Dim mois1 As Char
        Dim mois2 As Char
        mois1 = calendar.Chars(4) ' myChar = "D"
        mois2 = calendar.Chars(5) ' myChar = "D"

        Dim jour1 As Char
        Dim jour2 As Char
        jour1 = calendar.Chars(6) ' myChar = "D"
        jour2 = calendar.Chars(7) ' myChar = "D"



        Dim value As String = annee + "-" + mois1 + mois2 + "-" + jour1 + jour2
        Dim time As DateTime = DateTime.Parse(value)


        Dim dt2 As DateTime = time
        Dim ts As TimeSpan = dt2.Subtract(dt1)

        Dim nbr = Convert.ToInt32(ts.Days) + 1

        Dim nom_fournisseur = Recupere_description(fournisseur)
        'Corrigé 31 aout

        nbr = nbr - 1

        'Dim newValueConverted As Integer = Val(fournisseur)

        Dim newValueConverted As String = fournisseur

        If cheque = "" Then
            cheque = "0"
        End If

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            Select Case nbr
                Case Is < 7
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "', '" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "', '" & total & "', 0, 0, 0, 0, 0)"
                Case 7 To 13
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "','" & nom_fournisseur & "', '" & DateCourant & "', '" & time & "','" & cheque & "' , 0, '" & total & "', 0, 0, 0, 0)"
                Case 14 To 20
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "', 0, 0, '" & total & "', 0, 0, 0)"
                Case 21 To 27
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "', '" & cheque & "', 0, 0, 0, '" & total & "', 0, 0)"
                Case 28 To 30
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "','" & cheque & "', 0, 0, 0, 0, '" & total & "', 0)"
                Case Is >= 31
                    sqlStr = "INSERT INTO check_direct (numFour, nom,quand,date_paiement,num_paiem, un,deux,trois,quatre,cinq, six) VALUES('" & newValueConverted & "','" & nom_fournisseur & "','" & DateCourant & "' , '" & time & "','" & cheque & "', 0, 0, 0, 0, 0,'" & total & "')"
            End Select



            cmd = New OleDb.OleDbCommand(sqlStr, con)

            cmd.ExecuteNonQuery()




            enregistrer = True



            cmd.Dispose()

            con.Close()

        Catch ex As Exception
            MessageBox.Show("Paiement no " + cheque + "  et montant : " + total + " : Error occurred: " & ex.Message)

        End Try

    End Sub



    Private Sub ListeDesTransactionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesTransactionsToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmTransactions()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            frmInterface.Panel1.Controls.Add(fc)
            'ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = False
            'frmInterface.btm_lister_paiement_preautorise.Enabled = False



        End If
    End Sub

    Function Numero_existant(ByVal numero As String) As Boolean

        Dim trouve As Boolean = False
        Dim id As Integer
        id = Val(numero)




        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            sqlStr = "Select * from check_emis where no_cheque=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            Dim dr As System.Data.OleDb.OleDbDataReader
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then

                trouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

        Return trouve

    End Function


    Function Fournisseur_existant(ByVal numero As String) As Boolean

        Dim trouve As Boolean = False




        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            sqlStr = "Select * from employeur where numFournisseur = '" & numero & "'"

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            Dim dr As System.Data.OleDb.OleDbDataReader
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then

                trouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

        Return trouve

    End Function

    Function Numero_existant_d(ByVal numero As String) As Boolean

        Dim trouve As Boolean = False
        Dim id As Integer
        id = Val(numero)




        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            sqlStr = "Select * from check_direct where num_paiem=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            Dim dr As System.Data.OleDb.OleDbDataReader
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then

                trouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

        Return trouve

    End Function


    Function Numero_existant_cheque(ByVal numero As String) As Boolean

        Dim trouve As Boolean = False
        Dim id As Integer
        id = Val(numero)




        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()


            sqlStr = "Select * from list_cheque where numCheque=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            Dim dr As System.Data.OleDb.OleDbDataReader
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then

                trouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

        Return trouve

    End Function

    Function Numero_existant_direct(ByVal numero As String) As Boolean

        Dim trouve As Boolean = False
        Dim id As Integer
        id = Val(numero)




        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()

            sqlStr = "Select * from list_cheque where numPaiem=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            Dim dr As System.Data.OleDb.OleDbDataReader
            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then

                trouve = True

            End If


            con.Close()
            cmd.Dispose()

        Catch ex As Exception
            MessageBox.Show("could Not find record")

        End Try

        Return trouve

    End Function

    Private Sub enlever_fournisseur(ByVal numFourn As String)

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM employeur WHERE numFournisseur = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numFourn)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            enregistrer = True

            myConnection.Close()
            myCommand.Dispose()


            'supprimer_cheque(obtenir_numCheque(i.Text))
        Catch ex As Exception

        End Try

    End Sub


    Private Sub enlever_cheque(ByVal numCheq As Integer)

        Try


            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE no_cheque = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numCheq)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            enregistrer = True

            myConnection.Close()
            myCommand.Dispose()


            'supprimer_cheque(obtenir_numCheque(i.Text))
        Catch ex As Exception

        End Try

    End Sub

    Private Sub enlever_paiement_direct(ByVal numPaiem As Integer)

        Try


            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_direct WHERE num_paiem = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", numPaiem)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            enregistrer = True

            myConnection.Close()
            myCommand.Dispose()


            'supprimer_cheque(obtenir_numCheque(i.Text))
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListeDesChèquesSupprimésToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesChèquesSupprimésToolStripMenuItem.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                frmInterface.lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            frmInterface.Panel1.Controls.Clear()

            frmInterface.lister_button(True)

            Dim fc As New frmRestore()
            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)

            frmInterface.Panel1.Controls.Add(fc)
            'ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = False
            'frmInterface.btm_lister_paiement_preautorise.Enabled = False



        End If
    End Sub

    Private Sub ListeDesRevenusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesRevenusToolStripMenuItem.Click
        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                FrmInterface.Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            FrmInterface.Panel1.Controls.Clear()

            FrmInterface.Lister_button(True)

            Dim fc As New frmAfficherListeRevenu()

            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            FrmInterface.Panel1.Controls.Add(fc)
            ListeDesPaiementsDirectsToolStripMenuItem.Enabled = True
            FrmInterface.btm_lister_paiement_direct.Enabled = True



        End If



        '********************************************************************

    End Sub

    Private Sub ListeDesDepensesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListeDesDepensesToolStripMenuItem.Click
        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                FrmInterface.Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            FrmInterface.Panel1.Controls.Clear()

            FrmInterface.Lister_button(True)

            Dim fc As New frmAfficherListeDepense()

            'InitializeComponent()
            fc.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            fc.Visible = True
            fc.Location = New Point(1, 0)

            'frmInterface.Panel1.Remove(fc)


            FrmInterface.Panel1.Controls.Add(fc)
            ListeDesPaiementsDirectsToolStripMenuItem.Enabled = True
            FrmInterface.btm_lister_paiement_direct.Enabled = True



        End If
    End Sub
End Class