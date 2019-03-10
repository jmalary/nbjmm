Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization
Imports System.Collections 'For arraylist

Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.html
Imports iTextSharp.text.html.simpleparser
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Public Class FrmInterface

    Dim myConnString As String = GlobalVariables.test2ConnectionString
    Dim runn_global As Boolean
    Dim lister As New ArrayList
    Dim compte As Integer = 0
    Dim no_lign As Integer
    Dim del_chaine(100) As String
    Dim tableau() As String
    Dim tab_rep() As String
    Dim con As System.Data.OleDb.OleDbConnection
    Dim sqlStr As String
    Dim cmd As System.Data.OleDb.OleDbCommand



    Dim CW As Integer = Me.Width ' Current Width
    Dim CH As Integer = Me.Height ' Current Height
    Dim IW As Integer = Me.Width ' Initial Width
    Dim IH As Integer = Me.Height ' Initial Height
    Dim nom_ficbier As String


    Private Sub FrmInterface_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IW = Me.Width
        IH = Me.Height

        Dim milieu As Integer = (IW - Label1.Width) / 2

        Label1.Location = New Point(milieu, 68)


        'Me.Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width

        'Me.Height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height

        Dim longueur As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 40
        Dim hauteur As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - 250



        Panel1.Size = New Size(longueur, hauteur)

        Dim nbr As Integer = Label1.Width
        Dim center As Integer = Screen.PrimaryScreen.Bounds.Width - Label1.Width

        center = center \ 2

        Label1.Location = New Point(center, 68)

        center = Screen.PrimaryScreen.Bounds.Width - Label2.Width
        center = center \ 2

        Label2.Location = New Point(center, 103)



        If backup = True Then
            'Button1.Enabled = False

            'Btn_Liste_Cheque.Enabled = False
            'Button2.Enabled = False
            'Btm_Paiement_direct.Enabled = False
            'Button3.Enabled = False
            'Btm_Preautorise.Enabled = False
            'Label3.Text = "Archive ... "
            'Label5.Text = ""
            'btm_archive.Enabled = True
        Else


            'Label3.Text = ""
            'Button1.Enabled = True
            'Btn_Liste_Cheque.Enabled = True
            'Button2.Enabled = True
            'Btm_Paiement_direct.Enabled = True
            'Button3.Enabled = True
            'Btm_Preautorise.Enabled = True
            'btm_archive.Enabled = False

        End If



        Dim string_date As String


        If backup = False Then


            frmMenu.aff_miseajour()


            Proc_load() 'preautorise
            Proc_load_cheque()
            Proc_load_direct()
            Proc_load_revenue()
            Proc_load_depense()

            runn_global = True


            While runn_global

                runn_global = False

            Proc_load_db()

            End While

            runn_global = True


        Dim list As New ArrayList

        While runn_global
            runn_global = False

            'Interface  dont work sur revenu
            Proc_load_db_revenu()
        End While

        runn_global = True


        While runn_global
            runn_global = False

            Proc_load_db_depense()

            'frmPaiement.proc_load_depense()
        End While


            'Supprimer automatiquement les paiement qui depassé la date

            Proc_load_suppprimer_automatiqu()
            Proc_load_suppprimer_automatique_direct()
            Proc_load_suppprimer_automatique_cheque()
            Proc_load_suppprimer_automatique_revenu()
            Proc_load_suppprimer_automatique_depense()

            'End While



            proc_repeat_preautorise()
            proc_repeat_depense()
            proc_repeat_revenu()


            changer_date = False

        End If






        string_date = DateCourant.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA"))



        Label1.Text = string_date

        Dim path As String = Application.StartupPath






        'SavePDF()



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Dim from_var As New frmEnregistrer()
        'InitializeComponent()
        from_var.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        from_var.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        from_var.Visible = True
        from_var.Location = New Point(1, 0)

        Panel1.Controls.Add(from_var)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

        FermerIcon()

        If faire_vider = False Then

            Panel1.Controls.Clear()

            Dim f As New frmCheque
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_cheque.Enabled = False

            fenetre = "cheque"


        End If





    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)



        FermerIcon()


        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show("Vous voulez garder les données", "Enregistrer", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False

            ElseIf result = DialogResult.Yes Then

            End If


        End If





        If faire_vider = False Then

            Panel1.Controls.Clear()



            Dim f As New frmPaiement()
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_paiement.Enabled = False
        End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)


        FermerIcon()


        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show("Vous voulez garder les données", "Enregistrer", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If


        If faire_vider = False Then

            Panel1.Controls.Clear()

            Dim f As New frmAfficherListeCheque()
            'InitializeComponent()

            f.TopLevel = False
            f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True


            Panel1.Controls.Add(f)


            btm_lister_cheque.Enabled = False
            Form_windows.ListeDesChèquesToolStripMenuItem.Enabled = False

        End If







    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

        Dim f As New frmAfficherListePaiementDirect()
        'InitializeComponent()

        f.TopLevel = False
        f.WindowState = FormWindowState.Maximized
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.Visible = True


        Panel1.Controls.Add(f)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

        Dim f As New frmAfficherListePaiementPreautorise()
        'InitializeComponent()

        f.TopLevel = False
        f.WindowState = FormWindowState.Maximized
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.Visible = True


        Panel1.Controls.Add(f)
    End Sub

    Sub Lister_button(ByVal active As Boolean)

        Dim rep As Boolean = False

        If active = True Then
            rep = True
        End If


        btm_enregistrer.Enabled = rep


        btm_fournisseur.Enabled = rep


        btm_cheque.Enabled = rep

        btm_paiement.Enabled = rep

        btm_lister_cheque.Enabled = rep
        Form_windows.ListeDesChèquesToolStripMenuItem.Enabled = rep


        btm_lister_paiement_direct.Enabled = rep
        Form_windows.ListeDesPaiementsDirectsToolStripMenuItem.Enabled = rep

        btm_lister_paiement_preautorise.Enabled = rep
        Form_windows.ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = rep



    End Sub

    Private Sub Btm_aide_Click(sender As Object, e As EventArgs)



        Me.Hide()
        frmLoading.Show()
        frmGestion.MdiParent = Form_windows
        frmGestion.Show()



    End Sub

    Private Sub Btm_backup_Click(sender As Object, e As EventArgs)

        ToolStripButton8.Enabled = False
        Label2.Text = "Initialisation terminée"
        Label2.ForeColor = Color.Black


    End Sub


    Private Sub Proc_load_cheque()



        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_emis ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()



            While myReader1.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

                Dim valueIntConverted As Integer = CInt(myReader1.Item(0).ToString)

                Dim montant As String = ""

                Dim runningVB As Boolean
                Dim cpt As Integer = 7


                runningVB = True

                'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
                'identifer quel cpt 

                While runningVB

                    montant = myReader1.Item(cpt).ToString()

                    If montant = 0 Then
                        cpt += 1
                    Else
                        runningVB = False

                    End If



                End While


                'Dim dt1 As DateTime = DateCourant
                Dim dt1 As DateTime = DateCourant



                Dim dt2 As DateTime = Convert.ToDateTime(myReader1.Item(4))
                Dim ts As TimeSpan = dt2.Subtract(dt1)

                'MsgBox(myReader1.Item(0))

                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1


                Dim myCommand As New OleDb.OleDbCommand("Update check_emis set un = @un, deux = @deux, trois = @trois, quatre = @quatre, cinq = @cinq, six = @six WHERE numero = @numero", myConnection)


                Select Case nbr
                    Case Is <= 6
                        myCommand.Parameters.AddWithValue("@un", montant)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 7 To 13
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", montant)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)

                    Case 14 To 20
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", montant)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)

                    Case 21 To 27
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", montant)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)

                    Case 28 To 30
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", montant)
                        myCommand.Parameters.AddWithValue("@six", 0)

                    Case Is >= 31
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", montant)

                End Select


                myCommand.Parameters.AddWithValue("@numero", valueIntConverted)


                myCommand.ExecuteNonQuery()


                myCommand.Dispose()


            End While


            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()



        Catch ex As Exception

        End Try


    End Sub

    Private Sub Proc_load_direct()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_direct ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()



            While myReader1.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

                Dim valueIntConverted As Integer = CInt(myReader1.Item(0).ToString)

                Dim montant As String = ""

                Dim runningVB As Boolean
                Dim cpt As Integer = 6


                runningVB = True

                'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
                'identifer quel cpt 

                While runningVB

                    montant = myReader1.Item(cpt).ToString()

                    If montant = 0 Then
                        cpt += 1
                    Else
                        runningVB = False

                    End If



                End While


                'Dim dt1 As DateTime = DateCourant
                Dim dt1 As DateTime = DateCourant



                Dim dt2 As DateTime = Convert.ToDateTime(myReader1.Item(4))
                Dim ts As TimeSpan = dt2.Subtract(dt1)


                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1


                Dim myCommand As New OleDb.OleDbCommand("Update check_direct set un = @un, deux = @deux, trois = @trois, quatre = @quatre, cinq = @cinq, six = @six WHERE numero = @numero", myConnection)


                Select Case nbr
                    Case Is <= 6
                        myCommand.Parameters.AddWithValue("@un", montant)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 7 To 13
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", montant)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 14 To 20
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", montant)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 21 To 27
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", montant)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 28 To 30
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", montant)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case Is >= 31
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", montant)
                End Select


                myCommand.Parameters.AddWithValue("@numero", valueIntConverted)


                myCommand.ExecuteNonQuery()

                myCommand.Dispose()


            End While


            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()



        Catch ex As Exception

        End Try

    End Sub

    Private Sub Proc_load_revenue()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_revenu ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()



            While myReader1.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

                Dim valueIntConverted As Integer = CInt(myReader1.Item(0).ToString)

                Dim montant As String = ""

                Dim runningVB As Boolean
                Dim cpt As Integer = 4


                runningVB = True

                'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
                'identifer quel cpt 

                While runningVB

                    montant = myReader1.Item(cpt).ToString()

                    If montant = 0 Then
                        cpt += 1
                    Else
                        runningVB = False

                    End If



                End While


                'Dim dt1 As DateTime = DateCourant
                Dim dt1 As DateTime = DateCourant



                Dim dt2 As DateTime = Convert.ToDateTime(myReader1.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)



                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1


                Dim myCommand As New OleDb.OleDbCommand("Update check_revenu set un = @un, deux = @deux, trois = @trois, quatre = @quatre, cinq = @cinq, six = @six WHERE numero = @numero", myConnection)


                Select Case nbr
                    Case Is <= 6
                        myCommand.Parameters.AddWithValue("@un", montant)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 7 To 13
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", montant)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 14 To 20
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", montant)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 21 To 27
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", montant)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 28 To 30
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", montant)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case Is >= 31
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", montant)
                End Select


                myCommand.Parameters.AddWithValue("@numero", valueIntConverted)


                myCommand.ExecuteNonQuery()

                'MsgBox("Update !!! ")
                myCommand.Dispose()


            End While


            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()



        Catch ex As Exception

        End Try

    End Sub

    Private Sub Proc_load_depense()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_depenses   ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()



            While myReader1.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

                Dim valueIntConverted As Integer = CInt(myReader1.Item(0).ToString)

                Dim montant As String = ""

                Dim runningVB As Boolean
                Dim cpt As Integer = 4


                runningVB = True

                'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
                'identifer quel cpt 

                While runningVB

                    montant = myReader1.Item(cpt).ToString()

                    If montant = 0 Then
                        cpt += 1
                    Else
                        runningVB = False

                    End If



                End While


                'Dim dt1 As DateTime = DateCourant
                Dim dt1 As DateTime = DateCourant



                Dim dt2 As DateTime = Convert.ToDateTime(myReader1.Item(4))
                Dim ts As TimeSpan = dt2.Subtract(dt1)

                'MsgBox(myReader1.Item(0))

                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1


                Dim myCommand As New OleDb.OleDbCommand("Update check_depenses set un = @un, deux = @deux, trois = @trois, quatre = @quatre, cinq = @cinq, six = @six WHERE numero = @numero", myConnection)


                Select Case nbr
                    Case Is <= 6
                        myCommand.Parameters.AddWithValue("@un", montant)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 7 To 13
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", montant)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 14 To 20
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", montant)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 21 To 27
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", montant)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 28 To 30
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", montant)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case Is >= 31
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", montant)
                End Select


                myCommand.Parameters.AddWithValue("@numero", valueIntConverted)


                myCommand.ExecuteNonQuery()


                myCommand.Dispose()


            End While


            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()



        Catch ex As Exception

        End Try

    End Sub

    Sub Proc_load()


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            'Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_preautorises WHERE payer = '" & "oui" & "'", myConnection)
            Dim myCommand1 As New OleDb.OleDbCommand("SELECT * FROM check_preautorises ", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")

                Dim valueIntConverted As Integer = CInt(myReader1.Item(0).ToString)

                Dim montant As String = ""

                Dim runningVB As Boolean
                Dim cpt As Integer = 4


                runningVB = True


                While runningVB

                    montant = myReader1.Item(cpt).ToString()

                    If montant = 0 Then
                        cpt += 1
                    Else
                        runningVB = False

                    End If



                End While


                'Dim dt1 As DateTime = DateCourant
                Dim dt1 As DateTime = DateCourant



                Dim dt2 As DateTime = Convert.ToDateTime(myReader1.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)



                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1


                Dim myCommand As New OleDb.OleDbCommand("Update check_preautorises set un = @un, deux = @deux, trois = @trois, quatre = @quatre, cinq = @cinq, six = @six WHERE numero = @numero", myConnection)






                Select Case nbr
                    Case Is <= 6
                        myCommand.Parameters.AddWithValue("@un", montant)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 7 To 13
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", montant)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 14 To 20
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", montant)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 21 To 27
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", montant)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case 28 To 30
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", montant)
                        myCommand.Parameters.AddWithValue("@six", 0)
                    Case Is >= 31
                        'Case Else
                        myCommand.Parameters.AddWithValue("@un", 0)
                        myCommand.Parameters.AddWithValue("@deux", 0)
                        myCommand.Parameters.AddWithValue("@trois", 0)
                        myCommand.Parameters.AddWithValue("@quatre", 0)
                        myCommand.Parameters.AddWithValue("@cinq", 0)
                        myCommand.Parameters.AddWithValue("@six", montant)
                End Select


                myCommand.Parameters.AddWithValue("@numero", valueIntConverted)


                myCommand.ExecuteNonQuery()


                myCommand.Dispose()



            End While


            myConnection.Close()
            myCommand1.Dispose()
            myReader1.Close()




        Catch ex As Exception

            MsgBox(ex.Message)
        End Try


    End Sub


    Sub Proc_load_db()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM mensuel", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            Dim str_mensuel As String = ""

            lister = New ArrayList()

            compte = 0

            While myReader.Read

                no_lign = myReader.GetInt32(0)

                str_mensuel = str_mensuel + myReader.GetString(1) + "?"
                str_mensuel = str_mensuel + myReader.GetString(2) + "?"

                del_chaine(compte) = myReader.GetString(2)
                str_mensuel = str_mensuel + myReader.GetString(3) + "/"
                str_mensuel = str_mensuel + myReader.GetString(4) + "/"
                str_mensuel = str_mensuel + Convert.ToString(myReader.GetInt32(0))

                lister.Add(str_mensuel)
                str_mensuel = ""
                compte += 1


            End While


            compte = 0

            'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
            'identifer quel cpt 




            For Each num In lister

                tableau = num.Split("?")
                Proc_noday_mensuel()
                compte += 1
            Next









            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try
        'MessageBox.Show("Dot Net Perls is awesome.222")

    End Sub






    Sub Proc_load_db_revenu()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM revenu", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            Dim str_mensuel As String = ""

            lister = New ArrayList()

            compte = 0

            While myReader.Read

                no_lign = myReader.GetInt32(0)

                str_mensuel = str_mensuel + myReader.GetString(1) + "?"
                str_mensuel = str_mensuel + myReader.GetString(5) + "?"

                del_chaine(compte) = myReader.GetString(5)
                str_mensuel = str_mensuel + myReader.GetDouble(4).ToString() + "/"
                str_mensuel = str_mensuel + myReader.GetDateTime(2).ToString() + "/"
                str_mensuel = str_mensuel + Convert.ToString(myReader.GetInt32(0))

                lister.Add(str_mensuel)
                str_mensuel = ""
                compte += 1


            End While


            compte = 0

            'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
            'identifer quel cpt 




            For Each num In lister

                tableau = num.Split("?")
                Proc_noday_mensuel_revenu()
                compte += 1
            Next



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try
        'MessageBox.Show("Dot Net Perls is awesome.222")

    End Sub

    Sub Proc_load_db_depense()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM depense", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            Dim str_mensuel As String = ""

            lister = New ArrayList()

            compte = 0

            While myReader.Read

                no_lign = myReader.GetInt32(0)

                str_mensuel = str_mensuel + myReader.GetString(1) + "?"
                str_mensuel = str_mensuel + myReader.GetString(5) + "?"

                del_chaine(compte) = myReader.GetString(5)
                str_mensuel = str_mensuel + myReader.GetDouble(4).ToString() + "/"
                str_mensuel = str_mensuel + myReader.GetDateTime(2).ToString() + "/"
                str_mensuel = str_mensuel + Convert.ToString(myReader.GetInt32(0))

                lister.Add(str_mensuel)
                str_mensuel = ""
                compte += 1


            End While


            compte = 0

            'Trouver un montant  pas zero  dans liste des 5 colonnes  courant, 7jrs, 14
            'identifer quel cpt 




            For Each num In lister

                tableau = num.Split("?")
                Proc_noday_mensuel_depense()
                compte += 1
            Next









            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try
        'MessageBox.Show("Dot Net Perls is awesome.222")

    End Sub



    Sub Proc_noday_mensuel()

        Dim offset = New Date(1, 1, 1)

        Dim laDate As String = tableau(1)


        Dim result As String

        result = Proc_arret_paiement(tableau(2))


        Dim time As Date = Date.Parse(laDate)

        Dim diff As TimeSpan = time - DateCourant


        Dim days = diff.Days

        'Dim days As Double = span.TotalDays





        Dim lngth As Integer = tableau.Length()

        Dim lstrLastValue As String = tableau(lngth - 1)



        tab_rep = lstrLastValue.Split("/")


        Dim value As Integer = Integer.Parse(days.ToString)

        no_lign = Integer.Parse(tab_rep(2))


        If value < 31 Then



            Try
                con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
                con.Open()
                sqlStr = ""
                Select Case value
                    Case Is <= 6
                        'Case 0 To 6
                        sqlStr = "INSERT INTO check_preautorises (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "','" & tab_rep(0) & "', 0, 0, 0, 0,'" & no_lign & "')"
                    Case 7 To 13
                        sqlStr = "INSERT INTO check_preautorises (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, '" & tab_rep(0) & "', 0, 0, 0,'" & no_lign & "')"
                    Case 14 To 20
                        sqlStr = "INSERT INTO check_preautorises (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, '" & tab_rep(0) & "', 0, 0,'" & no_lign & "')"
                    Case 21 To 27
                        sqlStr = "INSERT INTO check_preautorises (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, '" & tab_rep(0) & "', 0,'" & no_lign & "')"
                    Case 28 To 30
                        'Case Else
                        sqlStr = "INSERT INTO check_preautorises (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, 0, '" & tab_rep(0) & "','" & no_lign & "')"
                End Select

                cmd = New OleDb.OleDbCommand(sqlStr, con)

                cmd.ExecuteNonQuery()


                If result = "termine" Then
                    Proc_remove_paiment()

                Else
                    Proc_remove_string()

                End If



                cmd.Dispose()
                con.Close()
            Catch ex As Exception
                'MessageBox.Show("Vous devez choisir un fournisseur et indiquer le montant")


            End Try
            'proc_load_db()


        End If

    End Sub

    Sub Proc_noday_mensuel_revenu()

        Dim offset = New Date(1, 1, 1)

        Dim laDate As String = tableau(1)


        Dim result As String

        result = Proc_arret_paiement(tableau(2))


        Dim time As Date = Date.Parse(laDate)

        Dim diff As TimeSpan = time - DateCourant


        Dim days = diff.Days

        'Dim days As Double = span.TotalDays




        Dim lngth As Integer = tableau.Length()

        Dim lstrLastValue As String = tableau(lngth - 1)



        tab_rep = lstrLastValue.Split("/")


        Dim value As Integer = Integer.Parse(days.ToString)

        no_lign = Integer.Parse(tab_rep(2))


        If value < 31 Then



            Try
                con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
                con.Open()

                Select Case value
                    Case Is <= 6
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "','" & tab_rep(0) & "', 0, 0, 0, 0,'" & no_lign & "')"
                    Case 7 To 13
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, '" & tab_rep(0) & "', 0, 0, 0,'" & no_lign & "')"
                    Case 14 To 20
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, '" & tab_rep(0) & "', 0, 0,'" & no_lign & "')"
                    Case 21 To 27
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, '" & tab_rep(0) & "', 0,'" & no_lign & "')"
                    Case 28 To 30
                        'Case Else
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, 0, '" & tab_rep(0) & "','" & no_lign & "')"
                End Select

                cmd = New OleDb.OleDbCommand(sqlStr, con)

                cmd.ExecuteNonQuery()

                If result = "termine" Then
                    Proc_remove_paiment_revenu()

                Else
                    Proc_remove_string_revenu_revenu()

                End If



                cmd.Dispose()
                con.Close()
            Catch ex As Exception
                'MessageBox.Show("Vous devez choisir un fournisseur et indiquer le montant")


            End Try
            'proc_load_db()


        End If

    End Sub


    Sub Proc_noday_mensuel_depense()

        Dim offset = New Date(1, 1, 1)

        Dim laDate As String = tableau(1)


        Dim result As String

        result = Proc_arret_paiement(tableau(2))


        Dim time As Date = Date.Parse(laDate)

        Dim diff As TimeSpan = time - DateCourant


        Dim days = diff.Days

        'Dim days As Double = span.TotalDays




        Dim lngth As Integer = tableau.Length()

        Dim lstrLastValue As String = tableau(lngth - 1)



        tab_rep = lstrLastValue.Split("/")


        Dim value As Integer = Integer.Parse(days.ToString)

        no_lign = Integer.Parse(tab_rep(2))


        If value < 31 Then

            sqlStr = ""

            Try
                con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
                con.Open()

                Select Case value
                    Case Is <= 6
                        sqlStr = "INSERT INTO check_depenses (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "','" & tab_rep(0) & "', 0, 0, 0, 0,'" & no_lign & "')"
                    Case 7 To 13
                        sqlStr = "INSERT INTO check_depenses (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, '" & tab_rep(0) & "', 0, 0, 0,'" & no_lign & "')"
                    Case 14 To 20
                        sqlStr = "INSERT INTO check_depenses (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, '" & tab_rep(0) & "', 0, 0,'" & no_lign & "')"
                    Case 21 To 27
                        sqlStr = "INSERT INTO check_depenses (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, '" & tab_rep(0) & "', 0,'" & no_lign & "')"
                    Case 28 To 30
                        'Case Else
                        sqlStr = "INSERT INTO check_depenses (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, 0, '" & tab_rep(0) & "','" & no_lign & "')"
                End Select

                cmd = New OleDb.OleDbCommand(sqlStr, con)

                cmd.ExecuteNonQuery()

                If result = "termine" Then
                    Proc_remove_paiment_depense()

                Else
                    Proc_remove_string_revenu_depense()

                End If



                cmd.Dispose()
                con.Close()
            Catch ex As Exception
                'MessageBox.Show("Vous devez choisir un fournisseur et indiquer le montant")


            End Try
            'proc_load_db()


        End If

    End Sub


    Sub Proc_noday_mensuel_revenue()

        Dim offset = New Date(1, 1, 1)

        Dim laDate As String = tableau(1)


        Dim result As String

        result = Proc_arret_paiement(tableau(2))


        Dim time As Date = Date.Parse(laDate)

        Dim diff As TimeSpan = time - DateCourant


        Dim days = diff.Days

        'Dim days As Double = span.TotalDays



        Dim lngth As Integer = tableau.Length()

        Dim lstrLastValue As String = tableau(lngth - 1)



        tab_rep = lstrLastValue.Split("/")


        Dim value As Integer = Integer.Parse(days.ToString)

        no_lign = Integer.Parse(tab_rep(2))


        If value < 31 Then



            Try
                con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
                con.Open()

                Select Case value
                    Case Is <= 6
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "','" & tab_rep(0) & "', 0, 0, 0, 0,'" & no_lign & "')"
                    Case 7 To 13
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, '" & tab_rep(0) & "', 0, 0, 0,'" & no_lign & "')"
                    Case 14 To 20
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, '" & tab_rep(0) & "', 0, 0,'" & no_lign & "')"
                    Case 21 To 27
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, '" & tab_rep(0) & "', 0,'" & no_lign & "')"
                    Case 28 To 30
                        'Case Else
                        sqlStr = "INSERT INTO check_revenu (nom,quand,date_paiement,un,deux,trois,quatre,cinq, payer) VALUES('" & tableau(0) & "','" & DateCourant & "' , '" & tableau(1) & "',0, 0, 0, 0, '" & tab_rep(0) & "','" & no_lign & "')"
                End Select

                cmd = New OleDb.OleDbCommand(sqlStr, con)

                cmd.ExecuteNonQuery()


                If result = "termine" Then
                    Proc_remove_paiment_revenu()

                Else
                    Proc_remove_string_revenu_revenu()

                End If



                cmd.Dispose()
                con.Close()
            Catch ex As Exception
                'MessageBox.Show("Vous devez choisir un fournisseur et indiquer le montant")


            End Try
            'proc_load_db()


        End If

    End Sub

    Private Function Proc_arret_paiement(ByVal charactere As String) As String


        Dim sqlStr As String
        Dim find As String = "-"
        Dim find2 As String = "/"


        Dim FirstCharacter As Integer = charactere.IndexOf(find)
        Dim FirstCharacter2 As Integer = charactere.IndexOf(find2)


        If FirstCharacter = "-1" And FirstCharacter2 <> "-1" Then
            sqlStr = "termine"
        Else
            sqlStr = "continue"

        End If


        Return sqlStr


    End Function

    Private Sub Proc_remove_paiment()


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM mensuel WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", no_lign)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            runn_global = True

            myCommand.Dispose()


        Catch ex As Exception

        End Try



    End Sub

    Private Sub Proc_remove_paiment_revenu()


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM revenu WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", no_lign)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            runn_global = True

            myCommand.Dispose()


        Catch ex As Exception

        End Try



    End Sub

    Private Sub Proc_remove_paiment_depense()


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("DELETE FROM depense WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@numero", no_lign)
            myCommand.ExecuteNonQuery()
            myConnection.Close()

            runn_global = True

            myCommand.Dispose()


        Catch ex As Exception

        End Try



    End Sub

    Private Sub Proc_remove_string()

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim rep_chaine As String

            rep_chaine = del_chaine(compte)


            rep_chaine = rep_chaine.Remove(0, 11)


            Dim myCommand As New OleDb.OleDbCommand("Update mensuel set date_mens = @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", rep_chaine)
            myCommand.Parameters.AddWithValue("@numero", no_lign)


            myCommand.ExecuteNonQuery()
            myConnection.Close()


            myCommand.Dispose()
            oledbAdapter.Dispose()


            runn_global = True

        Catch ex As Exception

        End Try



    End Sub

    Private Sub Proc_remove_string_revenu_revenu()

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim rep_chaine As String

            rep_chaine = del_chaine(compte)


            rep_chaine = rep_chaine.Remove(0, 11)


            Dim myCommand As New OleDb.OleDbCommand("Update revenu set sequence = @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", rep_chaine)
            myCommand.Parameters.AddWithValue("@numero", no_lign)


            myCommand.ExecuteNonQuery()
            myConnection.Close()


            myCommand.Dispose()
            oledbAdapter.Dispose()


            runn_global = True

        Catch ex As Exception

        End Try



    End Sub

    Private Sub Proc_remove_string_revenu_depense()

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim rep_chaine As String

            rep_chaine = del_chaine(compte)


            rep_chaine = rep_chaine.Remove(0, 11)


            Dim myCommand As New OleDb.OleDbCommand("Update depense set sequence = @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", rep_chaine)
            myCommand.Parameters.AddWithValue("@numero", no_lign)


            myCommand.ExecuteNonQuery()
            myConnection.Close()


            myCommand.Dispose()
            oledbAdapter.Dispose()


            runn_global = True

        Catch ex As Exception

        End Try



    End Sub


    Private Sub FermerIcon()


        faire_vider = False



        For Each ctrl As Control In Panel1.Controls



            If ctrl.GetType.ToString = "Comptable_NBJMM.frmCheque" Then


                'If ctrl.HasChildren Then
                '    ' Recursively call this method for each child control.
                '    Dim childControl As Control
                '    For Each childControl In ctrl.Controls

                '    Next childControl
                'End If

                For Each c As Control In ctrl.Controls
                    If c.GetType Is GetType(TextBox) Then

                        If c.Text <> "" Then
                            faire_vider = True
                        End If

                    End If




                    If c.GetType Is GetType(ComboBox) Then

                        If c.Text <> "Choisir un employeur" Then
                            faire_vider = True
                        End If

                    End If


                    If c.GetType Is GetType(DateTimePicker) Then

                        If c.Text <> DateCourant.ToString("dd MMMM yyyy", CultureInfo.CreateSpecificCulture("fr-CA")) Then
                            faire_vider = True
                        End If



                    End If

                Next

                If faire_vider = False Then
                    btm_cheque.Enabled = True
                End If


            End If

            If ctrl.GetType.ToString = "Comptable_NBJMM.frmPaiement" Then


                For Each c As Control In ctrl.Controls





                    If c.GetType Is GetType(TextBox) Then


                        If c.Text <> "" Then

                            faire_vider = True
                        End If

                    End If

                    If c.GetType Is GetType(ComboBox) Then


                        'If c.Text <> "Choisir un employeur" Then
                        '    MsgBox(c.Text)
                        '    faire_vider = True
                        'End If


                        If c.Text <> "Choisir un employeur" And c.Text <> "choisir une date" And c.Text <> "choisir la fréquence" Then
                            'MsgBox(c.Text)
                            faire_vider = True
                        End If

                    End If


                    If c.GetType Is GetType(DateTimePicker) Then

                        If c.Text <> DateCourant.ToString("dd MMMM yyyy", CultureInfo.CreateSpecificCulture("fr-CA")) And c.Text <> "" Then
                            'MsgBox(c.Text)
                            'MsgBox(DateCourant.ToString("dd MMMM yyyy", CultureInfo.CreateSpecificCulture("fr-CA")))
                            faire_vider = True
                        End If



                    End If

                Next

                If faire_vider = False Then
                    btm_paiement.Enabled = True
                End If


            End If


            If ctrl.GetType.ToString = "Comptable_NBJMM.frmUpdate_Paiement_preautorise" Then




                For Each c As Control In ctrl.Controls


                    If c.GetType Is GetType(DateTimePicker) Then

                        If String.Compare(c.Text, date_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If

                    If c.GetType Is GetType(ComboBox) Then



                        If String.Compare(c.Text, combox_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If

                    End If


                    If c.GetType Is GetType(TextBox) Then

                        If String.Compare(c.Text, montant_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If


                Next

                If faire_vider = False Then
                    btm_lister_paiement_preautorise.Enabled = True
                    Form_windows.ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = True
                End If


            End If

            If ctrl.GetType.ToString = "Comptable_NBJMM.frmUpdate_Paiement_direct" Then




                For Each c As Control In ctrl.Controls


                    If c.GetType Is GetType(DateTimePicker) Then

                        If String.Compare(c.Text, date_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If

                    If c.GetType Is GetType(ComboBox) Then



                        If String.Compare(c.Text, combox_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If

                    End If


                    If c.GetType Is GetType(TextBox) Then

                        If String.Compare(c.Text, montant_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If


                Next

                If faire_vider = False Then
                    btm_lister_paiement_direct.Enabled = True
                    Form_windows.ListeDesPaiementsDirectsToolStripMenuItem.Enabled = True

                End If


            End If

            If ctrl.GetType.ToString = "Comptable_NBJMM.frmUpdate_check" Then




                For Each c As Control In ctrl.Controls


                    If c.GetType Is GetType(DateTimePicker) Then

                        If String.Compare(c.Text, date_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If

                    If c.GetType Is GetType(ComboBox) Then



                        If String.Compare(c.Text, combox_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If

                    End If


                    If c.GetType Is GetType(TextBox) Then

                        If String.Compare(c.Text, montant_temp) = 0 Then

                        Else
                            faire_vider = True
                        End If


                    End If


                Next

                If faire_vider = False Then
                    btm_lister_cheque.Enabled = True
                    Form_windows.ListeDesChèquesToolStripMenuItem.Enabled = True
                End If


            End If


        Next



        'MsgBox(frmCheque.TextBox2.Text)


        If frmCheque.TextBox2.Text <> "" Then

            Dim result1 As DialogResult = MessageBox.Show("Is Dot Net Perls awesome?",
                                                      "Important Question",
                                                      MessageBoxButtons.YesNo)
            '


        End If



    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub Btm_enregistrer_Click(sender As Object, e As EventArgs) Handles btm_enregistrer.Click

        Dim f As New frmEnregistrer()
        'InitializeComponent()
        f.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.Visible = True
        f.Location = New Point(1, 0)

        Panel1.Controls.Add(f)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)

        FermerIcon()

        If faire_vider = False Then

            Panel1.Controls.Clear()


            Dim f As New frmCheque()
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_cheque.Enabled = False

            fenetre = "cheque"


        End If



    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles btm_paiement.Click
        FermerIcon()


        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If





        If faire_vider = False Then

            Panel1.Controls.Clear()
            Lister_button(True)


            Dim f As New frmPaiement()
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_paiement.Enabled = False
        End If

    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click

        FermerIcon()


        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If





        If faire_vider = False Then

            Panel1.Controls.Clear()
            Lister_button(True)


            Me.Hide()
            frmLoading.Show()
            frmGestion.MdiParent = Form_windows
            frmGestion.Show()
        End If






    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Lister_button(True)
        ToolStripButton8.Enabled = False

        Label2.Text = "Initialisation terminée"
        Label2.ForeColor = Color.Black
        Form_windows.ListeDesArchivesToolStripMenuItem.Enabled = True
        ToolStripButton8.BackColor = Color.Transparent
        backup = False
        DateCourant = Date.Now

    End Sub

    Private Sub Btm_cheque_Click(sender As Object, e As EventArgs) Handles btm_cheque.Click
        FermerIcon()



        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            Panel1.Controls.Clear()
            Lister_button(True)
            Dim from_var As New frmCheque()
            'InitializeComponent()
            from_var.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            from_var.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            from_var.Visible = True
            from_var.Location = New Point(1, 0)

            Panel1.Controls.Add(from_var)

            btm_cheque.Enabled = False

            fenetre = "cheque"


        End If
    End Sub



    Private Sub Btm_lister_paiement_direct_Click(sender As Object, e As EventArgs) Handles btm_lister_paiement_direct.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            Panel1.Controls.Clear()

            Lister_button(True)

            Dim f As New frmAfficherListePaiementDirect()
            ' InitializeComponent()
            f.TopLevel = False
            f.Visible = True
            f.WindowState = FormWindowState.Normal
            Panel1.Controls.Add(f)
            'f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_lister_paiement_direct.Enabled = False
            Form_windows.ListeDesPaiementsDirectsToolStripMenuItem.Enabled = False

        End If
    End Sub



    Private Sub Btm_lister_cheque_Click(sender As Object, e As EventArgs) Handles btm_lister_cheque.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            Panel1.Controls.Clear()

            Lister_button(True)

            Dim f As New frmAfficherListeCheque()

            f.TopLevel = False

            f.Visible = True
            f.WindowState = FormWindowState.Maximized
            Panel1.Controls.Add(f)

            btm_lister_cheque.Enabled = False

            Form_windows.ListeDesChèquesToolStripMenuItem.Enabled = False


        End If

    End Sub

    Private Sub Btm_lister_paiement_preautorise_Click(sender As Object, e As EventArgs) Handles btm_lister_paiement_preautorise.Click
        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If





        If faire_vider = False Then

            Panel1.Controls.Clear()

            Lister_button(True)

            Dim f As New frmAfficherListePaiementPreautorise()
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_lister_paiement_preautorise.Enabled = False
            Form_windows.ListeDesPaiementsPréautoriséToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub Btm_fournisseur_Click(sender As Object, e As EventArgs) Handles btm_fournisseur.Click

        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If





        If faire_vider = False Then

            Panel1.Controls.Clear()
            Lister_button(True)


            Dim f As New frmEmployeur()
            'InitializeComponent()
            f.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            f.Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
            f.Height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height
            f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            f.Visible = True
            f.Location = New Point(1, 0)

            Panel1.Controls.Add(f)

            btm_fournisseur.Enabled = False
            Form_windows.ToolStripMenuItemFournisseur.Enabled = False

        End If




    End Sub

    'Private Sub Changer_resolution()

    '    Dim DesignScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    '    Dim DesignScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height


    '    Dim CurrentScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    '    Dim CurrentScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
    '    Dim RatioX As Double = CurrentScreenWidth / DesignScreenWidth
    '    Dim RatioY As Double = CurrentScreenHeight / DesignScreenHeight
    '    For Each iControl In Me.Controls
    '        With iControl
    '            If (.GetType.GetProperty("Width").CanRead) Then .Width = CInt(.Width * RatioX)
    '            If (.GetType.GetProperty("Height").CanRead) Then .Height = CInt(.Height * RatioY)
    '            If (.GetType.GetProperty("Top").CanRead) Then .Top = CInt(.Top * RatioX)
    '            If (.GetType.GetProperty("Left").CanRead) Then .Left = CInt(.Left * RatioY)
    '        End With
    '    Next


    'End Sub

    Private Sub FrmInterface_Resize(sender As Object, e As EventArgs) Handles Me.Resize
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

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles bouton_excel.Click

        FermerIcon()



        If faire_vider = True Then

            Dim result As DialogResult

            result = MessageBox.Show(error_enregister, error_titre, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                faire_vider = False
                Lister_button(True)

            ElseIf result = DialogResult.Yes Then

            End If


        End If

        If faire_vider = False Then

            Panel1.Controls.Clear()
            Lister_button(True)
            Dim from_var As New frmExcel()
            'InitializeComponent()
            from_var.TopLevel = False
            'f.WindowState = FormWindowState.Maximized
            from_var.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            from_var.Visible = True
            from_var.Location = New Point(1, 0)

            Panel1.Controls.Add(from_var)

            'btm_cheque.Enabled = False

            'fenetre = "cheque"


        End If
        'testpdf()

        'SaveFileDialog1.FileName = ""

        'Dim Path As String
        'Path = SaveFileDialog1.FileName


        'If SaveFileDialog1.ShowDialog = DialogResult.OK Then
        '    'Appel d un fonction
        '    MsgBox(SaveFileDialog1.FileName)


        '    nom_ficbier = SaveFileDialog1.FileName
        '    Dim pdfTable As New PdfPTable(DataGridView1.ColumnCount)
        '    pdfTable.DefaultCell.Padding = 3
        '    pdfTable.WidthPercentage = 100
        '    pdfTable.HorizontalAlignment = Element.ALIGN_MIDDLE
        '    pdfTable.DefaultCell.BorderWidth = 3

        '    'Adding Header row
        '    For Each column As DataGridViewColumn In DataGridView1.Columns

        '        Dim cell As New PdfPCell(New Phrase(column.HeaderText))
        '        cell.BackgroundColor = New iTextSharp.text.Color(240, 240, 240)
        '        pdfTable.AddCell(cell)
        '    Next



        '    'Adding DataRow
        '    For Each row As DataGridViewRow In DataGridView1.Rows
        '        For Each cell As DataGridViewCell In row.Cells
        '            If cell.Value Is Nothing Then

        '            Else
        '                pdfTable.AddCell(cell.Value.ToString())
        '            End If
        '        Next
        '    Next



        '    Using Stream As New FileStream(nom_ficbier, FileMode.Create)
        '        Dim pdfDoc As New Document(PageSize.A2, 10.0F, 10.0F, 10.0F, 0.0F)
        '        PdfWriter.GetInstance(pdfDoc, Stream)
        '        pdfDoc.Open()
        '        pdfDoc.Add(pdfTable)
        '        pdfDoc.Close()
        '        Stream.Close()
        '    End Using

        '    MsgBox("Le fichier PDF a été crée")
        'End If







    End Sub

    Private Sub SavePDF()

        DataGridView1.DataSource = Nothing
        Dim datelist As New List(Of Date)

        Dim dt As New DataTable()
        Try
            'Dim myConnection As New OleDb.OleDbConnection(myConnString)
            'myConnection.Open()

            'Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_emis", myConnection)

            'Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            'Dim newDate As Date = Date.Now.AddMonths(1)

            For i = 0 To 11
                datelist.Add(Date.Now.AddMonths(i))
            Next





            'dt.Columns.AddRange(New DataColumn() {New DataColumn(datelist(0).ToString("MMMM yyyy"), GetType(String)),
            '                                   New DataColumn(datelist(1).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(2).ToString("MMMM yyyy"), GetType(String)),
            '                                   New DataColumn(datelist(3).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(4).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(5).ToString("MMMM yyyy"), GetType(String)),
            '                                   New DataColumn(datelist(6).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(7).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(8).ToString("MMMM yyyy"), GetType(String)),
            '                                   New DataColumn(datelist(9).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(10).ToString("MMMM yyyy"), GetType(Integer)),
            '                                   New DataColumn(datelist(11).ToString("MMMM yyyy"), GetType(Double))})

            'dt.Columns.AddRange(New DataColumn() {New DataColumn("Id", GetType(Integer)),
            '                                   New DataColumn("Name", GetType(String)),
            '                                   New DataColumn("Country", GetType(String))})

            Dim adc1 As DataColumn
            Dim adc2 As DataColumn
            Dim adc3 As DataColumn
            Dim adc4 As DataColumn
            Dim adc5 As DataColumn
            Dim adc6 As DataColumn


            adc1 = New DataColumn("Name", GetType(String))
            adc2 = New DataColumn(datelist(1).ToString("MMMM yyyy"), GetType(String))
            adc3 = New DataColumn(datelist(2).ToString("MMMM yyyy"), GetType(String))
            adc4 = New DataColumn(datelist(3).ToString("MMMM yyyy"), GetType(String))
            adc5 = New DataColumn(datelist(4).ToString("MMMM yyyy"), GetType(String))
            adc6 = New DataColumn(datelist(5).ToString("MMMM yyyy"), GetType(String))

            dt.Columns.Add(adc1)
            dt.Columns.Add(adc2)
            dt.Columns.Add(adc3)
            dt.Columns.Add(adc4)
            dt.Columns.Add(adc5)
            dt.Columns.Add(adc6)




        Catch ex As Exception

        End Try

        'charger_table()
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.Filter = "PDF (*.pdf)|*.pdf"





        Me.DataGridView1.DataSource = dt



    End Sub






    Function compter_revenu() As ArrayList

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM revenu ", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                list.Add(myReader.GetInt32(0))

            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try



        Return list
    End Function


    Sub Proc_load_suppprimer_automatiqu()

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_preautorises ", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim dt1 As DateTime = DateCourant
                Dim dt2 As DateTime = Convert.ToDateTime(myReader.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)
                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1

                If nbr < 0 Then
                    list.Add(valueIntConverted)
                    enregistrer = True

                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer
        For Each num In list
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_preautorises WHERE numero = @numero", myConnection)

                myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next

    End Sub

    Sub Proc_load_suppprimer_automatique_direct()

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_direct ", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim dt1 As DateTime = DateCourant
                Dim dt2 As DateTime = Convert.ToDateTime(myReader.Item(4))
                Dim ts As TimeSpan = dt2.Subtract(dt1)
                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1

                If nbr < 0 Then
                    list.Add(valueIntConverted)
                    enregistrer = True

                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer
        For Each num In list
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_direct WHERE numero = @numero", myConnection)

                myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next

    End Sub

    Sub Proc_load_suppprimer_automatique_cheque()

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_emis ", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim dt1 As DateTime = DateCourant
                Dim dt2 As DateTime = Convert.ToDateTime(myReader.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)
                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1

                If nbr < 0 Then
                    list.Add(valueIntConverted)
                    enregistrer = True

                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer
        For Each num In list
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE numero = @numero", myConnection)

                myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next

    End Sub

    Sub Proc_load_suppprimer_automatique_revenu()

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_revenu ", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim dt1 As DateTime = DateCourant
                Dim dt2 As DateTime = Convert.ToDateTime(myReader.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)
                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1

                If nbr < 0 Then
                    list.Add(valueIntConverted)
                    enregistrer = True

                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer
        For Each num In list
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_revenu WHERE numero = @numero", myConnection)

                myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception

            End Try

        Next

    End Sub

    Sub Proc_load_suppprimer_automatique_depense()

        Dim list As New ArrayList

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_depenses ", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim dt1 As DateTime = DateCourant
                Dim dt2 As DateTime = Convert.ToDateTime(myReader.Item(3))
                Dim ts As TimeSpan = dt2.Subtract(dt1)
                Dim nbr = Convert.ToInt32(ts.Days) + 1

                nbr = nbr - 1

                If nbr < 0 Then
                    list.Add(valueIntConverted)
                    enregistrer = True

                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer
        For Each num In list
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_depenses WHERE numero = @numero", myConnection)

                myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next

    End Sub

    Private Sub proc_repeat_preautorise()


        Dim list As New ArrayList
        Dim list_nbr As New ArrayList
        Dim total_string As String


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_preautorises", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim fournisseur As String = myReader.Item(1)
                Dim dt1 As String = myReader.Item(2).ToString()

                Dim dt2 As String = myReader.Item(3).ToString()

                total_string = fournisseur + ";" + dt1 + ";" + dt2



                'Dim list As New ArrayList
                'Dim list_nbr As New ArrayList
                'Dim total_string As String

                If list.Count > 0 Then
                    Dim trouve As Boolean = False

                    Dim total As String

                    For Each total In list
                        If total_string = total Then
                            trouve = True
                        End If
                    Next

                    If trouve = False Then
                        list.Add(total_string)
                    Else
                        list_nbr.Add(valueIntConverted)
                    End If
                Else
                    list.Add(total_string)
                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer

        For Each num In list_nbr
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_preautorises WHERE numero = @numero", myConnection)

                'myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.Parameters.AddWithValue("@numero", num)
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next


    End Sub

    Private Sub proc_repeat_depense()


        Dim list As New ArrayList
        Dim list_nbr As New ArrayList
        Dim total_string As String


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_depenses", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim fournisseur As String = myReader.Item(1)
                Dim dt1 As String = myReader.Item(2).ToString()

                Dim dt2 As String = myReader.Item(3).ToString()

                total_string = fournisseur + ";" + dt1 + ";" + dt2



                'Dim list As New ArrayList
                'Dim list_nbr As New ArrayList
                'Dim total_string As String

                If list.Count > 0 Then
                    Dim trouve As Boolean = False

                    Dim total As String

                    For Each total In list
                        If total_string = total Then
                            trouve = True
                        End If
                    Next

                    If trouve = False Then
                        list.Add(total_string)
                    Else
                        list_nbr.Add(valueIntConverted)
                    End If
                Else
                    list.Add(total_string)
                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer

        For Each num In list_nbr
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_depenses WHERE numero = @numero", myConnection)

                'myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.Parameters.AddWithValue("@numero", num)
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next



    End Sub

    Private Sub proc_repeat_revenu()


        Dim list As New ArrayList
        Dim list_nbr As New ArrayList
        Dim total_string As String


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM  check_revenu", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read
                'Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
                Dim valueIntConverted As Integer = CInt(myReader.Item(0).ToString)

                Dim fournisseur As String = myReader.Item(1)
                Dim dt1 As String = myReader.Item(2).ToString()

                Dim dt2 As String = myReader.Item(3).ToString()

                total_string = fournisseur + ";" + dt1 + ";" + dt2



                'Dim list As New ArrayList
                'Dim list_nbr As New ArrayList
                'Dim total_string As String

                If list.Count > 0 Then
                    Dim trouve As Boolean = False

                    Dim total As String

                    For Each total In list
                        If total_string = total Then
                            trouve = True
                        End If
                    Next

                    If trouve = False Then
                        list.Add(total_string)
                    Else
                        list_nbr.Add(valueIntConverted)
                    End If
                Else
                    list.Add(total_string)
                End If


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        Dim num As Integer

        For Each num In list_nbr
            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_revenu WHERE numero = @numero", myConnection)

                'myCommand.Parameters.AddWithValue("@numero", num.ToString())
                myCommand.Parameters.AddWithValue("@numero", num)
                myCommand.ExecuteNonQuery()
                myConnection.Close()


                myConnection.Close()
                myCommand.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Next



    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        DateCourant = Date.Now

        MsgBox(DateCourant.ToString())

        FrmMenu.Aff_miseajour()


        Proc_load() 'preautorise
        Proc_load_cheque()
        Proc_load_direct()
        Proc_load_revenue()
        Proc_load_depense()

        runn_global = True


        While runn_global

            runn_global = False

            Proc_load_db()

        End While

        runn_global = True


        Dim list As New ArrayList

        While runn_global
            runn_global = False

            'Interface  dont work sur revenu
            Proc_load_db_revenu()
        End While

        runn_global = True


        While runn_global
            runn_global = False

            Proc_load_db_depense()

            'frmPaiement.proc_load_depense()
        End While


        'Supprimer automatiquement les paiement qui depassé la date

        Proc_load_suppprimer_automatiqu()
        Proc_load_suppprimer_automatique_direct()
        Proc_load_suppprimer_automatique_cheque()
        Proc_load_suppprimer_automatique_revenu()
        Proc_load_suppprimer_automatique_depense()

        'End While

        changer_date = False

        proc_repeat_preautorise()


        MsgBox("Termine")


    End Sub
End Class