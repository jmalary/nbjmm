Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb

Public Class frmGestion

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.User_Access
    Dim AL As New Collection
    Dim somme_solde_enc As Double
    Dim num_temporaire As String

    Private Property OleDbDataReader As OleDbDataReader

    Private Sub frmGestion_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Me.WindowState = FormWindowState.Maximized
        frmMenu.MdiParent = Form_windows

        frmMenu.Show()

    End Sub




    Private Sub frmGestion_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub




    Private Sub frmGestion_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        If backup = False Then

            proc_ListerText()

            Dim currentDate As DateTime
            currentDate = Now.Date
            currentDate = currentDate.AddDays(6)

            proc_tout()
            proc_restaurantion()
        Else


            recurer_backup()

        End If

        desactive()

    End Sub

    Sub proc_ListerText()
        AL.Add(TextBox1)

    End Sub


    Sub proc_tout()

        proc_A()

        'proc_AC()
        'proc_AD()

        ' ******************************************************
        ' Apports
        '-********************************************************


        proc_B()
        proc_C()
        proc_D()
        proc_E()
        proc_E5()

        ' ******************************************************
        ' RETRAIRE MANUEL
        '-********************************************************

        proc_F()
        proc_G()
        proc_H()
        proc_I()
        proc_I5()

        ' ******************************************************
        'Cheque Direct
        '-********************************************************
        proc_J()
        proc_K()
        proc_L()
        proc_O()
        proc_O5()

        '' ******************************************************
        '' Cheque Preautorise
        ''-********************************************************
        proc_M()
        proc_N()
        proc_P()
        proc_Q()
        proc_Q5()

        '' ******************************************************
        '' Cheque Circulation
        ''-********************************************************
        proc_R()
        proc_S()
        proc_T()
        proc_U()
        proc_U5()


        '' ******************************************************
        '' Salaire paie courant
        ' ''-********************************************************
        'TODO Tabarnak
        proc_V()
        proc_W()
        proc_X()
        proc_Y()
        proc_Y5()

        '' ******************************************************
        '' Solde encaisse
        ''-********************************************************
        proc_AA()
        'proc_AB()
        'proc_AC()
        'proc_AD()

        '' ******************************************************
        '' Aregent à recevoir
        ''-********************************************************
        proc_BA()
        proc_BB()
        proc_BC()
        proc_BD()
        proc_BD5()


        '' ******************************************************
        '' Solde encaisse avant M/C
        ''-********************************************************

        proc_CA()
        proc_CB()
        proc_CC()
        'proc_CD()


        '' ******************************************************
        '' Marge de crédit utilisée
        ''-********************************************************
        proc_DA()


        '' ******************************************************
        '' Besoin d'emprunt à CT
        proc_DAA()
        ''-********************************************************

        proc_EA()
        proc_E25_1()  ' txt19
        proc_E25()  'lbl52


        proc_F25()




        '' ******************************************************
        '' Encaisse minimum requised
        ''-********************************************************

        proc_enc_min_req1()
        proc_enc_min_req2()
        proc_enc_min_req3()
        proc_enc_min_req4()
        proc_enc_min_req45()



        '' ******************************************************
        '' Marge de crédit disponible
        ''-********************************************************

        'vider
        proc_marge_credit_disp_req1()
        proc_marge_credit_disp_req2()
        proc_marge_credit_disp_req3()
        'vider



        '' ******************************************************
        '' Besoin d'emprunt à CT
        ''-********************************************************


        proc_besoin_emp_req1()
        'proc_besoin_emp_req2()
        'proc_besoin_emp_req3()
        proc_besoin_emp_req4()




        '' ******************************************************
        '' Solde encaisse début (DEUXIENE)
        ''-********************************************************

        proc_solde_enc_req1()
        proc_besoin_emp_req2()
        proc_solde_enc_req2()
        proc_besoin_emp_req3()

        proc_solde_enc_req3()
        proc_solde_enc_req35()


        '' ******************************************************
        '' Solde encaisse début
        ''-********************************************************


        proc_solde_enc_req1()

        ElementPlusMax()

        calculer_somme_encaisse_mc()


        '' ******************************************************
        '' Marge de crédit disponible à la fin
        ''-********************************************************




        corriger_marge_credit_debut()

        corriger_liquidite_immediat()

        proc_marge_credit_dispo_fin()


    End Sub

    Sub proc_A()


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Solde_encaisses_debut where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox1.Text = myReader1.Item(0).ToString

        End While
        TextBox1.Text = conventir_format(TextBox1.Text)
        myConnection.Close()

    End Sub

    Sub proc_A1()


        Dim num1 As Double
        Dim num2 As Double
        Dim SOMME As Double


        'TODO Label14.text et Lbl60

        Label14.Text = Format$(Label60.Text, "Currency")


        proc_AB()

    End Sub

    ' ******************************************************
    ' Apports
    '-********************************************************
    Sub proc_B()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Apports where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox12.Text = myReader1.Item(0).ToString

        End While

        TextBox12.Text = conventir_format(TextBox12.Text)
        myConnection.Close()


    End Sub

    Sub proc_C()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sept from Apports where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read


            TextBox11.Text = myReader1.Item(0).ToString
            'TextBox11.Text = myReader1.Item(0).ToString



        End While
        TextBox11.Text = conventir_format(TextBox11.Text)

        myConnection.Close()


    End Sub

    Sub proc_D()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select quatorze from Apports where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox10.Text = myReader1.Item(0).ToString


        End While

        TextBox10.Text = conventir_format(TextBox10.Text)
        myConnection.Close()


    End Sub

    Sub proc_E()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Apports where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox9.Text = myReader1.Item(0).ToString


        End While

        TextBox9.Text = conventir_format(TextBox9.Text)
        myConnection.Close()

    End Sub

    Sub proc_E5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select cinq from Apports where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox20.Text = myReader1.Item(0).ToString


        End While

        TextBox20.Text = conventir_format(TextBox20.Text)
        myConnection.Close()

    End Sub


    ' ******************************************************
    ' RETRAIRE MANUEL
    '-********************************************************

    Sub proc_F()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Retrait_manuel where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox18.Text = myReader1.Item(0).ToString

        End While

        TextBox18.Text = conventir_format(TextBox18.Text)
        myConnection.Close()

    End Sub

    Sub proc_G()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sept from Retrait_manuel where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox17.Text = myReader1.Item(0).ToString

        End While

        TextBox17.Text = conventir_format(TextBox17.Text)
        myConnection.Close()


    End Sub

    Sub proc_H()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select quatorze from Retrait_manuel where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox16.Text = myReader1.Item(0).ToString

        End While

        TextBox16.Text = conventir_format(TextBox16.Text)
        myConnection.Close()

    End Sub

    Sub proc_I()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Retrait_manuel where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox15.Text = myReader1.Item(0).ToString

        End While

        TextBox15.Text = conventir_format(TextBox15.Text)
        myConnection.Close()


    End Sub

    Sub proc_I5()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select cinq from Retrait_manuel where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox21.Text = myReader1.Item(0).ToString

        End While

        TextBox21.Text = conventir_format(TextBox21.Text)
        myConnection.Close()


    End Sub


    ' ******************************************************
    ' Cheque direct
    '-********************************************************
    Sub proc_J()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read

            Label17.Text = myReader1.Item(0).ToString


        End While

        Label17.Text = check_zero(Label17.Text)

        Label17.Text = conventir_format(Label17.Text)

        myConnection.Close()


    End Sub

    Sub proc_K()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label19.Text = myReader1.Item(0).ToString
        End While

        Label19.Text = check_zero(Label19.Text)
        Label19.Text = conventir_format(Label19.Text)

        myConnection.Close()


    End Sub

    Sub proc_L()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label20.Text = myReader1.Item(0).ToString
        End While

        Label20.Text = check_zero(Label20.Text)
        Label20.Text = conventir_format(Label20.Text)

        myConnection.Close()


    End Sub

    Sub proc_O()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label21.Text = myReader1.Item(0).ToString
        End While

        Label21.Text = check_zero(Label21.Text)
        Label21.Text = conventir_format(Label21.Text)

        myConnection.Close()


    End Sub

    Sub proc_O5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label36.Text = myReader1.Item(0).ToString
        End While

        Label36.Text = check_zero(Label36.Text)
        Label36.Text = conventir_format(Label36.Text)


        myConnection.Close()


    End Sub

    ' ******************************************************
    ' Cheque PREAUTORISE
    '-********************************************************
    Sub proc_M()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label25.Text = myReader1.Item(0).ToString


        End While

        Label25.Text = check_zero(Label25.Text)
        Label25.Text = conventir_format(Label25.Text)

        myConnection.Close()

    End Sub

    Sub proc_N()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_preautorises", myConnection)

        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label24.Text = myReader1.Item(0).ToString
        End While

        Label24.Text = check_zero(Label24.Text)
        Label24.Text = conventir_format(Label24.Text)


        myConnection.Close()

    End Sub

    Sub proc_P()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label23.Text = myReader1.Item(0).ToString
        End While

        Label23.Text = check_zero(Label23.Text)
        Label23.Text = conventir_format(Label23.Text)


        myConnection.Close()

    End Sub

    Sub proc_Q()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label22.Text = myReader1.Item(0).ToString
        End While

        Label22.Text = check_zero(Label22.Text)
        Label22.Text = conventir_format(Label22.Text)

        myConnection.Close()

    End Sub

    Sub proc_Q5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label37.Text = myReader1.Item(0).ToString
        End While

        Label37.Text = check_zero(Label37.Text)
        Label37.Text = conventir_format(Label37.Text)
        myConnection.Close()

    End Sub

    ' ******************************************************
    ' Cheque Circulation
    '-********************************************************
    Sub proc_R()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label29.Text = myReader1.Item(0).ToString


        End While

        Label29.Text = check_zero(Label29.Text)
        Label29.Text = conventir_format(Label29.Text)

        myConnection.Close()

    End Sub

    Sub proc_S()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label28.Text = myReader1.Item(0).ToString
        End While


        Label28.Text = check_zero(Label28.Text)
        Label28.Text = conventir_format(Label28.Text)

    End Sub

    Sub proc_T()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label27.Text = myReader1.Item(0).ToString
        End While

        Label27.Text = check_zero(Label27.Text)

        Label27.Text = conventir_format(Label27.Text)

        myConnection.Close()


    End Sub

    Sub proc_U()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label26.Text = myReader1.Item(0).ToString
        End While
        Label26.Text = check_zero(Label26.Text)
        Label26.Text = conventir_format(Label26.Text)

        myConnection.Close()


    End Sub


    Sub proc_U5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label41.Text = myReader1.Item(0).ToString
        End While

        Label41.Text = check_zero(Label41.Text)
        Label41.Text = conventir_format(Label41.Text)
        myConnection.Close()


    End Sub

    ' ******************************************************
    ' Salaire paie courant
    '-********************************************************
    Sub proc_V()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox72.Text = myReader1.Item(0).ToString


        End While

        TextBox72.Text = conventir_format(TextBox72.Text)

        myConnection.Close()


    End Sub

    Sub proc_W()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sept from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox71.Text = myReader1.Item(0).ToString

        End While

        TextBox71.Text = conventir_format(TextBox71.Text)

        myConnection.Close()


    End Sub

    Sub proc_X()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select quatorze from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox70.Text = myReader1.Item(0).ToString


        End While

        TextBox70.Text = conventir_format(TextBox70.Text)
        myConnection.Close()


    End Sub

    Sub proc_Y()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox69.Text = myReader1.Item(0).ToString


        End While

        TextBox69.Text = conventir_format(TextBox69.Text)
        myConnection.Close()


    End Sub

    Sub proc_Y5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select cinq from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox22.Text = myReader1.Item(0).ToString


        End While


        TextBox22.Text = conventir_format(TextBox22.Text)
        myConnection.Close()


    End Sub


    Sub proc_AB1()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Salaires_paie_courante where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox69.Text = myReader1.Item(0).ToString


        End While


        TextBox69.Text = conventir_format(TextBox69.Text)
        myConnection.Close()


    End Sub


    ' ******************************************************
    ' Solde encaisse
    '-********************************************************

    Sub proc_AA()

        Dim nbr1 As Double
        Dim nbr2 As Double
        Dim nbr3 As Double
        Dim nbr4 As Double
        Dim nbr5 As Double
        Dim nbr6 As Double
        Dim nbr7 As Double

        Dim c As Double

        nbr1 = TextBox1.Text
        nbr2 = TextBox12.Text
        nbr3 = TextBox18.Text
        nbr4 = Label17.Text
        nbr5 = Label25.Text
        nbr6 = Label29.Text
        nbr7 = TextBox72.Text




        c = nbr1 + nbr2 - nbr3 - nbr4 - nbr5 - nbr6 - nbr7

        'Label68.Text = Format$(c, "Currency")

        Label68.Text = conventir_format(CStr(c))


        '********************************************************************************

        '******************************************************************************************

    End Sub

    Sub proc_AB()

        Dim nbr1 As Double
        Dim nbr2 As Double
        Dim nbr3 As Double
        Dim nbr4 As Double
        Dim nbr5 As Double
        Dim nbr6 As Double
        Dim nbr7 As Double

        Dim c As Double



        nbr1 = Label14.Text
        nbr2 = TextBox11.Text
        nbr3 = TextBox17.Text
        nbr4 = Label19.Text
        nbr5 = Label24.Text
        nbr6 = Label28.Text
        nbr7 = TextBox71.Text




        c = nbr1 + nbr2 - nbr3 - nbr4 - nbr5 - nbr6 - nbr7

        'MsgBox(c)



        'Label67.Text = Math.Round(c, 2)
        Label67.Text = conventir_format(CStr(c))


        'Label67.Text = Format$(Label67.Text, "Currency")

        proc_CB()

        proc_AC()



    End Sub

    Sub proc_AC()

        Dim nbr1 As Double
        Dim nbr2 As Double
        Dim nbr3 As Double
        Dim nbr4 As Double
        Dim nbr5 As Double
        Dim nbr6 As Double
        Dim nbr7 As Double

        Dim c As Double

        nbr1 = Label15.Text
        nbr2 = TextBox10.Text
        nbr3 = TextBox16.Text
        nbr4 = Label20.Text
        nbr5 = Label23.Text
        nbr6 = Label27.Text
        nbr7 = TextBox70.Text




        c = nbr1 + nbr2 - nbr3 - nbr4 - nbr5 - nbr6 - nbr7

        Label66.Text = conventir_format(CStr(c))

        'Label66.Text = c
        proc_CC()

    End Sub

    Sub proc_AD()

        Dim nbr1 As Double
        Dim nbr2 As Double
        Dim nbr3 As Double
        Dim nbr4 As Double
        Dim nbr5 As Double
        Dim nbr6 As Double
        Dim nbr7 As Double

        Dim c As Double

        nbr1 = Label16.Text
        'nbr1 = 0
        nbr2 = TextBox9.Text
        nbr3 = TextBox15.Text
        nbr4 = Label21.Text
        nbr5 = Label22.Text
        nbr6 = Label26.Text
        nbr7 = TextBox69.Text




        c = nbr1 + nbr2 - nbr3 - nbr4 - nbr5 - nbr6 - nbr7



        'Label65.Text = c
        Label65.Text = conventir_format(CStr(c))


        proc_DD()


    End Sub
    ' ******************************************************
    ' Argent à recevoir
    '-********************************************************
    Sub proc_BA()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Argent_a_recevoir where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox14.Text = myReader1.Item(0).ToString


        End While

        TextBox14.Text = conventir_format(TextBox14.Text)
        myConnection.Close()

    End Sub

    Sub proc_BB()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sept from Argent_a_recevoir where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox13.Text = myReader1.Item(0).ToString


        End While

        TextBox13.Text = conventir_format(TextBox13.Text)
        myConnection.Close()


    End Sub

    Sub proc_BC()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select quatorze from Argent_a_recevoir where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox8.Text = myReader1.Item(0).ToString


        End While

        TextBox8.Text = conventir_format(TextBox8.Text)

        myConnection.Close()


    End Sub

    Sub proc_BD()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Argent_a_recevoir where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox7.Text = myReader1.Item(0).ToString


        End While

        TextBox7.Text = conventir_format(TextBox7.Text)
        myConnection.Close()


    End Sub

    Sub proc_BD5()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select cinq from Argent_a_recevoir where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox23.Text = myReader1.Item(0).ToString


        End While

        TextBox23.Text = conventir_format(TextBox23.Text)
        myConnection.Close()



    End Sub




    ' ******************************************************
    ' Solde encaisse avant M/C
    '-********************************************************
    Sub proc_CA()

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox14.Text
        nbrB = Label68.Text

        c1 = nbrA + nbrB
        '***********
        'Label60.Text = Format$(c1, "Currency")
        Label60.Text = conventir_format(CStr(c1))




        proc_A1()




    End Sub

    Sub proc_CB()

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox13.Text
        nbrB = Label67.Text


        c1 = nbrA + nbrB


        Label59.Text = conventir_format(CStr(c1))

        'Label59.Text = Format$(c1, "Currency")


        '********************************************************************************


        '******************************************************************************************


        Label15.Text = conventir_format(CStr(c1))

        'Label15.Text = Format$(c1, "Currency")

        'MsgBox(Label15.Text)   'TODO

        'Label15.Text = c1





    End Sub

    Sub proc_CC()

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox8.Text
        nbrB = Label66.Text


        c1 = nbrA + nbrB



        'Label58.Text = c1
        Label58.Text = conventir_format(CStr(c1))

        'Label16.Text = c1
        Label16.Text = conventir_format(CStr(c1))

        'MsgBox(c1)

        proc_AD()


    End Sub

    Sub proc_DD()

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox7.Text
        nbrB = Label65.Text


        c1 = nbrA + nbrB

        'Label57.Text = c1
        Label57.Text = conventir_format(CStr(c1))



        'MsgBox(c1)



    End Sub


    ' ******************************************************
    ' Besoin d'emprunt à CT
    '-********************************************************


    Sub proc_EA()

        Dim num As Double

        num = Label60.Text



    End Sub

    Sub proc_E25()

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double


        'Label56.Text = TextBox5.Text

        nbrA = Label60.Text

        nbrB = TextBox19.Text


        c1 = nbrA + nbrB


        'Label52.Text = c1
        Label52.Text = conventir_format(CStr(c1))

        'Label52.Text = FormatCurrency(Label52.Text)
    End Sub

    Sub proc_E25_1()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from marge_credit_disponible where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox19.Text = myReader1.Item(0).ToString
            'TextBox1.Text = myReader1.Item(0).ToString
            'TextBox1.Text = FormatCurrency(myReader1.Item(0).ToString)


        End While

        TextBox19.Text = conventir_format(TextBox19.Text)


    End Sub


    Sub proc_F25()
        'Dim nbrA As Double
        'Dim nbrB As Double

        'Dim c1 As Double


        'Label56.Text = TextBox5.Text

        'nbrA = Label60.Text
        'nbrB = Label56.Text


        'c1 = nbrA + nbrB
        'Label52.Text = c1


    End Sub

    ' ******************************************************
    ' Marge de crédit utilisée
    '-********************************************************

    Sub proc_DA()


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from marge_de_credit_utilise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox5.Text = myReader1.Item(0).ToString
            'TextBox1.Text = myReader1.Item(0).ToString
            'TextBox1.Text = FormatCurrency(myReader1.Item(0).ToString)


        End While

        TextBox5.Text = conventir_format(TextBox5.Text)

    End Sub

    Sub proc_DAA()

        Dim num As Double
        Dim num1 As Double
        Dim num2 As Double
        Dim num3 As Double

        num = Label60.Text
        num1 = TextBox5.Text



        If num < 0 Then ' => SI condition validée ALORS
            'Instructions si vrai
        Else ' => SINON
            'Instructions si faux
        End If



    End Sub


    ' ******************************************************
    ' Encaisse minimum requise
    '-********************************************************


    Sub proc_enc_min_req1()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Encaisse_minimum_requise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox6.Text = myReader1.Item(0).ToString


        End While

        TextBox6.Text = conventir_format(TextBox6.Text)

    End Sub

    Sub proc_enc_min_req2()


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sept from Encaisse_minimum_requise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox4.Text = myReader1.Item(0).ToString


        End While

        TextBox4.Text = conventir_format(TextBox4.Text)
    End Sub

    Sub proc_enc_min_req3()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select quatorze from Encaisse_minimum_requise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox3.Text = myReader1.Item(0).ToString


        End While

        TextBox3.Text = conventir_format(TextBox3.Text)

    End Sub

    Sub proc_enc_min_req4()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select vingt from Encaisse_minimum_requise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox2.Text = myReader1.Item(0).ToString


        End While

        TextBox2.Text = conventir_format(TextBox2.Text)
    End Sub

    Sub proc_enc_min_req45()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select cinq from Encaisse_minimum_requise where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            TextBox24.Text = myReader1.Item(0).ToString


        End While

        TextBox24.Text = conventir_format(TextBox24.Text)
    End Sub


    ' ******************************************************
    'Besoin d'emprunt à CT
    '-********************************************************


    Sub proc_besoin_emp_req1()

        Dim num As Double
        Dim num1 As Double
        Dim Somme As Double


        'e23
        num = Label60.Text

        'e29
        num1 = TextBox6.Text

        If num < num1 Then
            Somme = num1 - num
        Else
            Somme = 0
        End If

        Label32.Text = conventir_format(CStr(Somme))
        'Label32.Text = Somme
        'Label32.Text = FormatCurrency(Label32.Text)

    End Sub



    Sub proc_besoin_emp_req2()

        Dim num As Double
        Dim num1 As Double
        Dim Somme As Double


        'e23
        num = Label59.Text

        'e29
        num1 = TextBox4.Text

        If num < num1 Then
            Somme = num1 - num
        Else
            Somme = 0
        End If

        'MsgBox(num)
        'MsgBox(num1)


        'Label31.Text = Somme
        Label31.Text = conventir_format(CStr(Somme))

    End Sub

    Sub proc_besoin_emp_req3()
        Dim num As Double
        Dim num1 As Double
        Dim Somme As Double


        'e23
        num = Label58.Text

        'e29
        num1 = TextBox3.Text

        If num < num1 Then
            Somme = num1 - num
        Else
            Somme = 0
        End If

        Label30.Text = conventir_format(CStr(Somme))
        'Label30.Text = Somme
        'Label30.Text = Format$(Label30.Text, "Currency")

    End Sub

    Sub proc_besoin_emp_req4()
        Dim num As Double
        Dim num1 As Double
        Dim Somme As Double


        'e23
        num = Label57.Text

        'e29
        num1 = TextBox2.Text

        If num < num1 Then
            Somme = num1 - num
        Else
            Somme = 0
        End If


        Label13.Text = conventir_format(CStr(Somme))

        'Label13.Text = Somme
        'Label13.Text = Format$(Label13.Text, "Currency")

    End Sub

    ' ******************************************************
    'Solde encaisse début (Deuxieme)
    '-********************************************************



    Sub proc_solde_enc_req1()

        Dim num1 As Double
        Dim num2 As Double
        Dim somme As Double

        num1 = Label60.Text
        num2 = Label32.Text

        somme = num1 + num2

        Label14.Text = conventir_format(CStr(somme))
        'Label14.Text = somme
        'Label14.Text = Format$(Label14.Text, "Currency")


        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox13.Text
        nbrB = Label67.Text


        c1 = nbrA + nbrB

        'Label59.Text = Format$(c1, "Currency")
        Label59.Text = conventir_format(CStr(c1))
        '-----------------------------------------------------------------------------


        '--------------------------------------------------------------------------

        Dim num3 As Double
        Dim num4 As Double
        Dim num5 As Double
        Dim num6 As Double
        Dim num7 As Double
        Dim num8 As Double
        Dim num9 As Double

        num3 = somme
        num4 = TextBox11.Text
        num5 = TextBox17.Text
        num6 = Label19.Text
        num7 = Label24.Text
        num8 = Label28.Text
        num9 = TextBox71.Text

        somme = num3 + num4 - num5 - num6 - num7 - num8 - num9

        Label67.Text = conventir_format(CStr(somme))
        'Label67.Text = somme


        'Label67.Text = Format$(Label67.Text, "Currency")

        'Math.Round(Label67.Text, 2)
        'Label67.Text = Label67.Text + " $"

        Dim num10 As Double
        Dim num11 As Double

        num10 = TextBox19.Text
        num11 = Label32.Text

        somme = num10 - num11


        Label55.Text = conventir_format(CStr(somme))
        'Label55.Text = somme
        'Label55.Text = Format$(Label55.Text, "Currency")

        nbrA = Label59.Text
        nbrB = Label55.Text

        c1 = nbrA + nbrB

        Label51.Text = conventir_format(CStr(c1))
        'Label51.Text = Format$(c1, "Currency")

        nbrA = Label59.Text


        nbrB = TextBox4.Text

        If nbrA < nbrB Then
            somme = nbrB - nbrA
        Else
            somme = 0
        End If


        'Label31.Text = Format$(somme, "Currency")
        Label31.Text = conventir_format(CStr(somme))




    End Sub

    Sub proc_solde_enc_req2()

        Dim num1 As Double
        Dim num2 As Double
        Dim somme As Double

        num1 = Label59.Text
        num2 = Label31.Text

        somme = num1 + num2


        Label15.Text = conventir_format(CStr(somme))
        'Label15.Text = somme
        'Label15.Text = Format$(Label15.Text, "Currency")





        '-----------------------------------------------------------------------------

        Dim num3 As Double
        Dim num4 As Double
        Dim num5 As Double
        Dim num6 As Double
        Dim num7 As Double
        Dim num8 As Double
        Dim num9 As Double

        num3 = somme
        num4 = TextBox10.Text
        num5 = TextBox16.Text
        num6 = Label20.Text
        num7 = Label23.Text
        num8 = Label27.Text
        num9 = TextBox70.Text

        somme = num3 + num4 - num5 - num6 - num7 - num8 - num9


        'Label66.Text = somme
        Label66.Text = conventir_format(CStr(somme))



        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double



        'Label66.Text = Label66.Text + " $"


        nbrA = TextBox8.Text
        nbrB = Label66.Text


        c1 = nbrA + nbrB

        Label58.Text = conventir_format(CStr(c1))
        'Label58.Text = Format$(c1, "Currency")



        Dim num10 As Double
        Dim num11 As Double

        num10 = Label55.Text


        '*********************************************

        Dim numC As Double
        Dim numD As Double
        Dim total As Double

        'e23
        numC = Label67.Text

        'e29
        numD = TextBox13.Text

        total = numC + numD

        num11 = total


        '**********************************************






        somme = num10 + num11

        'MsgBox(num11)


        Label54.Text = conventir_format(CStr(somme))
        'Label54.Text = somme
        'Label54.Text = Format$(Label54.Text, "Currency")


        nbrA = Label58.Text
        nbrB = Label54.Text

        c1 = nbrA + nbrB

        Label50.Text = conventir_format(CStr(c1))
        'Label50.Text = Format$(c1, "Currency")

        nbrA = Label58.Text


        nbrB = TextBox3.Text

        If nbrA < nbrB Then
            somme = nbrB - nbrA
        Else
            somme = 0
        End If

        Label30.Text = conventir_format(CStr(somme))
        'Label30.Text = somme
        'Label30.Text = Format$(Label30.Text, "Currency")

    End Sub

    Sub proc_solde_enc_req3()


        Dim num1 As Double
        Dim num2 As Double
        Dim somme As Double

        num1 = Label58.Text
        num2 = Label30.Text

        somme = num1 + num2

        Label16.Text = conventir_format(CStr(somme))
        'Label16.Text = somme
        'Label16.Text = Format$(Label16.Text, "Currency")




        '-----------------------------------------------------------------------------
        'TODO Label57???

        'Dim S, s2 As String
        'S = Label57.Text
        'Dim answer As Char
        'answer = S.Substring(0, 1)
        'If answer = "(" Then

        '    s2 = S.Substring(0, S.Length - 1)

        '    Label57.Text = s2.Replace(CChar("("), CChar("-"))

        'Else
        '    'MsgBox("Positif")
        'End If

        '--------------------------------------------------------------------------

        Dim num3 As Double
        Dim num4 As Double
        Dim num5 As Double
        Dim num6 As Double
        Dim num7 As Double
        Dim num8 As Double
        Dim num9 As Double

        num3 = somme
        num4 = TextBox9.Text
        num5 = TextBox15.Text
        num6 = Label21.Text
        num7 = Label22.Text
        num8 = Label26.Text
        num9 = TextBox69.Text

        somme = num3 + num4 - num5 - num6 - num7 - num8 - num9


        Label65.Text = conventir_format(CStr(somme))
        'Label65.Text = somme


        'Label65.Text = Label65.Text + " $"


        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox7.Text
        nbrB = Label65.Text


        c1 = nbrA + nbrB

        Label57.Text = conventir_format(CStr(c1))
        'Label57.Text = Format$(c1, "Currency")



        Dim num10 As Double
        Dim num11 As Double

        num10 = Label54.Text


        '*********************************************

        Dim numC As Double
        Dim numD As Double
        Dim total As Double

        'e23
        numC = Label66.Text



        'e29
        numD = TextBox8.Text

        total = numC + numD

        num11 = total


        '**********************************************






        somme = num10 + num11

        'MsgBox(num11)


        Label53.Text = conventir_format(CStr(somme))
        'Label53.Text = somme
        'Label53.Text = Format$(Label53.Text, "Currency")



        nbrA = Label57.Text
        nbrB = Label53.Text

        c1 = nbrA + nbrB

        Label49.Text = conventir_format(CStr(c1))
        'Label49.Text = Format$(c1, "Currency")

        nbrA = Label57.Text


        nbrB = TextBox2.Text

        If nbrA < nbrB Then
            somme = nbrB - nbrA
        Else
            somme = 0
        End If

        Label13.Text = conventir_format(CStr(somme))
        'Label13.Text = somme
        'Label13.Text = Format$(Label13.Text, "Currency")

    End Sub

    Sub proc_solde_enc_req35()

        Dim num1 As Double
        Dim num2 As Double
        Dim somme As Double

        num1 = Label57.Text
        num2 = Label13.Text

        somme = num1 + num2


        Label35.Text = conventir_format(CStr(somme))
        'Label35.Text = somme

        'Label35.Text = Format$(Label35.Text, "Currency")


        Dim num3 As Double
        Dim num4 As Double
        Dim num5 As Double
        Dim num6 As Double
        Dim num7 As Double
        Dim num8 As Double
        Dim num9 As Double

        num3 = somme
        num4 = TextBox20.Text
        num5 = TextBox21.Text
        num6 = Label36.Text
        num7 = Label37.Text
        num8 = Label41.Text
        num9 = TextBox22.Text

        somme = num3 + num4 - num5 - num6 - num7 - num8 - num9

        'Math.Round(somme, 2)

        Label42.Text = conventir_format(CStr(somme))
        'Label42.Text = somme



        'Label42.Text = Label42.Text + " $"

        Dim nbrA As Double
        Dim nbrB As Double

        Dim c1 As Double

        nbrA = TextBox23.Text
        nbrB = Label42.Text


        c1 = nbrA + nbrB

        Label43.Text = conventir_format(CStr(c1))
        'Label43.Text = Format$(c1, "Currency")


        Dim num10 As Double
        Dim num11 As Double

        num10 = Label54.Text


        '*********************************************

        Dim numC As Double
        Dim numD As Double
        Dim total As Double

        'e23
        numC = Label65.Text

        'e29
        numD = TextBox7.Text

        total = numC + numD

        num11 = total


        '**********************************************


        num10 = Label53.Text

        num11 = Label13.Text



        'somme = num10 + num11
        somme = num10 - num11
        'MsgBox(num11)


        Label44.Text = conventir_format(CStr(somme))
        'Label44.Text = somme
        'Label44.Text = Format$(Label44.Text, "Currency")

        nbrA = Label43.Text
        nbrB = Label44.Text

        c1 = nbrA + nbrB

        Label45.Text = conventir_format(CStr(c1))
        'Label45.Text = Format$(c1, "Currency")


        nbrA = Label43.Text


        nbrB = TextBox24.Text

        If nbrA < nbrB Then
            somme = nbrB - nbrA
        Else
            somme = 0
        End If

        Label46.Text = conventir_format(CStr(somme))
        'Label46.Text = somme

        'Label46.Text = Format$(Label46.Text, "Currency")
        'ElementPlusMax()


    End Sub

    Private Sub ElementPlusMax()



        Dim element1 As Double = Label32.Text
        Dim element2 As Double = Label31.Text
        Dim element3 As Double = Label30.Text
        Dim element4 As Double = Label13.Text
        Dim element5 As Double = Label46.Text
        Dim reponse As Double

        If element1 > element2 Then
            reponse = element1
        Else
            reponse = element2
        End If

        If reponse < element3 Then
            reponse = element3

        End If

        If reponse < element4 Then
            reponse = element4

        End If

        If reponse < element5 Then
            reponse = element5

        End If

        Label39.Text = conventir_format(CStr(reponse))
        'Label39.Text = Format$(reponse, "Currency")


    End Sub

    Private Sub calculer_somme_encaisse_mc()

        Dim element1 As Double = Label60.Text
        Dim element2 As Double = Label59.Text
        Dim element3 As Double = Label58.Text
        Dim element4 As Double = Label57.Text
        Dim element5 As Double = Label43.Text
        Dim somme As Double

        somme = element1 + element2 + element3 + element4 + element5

        Label47.Text = conventir_format(CStr(somme))
        'Label47.Text = Format$(somme, "Currency")




    End Sub


    Private Sub proc_marge_credit_dispo_fin()

        Dim num1 As Double
        Dim num2 As Double
        Dim somme As Double

        num3 = somme
        num1 = TextBox19.Text
        num2 = Label32.Text

        somme = num1 - num2

        somme = replacer_zero(somme)

        Label70.Text = conventir_format(CStr(somme))
        'Label70.Text = somme

        num1 = Label55.Text
        num2 = Label31.Text

        somme = num1 - num2

        somme = replacer_zero(somme)
        Label71.Text = conventir_format(CStr(somme))
        'Label71.Text = somme

        num1 = Label54.Text
        num2 = Label30.Text

        somme = num1 - num2

        somme = replacer_zero(somme)
        Label72.Text = conventir_format(CStr(somme))
        'Label72.Text = somme

        num1 = Label53.Text
        num2 = Label13.Text
        somme = num1 - num2

        somme = replacer_zero(somme)
        Label73.Text = conventir_format(CStr(somme))
        'Label73.Text = somme

        num1 = Label44.Text
        num2 = Label46.Text
        somme = num1 - num2

        somme = replacer_zero(somme)
        Label74.Text = conventir_format(CStr(somme))
        'Label74.Text = somme




    End Sub

    Private Sub corriger_marge_credit_debut()

        Dim numA As Double
        Dim numB As Double
        Dim somme As Double


        numA = TextBox19.Text
        numB = Label32.Text

        somme = numA - numB
        Label55.Text = conventir_format(CStr(somme))

        numA = Label55.Text
        numB = Label31.Text

        somme = numA - numB
        Label54.Text = conventir_format(CStr(somme))

        numA = Label54.Text
        numB = Label30.Text

        somme = numA - numB
        Label53.Text = conventir_format(CStr(somme))

        numA = Label53.Text
        numB = Label13.Text

        somme = numA - numB
        Label44.Text = conventir_format(CStr(somme))





    End Sub


    Private Sub corriger_liquidite_immediat()

        Dim numA As Double
        Dim numB As Double
        Dim somme As Double


        numA = Label60.Text
        numB = TextBox19.Text

        somme = numA + numB
        Label52.Text = conventir_format(CStr(somme))
        '**************************************
        numA = Label59.Text
        numB = Label55.Text

        somme = numA + numB
        Label51.Text = conventir_format(CStr(somme))
        '**************************************
        numA = Label58.Text
        numB = Label54.Text

        somme = numA + numB
        Label50.Text = conventir_format(CStr(somme))
        '**************************************

        numA = Label57.Text
        numB = Label53.Text

        somme = numA + numB
        Label49.Text = conventir_format(CStr(somme))
        '**************************************

        numA = Label43.Text
        numB = Label44.Text

        somme = numA + numB
        Label45.Text = conventir_format(CStr(somme))
        '**************************************




    End Sub





    Private Sub proc_restaurantion()



        proc_tab_paiement_direct()
        'proc_tab_paiement_preautorise()
        'proc_tab_cheque_emis()

        Dim myList As New ArrayList

        myList.Add(CDbl(TextBox1.Text))
        myList.Add(CDbl(TextBox12.Text))
        myList.Add(CDbl(TextBox11.Text))
        myList.Add(CDbl(TextBox10.Text))
        myList.Add(CDbl(TextBox9.Text))
        myList.Add(CDbl(TextBox20.Text))
        myList.Add(CDbl(TextBox18.Text))
        myList.Add(CDbl(TextBox17.Text))
        myList.Add(CDbl(TextBox16.Text))
        myList.Add(CDbl(TextBox15.Text))
        myList.Add(CDbl(TextBox21.Text))
        myList.Add(CDbl(TextBox72.Text))
        myList.Add(CDbl(TextBox71.Text))
        myList.Add(CDbl(TextBox70.Text))
        myList.Add(CDbl(TextBox69.Text))
        myList.Add(CDbl(TextBox22.Text))
        myList.Add(CDbl(TextBox14.Text))
        myList.Add(CDbl(TextBox13.Text))
        myList.Add(CDbl(TextBox8.Text))
        myList.Add(CDbl(TextBox7.Text))
        myList.Add(CDbl(TextBox23.Text))
        myList.Add(CDbl(TextBox19.Text))
        myList.Add(CDbl(TextBox6.Text))
        myList.Add(CDbl(TextBox4.Text))
        myList.Add(CDbl(TextBox3.Text))
        myList.Add(CDbl(TextBox2.Text))
        myList.Add(CDbl(TextBox24.Text))
        myList.Add(CDbl(TextBox5.Text))

        myList.Sort()
        Dim ligne As String = ""

        For Each item In myList
            System.Diagnostics.Debug.WriteLine(item)

            Dim DblA As Double = item
            Dim VarA As String = FormatNumber(DblA)
            ligne = ligne + VarA + "/"

            'MsgBox(item)

        Next






    End Sub

    Private Sub proc_tab_paiement_direct()

        'Dans la fonction proc_tab_paiement_direct()



    End Sub

    Private Sub TextBox12_DragLeave(sender As Object, e As EventArgs) Handles TextBox12.DragLeave

    End Sub

    Private Sub TextBox12_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.Enter, TextBox11.Enter,
        TextBox10.Enter, TextBox9.Enter, TextBox20.Enter, TextBox1.Enter, TextBox18.Enter, TextBox17.Enter, TextBox16.Enter,
        TextBox15.Enter, TextBox21.Enter, TextBox72.Enter, TextBox71.Enter, TextBox70.Enter, TextBox69.Enter, TextBox22.Enter,
        TextBox14.Enter, TextBox13.Enter, TextBox8.Enter, TextBox7.Enter, TextBox23.Enter, TextBox19.Enter, TextBox6.Enter,
        TextBox4.Enter, TextBox3.Enter, TextBox2.Enter, TextBox24.Enter, TextBox5.Enter


        Dim tb As TextBox = CType(sender, TextBox)

        tb = CType(sender, TextBox)

        num_temporaire = tb.Text




    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown


        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)


        If num_temporaire <> "" Then



        End If

        'If rep = "Textbox1" Then
        '    categorie = "Solde_encaisses_debut"
        'End If



        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            TextBox12.Focus()

            
            periode_paiement_direct()
            periode_paiement_preautorise()
            periode_cheque_emis()


            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Solde_encaisses_debut set courant = @description WHERE numero = @numero", myConnection)

            myCommand.Parameters.AddWithValue("@description", TextBox1.Text)
            num_temporaire = tb.Text


            'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If

    End Sub


    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown, TextBox11.KeyDown, TextBox10.KeyDown, TextBox9.KeyDown, TextBox20.KeyDown



        Dim tb As TextBox = New TextBox
        Dim categorie As String
        Dim nom_table As String


        tb = CType(sender, TextBox)


        If num_temporaire <> "" Then

            Select Case tb.Name
                Case "TextBox12"
                    categorie = "courant"
                    nom_table = "Apports"
                Case "TextBox11"
                    categorie = "sept"
                    nom_table = "Apports"
                Case "TextBox10"
                    categorie = "quatorze"
                    nom_table = "Apports"
                Case "TextBox9"
                    categorie = "vingt"
                    nom_table = "Apports"
                Case "TextBox20"
                    categorie = "cinq"
                    nom_table = "Apports"
            End Select

        End If

        'If rep = "Textbox1" Then
        '    categorie = "Solde_encaisses_debut"
        'End If

        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            Select Case tb.Name
                Case "TextBox12"
                    TextBox11.Focus()
                Case "TextBox11"
                    TextBox10.Focus()
                Case "TextBox10"
                    TextBox9.Focus()
                Case "TextBox9"
                    TextBox20.Focus()
                Case "TextBox20"
                    TextBox18.Focus()


            End Select

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Apports set " & categorie & " = @description WHERE numero = @numero", myConnection)




            Select Case tb.Name
                Case "TextBox12"
                    myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
                    num_temporaire = tb.Text
                Case "TextBox11"
                    myCommand.Parameters.AddWithValue("@description", TextBox11.Text)
                    num_temporaire = tb.Text
                Case "TextBox10"
                    myCommand.Parameters.AddWithValue("@description", TextBox10.Text)
                    num_temporaire = tb.Text
                Case "TextBox9"
                    myCommand.Parameters.AddWithValue("@description", TextBox9.Text)
                    num_temporaire = tb.Text
                Case "TextBox20"
                    myCommand.Parameters.AddWithValue("@description", TextBox20.Text)
                    num_temporaire = tb.Text

            End Select


            'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)

            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If

    End Sub

    ' ******************************************************
    ' Encaisse minimum requise
    '-********************************************************




    ' ******************************************************
    '*********************************************************
    '-********************************************************

    ' ******************************************************
    ' Marge de crédit disponible
    '-********************************************************


    Sub proc_marge_credit_disp_req1()

    End Sub

    Sub proc_marge_credit_disp_req2()

    End Sub

    Sub proc_marge_credit_disp_req3()

    End Sub

  


    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress, TextBox11.KeyPress,
        TextBox10.KeyPress, TextBox9.KeyPress, TextBox20.KeyPress, TextBox1.KeyPress, TextBox18.KeyPress, TextBox17.KeyPress,
        TextBox16.KeyPress, TextBox15.KeyPress, TextBox21.KeyPress, TextBox72.KeyPress, TextBox71.KeyPress, TextBox70.KeyPress,
        TextBox69.KeyPress, TextBox22.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress,
        TextBox23.KeyPress, TextBox19.KeyPress, TextBox6.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress,
        TextBox24.KeyPress, TextBox5.KeyPress

        'MsgBox("retour")
        'num_temporaire = ""
        Dim allowedChars As String = "0123456789,"
        If allowedChars.IndexOf(e.KeyChar) = -1 AndAlso
            Not e.KeyChar = ChrW(8) Then
            ' Invalid Character
            e.Handled = True
        End If






        'If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
        '    'SendKeys.Send("{TAB}")
        '    e.Handled = True
        '    Me.Controls.Clear() 'removes all the controls on the form
        '    'TODO How to reload windows form without closing it using VB.NET

        '    InitializeComponent() 'load all the controls again

        '    frmGestion_Load(e, e) 'Load everything in your form load event again'
        '    MsgBox("Changed Success")


        'End If

        enregistrer = True



    End Sub

    Private Sub TextBox12_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.Leave, TextBox11.Leave,
        TextBox10.Leave, TextBox9.Leave, TextBox20.Leave, TextBox1.Leave, TextBox12.Leave, TextBox11.Leave, TextBox10.Leave,
        TextBox9.Leave, TextBox20.Leave, TextBox18.Leave, TextBox17.Leave, TextBox16.Leave, TextBox15.Leave, TextBox21.Leave,
        TextBox72.Leave, TextBox71.Leave, TextBox70.Leave, TextBox69.Leave, TextBox22.Leave, TextBox14.Leave, TextBox13.Leave,
        TextBox8.Leave, TextBox7.Leave, TextBox23.Leave, TextBox19.Leave, TextBox6.Leave, TextBox4.Leave, TextBox3.Leave, TextBox2.Leave,
        TextBox24.Leave, TextBox5.Leave




        Dim tb As TextBox = New TextBox

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then
            'Select Case tb.Name
            '    Case "TextBox12"
            '        'MsgBox(tb.Text)
            '        'TextBox12.Text = Format(Double.Parse(TextBox12.Text), "###,###,##0.00")
            '        ' TextBox12.Text = Format(Double.Parse(num_temporaire), "###,###,##0.00")
            '        TextBox12.Text = num_temporaire

            '    Case "TextBox11"
            '        TextBox11.Text = num_temporaire
            '    Case "TextBox10"
            '        TextBox10.Text = num_temporaire
            '    Case "TextBox9"
            '        TextBox9.Text = num_temporaire
            '    Case "TextBox20"
            '        TextBox20.Text = num_temporaire
            '    Case "TextBox1"
            '        TextBox1.Text = num_temporaire
            '    Case "TextBox18"
            '        TextBox18.Text = num_temporaire
            '    Case "TextBox17"
            '        TextBox17.Text = num_temporaire
            '    Case "TextBox16"
            '        TextBox16.Text = num_temporaire
            '    Case "TextBox15"
            '        TextBox15.Text = num_temporaire
            '    Case "TextBox21"
            '        TextBox21.Text = num_temporaire
            '    Case "TextBox72"
            '        TextBox72.Text = num_temporaire
            '    Case "TextBox71"
            '        TextBox71.Text = num_temporaire
            '    Case "TextBox70"
            '        TextBox70.Text = num_temporaire
            '    Case "TextBox69"
            '        TextBox69.Text = num_temporaire
            '    Case "TextBox22"
            '        TextBox22.Text = num_temporaire
            '    Case "TextBox14"
            '        TextBox14.Text = num_temporaire
            '    Case "TextBox13"
            '        TextBox13.Text = num_temporaire
            '    Case "TextBox8"
            '        TextBox8.Text = num_temporaire
            '    Case "TextBox7"
            '        TextBox7.Text = num_temporaire
            '    Case "TextBox23"
            '        TextBox23.Text = num_temporaire
            '    Case "TextBox6"
            '        TextBox6.Text = num_temporaire
            '    Case "TextBox4"
            '        TextBox4.Text = num_temporaire
            '    Case "TextBox3"
            '        TextBox3.Text = num_temporaire
            '    Case "TextBox2"
            '        TextBox2.Text = num_temporaire
            '    Case "TextBox24"
            '        TextBox24.Text = num_temporaire
            '    Case "TextBox5"
            '        TextBox5.Text = num_temporaire
            'End Select

            tb.Text = num_temporaire




        End If

        'Try
        '    ' Try to divide by zero.

        '    tb = CType(sender, TextBox)
        '    ' This statement is sadly not reached.
        '    'Console.WriteLine("Hi")
        'Catch ex As Exception
        '    ' Display the message.
        '    Select Case tb.Name
        '        Case "TextBox12"
        '            'MsgBox(tb.Text)
        '            'TextBox12.Text = Format(Double.Parse(TextBox12.Text), "###,###,##0.00")
        '            MsgBox("hey")
        '            TextBox12.Text = num_temporaire
        '        Case "TextBox11"
        '            'TextBox11.Text = Format(Double.Parse(TextBox11.Text), "###,###,##0.00")
        '        Case "TextBox10"
        '            'TextBox10.Text = Format(Double.Parse(TextBox10.Text), "###,###,##0.00")
        '        Case "TextBox9"
        '            'TextBox9.Text = Format(Double.Parse(TextBox9.Text), "###,###,##0.00")
        '        Case "TextBox20"
        '            'TextBox20.Text = Format(Double.Parse(TextBox20.Text), "###,###,##0.00 $")
        '    End Select
        'End Try







    End Sub


    Private Sub TextBox12_LostFocus(sender As Object, e As EventArgs) Handles TextBox12.LostFocus

        'Dim myConnection As New OleDb.OleDbConnection(myConnString)
        'Dim oledbAdapter As New OleDb.OleDbDataAdapter
        'myConnection.Open()

        'Dim myCommand As New OleDb.OleDbCommand("update Apports set courant = @description WHERE numero = @numero", myConnection)
        'myCommand.Parameters.AddWithValue("@description", Val(TextBox12.Text))
        'myCommand.Parameters.AddWithValue("@numero", 1)

        'myCommand.ExecuteNonQuery()
        'myConnection.Close()

    End Sub

    'Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        TextBox10.Focus()
    '        Dim myConnection As New OleDb.OleDbConnection(myConnString)
    '        Dim oledbAdapter As New OleDb.OleDbDataAdapter
    '        myConnection.Open()

    '        Dim myCommand As New OleDb.OleDbCommand("update Apports set sept = @description WHERE numero = @numero", myConnection)
    '        myCommand.Parameters.AddWithValue("@description", TextBox11.Text)
    '        myCommand.Parameters.AddWithValue("@numero", 1)

    '        myCommand.ExecuteNonQuery()
    '        myConnection.Close()

    '    End If

    'End Sub


    Private Sub TextBox11_LostFocus(sender As Object, e As EventArgs) Handles TextBox11.LostFocus

        'Dim myConnection As New OleDb.OleDbConnection(myConnString)
        'Dim oledbAdapter As New OleDb.OleDbDataAdapter
        'myConnection.Open()

        'Dim myCommand As New OleDb.OleDbCommand("update Apports set sept = @description WHERE numero = @numero", myConnection)
        'myCommand.Parameters.AddWithValue("@description", Val(TextBox11.Text))
        'myCommand.Parameters.AddWithValue("@numero", 1)

        'myCommand.ExecuteNonQuery()
        'myConnection.Close()

    End Sub





    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown, TextBox17.KeyDown, TextBox16.KeyDown, TextBox15.KeyDown, TextBox21.KeyDown



        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then

            Select Case tb.Name
                Case "TextBox18"
                    categorie = "courant"
                Case "TextBox17"
                    categorie = "sept"
                Case "TextBox16"
                    categorie = "quatorze"
                Case "TextBox15"
                    categorie = "vingt"
                Case "TextBox21"
                    categorie = "cinq"
            End Select

        End If



        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            Select Case tb.Name
                Case "TextBox18"
                    TextBox17.Focus()
                Case "TextBox17"
                    TextBox16.Focus()
                Case "TextBox16"
                    TextBox15.Focus()
                Case "TextBox15"
                    TextBox21.Focus()
                Case "TextBox21"
                    TextBox72.Focus()


            End Select

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Retrait_manuel set " & categorie & "= @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", tb.Text)
            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If
    End Sub

    ' ******************************************************
    ' RETRAIRE MANUEL
    '-********************************************************




   





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'test
       
        'frmMenu.Show()
        'Me.Dispose()

        'Me.Close()

        Me.Close()



        Me.WindowState = FormWindowState.Maximized
        frmMenu.MdiParent = Form_windows
        frmMenu.Show()



    End Sub

    Private Sub TextBox72_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox72.KeyDown, TextBox71.KeyDown, TextBox70.KeyDown, TextBox69.KeyDown, TextBox22.KeyDown

        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then

            Select Case tb.Name
                Case "TextBox72"
                    categorie = "courant"

                Case "TextBox71"
                    categorie = "sept"

                Case "TextBox70"
                    categorie = "quatorze"

                Case "TextBox69"
                    categorie = "vingt"

                Case "TextBox22"
                    categorie = "cinq"

            End Select

        End If


        If e.KeyCode = Keys.Enter Then

            num_temporaire = String.Empty

            Select Case tb.Name
                Case "TextBox72"
                    TextBox71.Focus()
                Case "TextBox71"
                    TextBox70.Focus()
                Case "TextBox70"
                    TextBox69.Focus()
                Case "TextBox69"
                    TextBox22.Focus()
                Case "TextBox22"
                    TextBox14.Focus()


            End Select


            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Salaires_paie_courante set " & categorie & " = @description WHERE numero = @numero", myConnection)


            Select Case tb.Name
                Case "TextBox72"
                    myCommand.Parameters.AddWithValue("@description", TextBox72.Text)
                    num_temporaire = tb.Text
                Case "TextBox71"
                    myCommand.Parameters.AddWithValue("@description", TextBox71.Text)
                    num_temporaire = tb.Text
                Case "TextBox70"
                    myCommand.Parameters.AddWithValue("@description", TextBox70.Text)
                    num_temporaire = tb.Text
                Case "TextBox69"
                    myCommand.Parameters.AddWithValue("@description", TextBox69.Text)
                    num_temporaire = tb.Text
                Case "TextBox22"
                    myCommand.Parameters.AddWithValue("@description", TextBox22.Text)
                    num_temporaire = tb.Text

            End Select



            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()

            proc_tout()


        End If

    End Sub

    ' ******************************************************
    ' Salaire paie courante
    '-********************************************************


    Private Sub TextBox72_LostFocus(sender As Object, e As EventArgs) Handles TextBox72.LostFocus
        'Dim myConnection As New OleDb.OleDbConnection(myConnString)
        'Dim oledbAdapter As New OleDb.OleDbDataAdapter
        'myConnection.Open()

        'Dim myCommand As New OleDb.OleDbCommand("update Salaires_paie_courante set courant = @description WHERE numero = @numero", myConnection)
        'myCommand.Parameters.AddWithValue("@description", Val(TextBox72.Text))
        'myCommand.Parameters.AddWithValue("@numero", 1)

        'myCommand.ExecuteNonQuery()
        'myConnection.Close()
    End Sub

    
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
    End Sub



    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        'Dim myConnection As New OleDb.OleDbConnection(myConnString)
        'Dim oledbAdapter As New OleDb.OleDbDataAdapter
        'myConnection.Open()

        'Dim myCommand As New OleDb.OleDbCommand("update Solde_encaisses_debut set courant = @description WHERE numero = @numero", myConnection)
        'myCommand.Parameters.AddWithValue("@description", Val(TextBox1.Text))
        'myCommand.Parameters.AddWithValue("@numero", 1)

        'myCommand.ExecuteNonQuery()

        'myConnection.Close()

        'TextBox1.Text = FormatCurrency(TextBox1.Text)

        'Dim myText = TextBox1.Text
        'Dim dIndex = myText.IndexOf("5")
        'If (dIndex > -1) Then
        '    MsgBox("Trouve")
        'Else
        '    MsgBox("non trouvé")
        'End If
    End Sub






    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown, TextBox13.KeyDown, TextBox8.KeyDown, TextBox7.KeyDown, TextBox23.KeyDown

        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then

            Select Case tb.Name
                Case "TextBox14"
                    categorie = "courant"

                Case "TextBox13"
                    categorie = "sept"

                Case "TextBox8"
                    categorie = "quatorze"

                Case "TextBox7"
                    categorie = "vingt"

                Case "TextBox23"
                    categorie = "cinq"

            End Select

        End If



        If e.KeyCode = Keys.Enter Then

            num_temporaire = String.Empty

            Select Case tb.Name
                Case "TextBox14"
                    TextBox13.Focus()
                Case "TextBox13"
                    TextBox8.Focus()
                Case "TextBox8"
                    TextBox7.Focus()
                Case "TextBox7"
                    TextBox23.Focus()
                Case "TextBox23"
                    TextBox19.Focus()


            End Select


            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Argent_a_recevoir set " & categorie & "  = @description WHERE numero = @numero", myConnection)

            Select Case tb.Name
                Case "TextBox14"
                    myCommand.Parameters.AddWithValue("@description", TextBox14.Text)
                    num_temporaire = tb.Text
                Case "TextBox13"
                    myCommand.Parameters.AddWithValue("@description", TextBox13.Text)
                    num_temporaire = tb.Text
                Case "TextBox8"
                    myCommand.Parameters.AddWithValue("@description", TextBox8.Text)
                    num_temporaire = tb.Text
                Case "TextBox7"
                    myCommand.Parameters.AddWithValue("@description", TextBox7.Text)
                    num_temporaire = tb.Text
                Case "TextBox23"
                    myCommand.Parameters.AddWithValue("@description", TextBox23.Text)
                    num_temporaire = tb.Text

            End Select



            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If
    End Sub

  


  
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown

        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            TextBox1.Focus()

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update marge_de_credit_utilise set courant = @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", TextBox5.Text)
            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()


        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown, TextBox4.KeyDown, TextBox3.KeyDown, TextBox2.KeyDown, TextBox24.KeyDown

        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then

            Select Case tb.Name
                Case "TextBox6"
                    categorie = "courant"

                Case "TextBox4"
                    categorie = "sept"

                Case "TextBox3"
                    categorie = "quatorze"

                Case "TextBox2"
                    categorie = "vingt"

                Case "TextBox24"
                    categorie = "cinq"

            End Select

        End If


        If e.KeyCode = Keys.Enter Then

            num_temporaire = String.Empty

            Select tb.Name
                Case "TextBox6"
                    TextBox4.Focus()
                Case "TextBox4"
                    TextBox3.Focus()
                Case "TextBox3"
                    TextBox2.Focus()
                Case "TextBox2"
                    TextBox24.Focus()
                Case "TextBox24"
                    TextBox5.Focus()
            End Select






            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update Encaisse_minimum_requise set " & categorie & " = @description WHERE numero = @numero", myConnection)

            Select Case tb.Name
                Case "TextBox6"
                    myCommand.Parameters.AddWithValue("@description", TextBox6.Text)
                    num_temporaire = tb.Text
                Case "TextBox4"
                    myCommand.Parameters.AddWithValue("@description", TextBox4.Text)
                    num_temporaire = tb.Text
                Case "TextBox3"
                    myCommand.Parameters.AddWithValue("@description", TextBox3.Text)
                    num_temporaire = tb.Text
                Case "TextBox2"
                    myCommand.Parameters.AddWithValue("@description", TextBox2.Text)
                    num_temporaire = tb.Text
                Case "TextBox24"
                    myCommand.Parameters.AddWithValue("@description", TextBox24.Text)
                    num_temporaire = tb.Text

            End Select


            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            TextBox6.Focus()

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            Dim oledbAdapter As New OleDb.OleDbDataAdapter
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("update marge_credit_disponible set courant = @description WHERE numero = @numero", myConnection)
            myCommand.Parameters.AddWithValue("@description", TextBox19.Text)
            myCommand.Parameters.AddWithValue("@numero", 1)

            myCommand.ExecuteNonQuery()
            myConnection.Close()
            proc_tout()

        End If

    End Sub

    Sub periode_paiement_direct()


        Dim numero As Integer
        Dim compteur As Integer = 0
        Dim montant As Double


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select * from check_direct", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()


        Dim tab As Double() = {0, 0, 0, 0, 0}


        While myReader1.Read
            numero = myReader1.Item(0)
            tab(0) = myReader1.Item(4)
            tab(1) = myReader1.Item(5)
            tab(2) = myReader1.Item(6)
            tab(3) = myReader1.Item(7)
            tab(4) = myReader1.Item(8)

            Dim periode As String = ""


            Dim date_paiement As DateTime


            If (tab(0) <> 0) Then
                periode = "un"
                montant = tab(0)
            End If

            If (tab(1) <> 0) Then
                periode = "deux"
                montant = tab(1)

            End If

            If (tab(2) <> 0) Then
                periode = "trois"
                montant = tab(2)
            End If

            If (tab(3) <> 0) Then
                periode = "quatre"
                montant = tab(3)

            End If

            If (tab(4) <> 0) Then
                periode = "cinq"
                montant = tab(4)

            End If

            date_paiement = myReader1.Item(3)

            fctable_paiement_direct(montant, numero, date_paiement)


        End While


        myConnection.Close()








    End Sub

    Sub periode_cheque_emis()

        Dim numero As Integer
        Dim compteur As Integer = 0
        Dim montant As Double


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select * from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()


        Dim tab As Double() = {0, 0, 0, 0, 0}


        While myReader1.Read
            numero = myReader1.Item(0)
            tab(0) = myReader1.Item(5)
            tab(1) = myReader1.Item(6)
            tab(2) = myReader1.Item(7)
            tab(3) = myReader1.Item(8)
            tab(4) = myReader1.Item(9)

            Dim periode As String = ""


            Dim date_paiement As DateTime


            If (tab(0) <> 0) Then
                periode = "un"
                montant = tab(0)
            End If

            If (tab(1) <> 0) Then
                periode = "deux"
                montant = tab(1)

            End If

            If (tab(2) <> 0) Then
                periode = "trois"
                montant = tab(2)
            End If

            If (tab(3) <> 0) Then
                periode = "quatre"
                montant = tab(3)

            End If

            If (tab(4) <> 0) Then
                periode = "cinq"
                montant = tab(4)

            End If

            date_paiement = myReader1.Item(3)

            fctable_cheque_emis(montant, numero, date_paiement)


        End While


        myConnection.Close()








    End Sub

    Sub periode_paiement_preautorise()

        Dim numero As Integer
        Dim compteur As Integer = 0
        Dim montant As Double


        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select * from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()


        Dim tab As Double() = {0, 0, 0, 0, 0}


        While myReader1.Read
            numero = myReader1.Item(0)
            tab(0) = myReader1.Item(4)
            tab(1) = myReader1.Item(5)
            tab(2) = myReader1.Item(6)
            tab(3) = myReader1.Item(7)
            tab(4) = myReader1.Item(8)

            Dim periode As String = ""


            Dim date_paiement As DateTime


            If (tab(0) <> 0) Then
                periode = "un"
                montant = tab(0)
            End If

            If (tab(1) <> 0) Then
                periode = "deux"
                montant = tab(1)

            End If

            If (tab(2) <> 0) Then
                periode = "trois"
                montant = tab(2)
            End If

            If (tab(3) <> 0) Then
                periode = "quatre"
                montant = tab(3)

            End If

            If (tab(4) <> 0) Then
                periode = "cinq"
                montant = tab(4)

            End If

            date_paiement = myReader1.Item(3)

            fctable_paiement_preautorise(montant, numero, date_paiement)


        End While


        myConnection.Close()








    End Sub




    Function fctable_paiement_direct(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)


        Dim reponse As String = ""
        Dim tab As Double() = {0, 0, 0, 0, 0}
        Dim cpt As Integer = 0
        'Dim currentDate As DateTime
        'currentDate = Now.Date


        ' Tester avant la date de paiement

        'Dim dtnow As Date = Date.Now
        Dim dtnow As Date = auj

        'If dtnow < date_ Then

        '    tab(0) = montant
        'End If

        Dim currentDate1 As Date

        'currentDate = currentDate.AddDays(1)

        'Dim currentDate As Date = Date.Now
        Dim currentDate As Date = auj



        If date_.Date <= currentDate.Date Then

            tab(cpt) = montant
        End If




        currentDate1 = currentDate.AddDays(6)



        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant
        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1


        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If


        If date_.Date >= currentDate.Date Then

            tab(cpt) = montant
        End If


        'If dtnow > date_ Then

        '    tab(4) = montant

        'End If



        'MsgBox("Test débogger")
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        Dim oledbAdapter As New OleDb.OleDbDataAdapter
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("update check_direct set un = @num1, deux = @num2, trois = @num3, quatre = @num4, cinq = @num5 WHERE numero = @numero", myConnection)

        myCommand.Parameters.AddWithValue("@num1", tab(0))
        myCommand.Parameters.AddWithValue("@num2", tab(1))
        myCommand.Parameters.AddWithValue("@num3", tab(2))
        myCommand.Parameters.AddWithValue("@num4", tab(3))
        myCommand.Parameters.AddWithValue("@num5", tab(4))
        'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
        myCommand.Parameters.AddWithValue("@numero", numero)

        myCommand.ExecuteNonQuery()
        myConnection.Close()











        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero ='" & numero & "'", myConnection)
        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero = 75", myConnection)


    End Function


    Function fctable_cheque_emis(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)


        Dim reponse As String = ""
        Dim tab As Double() = {0, 0, 0, 0, 0}
        Dim cpt As Integer = 0
        'Dim currentDate As DateTime
        'currentDate = Now.Date


        ' Tester avant la date de paiement

        'Dim dtnow As Date = Date.Now
        Dim dtnow As Date = auj


        'If dtnow < date_ Then

        '    tab(0) = montant
        'End If

        Dim currentDate1 As Date

        'currentDate = currentDate.AddDays(1)

        Dim currentDate As Date = auj
        'Dim currentDate As Date = Date.Now


        If date_.Date <= currentDate.Date Then

            tab(cpt) = montant
        End If




        currentDate1 = currentDate.AddDays(6)



        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant
        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1


        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If


        If date_.Date >= currentDate.Date Then

            tab(cpt) = montant
        End If


        'If dtnow > date_ Then

        '    tab(4) = montant

        'End If



        'MsgBox("Test débogger")
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        Dim oledbAdapter As New OleDb.OleDbDataAdapter
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("update check_emis set un = @num1, deux = @num2, trois = @num3, quatre = @num4, cinq = @num5 WHERE numero = @numero", myConnection)

        myCommand.Parameters.AddWithValue("@num1", tab(0))
        myCommand.Parameters.AddWithValue("@num2", tab(1))
        myCommand.Parameters.AddWithValue("@num3", tab(2))
        myCommand.Parameters.AddWithValue("@num4", tab(3))
        myCommand.Parameters.AddWithValue("@num5", tab(4))
        'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
        myCommand.Parameters.AddWithValue("@numero", numero)

        myCommand.ExecuteNonQuery()
        myConnection.Close()











        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero ='" & numero & "'", myConnection)
        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero = 75", myConnection)


    End Function


    Function fctable_paiement_preautorise(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)


        Dim reponse As String = ""
        Dim tab As Double() = {0, 0, 0, 0, 0}
        Dim cpt As Integer = 0
        'Dim currentDate As DateTime
        'currentDate = Now.Date


        ' Tester avant la date de paiement

        'Dim dtnow As Date = Date.Now7
        Dim dtnow As Date = auj

        'If dtnow < date_ Then

        '    tab(0) = montant
        'End If

        Dim currentDate1 As Date

        'currentDate = currentDate.AddDays(1)

        'Dim currentDate As Date =  
        Dim currentDate As Date = auj




        If date_.Date <= currentDate.Date Then

            tab(cpt) = montant
        End If




        currentDate1 = currentDate.AddDays(6)



        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant
        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1

        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If

        currentDate = currentDate1.AddDays(1)
        currentDate1 = currentDate.AddDays(6)
        cpt = cpt + 1


        If date_.Date >= currentDate.Date And date_.Date <= currentDate1.Date And cpt < 5 Then

            tab(cpt) = montant

        End If


        If date_.Date >= currentDate.Date Then

            tab(cpt) = montant
        End If


        'If dtnow > date_ Then

        '    tab(4) = montant

        'End If



        'MsgBox("Test débogger")
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        Dim oledbAdapter As New OleDb.OleDbDataAdapter
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("update check_preautorises set un = @num1, deux = @num2, trois = @num3, quatre = @num4, cinq = @num5 WHERE numero = @numero", myConnection)

        myCommand.Parameters.AddWithValue("@num1", tab(0))
        myCommand.Parameters.AddWithValue("@num2", tab(1))
        myCommand.Parameters.AddWithValue("@num3", tab(2))
        myCommand.Parameters.AddWithValue("@num4", tab(3))
        myCommand.Parameters.AddWithValue("@num5", tab(4))
        'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
        myCommand.Parameters.AddWithValue("@numero", numero)

        myCommand.ExecuteNonQuery()
        myConnection.Close()











        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero ='" & numero & "'", myConnection)
        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero = 75", myConnection)


    End Function


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox12_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox12.MouseClick, TextBox1.MouseClick, TextBox11.MouseClick, TextBox10.MouseClick, TextBox9.MouseClick, TextBox20.MouseClick, TextBox18.MouseClick,
        TextBox17.MouseClick, TextBox16.MouseClick, TextBox15.MouseClick, TextBox21.MouseClick, TextBox72.MouseClick, TextBox71.MouseClick,
         TextBox70.MouseClick, TextBox69.MouseClick, TextBox22.MouseClick, TextBox14.MouseClick, TextBox13.MouseClick, TextBox8.MouseClick,
         TextBox7.MouseClick, TextBox23.MouseClick, TextBox19.MouseClick, TextBox6.MouseClick, TextBox4.MouseClick, TextBox3.MouseClick,
         TextBox2.MouseClick, TextBox24.MouseClick, TextBox5.MouseClick

        Dim tb As TextBox = New TextBox
        Dim categorie As String

        tb = CType(sender, TextBox)

        tb.SelectAll()





        'TextBox12.SelectAll()

    End Sub

    Sub recurer_backup()


        Dim num_tab As String

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM bidon where numero = " & num_backup & "", myConnection)

        Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


        While myReader.Read
            num_tab = myReader.Item(3).ToString


        End While

        myConnection.Close()

        Dim strArr() As String

        strArr = num_tab.Split("/")

        placer_espace(strArr)


    End Sub

    Sub placer_espace(ByRef tab() As String)

        TextBox1.Text = Format$(tab(0), "Currency")
        TextBox12.Text = Format$(tab(1), "Currency")
        TextBox11.Text = Format$(tab(2), "Currency")
        TextBox10.Text = Format$(tab(3), "Currency")
        TextBox9.Text = Format$(tab(4), "Currency")
        TextBox20.Text = Format$(tab(5), "Currency")
        TextBox18.Text = Format$(tab(6), "Currency")
        TextBox17.Text = Format$(tab(7), "Currency")
        TextBox16.Text = Format$(tab(8), "Currency")
        TextBox15.Text = Format$(tab(9), "Currency")
        TextBox21.Text = Format$(tab(10), "Currency")


        Label17.Text = Format$(tab(11), "Currency")
        Label17.Text = replacer_negative(Label17.Text)

        Label19.Text = Format$(tab(12), "Currency")
        Label19.Text = replacer_negative(Label19.Text)
        Label20.Text = Format$(tab(13), "Currency")
        Label20.Text = replacer_negative(Label20.Text)
        Label21.Text = Format$(tab(14), "Currency")
        Label21.Text = replacer_negative(Label21.Text)
        Label36.Text = Format$(tab(15), "Currency")
        Label36.Text = replacer_negative(Label36.Text)
        Label25.Text = Format$(tab(16), "Currency")
        Label25.Text = replacer_negative(Label25.Text)
        Label24.Text = Format$(tab(17), "Currency")
        Label24.Text = replacer_negative(Label24.Text)
        Label23.Text = Format$(tab(18), "Currency")
        Label23.Text = replacer_negative(Label23.Text)
        Label22.Text = Format$(tab(19), "Currency")
        Label22.Text = replacer_negative(Label22.Text)
        Label37.Text = Format$(tab(20), "Currency")
        Label37.Text = replacer_negative(Label37.Text)
        Label29.Text = Format$(tab(21), "Currency")
        Label29.Text = replacer_negative(Label29.Text)
        Label28.Text = Format$(tab(22), "Currency")
        Label28.Text = replacer_negative(Label28.Text)
        Label27.Text = Format$(tab(23), "Currency")
        Label27.Text = replacer_negative(Label27.Text)
        Label26.Text = Format$(tab(24), "Currency")
        Label26.Text = replacer_negative(Label26.Text)
        Label41.Text = Format$(tab(25), "Currency")
        Label41.Text = replacer_negative(Label41.Text)
        TextBox72.Text = Format$(tab(26), "Currency")
        TextBox71.Text = Format$(tab(27), "Currency")
        TextBox70.Text = Format$(tab(28), "Currency")
        TextBox69.Text = Format$(tab(29), "Currency")
        TextBox22.Text = Format$(tab(30), "Currency")

        Label68.Text = Format$(tab(31), "Currency")
        Label68.Text = replacer_negative(Label68.Text)
        Label67.Text = Format$(tab(39), "Currency")
        Label67.Text = replacer_negative(Label67.Text)
        Label66.Text = Format$(tab(42), "Currency")
        Label66.Text = replacer_negative(Label66.Text)
        Label65.Text = Format$(tab(45), "Currency")
        Label65.Text = replacer_negative(Label65.Text)
        Label42.Text = Format$(tab(65), "Currency")
        Label42.Text = replacer_negative(Label42.Text)

        TextBox14.Text = Format$(tab(32), "Currency")
        TextBox13.Text = Format$(tab(33), "Currency")
        TextBox8.Text = Format$(tab(34), "Currency")
        TextBox7.Text = Format$(tab(35), "Currency")
        TextBox23.Text = Format$(tab(36), "Currency")

        Label60.Text = Format$(tab(37), "Currency")
        Label60.Text = replacer_negative(Label60.Text)
        Label59.Text = Format$(tab(40), "Currency")
        Label59.Text = replacer_negative(Label59.Text)
        Label58.Text = Format$(tab(43), "Currency")
        Label58.Text = replacer_negative(Label58.Text)
        Label57.Text = Format$(tab(46), "Currency")
        Label57.Text = replacer_negative(Label57.Text)
        Label43.Text = Format$(tab(66), "Currency")
        Label43.Text = replacer_negative(Label43.Text)

        TextBox19.Text = Format$(tab(48), "Currency")
        Label55.Text = Format$(tab(56), "Currency")
        Label55.Text = replacer_negative(Label55.Text)
        Label54.Text = Format$(tab(59), "Currency")
        Label54.Text = replacer_negative(Label54.Text)
        Label53.Text = Format$(tab(62), "Currency")
        Label53.Text = replacer_negative(Label53.Text)
        Label44.Text = Format$(tab(67), "Currency")
        Label44.Text = replacer_negative(Label44.Text)

        Label52.Text = Format$(tab(49), "Currency")
        Label52.Text = replacer_negative(Label52.Text)
        Label51.Text = Format$(tab(57), "Currency")
        Label51.Text = replacer_negative(Label51.Text)
        Label50.Text = Format$(tab(60), "Currency")
        Label50.Text = replacer_negative(Label50.Text)
        Label49.Text = Format$(tab(63), "Currency")
        Label49.Text = replacer_negative(Label49.Text)
        Label45.Text = Format$(tab(68), "Currency")
        Label45.Text = replacer_negative(Label45.Text)

        TextBox6.Text = Format$(tab(50), "Currency")
        TextBox4.Text = Format$(tab(51), "Currency")
        TextBox3.Text = Format$(tab(52), "Currency")
        TextBox2.Text = Format$(tab(53), "Currency")
        TextBox24.Text = Format$(tab(54), "Currency")

        Label32.Text = Format$(tab(55), "Currency")
        Label32.Text = replacer_negative(Label32.Text)
        Label31.Text = Format$(tab(58), "Currency")
        Label31.Text = replacer_negative(Label31.Text)
        Label30.Text = Format$(tab(61), "Currency")
        Label30.Text = replacer_negative(Label30.Text)
        Label13.Text = Format$(tab(56), "Currency")
        Label13.Text = replacer_negative(Label13.Text)
        Label46.Text = Format$(tab(69), "Currency")
        Label46.Text = replacer_negative(Label46.Text)

        TextBox5.Text = Format$(tab(47), "Currency")
        Label39.Text = Format$(tab(70), "Currency")
        Label39.Text = replacer_negative(Label39.Text)

        Label47.Text = Format$(tab(71), "Currency")
        Label47.Text = replacer_negative(Label47.Text)

        Label70.Text = Format$(tab(72), "Currency")
        Label70.Text = replacer_negative(Label70.Text)
        Label71.Text = Format$(tab(73), "Currency")
        Label71.Text = replacer_negative(Label71.Text)
        Label72.Text = Format$(tab(74), "Currency")
        Label72.Text = replacer_negative(Label72.Text)
        Label73.Text = Format$(tab(75), "Currency")
        Label73.Text = replacer_negative(Label73.Text)
        Label74.Text = Format$(tab(76), "Currency")
        Label74.Text = replacer_negative(Label74.Text)







        Label14.Text = Format$(tab(38), "Currency")
        Label14.Text = replacer_negative(Label14.Text)
        Label15.Text = Format$(tab(41), "Currency")
        Label15.Text = replacer_negative(Label15.Text)
        Label16.Text = Format$(tab(44), "Currency")
        Label16.Text = replacer_negative(Label16.Text)
        Label35.Text = Format$(tab(64), "Currency")
        Label35.Text = replacer_negative(Label35.Text)

        desactive()



    End Sub


    Function replacer_negative(ByVal nom_label As String) As String

        Dim S, s2 As String
        S = nom_label
        Dim answer As Char
        answer = S.Substring(0, 1)
        If answer = "(" Then

            s2 = S.Substring(0, S.Length - 1)

            nom_label = s2.Replace(CChar("("), CChar("-"))

        Else
            'MsgBox("Positif")
        End If




        Return nom_label


    End Function

    Sub desactive()

        If backup = True Then

            TextBox1.Enabled = False
            TextBox12.Enabled = False
            TextBox11.Enabled = False
            TextBox10.Enabled = False
            TextBox9.Enabled = False
            TextBox20.Enabled = False
            TextBox18.Enabled = False
            TextBox17.Enabled = False
            TextBox16.Enabled = False
            TextBox15.Enabled = False
            TextBox21.Enabled = False

            TextBox72.Enabled = False
            TextBox71.Enabled = False
            TextBox70.Enabled = False
            TextBox69.Enabled = False
            TextBox22.Enabled = False
            TextBox14.Enabled = False
            TextBox13.Enabled = False
            TextBox8.Enabled = False
            TextBox7.Enabled = False
            TextBox23.Enabled = False
            TextBox19.Enabled = False
            TextBox6.Enabled = False
            TextBox4.Enabled = False
            TextBox3.Enabled = False
            TextBox2.Enabled = False
            TextBox24.Enabled = False
            TextBox5.Enabled = False

        Else
            TextBox1.Enabled = True
            TextBox12.Enabled = True
            TextBox11.Enabled = True
            TextBox10.Enabled = True
            TextBox9.Enabled = True
            TextBox20.Enabled = True
            TextBox18.Enabled = True
            TextBox17.Enabled = True
            TextBox16.Enabled = True
            TextBox15.Enabled = True
            TextBox21.Enabled = True

            TextBox72.Enabled = True
            TextBox71.Enabled = True
            TextBox70.Enabled = True
            TextBox69.Enabled = True
            TextBox22.Enabled = True
            TextBox14.Enabled = True
            TextBox13.Enabled = True
            TextBox8.Enabled = True
            TextBox7.Enabled = True
            TextBox23.Enabled = True
            TextBox19.Enabled = True
            TextBox6.Enabled = True
            TextBox4.Enabled = True
            TextBox3.Enabled = True
            TextBox2.Enabled = True
            TextBox24.Enabled = True
            TextBox5.Enabled = True

        End If



    End Sub

    Function check_zero(nbr As String) As String

        If nbr = "" Then
            nbr = "0"
        End If

        nbr = Format$(nbr, "Currency")

        Return nbr

    End Function

    Function conventir_format(nbr As String) As String

        Dim dblValue As Double = nbr
        Dim after2 As String = dblValue.ToString("#.00")

        after2 = Format$(after2, "Currency")

        Dim S, s2 As String

        S = after2
        Dim answer As Char
        answer = S.Substring(0, 1)
        If answer = "(" Then

            s2 = S.Substring(0, S.Length - 1)

            after2 = s2.Replace(CChar("("), CChar("-"))

        End If




        Return after2

    End Function



End Class
