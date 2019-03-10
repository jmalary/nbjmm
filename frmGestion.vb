Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb
Imports System.Globalization

Public Class frmGestion

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.test2ConnectionString


    Dim AL As New Collection
    Dim lister_cellule As New ArrayList
    Dim tab_type() As String
    Dim lister_arrt As New ArrayList
    Dim lister_arra As New ArrayList


    Dim somme_solde_enc As Double
    Dim num_temporaire As String

    Private Property OleDbDataReader As OleDbDataReader

    Dim Al_un, Al_deux, Al_trois, Al_quatre,
        Al_cinq, Al_six, Al_sept, Al_huit, Al_neuf,
        Al_dix, Al_onze, Al_douze, Al_trieze,
        Al_quatorze, Al_cinquieme, Al_sixieme,
        Al_septieme, Al_huitieme As New Collection

    Dim arrt_un(), arrt_deux(), arrt_trois(), arrt_quatre(),
        arrt_cinq(), arrt_six(), arrt_sept(), arrt_huit(), arrt_neuf(),
        arrt_dix(), arrt_onze(), arrt_douze(), arrt_treize(),
        arrt_quatieme(), arrt_cinquieme(), arrt_sixieme(), arrt_septieme(),
        arrt_huitieme() As String


    Dim arra_un(), arra_deux(), arra_trois(), arra_quatre(),
        arra_cinq(), arra_six(), arra_sept(), arra_huit(), arra_neuf(),
        arra_dix(), arra_onze(), arra_douze(), arra_treize(),
        arra_quatieme(), arra_cinquieme(), arra_sixieme(), arra_septieme(),
        arra_huitieme() As String



    Private Sub frmGestion_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub




    Private Sub frmGestion_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        'Me.Size = New Size(1378, 900)

        frmExcel.testpdf()

        If backup = False Then

            proc_ListerText()

            Dim currentDate As DateTime
            currentDate = Now.Date
            currentDate = currentDate.AddDays(6)


            Initialisation()
            Initialiser_type()
            Initialiser_afficher()
            proc_gestion()
            proc_restaurantion()




        Else

            frmLoading.Close()
            proc_negative()
            recurer_backup() 'sixieme

        End If

        desactive()

        show_tooltip_on_textbox()




    End Sub




    Private Sub Initialisation()
        Al_un.Add(TextBox1)

        Al_un.Add(Label14)
        Al_un.Add(Label15)
        Al_un.Add(Label16)
        Al_un.Add(Label35)
        Al_deux.Add(TextBox12)
        Al_deux.Add(TextBox11)
        Al_deux.Add(TextBox10)
        Al_deux.Add(TextBox9)
        Al_deux.Add(TextBox20)

        Al_trois.Add(Label82)
        Al_trois.Add(Label83)
        Al_trois.Add(Label84)
        Al_trois.Add(Label85)
        Al_trois.Add(Label86)
        '--------------------
        Al_quatre.Add(Label80)
        Al_quatre.Add(Label79)
        Al_quatre.Add(Label78)
        Al_quatre.Add(Label77)
        Al_quatre.Add(Label76)
        Al_cinq.Add(Label17)
        Al_cinq.Add(Label19)
        Al_cinq.Add(Label20)
        Al_cinq.Add(Label21)
        Al_cinq.Add(Label36)
        Al_six.Add(Label25)
        Al_six.Add(Label24)
        Al_six.Add(Label23)
        Al_six.Add(Label22)
        Al_six.Add(Label37)
        Al_sept.Add(Label29)
        Al_sept.Add(Label28)
        Al_sept.Add(Label27)
        Al_sept.Add(Label26)
        Al_sept.Add(Label41)

        'Al_huit.Add(TextBox72)
        'Al_huit.Add(TextBox71)
        'Al_huit.Add(TextBox70)
        'Al_huit.Add(TextBox69)
        'Al_huit.Add(TextBox22)
        '

        Al_huit.Add(Label87)
        Al_huit.Add(Label88)
        Al_huit.Add(Label89)
        Al_huit.Add(Label90)
        Al_huit.Add(Label91)

        '--------------------
        Al_neuf.Add(TextBox14)
        Al_neuf.Add(TextBox13)
        Al_neuf.Add(TextBox8)
        Al_neuf.Add(TextBox7)
        Al_neuf.Add(TextBox23)
        Al_dix.Add(Label68)
        Al_dix.Add(Label67)
        Al_dix.Add(Label66)
        Al_dix.Add(Label65)
        Al_dix.Add(Label42)
        Al_onze.Add(Label60)
        Al_onze.Add(Label59)
        Al_onze.Add(Label58)
        Al_onze.Add(Label57)
        Al_onze.Add(Label43)
        'Al_onze.Add(Label47)
        Al_douze.Add(TextBox19)
        Al_douze.Add(Label55)
        Al_douze.Add(Label54)
        Al_douze.Add(Label53)
        Al_douze.Add(Label44)
        Al_trieze.Add(Label52)
        Al_trieze.Add(Label51)
        Al_trieze.Add(Label50)
        Al_trieze.Add(Label49)
        Al_trieze.Add(Label45)
        Al_quatorze.Add(TextBox6)
        Al_quatorze.Add(TextBox4)
        Al_quatorze.Add(TextBox3)
        Al_quatorze.Add(TextBox2)
        Al_quatorze.Add(TextBox24)
        Al_cinquieme.Add(Label32)
        Al_cinquieme.Add(Label31)
        Al_cinquieme.Add(Label30)
        Al_cinquieme.Add(Label13)
        Al_cinquieme.Add(Label46)
        Al_sixieme.Add(Label70)
        Al_sixieme.Add(Label71)
        Al_sixieme.Add(Label72)
        Al_sixieme.Add(Label73)
        Al_sixieme.Add(Label74)
        'Al_septieme.Add(TextBox5)
        'Al_huitieme.Add(Label39)

        lister_cellule.Add(Al_un)
        lister_cellule.Add(Al_deux)
        lister_cellule.Add(Al_trois)
        lister_cellule.Add(Al_quatre)
        lister_cellule.Add(Al_cinq)
        lister_cellule.Add(Al_six)
        lister_cellule.Add(Al_sept)
        lister_cellule.Add(Al_huit)
        lister_cellule.Add(Al_neuf)
        lister_cellule.Add(Al_dix)
        lister_cellule.Add(Al_onze)
        lister_cellule.Add(Al_douze)
        lister_cellule.Add(Al_trieze)
        lister_cellule.Add(Al_quatorze)
        lister_cellule.Add(Al_cinquieme)
        lister_cellule.Add(Al_sixieme)
        'lister_cellule.Add(Al_septieme)
        'lister_cellule.Add(Al_huitieme)


    End Sub


    Private Sub Initialiser_type()

        arrt_un = {"textbox", "label", "label", "label", "label"} '0
        arrt_deux = {"textbox", "textbox", "textbox", "textbox", "textbox"} '1
        arrt_trois = {"textbox", "textbox", "textbox", "textbox", "textbox"} '2


        arrt_quatre = {"label", "label", "label", "label", "label"} '3



        arrt_cinq = {"textbox", "textbox", "textbox", "textbox", "textbox"} '4
        arrt_six = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_sept = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_huit = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_neuf = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_dix = {"label", "label", "label", "label", "label"}
        'arrt_onze = {"label", "label", "label", "label", "label", "label"}
        arrt_onze = {"label", "label", "label", "label", "label"}
        arrt_douze = {"textbox", "label", "label", "label", "label"}
        arrt_treize = {"label", "label", "label", "label", "label"}
        arrt_quatieme = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_cinquieme = {"label", "label", "label", "label", "label"}
        arrt_sixieme = {"label", "label", "label", "label", "label"}
        'arrt_septieme = {"textbox"}
        'arrt_huitieme = {"label"}

        lister_arrt.Add(arrt_un)
        lister_arrt.Add(arrt_deux)
        lister_arrt.Add(arrt_trois)
        lister_arrt.Add(arrt_quatre)
        lister_arrt.Add(arrt_cinq)
        lister_arrt.Add(arrt_six)
        lister_arrt.Add(arrt_sept)
        lister_arrt.Add(arrt_huit)
        lister_arrt.Add(arrt_neuf)
        lister_arrt.Add(arrt_dix)
        lister_arrt.Add(arrt_onze)
        lister_arrt.Add(arrt_douze)
        lister_arrt.Add(arrt_treize)
        lister_arrt.Add(arrt_quatieme)
        lister_arrt.Add(arrt_cinquieme)
        lister_arrt.Add(arrt_sixieme)
        'lister_arrt.Add(arrt_septieme)
        'lister_arrt.Add(arrt_huitieme)


    End Sub

    Private Sub Initialiser_afficher()

        arra_un = {"Solde_encaisses_debut", "$10&1?13&1", "$10&2?13&2", "$10&3?13&3", "$10&4?13&4"}
        arra_deux = {"Apports", "Apports", "Apports", "Apports", "Apports"}
        arra_trois = {"check_revenu", "check_revenu", "check_revenu", "check_revenu", "check_revenu"}


        'arra_quatre = {"+1&1?2&1?3&1", "+1&2?2&2?3&2", "+1&3?2&3?3&3", "+1&4?2&4?3&4", "+1&5?2&5?3&5"}
        'arra_quatre = {"+0&0?1&0?2&0", "+0&1?1&1?2&1", "+0&2?1&2?2&2", "+0&3?1&3?2&3", "+0&4?1&4?2&4"}
        arra_quatre = {"+0&1?1&1?2&1", "+0&2?1&2?2&2", "+0&3?1&3?2&3", "+0&4?1&4?2&4", "+0&5?1&5?2&5"}


        arra_cinq = {"check_direct", "check_direct", "check_direct", "check_direct", "check_direct"}
        arra_six = {"check_preautorises", "check_preautorises", "check_preautorises", "check_preautorises", "check_preautorises"}
        arra_sept = {"check_emis", "check_emis", "check_emis", "check_emis", "check_emis"}
        arra_huit = {"check_depenses", "check_depenses", "check_depenses", "check_depenses", "check_depenses"}
        arra_neuf = {"Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir"}

        arra_dix = {"+4&1?5&1?6&1?7&1?8&1", "+4&2?5&2?6&2?7&2?8&2", "+4&3?5&3?6&3?7&3?8&3", "+4&4?5&4?6&4?7&4?8&4", "+4&5?5&5?6&5?7&5?8&5"}
        'arra_onze = {"-4&1?10&1", "-4&2?10&2", "-4&3?10&3", "-4&4?10&4", "-4&5?10&5", "+10&1?10&2?10&3?10&4?10&5"}
        'arra_onze = {"-4&1?10&1", "-4&2?10&2", "-4&3?10&3", "-4&4?10&4", "-4&5?10&5"}
        arra_onze = {"-3&1?9&1", "-3&2?9&2", "-3&3?9&3", "-3&4?9&4", "-3&5?9&5"}

        arra_douze = {"marge_credit_disponible", "-11&1?14&1", "-11&2?14&2", "-11&3?14&3", "-11&4?14&4"}
        'arra_treize = {"+11&1?12&1", "+11&2?12&2", "+11&3?12&3", "+11&4?12&4", "+11&4?12&4"}
        arra_treize = {"+10&1?11&1", "+10&2?11&2", "+10&3?11&3", "+10&4?11&4", "+10&5?11&5"}
        arra_quatieme = {"Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise"}
        'arra_cinquieme = {"*11&1?14&1", "*11&2?14&2", "*11&3?14&3", "*11&4?14&4", "*11&5?14&5"}
        ' modifier CT
        'ancien
        arra_cinquieme = {"*10&1?13&1", "*10&2?13&2", "*10&3?13&3", "*10&4?13&4", "*10&5?13&5"}

        'nouveau
        'arra_cinquieme = {"*12&1?13&1", "*12&2?13&2", "*12&3?13&3", "*12&4?13&4", "*12&5?13&5"}

        'arra_sixieme = {"@12&1?15&1", "@12&2?15&2", "@12&3?15&3", "@12&4?15&4", "@12&5?15&5"}
        arra_sixieme = {"@11&1?14&1", "@11&2?14&2", "@11&3?14&3", "@11&4?14&4", "@11&5?14&5"}
        'arra_septieme = {"marge_de_credit_utilise"}
        'arra_huitieme = {"15&1?15&2?15&3?15&4?15&5w?"}

        lister_arra.Add(arra_un)
        lister_arra.Add(arra_deux)
        lister_arra.Add(arra_trois)
        lister_arra.Add(arra_quatre)
        lister_arra.Add(arra_cinq)
        lister_arra.Add(arra_six)
        lister_arra.Add(arra_sept)
        lister_arra.Add(arra_huit)
        lister_arra.Add(arra_neuf)
        lister_arra.Add(arra_dix)
        lister_arra.Add(arra_onze)
        lister_arra.Add(arra_douze)
        lister_arra.Add(arra_treize)
        lister_arra.Add(arra_quatieme)
        lister_arra.Add(arra_cinquieme)
        lister_arra.Add(arra_sixieme)
        'lister_arra.Add(arra_septieme)
        'lister_arra.Add(arra_huitieme)

    End Sub

    Sub proc_ListerText()
        AL.Add(TextBox1)

    End Sub


    Private Sub proc_gestion()

        Dim AL As New Collection

        Dim num As Integer = lister_cellule.Count

        For value As Integer = 0 To 4

            For value1 As Integer = 0 To num - 1
                AL = lister_cellule(value1)
                Dim array1() As String = lister_arrt(value1)
                Dim array2() As String = lister_arra(value1)

                proc_tableau(AL, array1, array2, value)
            Next

        Next
        proc_marge_credit_dispo_a_la_fin()
        ElementPlusMax()
        proc_calculer_somme_encaisse_mc()



        'afficher_depense_et_revenu()


        If processeur = True Then
            frmLoading.Close()
            proc_negative()
            processeur = False
            Me.Show()
        Else
            frmLoading.Close()
            proc_negative()

        End If


    End Sub


    Private Sub proc_tableau(ByVal tab1 As Collection, ByVal tab2() As String, ByVal tab3() As String, ByVal nbr As Integer)

        Dim chaine1 As String = tab2(nbr)
        Dim chaine2 As String = tab3(nbr)
        If (chaine1 = "textbox") Then
            proc_resultat_textbox(tab1, nbr, chaine2)
        Else
            proc_resultat_label(tab1, nbr, chaine2)
        End If



    End Sub

    Private Sub proc_resultat_textbox(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)

        Dim periode As String = ""
        Dim no_element As Integer
        Dim statement As String = ""

        If (chaine <> "check_direct" And chaine <> "check_preautorises" And chaine <> "check_emis" And chaine <> "check_revenu" And chaine <> "check_depenses") Then

            Select Case numero
                Case 0
                    'Debug.WriteLine("Between 1 and 5, inclusive")
                    periode = "courant"
                Case 1
                    periode = "sept"
                Case 2
                    periode = "quatorze"
                Case 3
                    periode = "vingt"
                Case 4
                    periode = "cinq"
            End Select
            statement = "Select " + periode + " from " + chaine + " where numero = 1"
        Else
            Select Case numero
                Case 0
                    'Debug.WriteLine("Between 1 and 5, inclusive")
                    periode = "sum(un)"
                Case 1
                    periode = "sum(deux)"
                Case 2
                    periode = "sum(trois)"
                Case 3
                    periode = "sum(quatre)"
                Case 4
                    periode = "sum(cinq)"
            End Select
            statement = "Select " + periode + " from " + chaine
        End If

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand(statement, myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            no_element = numero + 1

            While myReader1.Read
                tab1(no_element).Text = myReader1.Item(0).ToString

            End While

            tab1(no_element).Text = conventir_format(tab1(no_element).Text)
            myConnection.Close()


            myCommand1.Dispose()
            myReader1.Close()
        Catch ex As Exception

        End Try


    End Sub


    Private Sub proc_resultat_label(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)

        Dim myChar As Char
        myChar = chaine.Chars(0) ' myChar = "D"

        chaine = chaine.Trim({"$"c, "+"c, "-"c, "*"c, "@"c, "#"c})

        Select Case myChar
            Case "$"
                proc_fonction_positif(tab1, numero, chaine)
            Case "+"
                proc_somme(tab1, numero, chaine)
            Case "-"
                proc_negative(tab1, numero, chaine)
            Case "*"
                proc_fonction_negative(tab1, numero, chaine)
            Case "@"
                proc_fonction_fin(tab1, numero, chaine)
            Case "#"
                Console.WriteLine("Invalid grade")
        End Select

    End Sub

    Private Sub proc_somme(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)

        Dim strarr() As String
        strarr = chaine.Split("?")

        Dim word As String
        Dim chr1() As String
        Dim ALs As New Collection
        Dim list_num As New ArrayList

        For Each word In strarr
            chr1 = word.Split("&")
            Dim nbr1 As Integer = chr1(0)
            Dim nbr2 As Integer = chr1(1)

            ALs = lister_cellule(nbr1)

            'If numero < 1 Then
            '    nbr2 = nbr2 + 1
            'End If

            Dim rep = ALs(nbr2).text
            rep = rep.Trim({"$"c, " "c})
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)

            list_num.Add(addition)
        Next

        Dim dblVal As Double = 0.0


        For i = 0 To list_num.Count - 1
            If dblVal = 0.0 Then
                dblVal = list_num.Item(i)
            Else
                dblVal = dblVal + list_num.Item(i)
            End If
        Next

        Dim no_element = numero + 1


        tab1(no_element).Text = conventir_format(CStr(dblVal))


    End Sub


    Private Sub proc_negative(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)


        Dim strarr() As String
        strarr = chaine.Split("?")

        Dim word As String
        Dim chr1() As String
        Dim ALs As New Collection
        Dim list_num As New ArrayList

        For Each word In strarr
            chr1 = word.Split("&")
            Dim nbr1 As Integer = chr1(0)
            Dim nbr2 As Integer = chr1(1)

            'ALs = lister_cellule(nbr1)
            'Dim addition As Double = Convert.ToDouble(ALs(nbr2 + 1).text)
            'list_num.Add(addition)
            ALs = lister_cellule(nbr1)
            'If numero < 1 Then
            '    nbr2 = nbr2 + 1
            'End If

            Dim rep = ALs(nbr2).text
            rep = rep.Trim({"$"c, " "c})
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)
            list_num.Add(addition)
        Next

        Dim dblVal As Double

        dblVal = list_num(0) - list_num(1)

        Dim no_element = numero + 1


        tab1(no_element).Text = conventir_format(CStr(dblVal))

    End Sub

    Private Sub proc_fonction_positif(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)



        Dim strarr() As String
        strarr = chaine.Split("?")

        Dim word As String
        Dim chr1() As String
        Dim ALs As New Collection
        Dim list_num As New ArrayList

        For Each word In strarr
            chr1 = word.Split("&")
            Dim nbr1 As Integer = chr1(0)
            Dim nbr2 As Integer = chr1(1)

            'ALs = lister_cellule(nbr1)
            'Dim addition As Double = Convert.ToDouble(ALs(nbr2).text)
            'list_num.Add(addition)
            ALs = lister_cellule(nbr1)

            'If numero < 1 Then
            '    nbr2 = nbr2 + 1
            'End If

            Dim rep = ALs(nbr2).text
            rep = rep.Trim({"$"c, " "c})
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)
            list_num.Add(addition)
        Next

        Dim dblVal As Double

        'dblVal = list_num(0) - list_num(1)

        If list_num(0) > list_num(1) Then
            dblVal = list_num(0) + list_num(1)
        Else
            dblVal = 0
        End If

        Dim no_element = numero + 1

        tab1(no_element).Text = conventir_format(CStr(dblVal))


    End Sub


    Private Sub proc_fonction_negative(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)

        Dim strarr() As String
        strarr = chaine.Split("?")

        Dim word As String
        Dim chr1() As String
        Dim ALs As New Collection
        Dim list_num As New ArrayList

        For Each word In strarr
            chr1 = word.Split("&")
            Dim nbr1 As Integer = chr1(0)
            Dim nbr2 As Integer = chr1(1)

            'ALs = lister_cellule(nbr1)
            'Dim addition As Double = Convert.ToDouble(ALs(nbr2 + 1).text)
            'list_num.Add(addition)
            ALs = lister_cellule(nbr1)
            'If numero < 1 Then
            '    nbr2 = nbr2 + 1
            'End If

            Dim rep = ALs(nbr2).text
            rep = rep.Trim({"$"c, " "c})
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)
            list_num.Add(addition)
        Next

        Dim dblVal As Double

        If list_num(0) < list_num(1) Then
            dblVal = list_num(1) - list_num(0)
        Else
            dblVal = 0
        End If

        Dim no_element = numero + 1

        tab1(no_element).Text = conventir_format(CStr(dblVal))

    End Sub

    Private Sub proc_fonction_fin(ByVal tab1 As Collection, ByVal numero As Integer, ByVal chaine As String)

        Dim strarr() As String
        strarr = chaine.Split("?")

        Dim word As String
        Dim chr1() As String
        Dim ALs As New Collection
        Dim list_num As New ArrayList

        For Each word In strarr
            chr1 = word.Split("&")
            Dim nbr1 As Integer = chr1(0)
            Dim nbr2 As Integer = chr1(1)

            'ALs = lister_cellule(nbr1)
            'Dim addition As Double = Convert.ToDouble(ALs(nbr2).text)
            'list_num.Add(addition)
            ALs = lister_cellule(nbr1)
            'If numero < 1 Then
            '    nbr2 = nbr2 + 1
            'End If
            Dim rep = ALs(nbr2).text
            rep = rep.Trim({"$"c, " "c})
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)
            list_num.Add(addition)
        Next

        Dim dblVal As Double

        If list_num(0) > list_num(1) Then
            dblVal = list_num(0) - list_num(1)
        Else
            dblVal = 0
        End If

        Dim no_element = numero + 1



        tab1(no_element).Text = conventir_format(CStr(dblVal))

    End Sub

    Sub proc_marge_credit_dispo_a_la_fin()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from marge_de_credit_utilise where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                TextBox5.Text = myReader1.Item(0).ToString


            End While

            TextBox5.Text = conventir_format(TextBox5.Text)
            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try


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


    Private Sub proc_calculer_somme_encaisse_mc()

        Dim element1 As Double = Label60.Text
        Dim element2 As Double = Label59.Text
        Dim element3 As Double = Label58.Text
        Dim element4 As Double = Label57.Text
        Dim element5 As Double = Label43.Text
        Dim somme As Double

        somme = element1 + element2 + element3 + element4 + element5

        Label47.Text = conventir_format(CStr(somme))

    End Sub


    Private Sub proc_restaurantion()




        'proc_tab_paiement_preautorise()
        'proc_tab_cheque_emis()

        Dim myList As New ArrayList

        myList.Add(CDbl(TextBox1.Text))
        myList.Add(CDbl(TextBox12.Text))
        myList.Add(CDbl(TextBox11.Text))
        myList.Add(CDbl(TextBox10.Text))
        myList.Add(CDbl(TextBox9.Text))
        myList.Add(CDbl(TextBox20.Text))
        'myList.Add(CDbl(TextBox18.Text))
        'myList.Add(CDbl(TextBox17.Text))
        'myList.Add(CDbl(TextBox16.Text))
        'myList.Add(CDbl(TextBox15.Text))
        'myList.Add(CDbl(TextBox21.Text))
        'myList.Add(CDbl(TextBox72.Text))
        'myList.Add(CDbl(TextBox71.Text))
        'myList.Add(CDbl(TextBox70.Text))
        'myList.Add(CDbl(TextBox69.Text))
        'myList.Add(CDbl(TextBox22.Text))
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


    Private Sub TextBox12_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.Enter, TextBox11.Enter,
        TextBox10.Enter, TextBox9.Enter, TextBox20.Enter, TextBox1.Enter,
        TextBox14.Enter, TextBox13.Enter, TextBox8.Enter, TextBox7.Enter, TextBox23.Enter, TextBox19.Enter, TextBox6.Enter,
        TextBox4.Enter, TextBox3.Enter, TextBox2.Enter, TextBox24.Enter, TextBox5.Enter


        Dim tb As TextBox = CType(sender, TextBox)

        tb = CType(sender, TextBox)

        num_temporaire = tb.Text





    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown


        Dim tb As TextBox = New TextBox
        'Dim categorie As String

        tb = CType(sender, TextBox)


        If num_temporaire <> "" Then



        End If

        'If rep = "Textbox1" Then
        '    categorie = "Solde_encaisses_debut"
        'End If



        If e.KeyCode = Keys.Enter Then
            num_temporaire = String.Empty
            TextBox12.Focus()
            frmLoading.Show()


            Me.Hide()
            processeur = True


            periode_paiement_direct()
            periode_paiement_preautorise()
            periode_cheque_emis()

            Try

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
                proc_gestion()


                myCommand.Dispose()

                oledbAdapter.Dispose()
            Catch ex As Exception

            End Try

        End If

    End Sub


    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown, TextBox11.KeyDown, TextBox10.KeyDown, TextBox9.KeyDown, TextBox20.KeyDown



        Dim tb As TextBox = New TextBox
        Dim categorie As String = ""
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
            'Dim SecondForm As New frmLoading


            nbr_negative = 0
            frmLoading.Show()

            Me.Hide()
            processeur = True

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
                    TextBox20.Focus()


            End Select

            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                Dim oledbAdapter As New OleDb.OleDbDataAdapter
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("update Apports set " & categorie & " = @description WHERE numero = @numero", myConnection)




                Select Case tb.Name
                    Case "TextBox12"
                        myCommand.Parameters.AddWithValue("@description", TextBox12.Text)
                        num_temporaire = tb.Text
                        TextBox12.Text = tb.Text

                    Case "TextBox11"
                        myCommand.Parameters.AddWithValue("@description", TextBox11.Text)
                        num_temporaire = tb.Text
                        TextBox11.Text = tb.Text
                    Case "TextBox10"
                        myCommand.Parameters.AddWithValue("@description", TextBox10.Text)
                        num_temporaire = tb.Text
                        TextBox10.Text = tb.Text
                    Case "TextBox9"
                        myCommand.Parameters.AddWithValue("@description", TextBox9.Text)
                        num_temporaire = tb.Text
                        TextBox9.Text = tb.Text
                    Case "TextBox20"
                        myCommand.Parameters.AddWithValue("@description", TextBox20.Text)
                        num_temporaire = tb.Text
                        TextBox20.Text = tb.Text
                End Select


                'myCommand.Parameters.AddWithValue("@description", TextBox12.Text)

                myCommand.Parameters.AddWithValue("@numero", 1)

                myCommand.ExecuteNonQuery()
                myConnection.Close()
                proc_gestion()


                myConnection.Close()

                myCommand.Dispose()
                oledbAdapter.Dispose()

            Catch ex As Exception

            End Try

        End If

    End Sub




    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress, TextBox11.KeyPress,
        TextBox10.KeyPress, TextBox9.KeyPress, TextBox20.KeyPress, TextBox1.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress,
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

        enregistrer = True

    End Sub

    Private Sub TextBox12_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.Leave, TextBox11.Leave,
        TextBox10.Leave, TextBox9.Leave, TextBox20.Leave, TextBox1.Leave, TextBox12.Leave, TextBox11.Leave, TextBox10.Leave,
        TextBox9.Leave, TextBox20.Leave, TextBox14.Leave, TextBox13.Leave,
        TextBox8.Leave, TextBox7.Leave, TextBox23.Leave, TextBox19.Leave, TextBox6.Leave, TextBox4.Leave, TextBox3.Leave, TextBox2.Leave,
        TextBox24.Leave, TextBox5.Leave




        Dim tb As TextBox = New TextBox

        tb = CType(sender, TextBox)

        If num_temporaire <> "" Then


            tb.Text = num_temporaire




        End If

    End Sub


    'Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs)

    '    Dim tb As TextBox = New TextBox
    '    Dim categorie As String = ""

    '    tb = CType(sender, TextBox)

    '    If num_temporaire <> "" Then

    '        Select Case tb.Name
    '            Case "TextBox18"
    '                categorie = "courant"
    '            Case "TextBox17"
    '                categorie = "sept"
    '            Case "TextBox16"
    '                categorie = "quatorze"
    '            Case "TextBox15"
    '                categorie = "vingt"
    '            Case "TextBox21"
    '                categorie = "cinq"
    '        End Select

    '    End If



    '    If e.KeyCode = Keys.Enter Then
    '        num_temporaire = String.Empty
    '        frmLoading.Show()
    '        Me.Hide()
    '        processeur = True

    '        Select Case tb.Name
    '            Case "TextBox18"
    '                TextBox17.Focus()
    '            Case "TextBox17"
    '                TextBox16.Focus()
    '            Case "TextBox16"
    '                TextBox15.Focus()
    '            Case "TextBox15"
    '                TextBox21.Focus()
    '            Case "TextBox21"
    '                TextBox72.Focus()


    '        End Select

    '        Try
    '            Dim myConnection As New OleDb.OleDbConnection(myConnString)
    '            Dim oledbAdapter As New OleDb.OleDbDataAdapter
    '            myConnection.Open()

    '            Dim myCommand As New OleDb.OleDbCommand("update Retrait_manuel set " & categorie & "= @description WHERE numero = @numero", myConnection)
    '            myCommand.Parameters.AddWithValue("@description", tb.Text)
    '            myCommand.Parameters.AddWithValue("@numero", 1)

    '            myCommand.ExecuteNonQuery()

    '            myCommand.Dispose()
    '            oledbAdapter.Dispose()
    '            myConnection.Close()

    '            proc_gestion()


    '        Catch ex As Exception

    '        End Try


    '    End If
    'End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Me.Close()


        FrmInterface.Show()


    End Sub

    'Private Sub TextBox72_KeyDown(sender As Object, e As KeyEventArgs)

    '    Dim tb As TextBox = New TextBox
    '    Dim categorie As String = ""

    '    tb = CType(sender, TextBox)

    '    If num_temporaire <> "" Then

    '        Select Case tb.Name
    '            Case "TextBox72"
    '                categorie = "courant"

    '            Case "TextBox71"
    '                categorie = "sept"

    '            Case "TextBox70"
    '                categorie = "quatorze"

    '            Case "TextBox69"
    '                categorie = "vingt"

    '            Case "TextBox22"
    '                categorie = "cinq"

    '        End Select

    '    End If


    '    If e.KeyCode = Keys.Enter Then

    '        num_temporaire = String.Empty
    '        frmLoading.Show()
    '        Me.Hide()
    '        processeur = True

    '        Select Case tb.Name
    '            Case "TextBox72"
    '                TextBox71.Focus()
    '            Case "TextBox71"
    '                TextBox70.Focus()
    '            Case "TextBox70"
    '                TextBox69.Focus()
    '            Case "TextBox69"
    '                TextBox22.Focus()
    '            Case "TextBox22"
    '                TextBox14.Focus()


    '        End Select

    '        Try
    '            Dim myConnection As New OleDb.OleDbConnection(myConnString)
    '            Dim oledbAdapter As New OleDb.OleDbDataAdapter
    '            myConnection.Open()

    '            Dim myCommand As New OleDb.OleDbCommand("update Salaires_paie_courante set " & categorie & " = @description WHERE numero = @numero", myConnection)


    '            Select Case tb.Name
    '                Case "TextBox72"
    '                    myCommand.Parameters.AddWithValue("@description", TextBox72.Text)
    '                    num_temporaire = tb.Text
    '                Case "TextBox71"
    '                    myCommand.Parameters.AddWithValue("@description", TextBox71.Text)
    '                    num_temporaire = tb.Text
    '                Case "TextBox70"
    '                    myCommand.Parameters.AddWithValue("@description", TextBox70.Text)
    '                    num_temporaire = tb.Text
    '                Case "TextBox69"
    '                    myCommand.Parameters.AddWithValue("@description", TextBox69.Text)
    '                    num_temporaire = tb.Text
    '                Case "TextBox22"
    '                    myCommand.Parameters.AddWithValue("@description", TextBox22.Text)
    '                    num_temporaire = tb.Text

    '            End Select



    '            myCommand.Parameters.AddWithValue("@numero", 1)

    '            myCommand.ExecuteNonQuery()


    '            myCommand.Dispose()
    '            oledbAdapter.Dispose()
    '            myConnection.Close()

    '            proc_gestion()

    '        Catch ex As Exception

    '        End Try


    '    End If

    'End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown, TextBox13.KeyDown, TextBox8.KeyDown, TextBox7.KeyDown, TextBox23.KeyDown

        Dim tb As TextBox = New TextBox
        Dim categorie As String = ""

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

        Try
            If e.KeyCode = Keys.Enter Then
                nbr_negative = 0
                frmLoading.Show()

                Me.Hide()
                processeur = True


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
                proc_gestion()
                myCommand.Dispose()
                oledbAdapter.Dispose()






            End If



        Catch ex As Exception

        End Try



    End Sub





    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then
                nbr_negative = 0
                frmLoading.Show()

                Me.Hide()
                processeur = True

                num_temporaire = String.Empty
                TextBox1.Focus()

                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                Dim oledbAdapter As New OleDb.OleDbDataAdapter
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("update marge_de_credit_utilise set courant = @description WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@description", TextBox5.Text)
                myCommand.Parameters.AddWithValue("@numero", 1)

                myCommand.ExecuteNonQuery()
                myCommand.Dispose()
                oledbAdapter.Dispose()
                myConnection.Close()
                proc_gestion()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown, TextBox4.KeyDown, TextBox3.KeyDown, TextBox2.KeyDown, TextBox24.KeyDown

        Dim tb As TextBox = New TextBox
        Dim categorie As String = ""

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

        Try

            If e.KeyCode = Keys.Enter Then

                num_temporaire = String.Empty
                nbr_negative = 0
                frmLoading.Show()

                Me.Hide()
                processeur = True

                Select Case tb.Name
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
                myCommand.Dispose()
                oledbAdapter.Dispose()
                myConnection.Close()
                proc_gestion()

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then
                num_temporaire = String.Empty
                TextBox6.Focus()
                nbr_negative = 0
                frmLoading.Show()

                Me.Hide()
                processeur = True


                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                Dim oledbAdapter As New OleDb.OleDbDataAdapter
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("update marge_credit_disponible set courant = @description WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@description", TextBox19.Text)
                myCommand.Parameters.AddWithValue("@numero", 1)

                myCommand.ExecuteNonQuery()

                myCommand.Dispose()
                oledbAdapter.Dispose()

                myConnection.Close()

                proc_gestion()


            End If


        Catch ex As Exception

        End Try

    End Sub

    Sub periode_paiement_direct()

        Try
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

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try


    End Sub

    Sub periode_cheque_emis()

        Try
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
                tab(0) = myReader1.Item(6)
                tab(1) = myReader1.Item(7)
                tab(2) = myReader1.Item(8)
                tab(3) = myReader1.Item(9)
                tab(4) = myReader1.Item(10)

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

            myCommand1.Dispose()
            myReader1.Close()
        Catch ex As Exception

        End Try


    End Sub

    Sub periode_paiement_preautorise()

        Try
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

            myCommand1.Dispose()
            myReader1.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub fctable_paiement_direct(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)

        Try

            Dim reponse As String = ""
            Dim tab As Double() = {0, 0, 0, 0, 0}
            Dim cpt As Integer = 0

            Dim dtnow As Date = DateCourant


            Dim currentDate1 As Date

            Dim currentDate As Date = DateCourant



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

            myCommand.Dispose()
            oledbAdapter.Dispose()

        Catch ex As Exception

        End Try


    End Sub


    Private Sub fctable_cheque_emis(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)


        Dim reponse As String = ""
        Dim tab As Double() = {0, 0, 0, 0, 0}
        Dim cpt As Integer = 0
        'Dim currentDate As DateTime
        'currentDate = Now.Date


        ' Tester avant la date de paiement

        'Dim dtnow As Date = Date.Now
        Dim dtnow As Date = DateCourant


        'If dtnow < date_ Then

        '    tab(0) = montant
        'End If

        Dim currentDate1 As Date

        'currentDate = currentDate.AddDays(1)

        Dim currentDate As Date = DateCourant
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


        Try
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


            myCommand.Dispose()
            oledbAdapter.Dispose()

        Catch ex As Exception

        End Try
        'MsgBox("Test débogger")








        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero ='" & numero & "'", myConnection)
        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero = 75", myConnection)


    End Sub


    Private Sub fctable_paiement_preautorise(ByVal montant As Double, ByVal numero As Integer, ByVal date_ As DateTime)


        Dim reponse As String = ""
        Dim tab As Double() = {0, 0, 0, 0, 0}
        Dim cpt As Integer = 0
        'Dim currentDate As DateTime
        'currentDate = Now.Date


        ' Tester avant la date de paiement

        'Dim dtnow As Date = Date.Now7
        Dim dtnow As Date = DateCourant

        'If dtnow < date_ Then

        '    tab(0) = montant
        'End If

        Dim currentDate1 As Date

        'currentDate = currentDate.AddDays(1)

        'Dim currentDate As Date =  
        Dim currentDate As Date = DateCourant




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

        Try
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

            myCommand.Dispose()
            oledbAdapter.Dispose()
        Catch ex As Exception

        End Try

        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero ='" & numero & "'", myConnection)
        'Dim myCommand2 As New OleDb.OleDbCommand("Select * from check_direct where numero = 75", myConnection)


    End Sub


    Private Sub TextBox12_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox12.MouseClick, TextBox1.MouseClick, TextBox11.MouseClick, TextBox10.MouseClick, TextBox9.MouseClick, TextBox20.MouseClick, TextBox14.MouseClick, TextBox13.MouseClick, TextBox8.MouseClick,
         TextBox7.MouseClick, TextBox23.MouseClick, TextBox19.MouseClick, TextBox6.MouseClick, TextBox4.MouseClick, TextBox3.MouseClick,
         TextBox2.MouseClick, TextBox24.MouseClick, TextBox5.MouseClick

        Dim tb As TextBox = New TextBox
        'Dim categorie As String

        tb = CType(sender, TextBox)

        tb.SelectAll()





        'TextBox12.SelectAll()

    End Sub

    Sub recurer_backup()

        Try
            Dim num_tab As String = ""

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

            placer_espace(strArr) ' septieme


            myConnection.Close()

            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try



    End Sub

    Sub placer_espace(ByRef tab() As String)


        Initialisation()

        Dim cpt As Integer = 0

        For value As Integer = 1 To 5

            For value1 As Integer = 0 To 15
                AL = lister_cellule(value1)
                AL(value).Text = tab(cpt)
                cpt = cpt + 1

            Next

        Next

        TextBox5.Text = tab(80)
        Label39.Text = tab(81)
        Label47.Text = tab(82)



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
            'TextBox18.Enabled = False
            'TextBox17.Enabled = False
            'TextBox16.Enabled = False
            'TextBox15.Enabled = False
            'TextBox21.Enabled = False

            'TextBox72.Enabled = False
            'TextBox71.Enabled = False
            'TextBox70.Enabled = False
            'TextBox69.Enabled = False
            'TextBox22.Enabled = False
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
            'TextBox18.Enabled = True
            'TextBox17.Enabled = True
            'TextBox16.Enabled = True
            'TextBox15.Enabled = True
            'TextBox21.Enabled = True

            'TextBox72.Enabled = True
            'TextBox71.Enabled = True
            'TextBox70.Enabled = True
            'TextBox69.Enabled = True
            'TextBox22.Enabled = True
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

        If nbr = "" Then
            nbr = "0"
        End If

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


    Private Sub show_tooltip_on_textbox()


        ToolTip2.SetToolTip(Me.TextBox1, "Inscrire le solde du début d'encaisse")
        'ToolTip2.SetToolTip(Me.TextBox18, "Inscrire tous les montants à recevoir pendant la période actuel")
        ToolTip2.SetToolTip(Me.TextBox19, "Inscrire la marge de crédit de début, Truc : si vous n'avez pas de luidités de début et seulement des déboursés, voir le maximum de la liquidité immédiate au cours des 4 prochaines semaines.")
        ToolTip2.SetToolTip(Me.TextBox6, "Par mesure de sécurité, il serait préférable de mettre un montant de sécurité en début d'année")
        ToolTip2.SetToolTip(Me.TextBox5, "Inscrire la marge de crédit négocié")



    End Sub

    Private Sub proc_negative()
        If nbr_negative = 1 Then

            'dater_negative.ToString
            'negative

            Dim abs As Double = Math.Abs(negative)

            MessageBox.Show("Votre solde est négative en " + dater_negative.ToString("MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA")) + " pour " + abs.ToString("#,##0.00 $") + ".",
                            "Aversitement")
            nbr_negative = 0

        ElseIf nbr_negative > 1 Then
            MessageBox.Show("Vous allez voir votre budget de caisse car plusieurs soldes négatives",
                            "Aversitement")
            nbr_negative = 0
        End If
    End Sub



End Class
