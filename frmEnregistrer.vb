Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb

Public Class frmEnregistrer

    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Dim myConnection As New OleDb.OleDbConnection(myConnString)
    Dim con As System.Data.OleDb.OleDbConnection
    Dim sqlStr As String
    Dim cmd As System.Data.OleDb.OleDbCommand

    Dim AL As New Collection
    Dim lister_cellule As New ArrayList
    Dim tab_type() As String
    Dim lister_arrt As New ArrayList
    Dim lister_arra As New ArrayList


    Dim message As String = ""


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

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Dim arra_un(), arra_deux(), arra_trois(), arra_quatre(),
        arra_cinq(), arra_six(), arra_sept(), arra_huit(), arra_neuf(),
        arra_dix(), arra_onze(), arra_douze(), arra_treize(),
        arra_quatieme(), arra_cinquieme(), arra_sixieme(), arra_septieme(),
        arra_huitieme() As String


    Private Sub Initialisation()

        Al_un.Add(frmGestion.TextBox1)
        Al_un.Add(frmGestion.Label14)
        Al_un.Add(frmGestion.Label15)
        Al_un.Add(frmGestion.Label16)
        Al_un.Add(frmGestion.Label35)
        Al_deux.Add(frmGestion.TextBox12)
        Al_deux.Add(frmGestion.TextBox11)
        Al_deux.Add(frmGestion.TextBox10)
        Al_deux.Add(frmGestion.TextBox9)
        Al_deux.Add(frmGestion.TextBox20)
        Al_trois.Add(frmGestion.Label82)
        Al_trois.Add(frmGestion.Label83)
        Al_trois.Add(frmGestion.Label84)
        Al_trois.Add(frmGestion.Label85)
        Al_trois.Add(frmGestion.Label86)
        Al_quatre.Add(frmGestion.Label80)
        Al_quatre.Add(frmGestion.Label79)
        Al_quatre.Add(frmGestion.Label78)
        Al_quatre.Add(frmGestion.Label77)
        Al_quatre.Add(frmGestion.Label76)
        Al_cinq.Add(frmGestion.Label17)
        Al_cinq.Add(frmGestion.Label19)
        Al_cinq.Add(frmGestion.Label20)
        Al_cinq.Add(frmGestion.Label21)
        Al_cinq.Add(frmGestion.Label36)
        Al_six.Add(frmGestion.Label25)
        Al_six.Add(frmGestion.Label24)
        Al_six.Add(frmGestion.Label23)
        Al_six.Add(frmGestion.Label22)
        Al_six.Add(frmGestion.Label37)
        Al_sept.Add(frmGestion.Label29)
        Al_sept.Add(frmGestion.Label28)
        Al_sept.Add(frmGestion.Label27)
        Al_sept.Add(frmGestion.Label26)
        Al_sept.Add(frmGestion.Label41)
        Al_huit.Add(frmGestion.Label87)
        Al_huit.Add(frmGestion.Label88)
        Al_huit.Add(frmGestion.Label89)
        Al_huit.Add(frmGestion.Label90)
        Al_huit.Add(frmGestion.Label91)
        Al_neuf.Add(frmGestion.TextBox14)
        Al_neuf.Add(frmGestion.TextBox13)
        Al_neuf.Add(frmGestion.TextBox8)
        Al_neuf.Add(frmGestion.TextBox7)
        Al_neuf.Add(frmGestion.TextBox23)
        Al_dix.Add(frmGestion.Label68)
        Al_dix.Add(frmGestion.Label67)
        Al_dix.Add(frmGestion.Label66)
        Al_dix.Add(frmGestion.Label65)
        Al_dix.Add(frmGestion.Label42)
        Al_onze.Add(frmGestion.Label60)
        Al_onze.Add(frmGestion.Label59)
        Al_onze.Add(frmGestion.Label58)
        Al_onze.Add(frmGestion.Label57)
        Al_onze.Add(frmGestion.Label43)
        'Al_onze.Add(Label47)
        Al_douze.Add(frmGestion.TextBox19)
        Al_douze.Add(frmGestion.Label55)
        Al_douze.Add(frmGestion.Label54)
        Al_douze.Add(frmGestion.Label53)
        Al_douze.Add(frmGestion.Label44)
        Al_trieze.Add(frmGestion.Label52)
        Al_trieze.Add(frmGestion.Label51)
        Al_trieze.Add(frmGestion.Label50)
        Al_trieze.Add(frmGestion.Label49)
        Al_trieze.Add(frmGestion.Label45)
        Al_quatorze.Add(frmGestion.TextBox6)
        Al_quatorze.Add(frmGestion.TextBox4)
        Al_quatorze.Add(frmGestion.TextBox3)
        Al_quatorze.Add(frmGestion.TextBox2)
        Al_quatorze.Add(frmGestion.TextBox24)
        Al_cinquieme.Add(frmGestion.Label32)
        Al_cinquieme.Add(frmGestion.Label31)
        Al_cinquieme.Add(frmGestion.Label30)
        Al_cinquieme.Add(frmGestion.Label13)
        Al_cinquieme.Add(frmGestion.Label46)
        Al_sixieme.Add(frmGestion.Label70)
        Al_sixieme.Add(frmGestion.Label71)
        Al_sixieme.Add(frmGestion.Label72)
        Al_sixieme.Add(frmGestion.Label73)
        Al_sixieme.Add(frmGestion.Label74)
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

        arrt_un = {"textbox", "label", "label", "label", "label"}
        arrt_deux = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_trois = {"textbox", "textbox", "textbox", "textbox", "textbox"}
        arrt_quatre = {"label", "label", "label", "label", "label"}
        arrt_cinq = {"textbox", "textbox", "textbox", "textbox", "textbox"}
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
        arra_trois = {"Retrait_manuel", "Retrait_manuel", "Retrait_manuel", "Retrait_manuel", "Retrait_manuel"}


        'arra_quatre = {"+1&1?2&1?3&1", "+1&2?2&2?3&2", "+1&3?2&3?3&3", "+1&4?2&4?3&4", "+1&5?2&5?3&5"}
        'arra_quatre = {"+0&0?1&0?2&0", "+0&1?1&1?2&1", "+0&2?1&2?2&2", "+0&3?1&3?2&3", "+0&4?1&4?2&4"}
        arra_quatre = {"+0&1?1&1?2&1", "+0&2?1&2?2&2", "+0&3?1&3?2&3", "+0&4?1&4?2&4", "+0&5?1&5?2&5"}


        arra_cinq = {"check_direct", "check_direct", "check_direct", "check_direct", "check_direct"}
        arra_six = {"check_preautorises", "check_preautorises", "check_preautorises", "check_preautorises", "check_preautorises"}
        arra_sept = {"check_emis", "check_emis", "check_emis", "check_emis", "check_emis"}
        arra_huit = {"Salaires_paie_courante", "Salaires_paie_courante", "Salaires_paie_courante", "Salaires_paie_courante", "Salaires_paie_courante"}
        arra_neuf = {"Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir", "Argent_a_recevoir"}
        'arra_dix = {"+5&1?6&1?7&1?8&1?9&1", "+5&2?6&2?7&2?8&2?9&2", "+5&3?6&3?7&3?8&3?9&3", "+5&4?6&4?7&4?8&4?9&4", "+5&5?6&5?7&5?8&5?9&5"}
        'arra_dix = {"+4&0?5&0?6&0?7&0?8&0", "+4&1?5&1?6&1?7&1?8&1", "+4&2?5&2?6&2?7&2?8&2", "+4&3?5&3?6&3?7&3?8&3", "+4&4?5&4?6&4?7&4?8&4"}
        arra_dix = {"+4&1?5&1?6&1?7&1?8&1", "+4&2?5&2?6&2?7&2?8&2", "+4&3?5&3?6&3?7&3?8&3", "+4&4?5&4?6&4?7&4?8&4", "+4&5?5&5?6&5?7&5?8&5"}
        'arra_onze = {"-4&1?10&1", "-4&2?10&2", "-4&3?10&3", "-4&4?10&4", "-4&5?10&5", "+10&1?10&2?10&3?10&4?10&5"}
        'arra_onze = {"-4&1?10&1", "-4&2?10&2", "-4&3?10&3", "-4&4?10&4", "-4&5?10&5"}
        arra_onze = {"-3&1?9&1", "-3&2?9&2", "-3&3?9&3", "-3&4?9&4", "-3&5?9&5"}

        arra_douze = {"marge_credit_disponible", "-11&1?14&1", "-11&2?14&2", "-11&3?14&3", "-11&4?14&4"}
        'arra_treize = {"+11&1?12&1", "+11&2?12&2", "+11&3?12&3", "+11&4?12&4", "+11&4?12&4"}
        arra_treize = {"+10&1?11&1", "+10&2?11&2", "+10&3?11&3", "+10&4?11&4", "+10&5?11&5"}
        arra_quatieme = {"Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise", "Encaisse_minimum_requise"}
        'arra_cinquieme = {"*11&1?14&1", "*11&2?14&2", "*11&3?14&3", "*11&4?14&4", "*11&5?14&5"}
        arra_cinquieme = {"*10&1?13&1", "*10&2?13&2", "*10&3?13&3", "*10&4?13&4", "*10&5?13&5"}
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

        'If processeur = True Then
        '    frmLoading.Close()
        '    processeur = False
        '    Me.Show()
        'Else
        '    frmLoading.Close()

        'End If

        frmLoading.Close()

        Me.Show()

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

        If (chaine <> "check_direct" And chaine <> "check_preautorises" And chaine <> "check_emis") Then

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

            message = message + tab1(no_element).Text + "/"

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
            'Test
            If rep = "" Then
                rep = 0
            End If
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

        message = message + tab1(no_element).Text + "/"

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
            If rep = "" Then
                rep = 0
            End If
            'Dim addition As Double = Convert.ToDouble(rep)
            Dim addition As Double = CDbl(rep)
            list_num.Add(addition)
        Next

        Dim dblVal As Double

        dblVal = list_num(0) - list_num(1)

        Dim no_element = numero + 1

        tab1(no_element).Text = conventir_format(CStr(dblVal))

        message = message + tab1(no_element).Text + "/"

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
            If rep = "" Then
                rep = 0
            End If
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

        message = message + tab1(no_element).Text + "/"

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
            If rep = "" Then
                rep = 0
            End If
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

        message = message + tab1(no_element).Text + "/"

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
            If rep = "" Then
                rep = 0
            End If
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

        message = message + tab1(no_element).Text + "/"

    End Sub

    Sub proc_marge_credit_dispo_a_la_fin()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from marge_de_credit_utilise where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                frmGestion.TextBox5.Text = myReader1.Item(0).ToString


            End While

            frmGestion.TextBox5.Text = conventir_format(frmGestion.TextBox5.Text)

            message = message + frmGestion.TextBox5.Text + "/"


            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try


    End Sub

    Private Sub ElementPlusMax()



        Dim element1 As Double = frmGestion.Label32.Text
        Dim element2 As Double = frmGestion.Label31.Text
        Dim element3 As Double = frmGestion.Label30.Text
        Dim element4 As Double = frmGestion.Label13.Text
        Dim element5 As Double = frmGestion.Label46.Text
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

        frmGestion.Label39.Text = conventir_format(CStr(reponse))

        message = message + frmGestion.Label39.Text + "/"

        'Label39.Text = Format$(reponse, "Currency")



    End Sub


    Private Sub proc_calculer_somme_encaisse_mc()

        Dim element1 As Double = frmGestion.Label60.Text
        Dim element2 As Double = frmGestion.Label59.Text
        Dim element3 As Double = frmGestion.Label58.Text
        Dim element4 As Double = frmGestion.Label57.Text
        Dim element5 As Double = frmGestion.Label43.Text
        Dim somme As Double

        somme = element1 + element2 + element3 + element4 + element5

        frmGestion.Label47.Text = conventir_format(CStr(somme))

        message = message + frmGestion.Label47.Text

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





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Me.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub frmEnregistrer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Label1.Text = Date.Now
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

        Label1.Text = courant + "   " + Format(Now(), "HH:mm:ss")

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim tab(76) As String


        message = ""

        frmLoading.Show()
        Me.Hide()

        Initialisation()
        Initialiser_type()
        Initialiser_afficher()
        proc_gestion()

        'For Each element As String In tab

        '    message = message + element + "/"

        '    'Console.Write(element)
        '    'Console.Write("... ")
        'Next

        Dim cnn As New OleDb.OleDbConnection

        Try
            cnn = New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            Dim command As String
            'command = "INSERT INTO listerdate(description, datenr, compression) VALUES (@description, @datenr, @compression)"
            command = "insert into bidon(nom, datenr, description) VALUES(@nom, @datenr, @description)"
            cnn.Open()
            Dim cmd As OleDbCommand
            cmd = New OleDbCommand(command, cnn)
            cmd.Parameters.AddWithValue("@nom", TextBox1.Text)
            cmd.Parameters.AddWithValue("@datenr", DateCourant)
            cmd.Parameters.AddWithValue("@description", message)
            cmd.ExecuteNonQuery()
            cmd.Dispose()

            cnn.Close()
            MessageBox.Show("Le fichier a été enregistré")
            enregistrer = False
            Me.Close()
            Me.WindowState = FormWindowState.Maximized
            'frmMenu.MdiParent = Form_windows
            'frmMenu.Show()
        Catch exceptionObject As Exception
            MessageBox.Show(exceptionObject.Message)
        Finally

        End Try

    End Sub

    Private Sub frmEnregistrer_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

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
End Class