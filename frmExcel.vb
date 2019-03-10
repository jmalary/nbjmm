Imports Comptable_NBJMM.GlobalVariables
Imports Comptable_NBJMM.GlobalLister

Imports System.Threading
Imports System.Globalization

Imports System.Data.OleDb
Imports System.Collections 'For arraylist

Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
#Disable Warning BC40056 ' L'espace de noms ou le type spécifié dans les Imports 'System.Data.SQLite' ne contient aucun membre public ou est introuvable. Vérifiez que l'espace de noms ou le type est défini et qu'il contient au moins un membre public. Vérifiez que le nom de l'élément importé n'utilise pas d'autres alias.
Imports System.Data.SQLite
#Enable Warning BC40056 ' L'espace de noms ou le type spécifié dans les Imports 'System.Data.SQLite' ne contient aucun membre public ou est introuvable. Vérifiez que l'espace de noms ou le type est défini et qu'il contient au moins un membre public. Vérifiez que le nom de l'élément importé n'utilise pas d'autres alias.

Imports iTextSharp.text.html
Imports iTextSharp.text.html.simpleparser
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient


Public Class frmExcel

    Dim myConnString As String = GlobalVariables.test2ConnectionString
    Dim nom_ficbier As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub frmExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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


        testpdf()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click





        SaveFileDialog1.FileName = ""

        Dim Path As String
        Path = SaveFileDialog1.FileName


        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            'Appel d un fonction
            nom_ficbier = SaveFileDialog1.FileName
            Dim pdfTable As New PdfPTable(DataGridView1.ColumnCount)

            pdfTable.DefaultCell.Padding = 1
            pdfTable.WidthPercentage = 80
            pdfTable.HorizontalAlignment = Element.ALIGN_CENTER
            pdfTable.DefaultCell.BorderWidth = 1



            Dim mesure As Boolean = True

            'Adding Header row
            Dim i As Integer = 0

            For Each column As DataGridViewColumn In DataGridView1.Columns


                Dim cell As New PdfPCell(New Phrase(column.HeaderText))
                cell.BackgroundColor = New iTextSharp.text.Color(240, 240, 240)


                pdfTable.AddCell(cell)
                cell.Colspan = 2

            Next



            'Adding DataRow
            For Each row As DataGridViewRow In DataGridView1.Rows
                For Each cell As DataGridViewCell In row.Cells

                    If cell.Value Is Nothing Then


                    Else

                        pdfTable.AddCell(cell.Value.ToString())
                    End If


                Next
            Next

            pdfTable.DefaultCell.BackgroundColor = BaseColor.LIGHT_GRAY



            Using Stream As New FileStream(nom_ficbier, FileMode.Create)
                Dim pdfDoc As New Document(PageSize.A2, 10.0F, 10.0F, 10.0F, 0.0F)
                PdfWriter.GetInstance(pdfDoc, Stream)




                pdfDoc.Open()
                'Paragraph Paragraph = New Paragraph("data Exported From DataGridview!")

                Dim para = New Paragraph("Budget de caisse 2018")
                para.SpacingBefore = 20
                para.SpacingAfter = 20
                para.Alignment = 1 '0-Left, 1 middle,2 Right
                pdfDoc.Add(para)
                pdfDoc.Add(pdfTable)


                pdfDoc.Close()
                Stream.Close()
            End Using

            MsgBox("Le fichier PDF a été crée")
        End If

        Me.Close()


    End Sub

    Function testpdf()
        'TODO testpdf()

        Dim total_hozontal As Double

        DataGridView1.DataSource = Nothing
        Dim datelist As New List(Of Date)

        Dim dt As New DataTable()
        Try

            For i = 0 To 11
                datelist.Add(Date.Now.AddMonths(i))
            Next

            Dim espace As DataColumn
            Dim adc1 As DataColumn
            Dim adc2 As DataColumn
            Dim adc3 As DataColumn
            Dim adc4 As DataColumn
            Dim adc5 As DataColumn
            Dim adc6 As DataColumn
            Dim adc7 As DataColumn
            Dim adc8 As DataColumn
            Dim adc9 As DataColumn
            Dim adc10 As DataColumn
            Dim adc11 As DataColumn
            Dim adc12 As DataColumn
            Dim total_hozontel As DataColumn


            Dim myString As String = "Description" & vbCrLf & vbCrLf
            Dim myString1 As String = "Total" & vbCrLf & vbCrLf

            espace = New DataColumn(myString, GetType(String))
            adc1 = New DataColumn(datelist(0).ToString("MMMM yyyy"), GetType(String))
            adc2 = New DataColumn(datelist(1).ToString("MMMM yyyy"), GetType(String))
            adc3 = New DataColumn(datelist(2).ToString("MMMM yyyy"), GetType(String))
            adc4 = New DataColumn(datelist(3).ToString("MMMM yyyy"), GetType(String))
            adc5 = New DataColumn(datelist(4).ToString("MMMM yyyy"), GetType(String))
            adc6 = New DataColumn(datelist(5).ToString("MMMM yyyy"), GetType(String))
            adc7 = New DataColumn(datelist(6).ToString("MMMM yyyy"), GetType(String))
            adc8 = New DataColumn(datelist(7).ToString("MMMM yyyy"), GetType(String))
            adc9 = New DataColumn(datelist(8).ToString("MMMM yyyy"), GetType(String))
            adc10 = New DataColumn(datelist(9).ToString("MMMM yyyy"), GetType(String))
            adc11 = New DataColumn(datelist(10).ToString("MMMM yyyy"), GetType(String))
            adc12 = New DataColumn(datelist(11).ToString("MMMM yyyy"), GetType(String))
            total_hozontel = New DataColumn(myString1, GetType(String))



            espace.MaxLength = 120
            'dt.Columns.Add("Birth Date")
            'dt.Rows.Add(New Object() {"Johnny", "23", Now.AddYears(-23).ToString})

            dt.Columns.Add(espace)
            dt.Columns.Add(adc1)
            dt.Columns.Add(adc2)
            dt.Columns.Add(adc3)
            dt.Columns.Add(adc4)
            dt.Columns.Add(adc5)
            dt.Columns.Add(adc6)
            dt.Columns.Add(adc7)
            dt.Columns.Add(adc8)
            dt.Columns.Add(adc9)
            dt.Columns.Add(adc10)
            dt.Columns.Add(adc11)
            dt.Columns.Add(adc12)
            dt.Columns.Add(total_hozontel)



            dt.Rows.Add("Revenus", "", "", "", "", "", "", "", "", "", "", "", "")

            Dim total_colonne() As Double = {0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00}


            Dim lister_arr As New List(Of String)
            Dim lister_array_revenu As New List(Of String)

            Dim element As String

            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()
                Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'revenu'", myConnection)
                Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


                While myReader.Read
                    element = ""
                    Dim nom As String = Recupere_description(myReader.GetString(2))
                    Dim numero_revenu As String = myReader.GetString(1)

                    'arr_details(nom, numero_revenu, myReader.GetString(6))

                    'arrays(0) = nom
                    'arrays(1) = numero_revenu
                    'arrays(2) = myReader.GetString(6)
                    element = nom + "?" + numero_revenu + "?" + myReader.GetString(6)

                    lister_arr.Add(element)
                    'lister_array_revenu.Add(myReader.GetString(9))
                    If String.IsNullOrEmpty(myReader.GetString(9)) Then
                        lister_array_revenu.Add("vide")
                    Else
                        lister_array_revenu.Add(myReader.GetString(9))
                    End If


                End While

                myConnection.Close()
                myCommand.Dispose()
                myReader.Close()

                Dim num As String
                Dim nbr As Integer = 0

                For Each num In lister_arr

                    Dim arr_d() As String
                    arr_d = num.Split("?")

                    Dim total_montant() As Double
                    arr_d(2) = arr_d(2).Replace(",", ".")
                    Dim d_montant As Double
                    d_montant = CDbl(Val(arr_d(2)))
                    Dim reponse_update As String = lister_array_revenu.Item(nbr)
                    nbr += 1
                    Dim rep1 As String = recuperer_liste_date_check_revenu(arr_d(1))
                    Dim rep2 As String = recuperer_liste_date_revenu(arr_d(1))

                    Dim rep_total As String
                    If rep1 <> "" And rep2 <> "" Then
                        rep_total = rep1 & "?" & rep2
                    ElseIf rep1 = "" And rep2 <> "" Then
                        rep_total = rep2
                    ElseIf rep1 <> "" And rep2 = "" Then
                        rep_total = rep1
                    End If


#Disable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                    rep_total = faire_ordre_date(rep_total)
#Enable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.

                    total_montant = proc_montant(rep_total, d_montant)

                    If reponse_update <> "vide" Then
                        Dim update_montant As String = arr_d(2)
                        total_montant = proc_update_montant(reponse_update, update_montant, total_montant)
                    End If


                    total_hozontal = 0.00
                    For i As Integer = 0 To 11
                        total_hozontal += total_montant(i)
                    Next

                    dt.Rows.Add(arr_d(0), total_montant(0).ToString("####0.00 $"), total_montant(1).ToString("####0.00 $"), total_montant(2).ToString("####0.00 $"), total_montant(3).ToString("####0.00 $"), total_montant(4).ToString("####0.00 $"), total_montant(5).ToString("####0.00 $"), total_montant(6).ToString("####0.00 $"), total_montant(7).ToString("####0.00 $"), total_montant(8).ToString("####0.00 $"), total_montant(9).ToString("####0.00 $"), total_montant(10).ToString("####0.00 $"), total_montant(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))

                    total_colonne(0) = total_colonne(0) + total_montant(0)
                    total_colonne(1) = total_colonne(1) + total_montant(1)
                    total_colonne(2) = total_colonne(2) + total_montant(2)
                    total_colonne(3) = total_colonne(3) + total_montant(3)
                    total_colonne(4) = total_colonne(4) + total_montant(4)
                    total_colonne(5) = total_colonne(5) + total_montant(5)
                    total_colonne(6) = total_colonne(6) + total_montant(6)
                    total_colonne(7) = total_colonne(7) + total_montant(7)
                    total_colonne(8) = total_colonne(8) + total_montant(8)
                    total_colonne(9) = total_colonne(9) + total_montant(9)
                    total_colonne(10) = total_colonne(10) + total_montant(10)
                    total_colonne(11) = total_colonne(11) + total_montant(11)

                Next



            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try

            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")

            total_hozontal = 0.00

            For i As Integer = 0 To 11
                total_hozontal += total_colonne(i)
            Next

            dt.Rows.Add("Encaissements", total_colonne(0).ToString("####0.00 $"), total_colonne(1).ToString("####0.00 $"), total_colonne(2).ToString("####0.00 $"), total_colonne(3).ToString("####0.00 $"), total_colonne(4).ToString("####0.00 $"), total_colonne(5).ToString("####0.00 $"), total_colonne(6).ToString("####0.00 $"), total_colonne(7).ToString("####0.00 $"), total_colonne(8).ToString("####0.00 $"), total_colonne(9).ToString("####0.00 $"), total_colonne(10).ToString("####0.00 $"), total_colonne(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            dt.Rows.Add("Déboursés", "", "", "", "", "", "", "", "", "", "", "", "")

            Dim array As Integer() = New Integer(3) {}

            Dim no_ligne As Integer = compter_check()

            Dim tab_montant As Double() = New Double(no_ligne) {}
            Dim nbr_ligne As Integer
            Dim list_fourn As String() = New String(no_ligne) {}
            Dim list_colonne As Integer() = New Integer(no_ligne) {}

            tab_montant = proc_cheque(nbr_ligne, list_fourn, list_colonne)

            'Dim colonne_vide, colonne_copie As String() = New String(12) {}
            Dim colonne_vide As Double() = New Double(12) {}
            Dim colonne_copie As Double() = New Double(12) {}


            For i As Integer = 0 To 11
                colonne_vide(i) = 0.00
            Next

            For i As Integer = 0 To 11
                colonne_copie(i) = 0.00
            Next

            Dim total_colonne_bas As Double() = New Double(12) {}

            nbr_ligne -= 1


            Dim vider As Boolean = True
            Dim compter As Integer = 0
            Dim s_nom As String
            Dim s_montant As Double
            Dim no_colonne As Integer



            While vider


                If list_fourn.Length = 1 Then
                    vider = False

                Else
                    s_nom = list_fourn(compter)
                    s_montant = tab_montant(compter)
                    no_colonne = list_colonne(compter)

                    enlever_cheque(list_fourn, list_colonne, tab_montant, False, compter)


                    colonne_vide(no_colonne) += s_montant

                    Dim indx As Integer

#Disable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.
                    indx = array.IndexOf(list_fourn, s_nom)
#Enable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.

                    If indx = -1 Then
                        compter = 0
                        total_hozontal = 0.00
                        For i As Integer = 0 To 11
                            total_hozontal += colonne_vide(i)
                        Next


                        dt.Rows.Add(s_nom, colonne_vide(0).ToString("####0.00 $"), colonne_vide(1).ToString("####0.00 $"), colonne_vide(2).ToString("####0.00 $"), colonne_vide(3).ToString("####0.00 $"), colonne_vide(4).ToString("####0.00 $"), colonne_vide(5).ToString("####0.00 $"), colonne_vide(6).ToString("####0.00 $"), colonne_vide(7).ToString("####0.00 $"), colonne_vide(8).ToString("####0.00 $"), colonne_vide(9).ToString("####0.00 $"), colonne_vide(10).ToString("####0.00 $"), colonne_vide(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
                        total_colonne_bas(no_colonne) += s_montant
#Disable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.
                        array.Copy(colonne_copie, colonne_vide, colonne_vide.Length)
#Enable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.

                    Else
                        compter = indx
                        'dt.Rows.Add(s_nom, colonne_vide(0).ToString(), colonne_vide(1), colonne_vide(2), colonne_vide(3), colonne_vide(4), colonne_vide(5), colonne_vide(6), colonne_vide(7), colonne_vide(8), colonne_vide(9), colonne_vide(10), colonne_vide(11), total_hozontal)
                        total_colonne_bas(no_colonne) += s_montant
                        'array.Copy(colonne_copie, colonne_vide, colonne_vide.Length)
                    End If



                End If


            End While


            For i As Integer = 0 To 11
                colonne_vide(i) = 0.00
            Next


            no_ligne = compter_direct()

            Dim tab_montant1 As Double() = New Double(no_ligne) {}
            Dim nbr_ligne1 As Integer
            Dim list_fourn1 As String() = New String(no_ligne) {}
            Dim list_colonne1 As Integer() = New Integer(no_ligne) {}

            tab_montant1 = proc_direct(nbr_ligne1, list_fourn1, list_colonne1)
            nbr_ligne1 -= 1

            vider = True

            While vider


                If list_fourn1.Length = 1 Then
                    vider = False

                Else
                    s_nom = list_fourn1(compter)
                    s_montant = tab_montant1(compter)
                    no_colonne = list_colonne1(compter)

                    enlever_direct(list_fourn1, list_colonne1, tab_montant1, False, compter)




                    colonne_vide(no_colonne) += s_montant

                        Dim indx As Integer

#Disable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.
                        indx = array.IndexOf(list_fourn1, s_nom)
#Enable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.

                        If indx = -1 Then
                            compter = 0
                            total_hozontal = 0.00
                            For i As Integer = 0 To 11
                                total_hozontal += colonne_vide(i)
                            Next
                            dt.Rows.Add(s_nom, colonne_vide(0).ToString("####0.00 $"), colonne_vide(1).ToString("####0.00 $"), colonne_vide(2).ToString("####0.00 $"), colonne_vide(3).ToString("####0.00 $"), colonne_vide(4).ToString("####0.00 $"), colonne_vide(5).ToString("####0.00 $"), colonne_vide(6).ToString("####0.00 $"), colonne_vide(7).ToString("####0.00 $"), colonne_vide(8).ToString("####0.00 $"), colonne_vide(9).ToString("####0.00 $"), colonne_vide(10).ToString("####0.00 $"), colonne_vide(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
                            total_colonne_bas(no_colonne) += s_montant
#Disable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.
                            array.Copy(colonne_copie, colonne_vide, colonne_vide.Length)
#Enable Warning BC42025 ' Accès d'un membre partagé, d'un membre de constante, d'un membre enum ou d'un type imbriqué via une instance ; l'expression qualifiante ne sera pas évaluée.

                        Else
                            'trouver oUI
                            compter = indx
                            'test
                            'dt.Rows.Add(s_nom, colonne_vide(0).ToString(), colonne_vide(1), colonne_vide(2), colonne_vide(3), colonne_vide(4), colonne_vide(5), colonne_vide(6), colonne_vide(7), colonne_vide(8), colonne_vide(9), colonne_vide(10), colonne_vide(11), total_hozontal)
                            total_colonne_bas(no_colonne) += s_montant
                            'array.Copy(colonne_copie, colonne_vide, colonne_vide.Length)
                        End If

                    End If




            End While



            Dim lister_arr_ddd As New List(Of String)

            Dim lister_array As New List(Of String)
            Dim lister_update_array As New List(Of String)


            Dim element_ddd As String
            Dim cpt As Integer = 0


            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()
                Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'preautorise'", myConnection)
                Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


                While myReader.Read

                    element_ddd = ""

                    Dim nom As String = Recupere_description(myReader.GetString(2))
                    Dim numero_preautorise As String = myReader.GetString(1)
                    'ajouter les upadates
                    'lister_array.Add(myReader.GetString(9))
                    If String.IsNullOrEmpty(myReader.GetString(9)) Then
                        lister_update_array.Add("vide")
                    Else
                        lister_update_array.Add(myReader.GetString(9))
                    End If
                    element_ddd = nom + "?" + numero_preautorise + "?" + myReader.GetString(6)

                    lister_arr_ddd.Add(element_ddd)

                End While

                Dim nbr As Integer = 0
                For Each num In lister_arr_ddd

                    Dim arr_ddd() As String
                    arr_ddd = num.Split("?")


                    Dim total_montant() As Double
                    Dim d_montant As Double
                    arr_ddd(2) = arr_ddd(2).Replace(",", ".")
                    d_montant = CDbl(Val(arr_ddd(2)))
                    Dim reponse_update As String = lister_update_array.Item(nbr)
                    nbr += 1
                    Dim rep1 As String = recuperer_liste_date_check_preautorise(arr_ddd(1))
                    Dim rep2 As String = recuperer_liste_date_preautorise(arr_ddd(1))


                    'TODO corriger ici 
                    Dim rep_total As String
                    If rep1 <> "" And rep2 <> "" Then
                        rep_total = rep1 & "?" & rep2
                    ElseIf rep1 = "" And rep2 <> "" Then
                        rep_total = rep2
                    ElseIf rep1 <> "" And rep2 = "" Then
                        rep_total = rep1
                    End If


#Disable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                    rep_total = faire_ordre_date(rep_total)
#Enable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.

                    total_montant = proc_montant(rep_total, d_montant)


                    If reponse_update <> "vide" Then
                        Dim update_montant As String = arr_ddd(2)
                        total_montant = proc_update_montant(reponse_update, update_montant, total_montant)
                    End If






                    total_hozontal = 0.00

                    For i As Integer = 0 To 11
                        total_hozontal += total_montant(i)
                    Next

                    dt.Rows.Add(arr_ddd(0), total_montant(0).ToString("####0.00 $"), total_montant(1).ToString("####0.00 $"), total_montant(2).ToString("####0.00 $"), total_montant(3).ToString("####0.00 $"), total_montant(4).ToString("####0.00 $"), total_montant(5).ToString("####0.00 $"), total_montant(6).ToString("####0.00 $"), total_montant(7).ToString("####0.00 $"), total_montant(8).ToString("####0.00 $"), total_montant(9).ToString("####0.00 $"), total_montant(10).ToString("####0.00 $"), total_montant(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))


                    For x As Integer = 0 To 11
                        total_colonne_bas(x) += total_montant(x)
                    Next


                Next

            Catch ex As Exception
                MsgBox(ex.Message + " " + where_erreur)

            End Try



            'Dim lister_arr_d As New ArrayList
            'Dim arrays_d(3) As String

            Dim lister_arr_dddd As New List(Of String)
            Dim lister_array_depense As New List(Of String)

            Dim element_dddd As String


            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()
                Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'depense'", myConnection)
                Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


                While myReader.Read
                    element_dddd = ""
                    Dim nom As String = Recupere_description(myReader.GetString(2))
                    Dim numero_depense As String = myReader.GetString(1)

                    'arrays_d(0) = nom
                    'arrays_d(1) = numero_depense
                    'arrays_d(2) = myReader.GetString(6)

                    'lister_arr_d.Add(arrays_d)
                    'lister_array_depense.Add(myReader.GetString(9))
                    If String.IsNullOrEmpty(myReader.GetString(9)) Then
                        lister_array_depense.Add("vide")
                    Else
                        lister_array_depense.Add(myReader.GetString(9))
                    End If
                    element_dddd = nom + "?" + numero_depense + "?" + myReader.GetString(6)

                    lister_arr_dddd.Add(element_dddd)

                End While

                Dim num As String
                Dim nbr As Integer = 0

                For Each num In lister_arr_dddd
                    Dim arr_dddd() As String
                    arr_dddd = num.Split("?")

                    Dim total_montant() As Double

                    arr_dddd(2) = arr_dddd(2).Replace(",", ".")

                    Dim d_montant As Double
                    d_montant = CDbl(Val(arr_dddd(2)))

                    Dim reponse_update As String = lister_array_depense.Item(nbr)


                    nbr += 1

                    Dim rep1 As String = recuperer_liste_date_check_depense(arr_dddd(1))
                    Dim rep2 As String = recuperer_liste_date_depense(arr_dddd(1))

                    Dim rep_total As String
                    If rep1 <> "" And rep2 <> "" Then
                        rep_total = rep1 & "?" & rep2
                    ElseIf rep1 = "" And rep2 <> "" Then
                        rep_total = rep2
                    ElseIf rep1 <> "" And rep2 = "" Then
                        rep_total = rep1
                    End If

#Disable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                    rep_total = faire_ordre_date(rep_total)
#Enable Warning BC42104 ' La variable 'rep_total' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.

                    total_montant = proc_montant(rep_total, d_montant)


                    If reponse_update <> "vide" Then
                        Dim update_montant As String = arr_dddd(2)
                        total_montant = proc_update_montant(reponse_update, update_montant, total_montant)
                    End If


                    total_hozontal = 0.00

                    For i As Integer = 0 To 11
                        total_hozontal += total_montant(i)
                    Next

                    dt.Rows.Add(arr_dddd(0), total_montant(0).ToString("####0.00 $"), total_montant(1).ToString("####0.00 $"), total_montant(2).ToString("####0.00 $"), total_montant(3).ToString("####0.00 $"), total_montant(4).ToString("####0.00 $"), total_montant(5).ToString("####0.00 $"), total_montant(6).ToString("####0.00 $"), total_montant(7).ToString("####0.00 $"), total_montant(8).ToString("####0.00 $"), total_montant(9).ToString("####0.00 $"), total_montant(10).ToString("####0.00 $"), total_montant(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))


                    For x As Integer = 0 To 11
                        total_colonne_bas(x) += total_montant(x)
                    Next



                Next


            Catch ex As Exception

            End Try



            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")

            total_hozontal = 0.00

            For i As Integer = 0 To 11
                total_hozontal += total_colonne_bas(i)
            Next

            dt.Rows.Add("TOtal", total_colonne_bas(0).ToString("####0.00 $"), total_colonne_bas(1).ToString("####0.00 $"), total_colonne_bas(2).ToString("####0.00 $"), total_colonne_bas(3).ToString("####0.00 $"), total_colonne_bas(4).ToString("####0.00 $"), total_colonne_bas(5).ToString("####0.00 $"), total_colonne_bas(6).ToString("####0.00 $"), total_colonne_bas(7).ToString("####0.00 $"), total_colonne_bas(8).ToString("####0.00 $"), total_colonne_bas(9).ToString("####0.00 $"), total_colonne_bas(10).ToString("####0.00 $"), total_colonne_bas(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")

            Dim total_surplus As Double() = New Double(12) {}

            For i As Integer = 0 To 11
                total_surplus(i) = total_colonne(i) - total_colonne_bas(i)
            Next

            'total surplus
            total_hozontal = 0
            For i As Integer = 0 To 11
                total_hozontal += total_surplus(i)
            Next
            Dim montant As Double = obtenir_solde_debut()
            Dim total_encaisse_fin As Double() = New Double(12) {}

            total_encaisse_fin(0) = total_surplus(0) + montant


            Dim encaisse_fin_retraite_manuel As Double = recuperer_bd_argent_a_recevoir()




            'total_encaisse_fin(0) = total_encaisse_fin(0) - encaisse_fin_retraite_manuel

            total_encaisse_fin(0) = recuperer_bd_solde_encaisse_debut() + recuperer_bd_solde_Apports() + recuperer_bd_marge_credit_disponible()

            'total_encaisse_fin(0) = total_encaisse_fin(0) + total_surplus(0)

            total_encaisse_fin(0) = total_encaisse_fin(0) - recuperer_bd_argent_a_recevoir()

            total_encaisse_fin(0) = total_encaisse_fin(0) + total_surplus(0)

            For i As Integer = 1 To 11
                total_encaisse_fin(i) = total_encaisse_fin(i - 1) + total_surplus(i)
            Next




            Dim marge_credit_utilise As Double() = New Double(12) {}

            For i As Integer = 1 To 11
                If total_encaisse_fin(i) > 0 Then
                    marge_credit_utilise(i) = 0
                Else
                    marge_credit_utilise(i) = total_encaisse_fin(i) * -1
                End If

            Next


            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            'dt.Rows.Add("TOtal", total_colonne_bas(0).ToString("####0.00 $"), total_colonne_bas(1).ToString("####0.00 $"), total_colonne_bas(2).ToString("####0.00 $"), total_colonne_bas(3).ToString("####0.00 $"), total_colonne_bas(4).ToString("####0.00 $"), total_colonne_bas(5).ToString("####0.00 $"), total_colonne_bas(6).ToString("####0.00 $"), total_colonne_bas(7).ToString("####0.00 $"), total_colonne_bas(8).ToString("####0.00 $"), total_colonne_bas(9).ToString("####0.00 $"), total_colonne_bas(10).ToString("####0.00 $"), total_colonne_bas(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
            dt.Rows.Add("SURPLUS DE CAISSE", total_surplus(0).ToString("####0.00 $"), total_surplus(1).ToString("####0.00 $"), total_surplus(2).ToString("####0.00 $"), total_surplus(3).ToString("####0.00 $"), total_surplus(4).ToString("####0.00 $"), total_surplus(5).ToString("####0.00 $"), total_surplus(6).ToString("####0.00 $"), total_surplus(7).ToString("####0.00 $"), total_surplus(8).ToString("####0.00 $"), total_surplus(9).ToString("####0.00 $"), total_surplus(10).ToString("####0.00 $"), total_surplus(11).ToString("####0.00 $"), total_hozontal.ToString("####0.00 $"))
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            dt.Rows.Add("ENCAISSE À LA FIN", total_encaisse_fin(0).ToString("####0.00 $"), total_encaisse_fin(1).ToString("####0.00 $"), total_encaisse_fin(2).ToString("####0.00 $"), total_encaisse_fin(3).ToString("####0.00 $"), total_encaisse_fin(4).ToString("####0.00 $"), total_encaisse_fin(5).ToString("####0.00 $"), total_encaisse_fin(6).ToString("####0.00 $"), total_encaisse_fin(7).ToString("####0.00 $"), total_encaisse_fin(8).ToString("####0.00 $"), total_encaisse_fin(9).ToString("####0.00 $"), total_encaisse_fin(10).ToString("####0.00 $"), total_encaisse_fin(11).ToString("####0.00 $"))
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")
            dt.Rows.Add("Marge de crédit utilisée", marge_credit_utilise(0).ToString("####0.00 $"), marge_credit_utilise(1).ToString("####0.00 $"), marge_credit_utilise(2).ToString("####0.00 $"), marge_credit_utilise(3).ToString("####0.00 $"), marge_credit_utilise(4).ToString("####0.00 $"), marge_credit_utilise(5).ToString("####0.00 $"), marge_credit_utilise(6).ToString("####0.00 $"), marge_credit_utilise(7).ToString("####0.00 $"), marge_credit_utilise(8).ToString("####0.00 $"), marge_credit_utilise(9).ToString("####0.00 $"), marge_credit_utilise(10).ToString("####0.00 $"), marge_credit_utilise(11).ToString("####0.00 $"))
            dt.Rows.Add("....", "", "", "", "", "", "", "", "", "", "", "", "")


            For i As Integer = 0 To 11
                If total_encaisse_fin(i) < 0 Then
                    nbr_negative += 1
                    negative = total_encaisse_fin(i)
                    dater_negative = datelist(i).ToString("MMMM yyyy")
                End If
            Next

        Catch ex As Exception

        End Try

        For i As Integer = 0 To 11

        Next


        'charger_table()
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.Filter = "PDF (*.pdf)|*.pdf"



        Me.DataGridView1.DataSource = dt



        'TextBox1.Text = ""
        'TextBox2.Text = ""
#Disable Warning BC42105 ' La fonction 'testpdf' ne retourne pas une valeur pour tous les chemins de code. Une exception de référence null peut se produire au moment de l'exécution lorsque le résultat est utilisé.
    End Function
#Enable Warning BC42105 ' La fonction 'testpdf' ne retourne pas une valeur pour tous les chemins de code. Une exception de référence null peut se produire au moment de l'exécution lorsque le résultat est utilisé.


    Function proc_montant(phrase As String, montant As Double) As Double()

        Dim total_montant() As Double = {0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00}

        Dim num_montant As Integer = 0

        Dim auj As Date = DateCourant
        Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
        end_auj = end_auj.AddDays(-1)
        end_auj = end_auj.AddYears(1)


        Dim strArr() As String
        strArr = phrase.Split("?")

        Dim count As Integer

        For count = 0 To strArr.Length - 1
            '2018-01-29
            Dim chaine_date As String = strArr(count)


            Dim expenddt As Date = Date.ParseExact(chaine_date, "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

            Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


            If expenddt <= end_auj Then

                If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then

                    total_montant(num_montant) = total_montant(num_montant) + montant
                Else

                    'auj = auj.AddMonths(1)
                    'num_montant = num_montant + 1
                    'total_montant(num_montant) = total_montant(num_montant) + montant
                    Dim check As Boolean = True

                    While check

                        firstDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                        lastDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)
                        If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then

                            total_montant(num_montant) = total_montant(num_montant) + montant
                            check = False
                        Else

                            num_montant = num_montant + 1
                            'auj = firstDayLastMonth
                            firstDayLastMonth = firstDayLastMonth.AddMonths(1)
                        End If

                    End While
                    auj = firstDayLastMonth
                End If
            End If

        Next


        Return total_montant

    End Function

    Private Sub enlever_cheque(ByRef list_fourn() As String, ByRef list_colonne() As Integer, ByRef list_montants() As Double, ByVal type As Boolean, ByVal element As Integer)

        Dim strList As List(Of String) = list_fourn.ToList()
        strList.RemoveAt(element)
        list_fourn = strList.ToArray()

        Dim strList1 As List(Of Integer) = list_colonne.ToList()
        strList1.RemoveAt(element)
        list_colonne = strList1.ToArray()

        Dim strList2 As List(Of Double) = list_montants.ToList()
        strList2.RemoveAt(element)
        list_montants = strList2.ToArray()




    End Sub

    Private Sub enlever_direct(ByRef list_fourn() As String, ByRef list_colonne() As Integer, ByRef list_montants() As Double, ByVal type As Boolean, ByVal element As Integer)


        Dim strList As List(Of String) = list_fourn.ToList()
        strList.RemoveAt(element)
        list_fourn = strList.ToArray()

        Dim strList1 As List(Of Integer) = list_colonne.ToList()
        strList1.RemoveAt(element)
        list_colonne = strList1.ToArray()

        Dim strList2 As List(Of Double) = list_montants.ToList()
        strList2.RemoveAt(element)
        list_montants = strList2.ToArray()




    End Sub

    Private Sub enlever_preautoriser(ByRef list_fourn() As String, ByRef list_colonne() As Integer, ByRef list_montants() As Double, ByVal type As Boolean, ByVal element As Integer)

        Dim strList As List(Of String) = list_fourn.ToList()
        strList.RemoveAt(element)
        list_fourn = strList.ToArray()

        Dim strList1 As List(Of Integer) = list_colonne.ToList()
        strList1.RemoveAt(element)
        list_colonne = strList1.ToArray()

        Dim strList2 As List(Of Double) = list_montants.ToList()
        strList2.RemoveAt(element)
        list_montants = strList2.ToArray()

    End Sub


    Function proc_cheque(ByRef nbr_ligne As Integer, ByRef list_fourn() As String, ByRef list_colonne() As Integer) As Double()


        nbr_ligne = 0
        Dim tab_montant As Double() = New Double(12) {}


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_emis", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read


                list_fourn(nbr_ligne) = myReader.GetString(2)


                Dim yValues As Double() = {myReader.GetDouble(7), myReader.GetDouble(8), myReader.GetDouble(9), myReader.GetDouble(10), myReader.GetDouble(11), myReader.GetDouble(12)}

                Dim bon_numero As Integer = 0
                For index As Integer = 0 To yValues.Length - 1


                    If yValues(index) = 0 Then
                        ''#This code will obviously never run, because the value was set to 24
                    Else
                        bon_numero = index
                    End If


                Next
                bon_numero = bon_numero + 7


                tab_montant(nbr_ligne) = myReader.GetDouble(bon_numero)

                Dim date_paiement As DateTime = myReader.GetDateTime(4)

                Dim nbr_colonne As Integer = 0
                Dim auj As Date = DateCourant
                Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                end_auj = end_auj.AddDays(-1)
                end_auj = end_auj.AddYears(1)

                For count = 0 To 11
                    '2018-01-29
                    Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                    Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                    lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)



                    If date_paiement < end_auj Then

                        If date_paiement > firstDayLastMonth AndAlso date_paiement < lastDayLastMonth Then
                            'total_montant(num_montant) = total_montant(num_montant) + montant
                        Else
                            auj = auj.AddMonths(1)
                            'it is not
                            nbr_colonne += 1
                        End If
                    End If

                Next

                list_colonne(nbr_ligne) = nbr_colonne
                nbr_ligne += 1

            End While



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()







        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try



        Return tab_montant
    End Function


    Function proc_direct(ByRef nbr_ligne As Integer, ByRef list_fourn() As String, ByRef list_colonne() As Integer) As Double()



        nbr_ligne = 0
        Dim tab_montant As Double() = New Double(12) {}


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_direct", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read


                list_fourn(nbr_ligne) = myReader.GetString(2)


                Dim yValues As Double() = {myReader.GetDouble(6), myReader.GetDouble(7), myReader.GetDouble(8), myReader.GetDouble(9), myReader.GetDouble(10), myReader.GetDouble(11)}

                Dim bon_numero As Integer = 0
                For index As Integer = 0 To yValues.Length - 1


                    If yValues(index) = 0 Then
                        ''#This code will obviously never run, because the value was set to 24
                    Else
                        bon_numero = index
                    End If


                Next
                bon_numero = bon_numero + 6


                tab_montant(nbr_ligne) = myReader.GetDouble(bon_numero)

                Dim date_paiement As DateTime = myReader.GetDateTime(4)

                Dim nbr_colonne As Integer = 0
                Dim auj As Date = DateCourant
                Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                end_auj = end_auj.AddDays(-1)
                end_auj = end_auj.AddYears(1)

                For count = 0 To 11
                    '2018-01-29
                    Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                    Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                    lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)



                    If date_paiement < end_auj Then

                        If date_paiement > firstDayLastMonth AndAlso date_paiement < lastDayLastMonth Then
                            'total_montant(num_montant) = total_montant(num_montant) + montant
                        Else
                            auj = auj.AddMonths(1)
                            'it is not
                            nbr_colonne += 1
                        End If
                    Else
                        '
                        nbr_colonne = 12
                    End If

                Next

                list_colonne(nbr_ligne) = nbr_colonne
                nbr_ligne += 1

            End While



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()







        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try


        Return tab_montant
    End Function

    Function proc_presautorise(ByRef nbr_ligne As Integer, ByRef list_fourn() As String, ByRef list_colonne() As Integer) As Double(,)

        'presautorsie
        Dim firstDayLastMonth As Date
        Dim lastDayLastMonth As Date

        Dim tab_montant(nbr_ligne, 12) As Double
        nbr_ligne = 0

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'preautorise'", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read

                list_fourn(nbr_ligne) = Recupere_description(myReader.GetString(2))
                Dim sMontant As String = myReader.GetString(6)

                Dim phrase As String = myReader.GetString(7)

                Dim tab_date As Date()
                Dim dimension As Integer = 0

                tab_date = ajouter_lister(phrase, dimension)

                Dim num_montant As Integer = 0

                Dim auj As Date = DateCourant

                Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                Dim debut_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                end_auj = end_auj.AddDays(-1)
                end_auj = end_auj.AddYears(1)



                Dim count As Integer

                For count = 0 To tab_date.Length - 1

                    Dim expenddt As Date = tab_date(count)



                    Dim passe As Boolean = True

                    While passe

                        firstDayLastMonth = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


                        If expenddt > firstDayLastMonth AndAlso expenddt < lastDayLastMonth Then
                            passe = False
                        Else
                            auj = auj.AddMonths(1)
                            'it is not
                            num_montant = num_montant + 1
                            tab_montant(nbr_ligne, num_montant) = 0
                        End If

                    End While




                    If expenddt <= end_auj Then

                        If expenddt > firstDayLastMonth AndAlso expenddt < lastDayLastMonth Then
                            'tab_montant(num_montant) = total_montant(num_montant) + sMontant
                            tab_montant(nbr_ligne, num_montant) = tab_montant(nbr_ligne, num_montant) + sMontant

                        Else
                            auj = auj.AddMonths(1)
                            'it is not
                            num_montant = num_montant + 1
                            tab_montant(nbr_ligne, num_montant) = tab_montant(nbr_ligne, num_montant) + sMontant
                        End If

                    End If
                    'Else
                    '    num_montant = num_montant + 1

                    'End If



                Next
                nbr_ligne += 1

            End While



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        Return tab_montant
    End Function

    Function proc_depense(ByRef nbr_ligne As Integer, ByRef list_fourn() As String, ByRef list_colonne() As Integer) As Double(,)

        'presautorsie
        Dim firstDayLastMonth As Date
        Dim lastDayLastMonth As Date

        Dim tab_montant(nbr_ligne, 12) As Double
        nbr_ligne = 0

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'depense'", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read

                list_fourn(nbr_ligne) = Recupere_description(myReader.GetString(2))
                Dim sMontant As String = myReader.GetString(6)

                Dim phrase As String = myReader.GetString(7)

                Dim tab_date As Date()
                Dim dimension As Integer = 0

                tab_date = ajouter_lister(phrase, dimension)

                Dim num_montant As Integer = 0

                Dim auj As Date = DateCourant

                Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                Dim debut_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                end_auj = end_auj.AddDays(-1)
                end_auj = end_auj.AddYears(1)



                Dim count As Integer

                For count = 0 To tab_date.Length - 1

                    Dim expenddt As Date = tab_date(count)



                    Dim passe As Boolean = True

                    While passe

                        firstDayLastMonth = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


                        If expenddt > firstDayLastMonth AndAlso expenddt < lastDayLastMonth Then
                            passe = False


                        Else

                            auj = auj.AddMonths(1)
                            'it is not
                            num_montant = num_montant + 1
                            tab_montant(nbr_ligne, num_montant) = 0
                        End If

                    End While




                    If expenddt <= end_auj Then

                        If expenddt > firstDayLastMonth AndAlso expenddt < lastDayLastMonth Then
                            'tab_montant(num_montant) = total_montant(num_montant) + sMontant
                            tab_montant(nbr_ligne, num_montant) = tab_montant(nbr_ligne, num_montant) + sMontant

                        Else
                            auj = auj.AddMonths(1)
                            'it is not
                            num_montant = num_montant + 1
                            tab_montant(nbr_ligne, num_montant) = tab_montant(nbr_ligne, num_montant) + sMontant
                        End If

                    End If
                    'Else
                    '    num_montant = num_montant + 1

                    'End If



                Next
                nbr_ligne += 1

            End While



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

        Return tab_montant
    End Function

    Function proc_presautorise1(ByRef nbr_ligne As Integer, ByRef list_fourn() As String, ByRef list_colonne() As Integer) As Double()

        nbr_ligne = 0
        Dim tab_montant As Double() = New Double(10000) {}


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read

                Dim nbr_colonne As Integer
                Dim phrase As String = myReader.GetString(6)

                Dim tab_date As Date()
                Dim dimension As Integer = 0
                Dim auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
                end_auj = end_auj.AddDays(-1)
                end_auj = end_auj.AddYears(1)
                nbr_colonne = 0

                tab_date = ajouter_lister(phrase, dimension)

                Dim debut As Integer = 0

                For Each element As Date In tab_date

                    If DateCourant <= element And element <= end_auj Then

                        Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                        Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


                    End If

                Next



                For Each element As Date In tab_date


                    Dim arret As Boolean = True

                    While arret

                        Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                        Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


                        If element < end_auj Then

                            If element >= firstDayLastMonth AndAlso element <= lastDayLastMonth Then

                                list_fourn(nbr_ligne) = Recupere_description(myReader.GetString(2))
                                Dim sMontant As String = myReader.GetString(5)
                                tab_montant(nbr_ligne) = Convert.ToDouble(sMontant)
                                list_colonne(nbr_ligne) = nbr_colonne
                                arret = False

                            Else
                                auj = auj.AddMonths(1)
                                'it is not
                                nbr_colonne += 1

                            End If



                        Else
                            arret = False

                        End If

                    End While

                    nbr_ligne += 1

                Next

            End While



            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try


        Return tab_montant
    End Function


    Function obtenir_solde_debut() As Double

        Dim montant As Double

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM Solde_encaisses_debut where numero = 1", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim nombre As Integer


            While myReader.Read


                montant = myReader.GetDouble(1)


            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()


        Catch ex As Exception

        End Try

        Return montant

    End Function

    Function ajouter_lister(ByVal chaine As String, ByRef nbr_date As Integer) As Date()


        Dim strArr() As String
        strArr = chaine.Split("?")

        Dim dimension As Integer = strArr.Length

        Dim tab_date As Date() = New Date(dimension) {}
        nbr_date = 0

        Dim count As Integer




        Dim listerp As New ArrayList()

        For count = 0 To strArr.Length - 1
            tab_date(count) = Date.Parse(strArr(count))
            nbr_date += 1
        Next

        Array.Resize(tab_date, tab_date.Length - 1)

        Return tab_date

    End Function

    Function proc_revenu_mois_courant(ByVal element As String) As String

        Dim desc As String = ""

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_revenu where payer like " & element & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim nombre As Integer

            Dim dater As String
            Dim strArr() As String

            While myReader.Read

                dater = myReader.GetDateTime(3).ToString()
                strArr = dater.Split(" ")

                desc = desc + strArr(0) + "?"


                'MsgBox(myReader.GetValue(5))
            End While

            desc = desc.Remove(desc.Length - 1)
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try

        Return desc

    End Function

    Private Sub proc_montant_deuxieme(ByVal chaine As String, ByRef total_montant() As Double, ByVal montant As Double)

        Dim num_montant As Integer = 0

        Dim auj As Date = DateCourant
        Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
        end_auj = end_auj.AddDays(-1)
        end_auj = end_auj.AddYears(1)


        Dim strArr() As String
        strArr = chaine.Split("?")

        Dim count As Integer

        For count = 0 To strArr.Length - 1
            '2018-01-29
            Dim chaine_date As String = strArr(count)


            Dim expenddt As Date = Date.ParseExact(chaine_date, "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

            Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)




            If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then
                total_montant(num_montant) = total_montant(num_montant) + montant
            Else
                auj = auj.AddMonths(1)
                'it is not
                num_montant = num_montant + 1
                total_montant(num_montant) = total_montant(num_montant) + montant
            End If


        Next


    End Sub

    Function compter_check() As Integer

        Dim compte As Integer = 0


        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_emis ", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                compte += 1
            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try



        Return compte
    End Function

    Function compter_preautoriser() As Integer

        Dim compte As Integer = 0


        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'preautorise'", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                compte += 1
            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try



        Return compte
    End Function

    Function compter_depense() As Integer

        Dim compte As Integer = 0


        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where categorie LIKE 'depense'", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                compte += 1
            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try



        Return compte
    End Function

    Function compter_direct() As Integer
        Dim compte As Integer = 0


        Try
            Dim myConnection As New OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_direct ", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read
                compte += 1
            End While

            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()



        Catch ex As Exception

        End Try


        Return compte
    End Function

    Function recuperer_liste_date_check_revenu(ByVal numero As String) As String



        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()


            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_revenu where payer like " & numero & " ORDER BY date_paiement", myConnection)


            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read
                Dim reponse As DateTime = myReader.GetDateTime(3)

                Dim rep As String = reponse.ToString("yyyy-MM-dd")

                code = code + rep + "?"

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try


        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If



        Return code

    End Function


    Function recuperer_liste_date_revenu(ByVal numero As String) As String



        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM revenu where numero = " & valueIntConverted & "", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read

                code = myReader.GetString(5)

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If

        Return code

    End Function

    Function recuperer_liste_date_check_depense(ByVal numero As String) As String


        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            'Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_revenu where numero = " & numero & "", myConnection)
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_depenses where payer like " & numero & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read
                Dim reponse As DateTime = myReader.GetDateTime(3)

                Dim rep As String = reponse.ToString("yyyy-MM-dd")

                code = code + rep + "?"

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If


        Return code

    End Function

    Function recuperer_liste_date_depense(ByVal numero As String) As String

        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM depense where numero = " & valueIntConverted & "", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read

                code = myReader.GetString(5)

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If

        Return code

    End Function

    Function recuperer_liste_date_check_preautorise(ByVal numero As String) As String

        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)


        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            'Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_revenu where numero = " & numero & "", myConnection)
            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_preautorises where payer like " & numero & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read
                Dim reponse As DateTime = myReader.GetDateTime(3)

                Dim rep As String = reponse.ToString("yyyy-MM-dd")

                code = code + rep + "?"

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If




        Return code

    End Function

    Function recuperer_liste_date_preautorise(ByVal numero As String) As String



        Dim code As String = ""
        Dim valueIntConverted As Integer = CInt(numero)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM mensuel where numero = " & valueIntConverted & "", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            'Dim tab_periode As String
            'Dim montant As Double
            'Dim check_oui As String


            While myReader.Read

                code = myReader.GetString(2)

            End While


            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()

        Catch ex As Exception

        End Try

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If

        Return code

    End Function

    Function recuperer_bd_argent_a_recevoir() As Double

        Dim reponse As String

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Argent_a_recevoir where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                reponse = myReader1.Item(0).ToString


            End While


            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
        Return CDbl(reponse)
#Enable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
    End Function

    Function recuperer_bd_marge_credit_disponible() As Double

        Dim reponse As String

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from marge_credit_disponible where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                reponse = myReader1.Item(0).ToString


            End While


            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
        Return CDbl(reponse)
#Enable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
    End Function

    Function recuperer_bd_solde_encaisse_debut() As Double

        Dim reponse As String

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Solde_encaisses_debut where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                reponse = myReader1.Item(0).ToString


            End While


            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
        Return CDbl(reponse)
#Enable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
    End Function


    Function recuperer_bd_solde_Apports() As Double

        Dim reponse As String

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()
            Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Apports where numero = 1", myConnection)
            Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

            While myReader1.Read
                reponse = myReader1.Item(0).ToString


            End While


            myConnection.Close()

            myCommand1.Dispose()
            myReader1.Close()

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
        Return CDbl(reponse)
#Enable Warning BC42104 ' La variable 'reponse' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
    End Function


    Function faire_ordre_date(ByVal numero As String) As String



        Dim code As String = ""


        Dim strArr() As String
        strArr = numero.Split("?")

        Dim ArrDates As New List(Of Date)

        For Each s As String In strArr
            Dim strArr1() As String
            strArr1 = s.Split("-")

            Dim mot As String = strArr1(0) + "-" + strArr1(1) + "-" + strArr1(2)
            Dim dt As DateTime = DateTime.ParseExact(mot, "yyyy-MM-dd", Nothing)

            'Dim dt As Date = Convert.ToDateTime(mot)
            ArrDates.Add(dt)
            'MsgBox(s)
        Next

        'For i As Integer = 0 To ArrAgentsReportCheck.Length - 1
        '    ArrDates.Add(ArrAgentsReportCheck(i))
        'Next

        'ArrDates.Sort()
        'Dim LatestDate As Date = ArrDates.Max()

        ArrDates.Sort()

        Dim ArrDates1 As New List(Of Date)
        Dim auj As Date = Date.Now

        Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
        'check faire un ordre
        For Each d As Date In ArrDates

            If d >= firstDayLastMonth Then
                ArrDates1.Add(d)
            End If

        Next
        '*************************************
        ArrDates1.Sort()

        For Each d As Date In ArrDates1
            Dim d2 As String = d.ToString("yyyy-MM-dd")
            Dim strArr2() As String
            strArr2 = d2.Split("-")
            'Dim rep As String = strArr2(0) + "-" + strArr2(2) + "-" + strArr2(1)
            Dim rep As String = strArr2(0) + "-" + strArr2(1) + "-" + strArr2(2)
            code = code + rep + "?"
        Next


        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If

        strArr = code.Split("?")

        Dim list As New ArrayList

        For Each s As String In strArr
            If list.Count > 0 Then
                Dim trouve As Boolean = False


                For Each total As String In list
                    If s = total Then
                        trouve = True
                    End If
                Next

                If trouve = False Then
                    list.Add(s)
                End If
            Else
                list.Add(s)
            End If


        Next

        Dim nouv_string As String = ""

        For Each d As String In list

            nouv_string = nouv_string + d + "?"

        Next

        code = nouv_string

        If code <> "" Then
            Dim lastLetter As Char = code(code.Length - 1)
            If lastLetter = "?" Then
                code = code.Remove(code.Length - 1)
            End If

        End If



        Return code

    End Function



    Function proc_update_montant(ByVal dater As String, ByVal update_montant As String, ByVal group_montant As Double()) As Double()

        Dim bVider As Boolean = False
        Dim dIndex
        Dim str_montants() As String
        Dim str_montant As String = ""
        Dim num_montant As Integer = 0


        dIndex = dater.IndexOf("?")

        If (dIndex > -1) Then
            str_montants = dater.Split("?")
            bVider = True
        Else
            str_montant = dater
        End If

        Dim auj As Date = DateCourant
        Dim end_auj As Date = New Date(DateCourant.Year, DateCourant.Month, 1)
        end_auj = end_auj.AddDays(-1)
        end_auj = end_auj.AddYears(1)

        If bVider = True Then
#Disable Warning BC42104 ' La variable 'str_montants' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            For count = 0 To str_montants.Length - 1
#Enable Warning BC42104 ' La variable 'str_montants' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                '2018-01-29

                Dim element() As String

                Dim chaine_date As String = str_montants(count)

                element = chaine_date.Split("/")


                Dim expenddt As Date = Date.ParseExact(element(0), "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

                Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
                lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


                If expenddt <= end_auj Then

                    If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then
                        Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                        ciClone.NumberFormat.NumberDecimalSeparator = "."
                        Dim argent As Double = 0.0
                        'argent = CDbl(update_montant)
                        argent = Convert.ToDouble(update_montant, ciClone)

                        Dim nouv As Double = 0.0

                        Dim ponct() As String
                        ponct = element(1).Split("/")
                        Dim frist_ponct As String = ponct(0).Replace(",", ".")
                        nouv = Convert.ToDouble(frist_ponct, ciClone)
                        'Dim nouv As Double = CDbl(element(1))
                        group_montant(num_montant) = group_montant(num_montant) - argent
                        group_montant(num_montant) = group_montant(num_montant) + nouv

                    Else

                        'auj = auj.AddMonths(1)
                        'num_montant = num_montant + 1
                        'total_montant(num_montant) = total_montant(num_montant) + montant
                        Dim check As Boolean = True

                        While check

                            firstDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                            lastDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                            lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)
                            where_erreur = firstDayLastMonth

                            If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then
                                Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                                ciClone.NumberFormat.NumberDecimalSeparator = "."
                                Dim argent As Double = 0.0
                                'argent = CDbl(update_montant)
                                argent = Convert.ToDouble(update_montant, ciClone)

                                Dim nouv As Double = 0.0

                                Dim ponct As String = element(1)
                                ponct = ponct.Substring(0, ponct.Length - 2)
                                Dim frist_ponct As String = ponct.Replace(",", ".")
                                nouv = Convert.ToDouble(frist_ponct, ciClone)
                                'Dim nouv As Double = CDbl(element(1))

                                group_montant(num_montant) = group_montant(num_montant) - argent
                                group_montant(num_montant) = group_montant(num_montant) + nouv
                                check = False
                            Else
                                num_montant = num_montant + 1
                                'auj = firstDayLastMonth
                                firstDayLastMonth = firstDayLastMonth.AddMonths(1)
                            End If

                        End While
                        auj = firstDayLastMonth
                    End If
                End If

            Next

        Else
            Dim element() As String

            Dim chaine_date As String = str_montant

            element = chaine_date.Split("/")


            Dim expenddt As Date = Date.ParseExact(element(0), "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

            Dim firstDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            Dim lastDayLastMonth As Date = New Date(auj.Year, auj.Month, 1)
            lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)


            If expenddt <= end_auj Then

                If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then
                    Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                    ciClone.NumberFormat.NumberDecimalSeparator = "."
                    Dim argent As Double = 0.0
                    'argent = CDbl(update_montant)
                    argent = Convert.ToDouble(update_montant, ciClone)

                    Dim nouv As Double = 0.0

                    Dim ponct As String = element(1)
                    ponct = ponct.Substring(0, ponct.Length - 2)
                    Dim frist_ponct As String = ponct.Replace(",", ".")
                    nouv = Convert.ToDouble(frist_ponct, ciClone)
                    'Dim nouv As Double = CDbl(element(1))

                    group_montant(num_montant) = group_montant(num_montant) - argent
                    group_montant(num_montant) = group_montant(num_montant) + nouv

                Else

                    'auj = auj.AddMonths(1)
                    'num_montant = num_montant + 1
                    'total_montant(num_montant) = total_montant(num_montant) + montant
                    Dim check As Boolean = True

                    While check

                        firstDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                        lastDayLastMonth = New Date(firstDayLastMonth.Year, firstDayLastMonth.Month, 1)
                        lastDayLastMonth = lastDayLastMonth.AddMonths(1).AddDays(-1)
                        If expenddt >= firstDayLastMonth AndAlso expenddt <= lastDayLastMonth Then

                            Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
                            ciClone.NumberFormat.NumberDecimalSeparator = "."
                            Dim argent As Double = 0.0
                            'argent = CDbl(update_montant)
                            argent = Convert.ToDouble(update_montant, ciClone)

                            Dim nouv As Double = 0.0

                            Dim ponct As String = element(1)
                            ponct = ponct.Substring(0, ponct.Length - 2)
                            Dim frist_ponct As String = ponct.Replace(",", ".")
                            nouv = Convert.ToDouble(frist_ponct, ciClone)
                            'Dim nouv As Double = CDbl(element(1))


                            group_montant(num_montant) = group_montant(num_montant) - argent
                            group_montant(num_montant) = group_montant(num_montant) + nouv
                            check = False
                        Else
                            num_montant = num_montant + 1
                            'auj = firstDayLastMonth
                            firstDayLastMonth = firstDayLastMonth.AddMonths(1)
                        End If

                    End While
                    auj = firstDayLastMonth
                End If
            End If


        End If



        Return group_montant






    End Function

End Class
