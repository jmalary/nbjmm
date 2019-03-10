Imports System.Data.OleDb
Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization


Public Class frmDetailsFinancement

    Dim choice As String
    Public myConnToAccess As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim tables As DataTableCollection

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim temp_montant As String
    Dim montant_global As String

    Dim sqlStr As String
    'Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;"

    Dim myConnString As String = GlobalVariables.test2ConnectionString


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Me.WindowState = FormWindowState.Maximized
        Me.Close()


        Dim f As New frmAfficherListeFinancements()
        f.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        f.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        f.Visible = True
        f.Location = New Point(1, 0)

        frmInterface.Panel1.Controls.Add(f)


        'frmLoading.Show()
        'frmAfficherListeFinancements.MdiParent = Form_windows
        'frmAfficherListeFinancements.Show()







    End Sub

    Private Sub frmDetailsFinancement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        obtenir_liste()
    End Sub

    Private Sub obtenir_liste()

        'Dim numero As String = Label2.Text
        Dim num_financement As String

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM list_financement where id = " & num_list_financement & "", myConnection)

            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




            While myReader.Read


                num_financement = myReader.GetString(1)
                Dim sNom As String = Recupere_description(myReader.GetString(2))

                lbl_fournisseur.Text = sNom
                Dim payer As Date = myReader.GetDateTime(3)
                lbl_date_paiement.Text = payer.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA"))
                lbl_mode.Text = myReader.GetString(5)
                lbl_montant.Text = myReader.GetString(6) + " $"
                temp_montant = myReader.GetString(9)
                montant_global = myReader.GetString(6)
                ajouter_lister(myReader.GetString(7), myReader.GetString(6))
                Dim payer1 As Date = myReader.GetDateTime(8)
                lbl_date_echances.Text = payer1.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA"))

                'MsgBox(myReader.GetValue(5))
            End While
            myConnection.Close()
            myCommand.Dispose()
            myReader.Close()


        Catch ex As Exception

        End Try

        'supprimer_mensuel(num_financement)

    End Sub


    Sub ajouter_lister(ByVal chaine As String, ByVal montant As String)

        Dim bmontant As Boolean = False

        Dim dIndex
        Dim str_montants() As String
        Dim str_montant As String = ""

        If temp_montant <> "" Then
            bmontant = True
        End If

        Dim strArr() As String
        Dim count As Integer

        strArr = chaine.Split("?")
        Dim listerp As New ArrayList

        lbl_nbr_transaction.Text = strArr.Length.ToString

        If bmontant = True Then
            dIndex = temp_montant.IndexOf("?")

            If (dIndex > -1) Then
                str_montants = temp_montant.Split("?")
            Else
                str_montant = temp_montant
            End If

        End If

        For count = 0 To strArr.Length - 1
            'MsgBox(strArr(count))
            Dim dater As Date = Date.Parse(strArr(count))
            'Dim list As New List(Of Date)

            listerp.Add(dater)
        Next


        Dim i As Integer = 0

        Dim num As Date
        For Each num In listerp
            Dim newListViewItem As New ListViewItem
            i += 1

            newListViewItem.Text = i

            If bmontant = True Then
                Dim str_dater() As String

                If str_montant = "" Then
#Disable Warning BC42104 ' La variable 'str_montants' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                    For Each s As String In str_montants
#Enable Warning BC42104 ' La variable 'str_montants' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                        str_dater = s.Split("/")
                        Dim expenddt As Date = Date.ParseExact(str_dater(0), "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If expenddt = num Then
                            montant = str_dater(1)
                        End If
                    Next

                Else
                    str_dater = str_montant.Split("/")
                    Dim expenddt As Date = Date.ParseExact(str_dater(0), "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    If expenddt = num Then
                        montant = str_dater(1)
                    End If

                End If

            End If



            newListViewItem.SubItems.Add(
                num.ToString("dddd dd MMMM yyyy",
              CultureInfo.CreateSpecificCulture("fr-CA")))
            '
            newListViewItem.SubItems.Add(montant + " $")
            Console.WriteLine(num)
            ListView1.Items.Add(newListViewItem)
            montant = montant_global

        Next
    End Sub
End Class