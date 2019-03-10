Imports Comptable_NBJMM.GlobalVariables
Imports System.Data.OleDb

Public Class calculerGestion

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.test2ConnectionString
    Dim AL As New Collection
    Dim somme_solde_enc As Double
    Dim num_temporaire As String
    Dim lister As New List(Of String)


    Private Property OleDbDataReader As OleDbDataReader


    Public Sub proc_tout()

        proc_solde_encaisse_a()
        proc_apport()


    End Sub

    Private Sub proc_solde_encaisse_a()

        proc_1A()

    End Sub

    Private Sub proc_1A()

        Dim ajouter As String = ""

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select courant from Solde_encaisses_debut where numero = 1", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            ajouter = myReader1.Item(0).ToString
        End While

        lister.Add(ajouter)

        myConnection.Close()

    End Sub


    Private Sub proc_apport()

    End Sub


    Private Sub proc_1B()

    End Sub

    Private Sub proc_1C()

    End Sub

    Private Sub proc_1D()

    End Sub

    Private Sub proc_1E()

    End Sub


End Class
