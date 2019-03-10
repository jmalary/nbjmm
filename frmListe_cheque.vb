
Imports System.Data.OleDb
Imports Comptable_NBJMM.GlobalVariables

Public Class frmListe_cheque

    Public myConnToAccess As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim tables As DataTableCollection

    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader

    Dim sqlStr As String
    'Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;"




    Dim myConnString As String = GlobalVariables.test2ConnectionString

    Private Sub frmListe_cheque_LocationChanged()
        MsgBox("hey")

    End Sub

    Private Sub frmListe_cheque_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_emis", myConnection)
        Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()




        While myReader.Read
            Dim newListViewItem As New ListViewItem

            newListViewItem.Text = myReader.GetInt32(0)
            'newListViewItem.SubItems.Add(myReader.GetString(1))



            '*************************************************************************************

            Dim today = myReader.GetDateTime(2)

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

            today = myReader.GetDateTime(3)



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


            newListViewItem.SubItems.Add(myReader.GetInt32(4))

            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(5), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(6), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(7), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(8), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetDouble(9), "Currency"))
            ListView1.Items.Add(newListViewItem)
        End While

        proc_A()
        proc_B()
        proc_C()
        proc_D()
        proc_E()


        myConnection.Close()

        proc_Vider()


    End Sub

  
    Sub proc_A()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")


        End While

    End Sub

    Sub proc_B()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label3.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While

    End Sub

    Sub proc_C()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label4.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_D()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label5.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_E()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(cinq) from check_emis", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label7.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'frmMenu.Show()
        Me.Close()


    End Sub

    Sub proc_Vider()



        myConnToAccess = New OleDbConnection(GlobalVariables.test2ConnectionString)
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        da = New OleDbDataAdapter("SELECT numero from check_emis", myConnToAccess)
        da.Fill(ds, "employeur")
        myConnToAccess.Dispose()

        myConnToAccess = Nothing



        Dim quantity As Integer = ds.Tables.Count

        'If ds.Tables.Count > 0 Then

        '    quantity += 1
        '    'MsgBox(ds.Tables(0).Rows(-1).Item(0))
        '    MsgBox(ds.Tables(0).Rows(quantity).Item(0))
        'Else

        '    MsgBox("aucunn")
        'End If




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If (ListView1.SelectedItems.Count > 0) Then
            For Each i As ListViewItem In ListView1.SelectedItems
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("DELETE FROM check_emis WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@numero", i.Text)

                myCommand.ExecuteNonQuery()
                myConnection.Close()

                ListView1.Items.Remove(i)
            Next
        Else
            MessageBox.Show("please select an item to delete")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (ListView1.SelectedItems.Count > 0) Then
            'MsgBox(ListView1.SelectedItems(0).SubItems(0).Text)
            frmUpdate_check.Show()
            frmUpdate_check.Label2.Text = ListView1.SelectedItems(0).SubItems(0).Text
            Me.Hide()



        Else
            MessageBox.Show("please select an item to update")
        End If
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub frmListe_cheque_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged



        'Dim x As Integer
        'Dim y As Integer
        'Dim r As Rectangle

        'If Not Parent Is Nothing Then
        '    r = Parent.ClientRectangle
        '    x = r.Width - Me.Width + Parent.Left
        '    y = r.Height - Me.Height + Parent.Top
        'Else
        '    r = Screen.PrimaryScreen.WorkingArea
        '    x = r.Width - Me.Width
        '    y = r.Height - Me.Height
        'End If

        'x = CInt(x / 2)
        'y = CInt(y / 2)

        'Me.StartPosition = FormStartPosition.Manual
        'Me.Location = New Point(x, y)
    End Sub


    'Private Sub ListView1_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles ListView1.ColumnWidthChanging
    '    e.Cancel = True;
    '    If e.ColumnIndex = 0 AndAlso ListView1.Columns(e.ColumnIndex).Width <> 10 Then
    '        ListView1.Columns(e.ColumnIndex).Width = 10
    '    End If

    'End Sub
    Private Sub listView1_ColumnWidthChanging(ByVal sender As Object, ByVal e As ColumnWidthChangingEventArgs)
        If e.ColumnIndex = 0 Then
            e.Cancel = True
        End If
    End Sub
    Private Sub listView1_ColumnWidthChanged(ByVal sender As Object, ByVal e As ColumnWidthChangedEventArgs)
        If e.ColumnIndex = 0 AndAlso ListView1.Columns(e.ColumnIndex).Width <> 10 Then
            ListView1.Columns(e.ColumnIndex).Width = 10
        End If
    End Sub

    Private Sub frmListe_cheque_MaximizedBoundsChanged(sender As Object, e As EventArgs) Handles Me.MaximizedBoundsChanged
        MsgBox("Fenetre ")
    End Sub

    Private Sub frmListe_cheque_MaximumSizeChanged(sender As Object, e As EventArgs) Handles Me.MaximumSizeChanged
        MsgBox("Fenetre ")
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class