Public Class frmListe1
    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\test2.mdb;"


    Private Sub frmListe1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()

        Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM check_preautorises", myConnection)
        Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

        While myReader.Read
            Dim newListViewItem As New ListViewItem
            newListViewItem.Text = myReader.GetInt32(0)
            newListViewItem.SubItems.Add(myReader.GetString(1))
            newListViewItem.SubItems.Add(myReader.GetDateTime(2))
            newListViewItem.SubItems.Add(myReader.GetDateTime(3))

            newListViewItem.SubItems.Add(Format$(myReader.GetInt32(4), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetInt32(5), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetInt32(6), "Currency"))
            newListViewItem.SubItems.Add(Format$(myReader.GetInt32(7), "Currency"))
            ListView1.Items.Add(newListViewItem)
        End While


        'Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_direct", myConnection)
        'Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()
        'Dim myCommand2 As New OleDb.OleDbCommand("Select sum(deux) from check_direct", myConnection)
        'Dim myReader2 As OleDb.OleDbDataReader = myCommand2.ExecuteReader()
        ''Dim myCommand3 As New OleDb.OleDbCommand("Select sum(trois) from check_direct", myConnection)
        'Dim myReader3 As OleDb.OleDbDataReader = myCommand3.ExecuteReader()
        'Dim myCommand4 As New OleDb.OleDbCommand("Select sum(quatre) from check_direct", myConnection)
        'Dim myReader4 As OleDb.OleDbDataReader = myCommand4.ExecuteReader()


        'While myReader1.Read
        '    Label2.Text = Format$(myReader1.Item(0).ToString, "Currency")
        '    Label3.Text = Format$(myReader2.Item(0).ToString, "Currency")
        '    'Label4.Text = Format$(myReader3.Item(0).ToString, "Currency")
        '    'Label5.Text = Format$(myReader4.Item(0).ToString, "Currency")

        'End While
        proc_A()
        proc_B()
        proc_C()
        proc_D()


        myConnection.Close()


    End Sub
    Sub proc_A()
        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(un) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label7.Text = Format$(myReader1.Item(0).ToString, "Currency")


        End While

    End Sub

    Sub proc_B()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(deux) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label3.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While

    End Sub

    Sub proc_C()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(trois) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label4.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Sub proc_D()

        Dim myConnection As New OleDb.OleDbConnection(myConnString)
        myConnection.Open()
        Dim myCommand1 As New OleDb.OleDbCommand("Select sum(quatre) from check_preautorises", myConnection)
        Dim myReader1 As OleDb.OleDbDataReader = myCommand1.ExecuteReader()

        While myReader1.Read
            Label5.Text = Format$(myReader1.Item(0).ToString, "Currency")
        End While


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmMenu.Show()
        Me.Hide()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class