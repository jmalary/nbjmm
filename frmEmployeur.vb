Imports System.Data.SqlClient
Imports Comptable_NBJMM.GlobalVariables


Public Class frmEmployeur


    Dim con As System.Data.OleDb.OleDbConnection
    Dim cmd As System.Data.OleDb.OleDbCommand
    Dim dr As System.Data.OleDb.OleDbDataReader
    Dim sqlStr As String
    Dim myConnString As String = GlobalVariables.test2ConnectionString
    Dim last_num As Integer


    Dim con1 As SqlConnection

    Dim stcon As String
    Dim ds As DataSet
    Dim da As SqlDataAdapter

    Dim CW As Integer = Me.Width ' Current Width
    Dim CH As Integer = Me.Height ' Current Height
    Dim IW As Integer = Me.Width ' Initial Width
    Dim IH As Integer = Me.Height ' Initial Height


    Private Sub btmAjouter_Click(sender As Object, e As EventArgs) Handles btmAjouter.Click

        If TextBox2.Text.Trim <> "" Then
            ajouter_bd()
            initialser()
        Else
            MsgBox("Please enter a value")

        End If







    End Sub


    Private Sub ajouter_bd()
        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()
            'copier la connection string above
            'ouvrir la connection 
            ' MsgBox("connection Ok")
            sqlStr = "INSERT INTO employeur (numFournisseur, description,adresse,ville, code) VALUES('" & TextBox7.Text & "','" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox6.Text & "')"
            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            cmd.ExecuteNonQuery()
            MsgBox("Record inserf ok")
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox6.Text = ""



            cmd.Dispose()
            con.Close()
            proc_A()
        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try


    End Sub

    Private Sub initialser()

        Dim nom As String

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            While myReader.Read


                nom = myReader.GetString(1)


            End While


            myCommand.Dispose()


            myReader.Close()
            myConnection.Close()

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' La variable 'nom' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
        If nom = "24" Then
#Enable Warning BC42104 ' La variable 'nom' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.


            last_num = 0

            last_num += 1
        ElseIf nom = "23" Then
            last_num = CInt(Int(nom))

            last_num += 2
        Else

            last_num = CInt(Int(nom))

            last_num += 1
        End If

        TextBox7.Text = CStr(last_num)

    End Sub


    Sub proc_A()

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            ListView1.Items.Clear()

            For Each selection As ListViewItem In ListView1.SelectedItems()

                If selection.Selected = True Then
                    ListView1.SelectedItems.Clear()
                End If
            Next

            myCommand.Dispose()


            myReader.Close()
            myConnection.Close()

            myConnection.Open()




        Catch ex As Exception

        End Try

        Try

            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()


            While myReader.Read






                Dim nom As String = myReader.GetString(2)

                If nom <> "Choisir un employeur" Then
                    Dim newListViewItem As New ListViewItem
                    newListViewItem.Text = myReader.GetString(1)


                    newListViewItem.SubItems.Add(myReader.GetString(2))
                    newListViewItem.SubItems.Add(myReader.GetString(3))
                    newListViewItem.SubItems.Add(myReader.GetString(4))
                    newListViewItem.SubItems.Add(myReader.GetString(5))
                    ListView1.Items.Add(newListViewItem)
                End If

            End While
            myCommand.Dispose()


            myReader.Close()
            myConnection.Close()

        Catch ex As Exception

        End Try
        ListView1.Refresh()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Dim id As Integer
        id = Val(InputBox("Enter the emp id"))

        'id  contain the value of ID

        Try
            con = New System.Data.OleDb.OleDbConnection(GlobalVariables.test2ConnectionString)
            con.Open()
            'copier la connection string above
            'ouvrir la connection 
            'MsgBox("connection Ok")
            sqlStr = "Select * from employeur where numero=" & id & ""

            'write the sql query
            cmd = New OleDb.OleDbCommand(sqlStr, con)
            'cmd.ExecuteNonQuery()

            dr = cmd.ExecuteReader()
            If dr.HasRows = True Then
                While dr.Read()
                    TextBox1.Text = dr(0)
                    TextBox2.Text = dr(1)
                End While
            End If
            'ExecuteReader fo
            MsgBox("Record Found")
            cmd.Dispose()
            dr.Close()
            con.Close()

        Catch ex As Exception
            MessageBox.Show("could not find record")

        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

        TextBox1.Text = ""
        TextBox2.Text = ""

        MsgBox("Please enter the data")

        TextBox1.Focus()

    End Sub

    Private Sub frmEmployeur_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed


    End Sub

    Private Sub frmEmployeur_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        IW = Me.Width
        IH = Me.Height

        Dim longueur As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 40
        Dim hauteur As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - 250




        Dim no_emp As Integer = 0

        'rs.FindAllControls(Me)
        'con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;")
        'stcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;"
        'con1 = New SqlConnection(stcon)
        'con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\test2.mdb;")
        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read

                'Dim nom As String companyCode = myReader.GetValue(0).ToString();
                Dim nom As String = myReader.GetString(2)

                If nom <> "Choisir un employeur" Then
                    Dim newListViewItem As New ListViewItem
                    newListViewItem.Text = myReader.GetString(1)


                    newListViewItem.SubItems.Add(myReader.GetString(2))
                    'newListViewItem.SubItems.Add(myReader.GetString(3))
                    'newListViewItem.SubItems.Add(myReader.GetString(4))
                    'newListViewItem.SubItems.Add(myReader.GetString(5))
                    ListView1.Items.Add(newListViewItem)
                    no_emp += 1
                End If

            End While
            myCommand.Dispose()
            myReader.Close()

            myConnection.Close()

        Catch ex As Exception

        End Try

        Dim AL As New Collection
        Dim tabs() As String = {"dog", "cat", "fish"}

        AL.Add(TextBox5)
        AL.Add(Label5)

        AL(1).TEXT = tabs(0)
        AL(2).text = tabs(1)



        Label7.Text = "Total = " + no_emp.ToString() + " données"
        initialser()


        Dim center As Integer = Screen.PrimaryScreen.Bounds.Width - Label7.Width

        center = center \ 2

        Label7.Location = New Point(center, 68)

        lblTItre.Location = New Point(center, 10)


    End Sub

    Private Sub btmRefresh_Click(sender As Object, e As EventArgs)

        Try
            Dim myConnection As New OleDb.OleDbConnection(myConnString)
            myConnection.Open()

            Dim myCommand As New OleDb.OleDbCommand("SELECT * FROM employeur", myConnection)
            Dim myReader As OleDb.OleDbDataReader = myCommand.ExecuteReader()

            For Each selection As ListViewItem In ListView1.SelectedItems()

                If selection.Selected = True Then
                    ListView1.SelectedItems.Clear()
                End If
            Next
            ListView1.Refresh()


            While myReader.Read
                Dim newListViewItem As New ListViewItem
                newListViewItem.Text = myReader.GetInt32(0)
                newListViewItem.SubItems.Add(myReader.GetString(1))
                ListView1.Items.Add(newListViewItem)
            End While


            myCommand.Dispose()
            myReader.Close()
            myConnection.Close()

        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

        If (ListView1.SelectedItems.Count > 0) Then
            'MsgBox(ListView1.SelectedItems(0).SubItems(0).Text)

            Try
                Dim myConnection As New OleDb.OleDbConnection(myConnString)
                myConnection.Open()

                Dim myCommand As New OleDb.OleDbCommand("UPDATE employeur SET numero = @numero, description = @description WHERE numero = @numero", myConnection)
                myCommand.Parameters.AddWithValue("@description", TextBox2.Text)
                myCommand.Parameters.AddWithValue("@numero", 1)

                myCommand.ExecuteNonQuery()

                myCommand.Dispose()
                myConnection.Close()




            Catch ex As Exception

            End Try
            'frmUpdate_check.Show()



        Else
            MessageBox.Show("please select an item to update")
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If (ListView1.SelectedItems.Count > 0) Then
            For Each i As ListViewItem In ListView1.SelectedItems
                Try
                    Dim myConnection As New OleDb.OleDbConnection(myConnString)
                    myConnection.Open()


                    Dim myCommand As New OleDb.OleDbCommand("DELETE FROM employeur WHERE numFournisseur = @numero", myConnection)
                    myCommand.Parameters.AddWithValue("@numero", i.Text)

                    myCommand.ExecuteNonQuery()
                    myConnection.Close()
                    ListView1.Items.Remove(i)

                    myCommand.Dispose()





                Catch ex As Exception
                    MsgBox("deletee non reussi")
                End Try

            Next
        Else
            MessageBox.Show("s'il vous plaît sélectionner un élément à supprimer")
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()

        frmInterface.btm_fournisseur.Enabled = True
        Form_windows.ToolStripMenuItemFournisseur.Enabled = True


        'frmMenu.Show()

    End Sub

    Private Sub frmEmployeur_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
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

    'Private Sub changer_resolution()

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

    Private Sub frmEmployeur_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim RW As Double = (Me.Width - CW) / CW ' Ratio change of width
        Dim RH As Double = (Me.Height - CH) / CH ' Ratio change of height

        For Each Ctrl As Control In Controls
            Ctrl.Width += CInt(Ctrl.Width * RW)
            Ctrl.Height += CInt(Ctrl.Height * RH)
            Ctrl.Left += CInt(Ctrl.Left * RW)
            Ctrl.Top += CInt(Ctrl.Top * RH)
        Next

        CW = Me.Width
        CH = Me.Height
    End Sub
End Class