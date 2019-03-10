Imports System.Data.OleDb
Imports Comptable_NBJMM.GlobalVariables
Imports System.Globalization


Public Class frmChoisir_date

    Dim dynamicDTP As New DateTimePicker()
    Dim dynamicDTP1 As New DateTimePicker()
    Dim dynamicDTP2 As New DateTimePicker()
    Dim dynamicDTP3 As New DateTimePicker()
    Dim dynamicDTP4 As New DateTimePicker()
    Dim dynamicDTP5 As New DateTimePicker()
    Dim dynamicDTP6 As New DateTimePicker()
    Dim dynamicDTP7 As New DateTimePicker()
    Dim dynamicDTP8 As New DateTimePicker()
    Dim dynamicDTP9 As New DateTimePicker()
    Dim dynamicDTP10 As New DateTimePicker()
    Dim dynamicDTP11 As New DateTimePicker()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        proc_afficher_nbr_bottom()


        Dim value As Integer = choisir_nbr_payer + 1

        Select Case value
            Case 2

                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)

            Case 3
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
            Case 4
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
            Case 5
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
            Case 6
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
            Case 7
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
            Case 8
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
                list_datepicker.Add(dynamicDTP7.Value.Date)
            Case 9
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
                list_datepicker.Add(dynamicDTP7.Value.Date)
                list_datepicker.Add(dynamicDTP8.Value.Date)

            Case 10
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
                list_datepicker.Add(dynamicDTP7.Value.Date)
                list_datepicker.Add(dynamicDTP8.Value.Date)
                list_datepicker.Add(dynamicDTP9.Value.Date)
            Case 11
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
                list_datepicker.Add(dynamicDTP7.Value.Date)
                list_datepicker.Add(dynamicDTP8.Value.Date)
                list_datepicker.Add(dynamicDTP9.Value.Date)
                list_datepicker.Add(dynamicDTP10.Value.Date)
            Case 12
                list_datepicker.Add(dynamicDTP.Value.Date)
                list_datepicker.Add(dynamicDTP1.Value.Date)
                list_datepicker.Add(dynamicDTP2.Value.Date)
                list_datepicker.Add(dynamicDTP3.Value.Date)
                list_datepicker.Add(dynamicDTP4.Value.Date)
                list_datepicker.Add(dynamicDTP5.Value.Date)
                list_datepicker.Add(dynamicDTP6.Value.Date)
                list_datepicker.Add(dynamicDTP7.Value.Date)
                list_datepicker.Add(dynamicDTP8.Value.Date)
                list_datepicker.Add(dynamicDTP9.Value.Date)
                list_datepicker.Add(dynamicDTP10.Value.Date)
                list_datepicker.Add(dynamicDTP11.Value.Date)


        End Select

        'check_menu = True
        Me.Close()
        'Me.WindowState = FormWindowState.Maximized
        'frmAfficherListeCheque.MdiParent = Form_windows
        'frmAfficherListeCheque.Show()
        ''frmLoading.Show()

        'frmAfficherListeCheque.Show()

        Dim fc As New frmPaiement()
        fc.TopLevel = False
        'f.WindowState = FormWindowState.Maximized
        fc.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        fc.Visible = True
        fc.Location = New Point(1, 0)
        fc.RadioButton4.Checked = True

        fc.ComboBox2.SelectedIndex = choisir_fournisseur
        fc.TextBox1.Text = choisir_montant
        fc.lister_nbr_paiement.SelectedIndex = choisir_nbr_payer
        fc.btm_choisir_date.Text = "Mois à choisir  " + aff_nbr_bottom.ToString() + "/" + change + ""
        'frmInterface.Panel1.Remove(fc)

        FrmInterface.Panel1.Controls.Add(fc)
    End Sub

    Private Sub frmChoisir_date_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim list_label As New ArrayList
        Dim list_date As New ArrayList



        Dim aff, aff1, aff2, aff3, aff4, aff5, aff6, aff7, aff8, aff9, aff10, aff11 As New Label
        Dim myfont As New Font("Sans Serif", 12, FontStyle.Regular)

        aff.Font = myfont
        aff.Text = "Paiement 1 "
        aff.Location = New Point(3, 50)
        '------------------------------------------
        aff1.Font = myfont
        aff1.Text = "Paiement 2 "
        aff1.Location = New Point(3, 80)
        '------------------------------------------
        aff2.Font = myfont
        aff2.Text = "Paiement 3 "
        aff2.Location = New Point(3, 110)
        '------------------------------------------
        aff3.Font = myfont
        aff3.Text = "Paiement 4 "
        aff3.Location = New Point(3, 140)
        '------------------------------------------
        aff4.Font = myfont
        aff4.Text = "Paiement 5 "
        aff4.Location = New Point(3, 170)
        '------------------------------------------
        aff5.Font = myfont
        aff5.Text = "Paiement 6 "
        aff5.Location = New Point(3, 200)
        '------------------------------------------
        aff6.Font = myfont
        aff6.Text = "Paiement 7 "
        aff6.Location = New Point(3, 230)
        '------------------------------------------
        aff7.Font = myfont
        aff7.Text = "Paiement 8 "
        aff7.Location = New Point(3, 260)
        '------------------------------------------
        aff8.Font = myfont
        aff8.Text = "Paiement 9 "
        aff8.Location = New Point(3, 290)
        '------------------------------------------
        aff9.Font = myfont
        aff9.Text = "Paiement 10 "
        aff9.Location = New Point(3, 320)
        '------------------------------------------
        aff10.Font = myfont
        aff10.Text = "Paiement 11 "
        aff10.Location = New Point(3, 350)
        '------------------------------------------
        aff11.Font = myfont
        aff11.Text = "Paiement 12 "
        aff11.Location = New Point(3, 380)
        '------------------------------------------

        dynamicDTP.Location = New Point(140, 50)
        dynamicDTP1.Location = New Point(140, 80)
        dynamicDTP2.Location = New Point(140, 110)
        dynamicDTP3.Location = New Point(140, 140)
        dynamicDTP4.Location = New Point(140, 170)
        dynamicDTP5.Location = New Point(140, 200)
        dynamicDTP6.Location = New Point(140, 230)
        dynamicDTP7.Location = New Point(140, 260)
        dynamicDTP8.Location = New Point(140, 290)
        dynamicDTP9.Location = New Point(140, 320)
        dynamicDTP10.Location = New Point(140, 350)
        dynamicDTP11.Location = New Point(140, 380)




        Dim value As Integer = choisir_nbr_payer + 1

        Select Case value
            Case 2
                Controls.Add(aff)
                Controls.Add(aff1)

            Case 3
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
            Case 4
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
            Case 5
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
            Case 6
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
            Case 7
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
            Case 8
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
                Controls.Add(aff7)
            Case 9
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
                Controls.Add(aff7)
                Controls.Add(aff8)

            Case 10
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
                Controls.Add(aff7)
                Controls.Add(aff8)
                Controls.Add(aff9)
            Case 11
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
                Controls.Add(aff7)
                Controls.Add(aff8)
                Controls.Add(aff9)
                Controls.Add(aff10)
            Case 12
                Controls.Add(aff)
                Controls.Add(aff1)
                Controls.Add(aff2)
                Controls.Add(aff3)
                Controls.Add(aff4)
                Controls.Add(aff5)
                Controls.Add(aff6)
                Controls.Add(aff7)
                Controls.Add(aff8)
                Controls.Add(aff9)
                Controls.Add(aff10)
                Controls.Add(aff11)


        End Select




        Select Case value
            Case 2
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)

            Case 3
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
            Case 4
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
            Case 5
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
            Case 6
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
            Case 7
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
            Case 8
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
                Controls.Add(dynamicDTP7)
            Case 9
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
                Controls.Add(dynamicDTP7)
                Controls.Add(dynamicDTP8)

            Case 10
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
                Controls.Add(dynamicDTP7)
                Controls.Add(dynamicDTP8)
                Controls.Add(dynamicDTP9)
            Case 11
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
                Controls.Add(dynamicDTP7)
                Controls.Add(dynamicDTP8)
                Controls.Add(dynamicDTP9)
                Controls.Add(dynamicDTP10)
            Case 12
                Controls.Add(dynamicDTP)
                Controls.Add(dynamicDTP1)
                Controls.Add(dynamicDTP2)
                Controls.Add(dynamicDTP3)
                Controls.Add(dynamicDTP4)
                Controls.Add(dynamicDTP5)
                Controls.Add(dynamicDTP6)
                Controls.Add(dynamicDTP7)
                Controls.Add(dynamicDTP8)
                Controls.Add(dynamicDTP9)
                Controls.Add(dynamicDTP10)
                Controls.Add(dynamicDTP11)


        End Select

        Select Case value
            Case 2
                Button1.Location = New Point(167, 110)
            Case 3
                Button1.Location = New Point(167, 140)
            Case 4
                Button1.Location = New Point(167, 170)
            Case 5
                Button1.Location = New Point(167, 200)
            Case 6
                Button1.Location = New Point(167, 230)
            Case 7
                Button1.Location = New Point(167, 260)
            Case 8
                Button1.Location = New Point(167, 290)
            Case 9
                Button1.Location = New Point(167, 320)

            Case 10
                Button1.Location = New Point(167, 350)
            Case 11
                Button1.Location = New Point(167, 380)
            Case 12
                Button1.Location = New Point(167, 410)


        End Select

    End Sub

    Private Sub frmChoisir_date_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
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
        y = CInt(30)

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(x, y)

    End Sub

    Private Sub proc_afficher_nbr_bottom()

        Dim avant As Date = Date.Now

        aff_nbr_bottom = 0

        If avant < dynamicDTP.Value.Date Then
            aff_nbr_bottom += 1
        End If

        If avant < dynamicDTP1.Value.Date Then
            aff_nbr_bottom += 1
        End If

        If avant < dynamicDTP2.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP3.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP4.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP5.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP6.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP7.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP8.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP9.Value.Date Then
            aff_nbr_bottom += 1
        End If
        If avant < dynamicDTP10.Value.Date Then
            aff_nbr_bottom += 1
        End If

        If avant < dynamicDTP11.Value.Date Then
            aff_nbr_bottom += 1
        End If


    End Sub
End Class