<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDetailsFinancement
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lbl_fournisseur = New System.Windows.Forms.Label()
        Me.lbl_date_paiement = New System.Windows.Forms.Label()
        Me.lbl_date_echances = New System.Windows.Forms.Label()
        Me.lbl_mode = New System.Windows.Forms.Label()
        Me.lbl_montant = New System.Windows.Forms.Label()
        Me.lbl_numero = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbl_nbr_transaction = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0!)
        Me.Label8.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label8.Location = New System.Drawing.Point(190, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(322, 31)
        Me.Label8.TabIndex = 304
        Me.Label8.Text = "La détails de financement"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label1.Location = New System.Drawing.Point(12, 94)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(138, 25)
        Me.Label1.TabIndex = 305
        Me.Label1.Text = "Fournisseur: "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label2.Location = New System.Drawing.Point(12, 130)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(193, 25)
        Me.Label2.TabIndex = 306
        Me.Label2.Text = "Date de paiement: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label3.Location = New System.Drawing.Point(568, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(190, 25)
        Me.Label3.TabIndex = 307
        Me.Label3.Text = "Mode de paiement"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(568, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 25)
        Me.Label4.TabIndex = 308
        Me.Label4.Text = "Montant:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label5.Location = New System.Drawing.Point(15, 173)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(172, 25)
        Me.Label5.TabIndex = 309
        Me.Label5.Text = "Date d'échances"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label6.Location = New System.Drawing.Point(15, 215)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(158, 25)
        Me.Label6.TabIndex = 310
        Me.Label6.Text = "Liste des dates"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(591, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(110, 31)
        Me.Button1.TabIndex = 312
        Me.Button1.Text = "Retour"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader9, Me.ColumnHeader1})
        Me.ListView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(17, 244)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(515, 242)
        Me.ListView1.TabIndex = 313
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "No"
        Me.ColumnHeader2.Width = 43
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Date "
        Me.ColumnHeader9.Width = 300
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "montant"
        Me.ColumnHeader1.Width = 200
        '
        'lbl_fournisseur
        '
        Me.lbl_fournisseur.AutoSize = True
        Me.lbl_fournisseur.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_fournisseur.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_fournisseur.Location = New System.Drawing.Point(156, 94)
        Me.lbl_fournisseur.Name = "lbl_fournisseur"
        Me.lbl_fournisseur.Size = New System.Drawing.Size(72, 24)
        Me.lbl_fournisseur.TabIndex = 314
        Me.lbl_fournisseur.Text = "Label7"
        '
        'lbl_date_paiement
        '
        Me.lbl_date_paiement.AutoSize = True
        Me.lbl_date_paiement.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_date_paiement.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_date_paiement.Location = New System.Drawing.Point(211, 130)
        Me.lbl_date_paiement.Name = "lbl_date_paiement"
        Me.lbl_date_paiement.Size = New System.Drawing.Size(72, 24)
        Me.lbl_date_paiement.TabIndex = 315
        Me.lbl_date_paiement.Text = "Label9"
        '
        'lbl_date_echances
        '
        Me.lbl_date_echances.AutoSize = True
        Me.lbl_date_echances.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_date_echances.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_date_echances.Location = New System.Drawing.Point(200, 174)
        Me.lbl_date_echances.Name = "lbl_date_echances"
        Me.lbl_date_echances.Size = New System.Drawing.Size(83, 24)
        Me.lbl_date_echances.TabIndex = 316
        Me.lbl_date_echances.Text = "Label10"
        '
        'lbl_mode
        '
        Me.lbl_mode.AutoSize = True
        Me.lbl_mode.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_mode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_mode.Location = New System.Drawing.Point(786, 94)
        Me.lbl_mode.Name = "lbl_mode"
        Me.lbl_mode.Size = New System.Drawing.Size(83, 24)
        Me.lbl_mode.TabIndex = 317
        Me.lbl_mode.Text = "Label11"
        '
        'lbl_montant
        '
        Me.lbl_montant.AutoSize = True
        Me.lbl_montant.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_montant.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_montant.Location = New System.Drawing.Point(670, 130)
        Me.lbl_montant.Name = "lbl_montant"
        Me.lbl_montant.Size = New System.Drawing.Size(83, 24)
        Me.lbl_montant.TabIndex = 318
        Me.lbl_montant.Text = "Label12"
        '
        'lbl_numero
        '
        Me.lbl_numero.AutoSize = True
        Me.lbl_numero.Location = New System.Drawing.Point(625, 227)
        Me.lbl_numero.Name = "lbl_numero"
        Me.lbl_numero.Size = New System.Drawing.Size(39, 13)
        Me.lbl_numero.TabIndex = 319
        Me.lbl_numero.Text = "Label7"
        Me.lbl_numero.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label7.Location = New System.Drawing.Point(555, 174)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(246, 25)
        Me.Label7.TabIndex = 320
        Me.Label7.Text = "Nombre de transactions:"
        '
        'lbl_nbr_transaction
        '
        Me.lbl_nbr_transaction.AutoSize = True
        Me.lbl_nbr_transaction.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_nbr_transaction.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_nbr_transaction.Location = New System.Drawing.Point(807, 173)
        Me.lbl_nbr_transaction.Name = "lbl_nbr_transaction"
        Me.lbl_nbr_transaction.Size = New System.Drawing.Size(83, 24)
        Me.lbl_nbr_transaction.TabIndex = 321
        Me.lbl_nbr_transaction.Text = "Label10"
        '
        'frmDetailsFinancement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1354, 652)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl_nbr_transaction)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lbl_numero)
        Me.Controls.Add(Me.lbl_montant)
        Me.Controls.Add(Me.lbl_mode)
        Me.Controls.Add(Me.lbl_date_echances)
        Me.Controls.Add(Me.lbl_date_paiement)
        Me.Controls.Add(Me.lbl_fournisseur)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label8)
        Me.Name = "frmDetailsFinancement"
        Me.Text = "frmDetailsFinancement"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label8 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents ColumnHeader9 As ColumnHeader
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents lbl_fournisseur As Label
    Friend WithEvents lbl_date_paiement As Label
    Friend WithEvents lbl_date_echances As Label
    Friend WithEvents lbl_mode As Label
    Friend WithEvents lbl_montant As Label
    Public WithEvents lbl_numero As Label
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents Label7 As Label
    Friend WithEvents lbl_nbr_transaction As Label
End Class
