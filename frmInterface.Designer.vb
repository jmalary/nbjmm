<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmInterface
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmInterface))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btm_enregistrer = New System.Windows.Forms.ToolStripButton()
        Me.btm_fournisseur = New System.Windows.Forms.ToolStripButton()
        Me.btm_cheque = New System.Windows.Forms.ToolStripButton()
        Me.btm_paiement = New System.Windows.Forms.ToolStripButton()
        Me.btm_lister_cheque = New System.Windows.Forms.ToolStripButton()
        Me.btm_lister_paiement_direct = New System.Windows.Forms.ToolStripButton()
        Me.btm_lister_paiement_preautorise = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton7 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton8 = New System.Windows.Forms.ToolStripButton()
        Me.bouton_excel = New System.Windows.Forms.ToolStripButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SaveFileDialog2 = New System.Windows.Forms.SaveFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ToolStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(360, 63)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 24)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Lundi 1 mai 2017"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.DarkRed
        Me.Label2.Location = New System.Drawing.Point(325, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(244, 20)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Nouvelle version:  12 février 2019 "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Font = New System.Drawing.Font("Segoe UI", 20.0!)
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(50, 50)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btm_enregistrer, Me.btm_fournisseur, Me.btm_cheque, Me.btm_paiement, Me.btm_lister_cheque, Me.btm_lister_paiement_direct, Me.btm_lister_paiement_preautorise, Me.ToolStripButton7, Me.ToolStripButton8, Me.bouton_excel})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1362, 57)
        Me.ToolStrip1.TabIndex = 13
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btm_enregistrer
        '
        Me.btm_enregistrer.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_enregistrer.Image = Global.Comptable_NBJMM.My.Resources.Resources.save1
        Me.btm_enregistrer.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_enregistrer.Name = "btm_enregistrer"
        Me.btm_enregistrer.Size = New System.Drawing.Size(54, 54)
        Me.btm_enregistrer.Text = "Enregistrer"
        '
        'btm_fournisseur
        '
        Me.btm_fournisseur.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_fournisseur.Image = CType(resources.GetObject("btm_fournisseur.Image"), System.Drawing.Image)
        Me.btm_fournisseur.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_fournisseur.Name = "btm_fournisseur"
        Me.btm_fournisseur.Size = New System.Drawing.Size(54, 54)
        Me.btm_fournisseur.Text = "ajouter données"
        '
        'btm_cheque
        '
        Me.btm_cheque.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_cheque.Image = Global.Comptable_NBJMM.My.Resources.Resources.cheque
        Me.btm_cheque.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_cheque.Name = "btm_cheque"
        Me.btm_cheque.Size = New System.Drawing.Size(54, 54)
        Me.btm_cheque.Text = "ajouter une cheque"
        '
        'btm_paiement
        '
        Me.btm_paiement.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_paiement.Image = Global.Comptable_NBJMM.My.Resources.Resources.add_paiement1
        Me.btm_paiement.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_paiement.Name = "btm_paiement"
        Me.btm_paiement.Size = New System.Drawing.Size(54, 54)
        Me.btm_paiement.Text = "ajouter une transaction "
        '
        'btm_lister_cheque
        '
        Me.btm_lister_cheque.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_lister_cheque.Image = Global.Comptable_NBJMM.My.Resources.Resources.list_check
        Me.btm_lister_cheque.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_lister_cheque.Name = "btm_lister_cheque"
        Me.btm_lister_cheque.Size = New System.Drawing.Size(54, 54)
        Me.btm_lister_cheque.Text = "ToolStripButton4"
        Me.btm_lister_cheque.ToolTipText = "Liste des cheques"
        '
        'btm_lister_paiement_direct
        '
        Me.btm_lister_paiement_direct.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_lister_paiement_direct.Image = Global.Comptable_NBJMM.My.Resources.Resources.direct1
        Me.btm_lister_paiement_direct.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_lister_paiement_direct.Name = "btm_lister_paiement_direct"
        Me.btm_lister_paiement_direct.Size = New System.Drawing.Size(54, 54)
        Me.btm_lister_paiement_direct.Text = "ToolStripButton5"
        Me.btm_lister_paiement_direct.ToolTipText = "liste des paiements directs"
        '
        'btm_lister_paiement_preautorise
        '
        Me.btm_lister_paiement_preautorise.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btm_lister_paiement_preautorise.Image = Global.Comptable_NBJMM.My.Resources.Resources.preautotrise1
        Me.btm_lister_paiement_preautorise.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btm_lister_paiement_preautorise.Name = "btm_lister_paiement_preautorise"
        Me.btm_lister_paiement_preautorise.Size = New System.Drawing.Size(54, 54)
        Me.btm_lister_paiement_preautorise.Text = "ToolStripButton6"
        Me.btm_lister_paiement_preautorise.ToolTipText = "liste des paiements préautorisés"
        '
        'ToolStripButton7
        '
        Me.ToolStripButton7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton7.Image = Global.Comptable_NBJMM.My.Resources.Resources.gestion1
        Me.ToolStripButton7.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton7.Name = "ToolStripButton7"
        Me.ToolStripButton7.Size = New System.Drawing.Size(54, 54)
        Me.ToolStripButton7.Text = "liquidités 30 jours"
        '
        'ToolStripButton8
        '
        Me.ToolStripButton8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton8.Enabled = False
        Me.ToolStripButton8.Image = Global.Comptable_NBJMM.My.Resources.Resources.backup1
        Me.ToolStripButton8.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton8.Name = "ToolStripButton8"
        Me.ToolStripButton8.Size = New System.Drawing.Size(54, 54)
        Me.ToolStripButton8.Text = "Arrêtez le backup "
        '
        'bouton_excel
        '
        Me.bouton_excel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.bouton_excel.Image = CType(resources.GetObject("bouton_excel.Image"), System.Drawing.Image)
        Me.bouton_excel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.bouton_excel.Name = "bouton_excel"
        Me.bouton_excel.Size = New System.Drawing.Size(54, 54)
        Me.bouton_excel.Text = "budget de caisse"
        Me.bouton_excel.ToolTipText = "budget de caisse"
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.AutoSize = True
        Me.Panel1.BackColor = System.Drawing.Color.DimGray
        Me.Panel1.BackgroundImage = Global.Comptable_NBJMM.My.Resources.Resources.WEFS_couleurs_grand_test1
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Location = New System.Drawing.Point(12, 124)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 8, 3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(0, 3, 0, 0)
        Me.Panel1.Size = New System.Drawing.Size(1306, 485)
        Me.Panel1.TabIndex = 12
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowDrop = True
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(539, 203)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(240, 150)
        Me.DataGridView1.TabIndex = 319
        Me.DataGridView1.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(673, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'FrmInterface
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1362, 651)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FrmInterface"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "."
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btm_enregistrer As ToolStripButton
    Friend WithEvents btm_cheque As ToolStripButton
    Friend WithEvents btm_paiement As ToolStripButton
    Friend WithEvents btm_lister_cheque As ToolStripButton
    Friend WithEvents btm_lister_paiement_direct As ToolStripButton
    Friend WithEvents btm_lister_paiement_preautorise As ToolStripButton
    Friend WithEvents ToolStripButton7 As ToolStripButton
    Friend WithEvents ToolStripButton8 As ToolStripButton
    Friend WithEvents btm_fournisseur As ToolStripButton
    Friend WithEvents bouton_excel As ToolStripButton
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents SaveFileDialog2 As SaveFileDialog
    Friend WithEvents Button1 As Button
End Class
