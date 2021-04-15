namespace Brevet_blanc
{
    partial class Principal
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Principal));
            this.BtnGénérerDiplomes = new System.Windows.Forms.Button();
            this.BtnGénérerStats = new System.Windows.Forms.Button();
            this.chkLb_Notes = new System.Windows.Forms.CheckedListBox();
            this.chkLb_Composantes = new System.Windows.Forms.CheckedListBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.ThreadDiplomes = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCompteur = new System.Windows.Forms.Label();
            this.lblClasse = new System.Windows.Forms.Label();
            this.btnSource = new System.Windows.Forms.Button();
            this.btnDestination = new System.Windows.Forms.Button();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblDestination = new System.Windows.Forms.Label();
            this.rdbDnb1 = new System.Windows.Forms.RadioButton();
            this.rdbDnb2 = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rdbSansOral = new System.Windows.Forms.RadioButton();
            this.rdbAvecOral = new System.Windows.Forms.RadioButton();
            this.numDelta = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.ThreadStatistiques = new System.ComponentModel.BackgroundWorker();
            this.lblClasses = new System.Windows.Forms.Label();
            this.btnDossierRésultats = new System.Windows.Forms.Button();
            this.txb_dates_brevet = new System.Windows.Forms.TextBox();
            this.lblDatesBrevet = new System.Windows.Forms.Label();
            this.lblAnnéeScolaire = new System.Windows.Forms.Label();
            this.cbxAnnéeScolaire = new System.Windows.Forms.ComboBox();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDelta)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnGénérerDiplomes
            // 
            this.BtnGénérerDiplomes.BackColor = System.Drawing.Color.Transparent;
            this.BtnGénérerDiplomes.Location = new System.Drawing.Point(632, 347);
            this.BtnGénérerDiplomes.Name = "BtnGénérerDiplomes";
            this.BtnGénérerDiplomes.Size = new System.Drawing.Size(199, 23);
            this.BtnGénérerDiplomes.TabIndex = 0;
            this.BtnGénérerDiplomes.Text = "Générer les diplômes";
            this.BtnGénérerDiplomes.UseVisualStyleBackColor = false;
            this.BtnGénérerDiplomes.Click += new System.EventHandler(this.BtnGénérerDnb);
            // 
            // BtnGénérerStats
            // 
            this.BtnGénérerStats.Location = new System.Drawing.Point(632, 388);
            this.BtnGénérerStats.Name = "BtnGénérerStats";
            this.BtnGénérerStats.Size = new System.Drawing.Size(199, 23);
            this.BtnGénérerStats.TabIndex = 1;
            this.BtnGénérerStats.Text = "Générer les statistiques";
            this.BtnGénérerStats.UseVisualStyleBackColor = true;
            this.BtnGénérerStats.Click += new System.EventHandler(this.BtnGénérerStatistiques);
            // 
            // chkLb_Notes
            // 
            this.chkLb_Notes.BackColor = System.Drawing.Color.Linen;
            this.chkLb_Notes.CheckOnClick = true;
            this.chkLb_Notes.FormattingEnabled = true;
            this.chkLb_Notes.Location = new System.Drawing.Point(76, 251);
            this.chkLb_Notes.Name = "chkLb_Notes";
            this.chkLb_Notes.Size = new System.Drawing.Size(247, 244);
            this.chkLb_Notes.TabIndex = 2;
            this.chkLb_Notes.SelectedIndexChanged += new System.EventHandler(this.chkLb_Notes_SelectedIndexChanged);
            // 
            // chkLb_Composantes
            // 
            this.chkLb_Composantes.BackColor = System.Drawing.Color.Linen;
            this.chkLb_Composantes.CheckOnClick = true;
            this.chkLb_Composantes.FormattingEnabled = true;
            this.chkLb_Composantes.Location = new System.Drawing.Point(350, 251);
            this.chkLb_Composantes.Name = "chkLb_Composantes";
            this.chkLb_Composantes.Size = new System.Drawing.Size(247, 244);
            this.chkLb_Composantes.TabIndex = 3;
            this.chkLb_Composantes.SelectedIndexChanged += new System.EventHandler(this.chkLb_Notes_SelectedIndexChanged);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(246, 518);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(186, 23);
            this.button3.TabIndex = 4;
            this.button3.Text = "Supprimer les fichiers sélectionnés";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.BtnSuppressionFichiers);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(113, 210);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "Epreuves écrites";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Blue;
            this.label2.Location = new System.Drawing.Point(390, 210);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "Contrôle continu";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Brush Script MT", 48F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.OliveDrab;
            this.label3.Location = new System.Drawing.Point(336, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(285, 79);
            this.label3.TabIndex = 7;
            this.label3.Text = "Brevet blanc";
            // 
            // ThreadDiplomes
            // 
            this.ThreadDiplomes.WorkerReportsProgress = true;
            this.ThreadDiplomes.DoWork += new System.ComponentModel.DoWorkEventHandler(this.ThreadDiplomesMéthode);
            this.ThreadDiplomes.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.ThreadDiplomesProgression);
            this.ThreadDiplomes.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.ThreadDiplomesTerminé);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(632, 456);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(199, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // lblCompteur
            // 
            this.lblCompteur.AutoSize = true;
            this.lblCompteur.BackColor = System.Drawing.Color.Transparent;
            this.lblCompteur.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompteur.ForeColor = System.Drawing.Color.Red;
            this.lblCompteur.Location = new System.Drawing.Point(629, 500);
            this.lblCompteur.Name = "lblCompteur";
            this.lblCompteur.Size = new System.Drawing.Size(0, 13);
            this.lblCompteur.TabIndex = 9;
            // 
            // lblClasse
            // 
            this.lblClasse.AutoSize = true;
            this.lblClasse.BackColor = System.Drawing.Color.Transparent;
            this.lblClasse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblClasse.ForeColor = System.Drawing.Color.Red;
            this.lblClasse.Location = new System.Drawing.Point(629, 431);
            this.lblClasse.Name = "lblClasse";
            this.lblClasse.Size = new System.Drawing.Size(0, 16);
            this.lblClasse.TabIndex = 10;
            // 
            // btnSource
            // 
            this.btnSource.Location = new System.Drawing.Point(66, 87);
            this.btnSource.Name = "btnSource";
            this.btnSource.Size = new System.Drawing.Size(106, 23);
            this.btnSource.TabIndex = 11;
            this.btnSource.Text = "Fichiers source";
            this.btnSource.UseVisualStyleBackColor = true;
            this.btnSource.Click += new System.EventHandler(this.BtnChoisirSource);
            // 
            // btnDestination
            // 
            this.btnDestination.Location = new System.Drawing.Point(66, 117);
            this.btnDestination.Name = "btnDestination";
            this.btnDestination.Size = new System.Drawing.Size(106, 23);
            this.btnDestination.TabIndex = 12;
            this.btnDestination.Text = "Fichiers destination";
            this.btnDestination.UseVisualStyleBackColor = true;
            this.btnDestination.Click += new System.EventHandler(this.BtnChoisirDestination);
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.BackColor = System.Drawing.Color.Transparent;
            this.lblSource.ForeColor = System.Drawing.Color.DeepPink;
            this.lblSource.Location = new System.Drawing.Point(188, 92);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(0, 13);
            this.lblSource.TabIndex = 13;
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.BackColor = System.Drawing.Color.Transparent;
            this.lblDestination.ForeColor = System.Drawing.Color.DeepPink;
            this.lblDestination.Location = new System.Drawing.Point(188, 122);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(0, 13);
            this.lblDestination.TabIndex = 14;
            // 
            // rdbDnb1
            // 
            this.rdbDnb1.AutoSize = true;
            this.rdbDnb1.Location = new System.Drawing.Point(15, 7);
            this.rdbDnb1.Name = "rdbDnb1";
            this.rdbDnb1.Size = new System.Drawing.Size(54, 17);
            this.rdbDnb1.TabIndex = 17;
            this.rdbDnb1.TabStop = true;
            this.rdbDnb1.Text = "DNB1";
            this.rdbDnb1.UseVisualStyleBackColor = true;
            // 
            // rdbDnb2
            // 
            this.rdbDnb2.AutoSize = true;
            this.rdbDnb2.Location = new System.Drawing.Point(15, 27);
            this.rdbDnb2.Name = "rdbDnb2";
            this.rdbDnb2.Size = new System.Drawing.Size(54, 17);
            this.rdbDnb2.TabIndex = 18;
            this.rdbDnb2.TabStop = true;
            this.rdbDnb2.Text = "DNB2";
            this.rdbDnb2.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Transparent;
            this.panel1.Controls.Add(this.rdbDnb1);
            this.panel1.Controls.Add(this.rdbDnb2);
            this.panel1.Location = new System.Drawing.Point(632, 251);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(94, 47);
            this.panel1.TabIndex = 19;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Transparent;
            this.panel2.Controls.Add(this.rdbSansOral);
            this.panel2.Controls.Add(this.rdbAvecOral);
            this.panel2.Location = new System.Drawing.Point(734, 251);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(97, 47);
            this.panel2.TabIndex = 20;
            // 
            // rdbSansOral
            // 
            this.rdbSansOral.AutoSize = true;
            this.rdbSansOral.Location = new System.Drawing.Point(12, 7);
            this.rdbSansOral.Name = "rdbSansOral";
            this.rdbSansOral.Size = new System.Drawing.Size(69, 17);
            this.rdbSansOral.TabIndex = 1;
            this.rdbSansOral.TabStop = true;
            this.rdbSansOral.Text = "Sans oral";
            this.rdbSansOral.UseVisualStyleBackColor = true;
            this.rdbSansOral.CheckedChanged += new System.EventHandler(this.rdbSansOral_CheckedChanged);
            // 
            // rdbAvecOral
            // 
            this.rdbAvecOral.AutoSize = true;
            this.rdbAvecOral.Location = new System.Drawing.Point(12, 27);
            this.rdbAvecOral.Name = "rdbAvecOral";
            this.rdbAvecOral.Size = new System.Drawing.Size(70, 17);
            this.rdbAvecOral.TabIndex = 0;
            this.rdbAvecOral.TabStop = true;
            this.rdbAvecOral.Text = "Avec oral";
            this.rdbAvecOral.UseVisualStyleBackColor = true;
            this.rdbAvecOral.CheckedChanged += new System.EventHandler(this.rdbAvecOral_CheckedChanged);
            // 
            // numDelta
            // 
            this.numDelta.BackColor = System.Drawing.Color.Linen;
            this.numDelta.Location = new System.Drawing.Point(734, 214);
            this.numDelta.Name = "numDelta";
            this.numDelta.Size = new System.Drawing.Size(47, 20);
            this.numDelta.TabIndex = 21;
            this.numDelta.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(677, 216);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 23;
            this.label4.Text = "Delta";
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox3.Image = global::Brevet_blanc.Properties.Resources.Sigle1;
            this.pictureBox3.Location = new System.Drawing.Point(709, 216);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(20, 13);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox3.TabIndex = 22;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Image = global::Brevet_blanc.Properties.Resources.ED1;
            this.pictureBox2.Location = new System.Drawing.Point(800, 31);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(69, 57);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 16;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = global::Brevet_blanc.Properties.Resources.LOGO1;
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(109, 61);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 15;
            this.pictureBox1.TabStop = false;
            // 
            // ThreadStatistiques
            // 
            this.ThreadStatistiques.WorkerReportsProgress = true;
            this.ThreadStatistiques.DoWork += new System.ComponentModel.DoWorkEventHandler(this.ThreadStatistiquesMéthode);
            this.ThreadStatistiques.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.ThreadStatistiquesProgression);
            this.ThreadStatistiques.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.ThreadStatistiquesTerminé);
            // 
            // lblClasses
            // 
            this.lblClasses.AutoSize = true;
            this.lblClasses.BackColor = System.Drawing.Color.Transparent;
            this.lblClasses.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblClasses.ForeColor = System.Drawing.Color.Red;
            this.lblClasses.Location = new System.Drawing.Point(644, 319);
            this.lblClasses.Name = "lblClasses";
            this.lblClasses.Size = new System.Drawing.Size(0, 16);
            this.lblClasses.TabIndex = 24;
            // 
            // btnDossierRésultats
            // 
            this.btnDossierRésultats.Location = new System.Drawing.Point(246, 547);
            this.btnDossierRésultats.Name = "btnDossierRésultats";
            this.btnDossierRésultats.Size = new System.Drawing.Size(186, 23);
            this.btnDossierRésultats.TabIndex = 25;
            this.btnDossierRésultats.Text = "Ouvrir le dossier des résultats";
            this.btnDossierRésultats.UseVisualStyleBackColor = true;
            this.btnDossierRésultats.Click += new System.EventHandler(this.btnDossierRésultats_Click);
            // 
            // txb_dates_brevet
            // 
            this.txb_dates_brevet.BackColor = System.Drawing.Color.Linen;
            this.txb_dates_brevet.Location = new System.Drawing.Point(191, 155);
            this.txb_dates_brevet.Name = "txb_dates_brevet";
            this.txb_dates_brevet.Size = new System.Drawing.Size(132, 20);
            this.txb_dates_brevet.TabIndex = 26;
            this.txb_dates_brevet.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // lblDatesBrevet
            // 
            this.lblDatesBrevet.AutoSize = true;
            this.lblDatesBrevet.BackColor = System.Drawing.Color.Transparent;
            this.lblDatesBrevet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDatesBrevet.Location = new System.Drawing.Point(68, 158);
            this.lblDatesBrevet.Name = "lblDatesBrevet";
            this.lblDatesBrevet.Size = new System.Drawing.Size(98, 13);
            this.lblDatesBrevet.TabIndex = 27;
            this.lblDatesBrevet.Text = "Dates du brevet";
            // 
            // lblAnnéeScolaire
            // 
            this.lblAnnéeScolaire.AutoSize = true;
            this.lblAnnéeScolaire.BackColor = System.Drawing.Color.Transparent;
            this.lblAnnéeScolaire.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnnéeScolaire.Location = new System.Drawing.Point(347, 158);
            this.lblAnnéeScolaire.Name = "lblAnnéeScolaire";
            this.lblAnnéeScolaire.Size = new System.Drawing.Size(91, 13);
            this.lblAnnéeScolaire.TabIndex = 28;
            this.lblAnnéeScolaire.Text = "Année scolaire";
            // 
            // cbxAnnéeScolaire
            // 
            this.cbxAnnéeScolaire.BackColor = System.Drawing.Color.Linen;
            this.cbxAnnéeScolaire.FormattingEnabled = true;
            this.cbxAnnéeScolaire.Items.AddRange(new object[] {
            "2020-2021",
            "2021-2022",
            "2022-2023",
            "2023-2024",
            "2024-2025",
            "2025-2026",
            "2026-2027",
            "2027-2028",
            "2028-2029",
            "2029-2030"});
            this.cbxAnnéeScolaire.Location = new System.Drawing.Point(457, 155);
            this.cbxAnnéeScolaire.Name = "cbxAnnéeScolaire";
            this.cbxAnnéeScolaire.Size = new System.Drawing.Size(121, 21);
            this.cbxAnnéeScolaire.TabIndex = 29;
            this.cbxAnnéeScolaire.SelectedIndexChanged += new System.EventHandler(this.cbxAnnéeScolaire_SelectedIndexChanged);
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImage = global::Brevet_blanc.Properties.Resources.Fond;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(915, 592);
            this.Controls.Add(this.cbxAnnéeScolaire);
            this.Controls.Add(this.lblAnnéeScolaire);
            this.Controls.Add(this.lblDatesBrevet);
            this.Controls.Add(this.txb_dates_brevet);
            this.Controls.Add(this.btnDossierRésultats);
            this.Controls.Add(this.lblClasses);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.numDelta);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblDestination);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.btnDestination);
            this.Controls.Add(this.btnSource);
            this.Controls.Add(this.lblClasse);
            this.Controls.Add(this.lblCompteur);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.chkLb_Composantes);
            this.Controls.Add(this.chkLb_Notes);
            this.Controls.Add(this.BtnGénérerStats);
            this.Controls.Add(this.BtnGénérerDiplomes);
            this.Controls.Add(this.panel1);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Principal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Brevet blanc";
            this.Load += new System.EventHandler(this.Principal_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDelta)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnGénérerDiplomes;
        private System.Windows.Forms.Button BtnGénérerStats;
        private System.Windows.Forms.CheckedListBox chkLb_Notes;
        private System.Windows.Forms.CheckedListBox chkLb_Composantes;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.ComponentModel.BackgroundWorker ThreadDiplomes;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCompteur;
        private System.Windows.Forms.Label lblClasse;
        private System.Windows.Forms.Button btnSource;
        private System.Windows.Forms.Button btnDestination;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.RadioButton rdbDnb1;
        private System.Windows.Forms.RadioButton rdbDnb2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rdbSansOral;
        private System.Windows.Forms.RadioButton rdbAvecOral;
        private System.Windows.Forms.NumericUpDown numDelta;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label4;
        private System.ComponentModel.BackgroundWorker ThreadStatistiques;
        private System.Windows.Forms.Label lblClasses;
        private System.Windows.Forms.Button btnDossierRésultats;
        private System.Windows.Forms.TextBox txb_dates_brevet;
        private System.Windows.Forms.Label lblDatesBrevet;
        private System.Windows.Forms.Label lblAnnéeScolaire;
        private System.Windows.Forms.ComboBox cbxAnnéeScolaire;
    }
}

