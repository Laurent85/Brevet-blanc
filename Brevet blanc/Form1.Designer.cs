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
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.rdbDnb1 = new System.Windows.Forms.RadioButton();
            this.rdbDnb2 = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rdbAvecOral = new System.Windows.Forms.RadioButton();
            this.rdbSansOral = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnGénérerDiplomes
            // 
            this.BtnGénérerDiplomes.Location = new System.Drawing.Point(631, 208);
            this.BtnGénérerDiplomes.Name = "BtnGénérerDiplomes";
            this.BtnGénérerDiplomes.Size = new System.Drawing.Size(132, 23);
            this.BtnGénérerDiplomes.TabIndex = 0;
            this.BtnGénérerDiplomes.Text = "Générer les diplômes";
            this.BtnGénérerDiplomes.UseVisualStyleBackColor = true;
            this.BtnGénérerDiplomes.Click += new System.EventHandler(this.BtnGénérerDnb);
            // 
            // BtnGénérerStats
            // 
            this.BtnGénérerStats.Location = new System.Drawing.Point(631, 249);
            this.BtnGénérerStats.Name = "BtnGénérerStats";
            this.BtnGénérerStats.Size = new System.Drawing.Size(132, 23);
            this.BtnGénérerStats.TabIndex = 1;
            this.BtnGénérerStats.Text = "Générer les statistiques";
            this.BtnGénérerStats.UseVisualStyleBackColor = true;
            this.BtnGénérerStats.Click += new System.EventHandler(this.BtnGénérerStatistiques);
            // 
            // chkLb_Notes
            // 
            this.chkLb_Notes.CheckOnClick = true;
            this.chkLb_Notes.FormattingEnabled = true;
            this.chkLb_Notes.Location = new System.Drawing.Point(45, 208);
            this.chkLb_Notes.Name = "chkLb_Notes";
            this.chkLb_Notes.Size = new System.Drawing.Size(247, 244);
            this.chkLb_Notes.TabIndex = 2;
            // 
            // chkLb_Composantes
            // 
            this.chkLb_Composantes.CheckOnClick = true;
            this.chkLb_Composantes.FormattingEnabled = true;
            this.chkLb_Composantes.Location = new System.Drawing.Point(319, 208);
            this.chkLb_Composantes.Name = "chkLb_Composantes";
            this.chkLb_Composantes.Size = new System.Drawing.Size(247, 244);
            this.chkLb_Composantes.TabIndex = 3;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(215, 475);
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
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(62, 167);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "Epreuves écrites";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Blue;
            this.label2.Location = new System.Drawing.Point(341, 167);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "Contrôle continu";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Crimson;
            this.label3.Location = new System.Drawing.Point(297, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(189, 33);
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
            this.progressBar1.Location = new System.Drawing.Point(631, 339);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(132, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // lblCompteur
            // 
            this.lblCompteur.AutoSize = true;
            this.lblCompteur.Location = new System.Drawing.Point(631, 377);
            this.lblCompteur.Name = "lblCompteur";
            this.lblCompteur.Size = new System.Drawing.Size(0, 13);
            this.lblCompteur.TabIndex = 9;
            // 
            // lblClasse
            // 
            this.lblClasse.AutoSize = true;
            this.lblClasse.Location = new System.Drawing.Point(631, 304);
            this.lblClasse.Name = "lblClasse";
            this.lblClasse.Size = new System.Drawing.Size(0, 13);
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
            this.lblSource.Location = new System.Drawing.Point(188, 92);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(0, 13);
            this.lblSource.TabIndex = 13;
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(188, 122);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(0, 13);
            this.lblDestination.TabIndex = 14;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::Brevet_blanc.Properties.Resources.ED;
            this.pictureBox2.Location = new System.Drawing.Point(688, 9);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(100, 69);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 16;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Brevet_blanc.Properties.Resources.LOGO_COLLEGE_SAINTJACQUES_GEANT;
            this.pictureBox1.Location = new System.Drawing.Point(12, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 50);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 15;
            this.pictureBox1.TabStop = false;
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
            this.panel1.Controls.Add(this.rdbDnb1);
            this.panel1.Controls.Add(this.rdbDnb2);
            this.panel1.Location = new System.Drawing.Point(589, 144);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(94, 47);
            this.panel1.TabIndex = 19;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.rdbSansOral);
            this.panel2.Controls.Add(this.rdbAvecOral);
            this.panel2.Location = new System.Drawing.Point(691, 144);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(97, 47);
            this.panel2.TabIndex = 20;
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
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 519);
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
            this.Name = "Principal";
            this.Text = "Brevet blanc";
            this.Load += new System.EventHandler(this.Principal_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
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
    }
}

