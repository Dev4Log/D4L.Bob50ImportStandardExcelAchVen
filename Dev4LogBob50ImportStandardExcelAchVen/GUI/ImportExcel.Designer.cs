namespace Dev4LogBob50ImportStandardExcelAchVen.GUI
{
    partial class ImportsExcel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportsExcel));
            ProgressBar1 = new System.Windows.Forms.ProgressBar();
            GroupBox1 = new System.Windows.Forms.GroupBox();
            cb_NumAutoVentes = new System.Windows.Forms.CheckBox();
            RechDossier = new System.Windows.Forms.Button();
            FichierImport = new System.Windows.Forms.TextBox();
            Label1 = new System.Windows.Forms.Label();
            GroupBox2 = new System.Windows.Forms.GroupBox();
            InfosImport = new System.Windows.Forms.TextBox();
            ButtExecuter = new System.Windows.Forms.Button();
            ButtQuitter = new System.Windows.Forms.Button();
            ButtExecuterMulti = new System.Windows.Forms.Button();
            statusStrip1 = new System.Windows.Forms.StatusStrip();
            toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuLogs = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            toolStripMenuLicence = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            toolStripMenuItemDev4LogMail = new System.Windows.Forms.ToolStripMenuItem();
            GroupBox1.SuspendLayout();
            GroupBox2.SuspendLayout();
            statusStrip1.SuspendLayout();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // ProgressBar1
            // 
            ProgressBar1.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            ProgressBar1.Location = new System.Drawing.Point(74, 695);
            ProgressBar1.Margin = new System.Windows.Forms.Padding(4);
            ProgressBar1.Name = "ProgressBar1";
            ProgressBar1.Size = new System.Drawing.Size(542, 32);
            ProgressBar1.TabIndex = 15;
            ProgressBar1.Visible = false;
            // 
            // GroupBox1
            // 
            GroupBox1.Controls.Add(cb_NumAutoVentes);
            GroupBox1.Controls.Add(RechDossier);
            GroupBox1.Controls.Add(FichierImport);
            GroupBox1.Controls.Add(Label1);
            GroupBox1.Location = new System.Drawing.Point(7, 29);
            GroupBox1.Margin = new System.Windows.Forms.Padding(5);
            GroupBox1.Name = "GroupBox1";
            GroupBox1.Padding = new System.Windows.Forms.Padding(5);
            GroupBox1.Size = new System.Drawing.Size(716, 110);
            GroupBox1.TabIndex = 9;
            GroupBox1.TabStop = false;
            // 
            // cb_NumAutoVentes
            // 
            cb_NumAutoVentes.AutoSize = true;
            cb_NumAutoVentes.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            cb_NumAutoVentes.Location = new System.Drawing.Point(116, 72);
            cb_NumAutoVentes.Margin = new System.Windows.Forms.Padding(4);
            cb_NumAutoVentes.Name = "cb_NumAutoVentes";
            cb_NumAutoVentes.Size = new System.Drawing.Size(244, 20);
            cb_NumAutoVentes.TabIndex = 3;
            cb_NumAutoVentes.Text = "Numérotation automatique des ventes";
            cb_NumAutoVentes.UseVisualStyleBackColor = true;
            // 
            // RechDossier
            // 
            RechDossier.BackgroundImage = (System.Drawing.Image)resources.GetObject("RechDossier.BackgroundImage");
            RechDossier.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            RechDossier.FlatAppearance.BorderSize = 0;
            RechDossier.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            RechDossier.Location = new System.Drawing.Point(649, 29);
            RechDossier.Margin = new System.Windows.Forms.Padding(4);
            RechDossier.Name = "RechDossier";
            RechDossier.Size = new System.Drawing.Size(30, 30);
            RechDossier.TabIndex = 2;
            RechDossier.UseVisualStyleBackColor = true;
            RechDossier.Click += RechDossier_Click;
            // 
            // FichierImport
            // 
            FichierImport.Location = new System.Drawing.Point(119, 34);
            FichierImport.Margin = new System.Windows.Forms.Padding(4);
            FichierImport.Name = "FichierImport";
            FichierImport.Size = new System.Drawing.Size(528, 22);
            FichierImport.TabIndex = 1;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Location = new System.Drawing.Point(8, 38);
            Label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(95, 16);
            Label1.TabIndex = 0;
            Label1.Text = "Fichier d'import";
            // 
            // GroupBox2
            // 
            GroupBox2.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            GroupBox2.Controls.Add(InfosImport);
            GroupBox2.Location = new System.Drawing.Point(7, 148);
            GroupBox2.Margin = new System.Windows.Forms.Padding(4);
            GroupBox2.Name = "GroupBox2";
            GroupBox2.Padding = new System.Windows.Forms.Padding(4);
            GroupBox2.Size = new System.Drawing.Size(716, 533);
            GroupBox2.TabIndex = 12;
            GroupBox2.TabStop = false;
            GroupBox2.Text = "Informations d'import";
            // 
            // InfosImport
            // 
            InfosImport.Dock = System.Windows.Forms.DockStyle.Fill;
            InfosImport.Location = new System.Drawing.Point(4, 19);
            InfosImport.Margin = new System.Windows.Forms.Padding(4);
            InfosImport.Multiline = true;
            InfosImport.Name = "InfosImport";
            InfosImport.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            InfosImport.Size = new System.Drawing.Size(708, 510);
            InfosImport.TabIndex = 0;
            // 
            // ButtExecuter
            // 
            ButtExecuter.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            ButtExecuter.BackgroundImage = (System.Drawing.Image)resources.GetObject("ButtExecuter.BackgroundImage");
            ButtExecuter.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            ButtExecuter.FlatAppearance.BorderSize = 0;
            ButtExecuter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            ButtExecuter.Location = new System.Drawing.Point(643, 695);
            ButtExecuter.Margin = new System.Windows.Forms.Padding(4);
            ButtExecuter.Name = "ButtExecuter";
            ButtExecuter.Size = new System.Drawing.Size(30, 30);
            ButtExecuter.TabIndex = 13;
            ButtExecuter.UseVisualStyleBackColor = true;
            ButtExecuter.Click += ButtExecuter_Click;
            // 
            // ButtQuitter
            // 
            ButtQuitter.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            ButtQuitter.BackgroundImage = (System.Drawing.Image)resources.GetObject("ButtQuitter.BackgroundImage");
            ButtQuitter.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            ButtQuitter.FlatAppearance.BorderSize = 0;
            ButtQuitter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            ButtQuitter.Location = new System.Drawing.Point(686, 695);
            ButtQuitter.Margin = new System.Windows.Forms.Padding(4);
            ButtQuitter.Name = "ButtQuitter";
            ButtQuitter.Size = new System.Drawing.Size(30, 30);
            ButtQuitter.TabIndex = 14;
            ButtQuitter.UseVisualStyleBackColor = true;
            ButtQuitter.Click += ButtQuitter_Click;
            // 
            // ButtExecuterMulti
            // 
            ButtExecuterMulti.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
            ButtExecuterMulti.BackgroundImage = (System.Drawing.Image)resources.GetObject("ButtExecuterMulti.BackgroundImage");
            ButtExecuterMulti.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            ButtExecuterMulti.FlatAppearance.BorderSize = 0;
            ButtExecuterMulti.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            ButtExecuterMulti.Location = new System.Drawing.Point(15, 695);
            ButtExecuterMulti.Margin = new System.Windows.Forms.Padding(4);
            ButtExecuterMulti.Name = "ButtExecuterMulti";
            ButtExecuterMulti.Size = new System.Drawing.Size(30, 30);
            ButtExecuterMulti.TabIndex = 26;
            ButtExecuterMulti.UseVisualStyleBackColor = true;
            ButtExecuterMulti.Visible = false;
            ButtExecuterMulti.Click += ButtExecuterMulti_Click;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripStatusLabel1 });
            statusStrip1.Location = new System.Drawing.Point(0, 739);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new System.Drawing.Size(729, 22);
            statusStrip1.TabIndex = 27;
            statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripMenuItem1 });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(729, 24);
            menuStrip1.TabIndex = 28;
            menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripMenuLogs, toolStripSeparator1, toolStripMenuLicence, toolStripSeparator2, toolStripMenuItemDev4LogMail });
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new System.Drawing.Size(24, 20);
            toolStripMenuItem1.Text = "?";
            // 
            // toolStripMenuLogs
            // 
            toolStripMenuLogs.Name = "toolStripMenuLogs";
            toolStripMenuLogs.Size = new System.Drawing.Size(180, 22);
            toolStripMenuLogs.Text = "Liste des logs";
            toolStripMenuLogs.Click += toolStripMenuLogs_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new System.Drawing.Size(177, 6);
            // 
            // toolStripMenuLicence
            // 
            toolStripMenuLicence.Name = "toolStripMenuLicence";
            toolStripMenuLicence.Size = new System.Drawing.Size(180, 22);
            toolStripMenuLicence.Text = "Licence";
            toolStripMenuLicence.Click += toolStripMenuLicence_Click;
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new System.Drawing.Size(177, 6);
            // 
            // toolStripMenuItemDev4LogMail
            // 
            toolStripMenuItemDev4LogMail.Name = "toolStripMenuItemDev4LogMail";
            toolStripMenuItemDev4LogMail.Size = new System.Drawing.Size(180, 22);
            toolStripMenuItemDev4LogMail.Text = "DevLog Mail";
            toolStripMenuItemDev4LogMail.Click += toolStripMenuItemDev4LogMail_Click;
            // 
            // ImportsExcel
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.SystemColors.Window;
            ClientSize = new System.Drawing.Size(729, 761);
            Controls.Add(statusStrip1);
            Controls.Add(menuStrip1);
            Controls.Add(ButtExecuterMulti);
            Controls.Add(ProgressBar1);
            Controls.Add(GroupBox1);
            Controls.Add(GroupBox2);
            Controls.Add(ButtExecuter);
            Controls.Add(ButtQuitter);
            Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            Margin = new System.Windows.Forms.Padding(4);
            Name = "ImportsExcel";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Dev4Log Import Excel Achats/Ventes vers Bob50";
            GroupBox1.ResumeLayout(false);
            GroupBox1.PerformLayout();
            GroupBox2.ResumeLayout(false);
            GroupBox2.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.GroupBox GroupBox2;
        private System.Windows.Forms.StatusStrip statusStrip1;
        public System.Windows.Forms.ProgressBar ProgressBar1;
        public System.Windows.Forms.Button RechDossier;
        public System.Windows.Forms.TextBox FichierImport;
        public System.Windows.Forms.TextBox InfosImport;
        public System.Windows.Forms.CheckBox cb_NumAutoVentes;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuLogs;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuLicence;
        public System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemDev4LogMail;
        public System.Windows.Forms.Button ButtExecuter;
        public System.Windows.Forms.Button ButtQuitter;
        public System.Windows.Forms.Button ButtExecuterMulti;
    }
}