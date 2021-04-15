namespace JurisUtilityBase
{
    partial class UtilityBaseMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UtilityBaseMain));
            this.JurisLogoImageBox = new System.Windows.Forms.PictureBox();
            this.LexisNexisLogoPictureBox = new System.Windows.Forms.PictureBox();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.listBoxCompanies = new System.Windows.Forms.ListBox();
            this.labelDescription = new System.Windows.Forms.Label();
            this.statusGroupBox = new System.Windows.Forms.GroupBox();
            this.labelCurrentStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelPercentComplete = new System.Windows.Forms.Label();
            this.OpenFileDialogOpen = new System.Windows.Forms.OpenFileDialog();
            this.buttonReport = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.TabControl1 = new System.Windows.Forms.TabControl();
            this.ClientTab = new System.Windows.Forms.TabPage();
            this.richTextBoxCliRemarks = new System.Windows.Forms.RichTextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBoxCliName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxCliBranch = new System.Windows.Forms.TextBox();
            this.buttonClientExcel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxClientText = new System.Windows.Forms.TextBox();
            this.MatterTab = new System.Windows.Forms.TabPage();
            this.richTextBoxRemarksMatter = new System.Windows.Forms.RichTextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxClientMatter = new System.Windows.Forms.TextBox();
            this.buttonMatterExcel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxMatterText = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).BeginInit();
            this.statusStrip.SuspendLayout();
            this.statusGroupBox.SuspendLayout();
            this.TabControl1.SuspendLayout();
            this.ClientTab.SuspendLayout();
            this.MatterTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // JurisLogoImageBox
            // 
            this.JurisLogoImageBox.Image = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.Image")));
            this.JurisLogoImageBox.InitialImage = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.InitialImage")));
            this.JurisLogoImageBox.Location = new System.Drawing.Point(0, 1);
            this.JurisLogoImageBox.Name = "JurisLogoImageBox";
            this.JurisLogoImageBox.Size = new System.Drawing.Size(104, 336);
            this.JurisLogoImageBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.JurisLogoImageBox.TabIndex = 0;
            this.JurisLogoImageBox.TabStop = false;
            // 
            // LexisNexisLogoPictureBox
            // 
            this.LexisNexisLogoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LexisNexisLogoPictureBox.Image")));
            this.LexisNexisLogoPictureBox.Location = new System.Drawing.Point(8, 343);
            this.LexisNexisLogoPictureBox.Name = "LexisNexisLogoPictureBox";
            this.LexisNexisLogoPictureBox.Size = new System.Drawing.Size(96, 28);
            this.LexisNexisLogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.LexisNexisLogoPictureBox.TabIndex = 1;
            this.LexisNexisLogoPictureBox.TabStop = false;
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 469);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(628, 22);
            this.statusStrip.TabIndex = 2;
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripStatusLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(134, 17);
            this.toolStripStatusLabel.Text = "Status: Ready to Execute";
            // 
            // listBoxCompanies
            // 
            this.listBoxCompanies.FormattingEnabled = true;
            this.listBoxCompanies.Location = new System.Drawing.Point(111, 1);
            this.listBoxCompanies.Name = "listBoxCompanies";
            this.listBoxCompanies.Size = new System.Drawing.Size(262, 69);
            this.listBoxCompanies.TabIndex = 0;
            this.listBoxCompanies.SelectedIndexChanged += new System.EventHandler(this.listBoxCompanies_SelectedIndexChanged);
            // 
            // labelDescription
            // 
            this.labelDescription.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelDescription.Location = new System.Drawing.Point(380, 1);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(224, 69);
            this.labelDescription.TabIndex = 1;
            this.labelDescription.Text = "Allows matters to be closed one at a time, per client or via an Excel file";
            // 
            // statusGroupBox
            // 
            this.statusGroupBox.Controls.Add(this.labelCurrentStatus);
            this.statusGroupBox.Controls.Add(this.progressBar);
            this.statusGroupBox.Controls.Add(this.labelPercentComplete);
            this.statusGroupBox.Location = new System.Drawing.Point(111, 77);
            this.statusGroupBox.Name = "statusGroupBox";
            this.statusGroupBox.Size = new System.Drawing.Size(493, 73);
            this.statusGroupBox.TabIndex = 5;
            this.statusGroupBox.TabStop = false;
            this.statusGroupBox.Text = "Utility Status:";
            // 
            // labelCurrentStatus
            // 
            this.labelCurrentStatus.AutoSize = true;
            this.labelCurrentStatus.Location = new System.Drawing.Point(7, 50);
            this.labelCurrentStatus.Name = "labelCurrentStatus";
            this.labelCurrentStatus.Size = new System.Drawing.Size(77, 13);
            this.labelCurrentStatus.TabIndex = 2;
            this.labelCurrentStatus.Text = "Current Status:";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(7, 27);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(468, 20);
            this.progressBar.TabIndex = 0;
            // 
            // labelPercentComplete
            // 
            this.labelPercentComplete.AutoSize = true;
            this.labelPercentComplete.Location = new System.Drawing.Point(413, 11);
            this.labelPercentComplete.Name = "labelPercentComplete";
            this.labelPercentComplete.Size = new System.Drawing.Size(62, 13);
            this.labelPercentComplete.TabIndex = 0;
            this.labelPercentComplete.Text = "% Complete";
            // 
            // OpenFileDialogOpen
            // 
            this.OpenFileDialogOpen.FileName = "openFileDialog1";
            // 
            // buttonReport
            // 
            this.buttonReport.BackColor = System.Drawing.Color.LightGray;
            this.buttonReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonReport.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonReport.Location = new System.Drawing.Point(133, 418);
            this.buttonReport.Name = "buttonReport";
            this.buttonReport.Size = new System.Drawing.Size(148, 38);
            this.buttonReport.TabIndex = 16;
            this.buttonReport.Text = "Exit";
            this.buttonReport.UseVisualStyleBackColor = false;
            this.buttonReport.Click += new System.EventHandler(this.buttonReport_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.LightGray;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.MidnightBlue;
            this.button1.Location = new System.Drawing.Point(438, 418);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(148, 38);
            this.button1.TabIndex = 17;
            this.button1.Text = "Run";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TabControl1
            // 
            this.TabControl1.Controls.Add(this.ClientTab);
            this.TabControl1.Controls.Add(this.MatterTab);
            this.TabControl1.Location = new System.Drawing.Point(121, 156);
            this.TabControl1.Name = "TabControl1";
            this.TabControl1.SelectedIndex = 0;
            this.TabControl1.Size = new System.Drawing.Size(465, 243);
            this.TabControl1.TabIndex = 18;
            this.TabControl1.SelectedIndexChanged += new System.EventHandler(this.TabControl1_SelectedIndexChanged);
            // 
            // ClientTab
            // 
            this.ClientTab.Controls.Add(this.richTextBoxCliRemarks);
            this.ClientTab.Controls.Add(this.label8);
            this.ClientTab.Controls.Add(this.label7);
            this.ClientTab.Controls.Add(this.textBoxCliName);
            this.ClientTab.Controls.Add(this.label6);
            this.ClientTab.Controls.Add(this.textBoxCliBranch);
            this.ClientTab.Controls.Add(this.buttonClientExcel);
            this.ClientTab.Controls.Add(this.label2);
            this.ClientTab.Controls.Add(this.label1);
            this.ClientTab.Controls.Add(this.textBoxClientText);
            this.ClientTab.Location = new System.Drawing.Point(4, 22);
            this.ClientTab.Name = "ClientTab";
            this.ClientTab.Padding = new System.Windows.Forms.Padding(3);
            this.ClientTab.Size = new System.Drawing.Size(457, 217);
            this.ClientTab.TabIndex = 0;
            this.ClientTab.Text = "Client";
            this.ClientTab.UseVisualStyleBackColor = true;
            // 
            // richTextBoxCliRemarks
            // 
            this.richTextBoxCliRemarks.Location = new System.Drawing.Point(241, 84);
            this.richTextBoxCliRemarks.Name = "richTextBoxCliRemarks";
            this.richTextBoxCliRemarks.Size = new System.Drawing.Size(194, 46);
            this.richTextBoxCliRemarks.TabIndex = 3;
            this.richTextBoxCliRemarks.Text = "";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label8.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label8.Location = new System.Drawing.Point(240, 65);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 16);
            this.label8.TabIndex = 25;
            this.label8.Text = "Remarks";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label7.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label7.Location = new System.Drawing.Point(11, 65);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(83, 16);
            this.label7.TabIndex = 23;
            this.label7.Text = "New Name";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // textBoxCliName
            // 
            this.textBoxCliName.Location = new System.Drawing.Point(12, 87);
            this.textBoxCliName.Name = "textBoxCliName";
            this.textBoxCliName.Size = new System.Drawing.Size(194, 20);
            this.textBoxCliName.TabIndex = 2;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label6.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label6.Location = new System.Drawing.Point(240, 9);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(90, 16);
            this.label6.TabIndex = 21;
            this.label6.Text = "New Branch";
            // 
            // textBoxCliBranch
            // 
            this.textBoxCliBranch.Location = new System.Drawing.Point(241, 31);
            this.textBoxCliBranch.Name = "textBoxCliBranch";
            this.textBoxCliBranch.Size = new System.Drawing.Size(194, 20);
            this.textBoxCliBranch.TabIndex = 1;
            // 
            // buttonClientExcel
            // 
            this.buttonClientExcel.BackColor = System.Drawing.Color.LightGray;
            this.buttonClientExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonClientExcel.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonClientExcel.Location = new System.Drawing.Point(121, 155);
            this.buttonClientExcel.Name = "buttonClientExcel";
            this.buttonClientExcel.Size = new System.Drawing.Size(180, 38);
            this.buttonClientExcel.TabIndex = 19;
            this.buttonClientExcel.Text = "Select CSV";
            this.buttonClientExcel.UseVisualStyleBackColor = false;
            this.buttonClientExcel.Click += new System.EventHandler(this.buttonClientExcel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label2.Location = new System.Drawing.Point(198, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "OR";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label1.Location = new System.Drawing.Point(11, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(181, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter ONE Client Number";
            // 
            // textBoxClientText
            // 
            this.textBoxClientText.Location = new System.Drawing.Point(12, 31);
            this.textBoxClientText.Name = "textBoxClientText";
            this.textBoxClientText.Size = new System.Drawing.Size(194, 20);
            this.textBoxClientText.TabIndex = 0;
            // 
            // MatterTab
            // 
            this.MatterTab.Controls.Add(this.richTextBoxRemarksMatter);
            this.MatterTab.Controls.Add(this.label9);
            this.MatterTab.Controls.Add(this.label5);
            this.MatterTab.Controls.Add(this.textBoxClientMatter);
            this.MatterTab.Controls.Add(this.buttonMatterExcel);
            this.MatterTab.Controls.Add(this.label3);
            this.MatterTab.Controls.Add(this.label4);
            this.MatterTab.Controls.Add(this.textBoxMatterText);
            this.MatterTab.Location = new System.Drawing.Point(4, 22);
            this.MatterTab.Name = "MatterTab";
            this.MatterTab.Padding = new System.Windows.Forms.Padding(3);
            this.MatterTab.Size = new System.Drawing.Size(457, 217);
            this.MatterTab.TabIndex = 1;
            this.MatterTab.Text = "Matter";
            this.MatterTab.UseVisualStyleBackColor = true;
            // 
            // richTextBoxRemarksMatter
            // 
            this.richTextBoxRemarksMatter.Location = new System.Drawing.Point(131, 86);
            this.richTextBoxRemarksMatter.Name = "richTextBoxRemarksMatter";
            this.richTextBoxRemarksMatter.Size = new System.Drawing.Size(194, 46);
            this.richTextBoxRemarksMatter.TabIndex = 6;
            this.richTextBoxRemarksMatter.Text = "";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label9.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label9.Location = new System.Drawing.Point(130, 67);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(70, 16);
            this.label9.TabIndex = 27;
            this.label9.Text = "Remarks";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label5.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label5.Location = new System.Drawing.Point(6, 12);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(181, 16);
            this.label5.TabIndex = 25;
            this.label5.Text = "Enter ONE Client Number";
            // 
            // textBoxClientMatter
            // 
            this.textBoxClientMatter.Location = new System.Drawing.Point(7, 34);
            this.textBoxClientMatter.Name = "textBoxClientMatter";
            this.textBoxClientMatter.Size = new System.Drawing.Size(194, 20);
            this.textBoxClientMatter.TabIndex = 4;
            // 
            // buttonMatterExcel
            // 
            this.buttonMatterExcel.BackColor = System.Drawing.Color.LightGray;
            this.buttonMatterExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonMatterExcel.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonMatterExcel.Location = new System.Drawing.Point(133, 165);
            this.buttonMatterExcel.Name = "buttonMatterExcel";
            this.buttonMatterExcel.Size = new System.Drawing.Size(180, 38);
            this.buttonMatterExcel.TabIndex = 23;
            this.buttonMatterExcel.Text = "Select CSV";
            this.buttonMatterExcel.UseVisualStyleBackColor = false;
            this.buttonMatterExcel.Click += new System.EventHandler(this.buttonMatterExcel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label3.Location = new System.Drawing.Point(205, 134);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 16);
            this.label3.TabIndex = 22;
            this.label3.Text = "OR";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label4.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label4.Location = new System.Drawing.Point(252, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(185, 16);
            this.label4.TabIndex = 21;
            this.label4.Text = "Enter ONE Matter Number";
            // 
            // textBoxMatterText
            // 
            this.textBoxMatterText.Location = new System.Drawing.Point(253, 34);
            this.textBoxMatterText.Name = "textBoxMatterText";
            this.textBoxMatterText.Size = new System.Drawing.Size(194, 20);
            this.textBoxMatterText.TabIndex = 5;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // UtilityBaseMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(628, 491);
            this.Controls.Add(this.TabControl1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.buttonReport);
            this.Controls.Add(this.statusGroupBox);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.listBoxCompanies);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.LexisNexisLogoPictureBox);
            this.Controls.Add(this.JurisLogoImageBox);
            this.ForeColor = System.Drawing.SystemColors.WindowText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UtilityBaseMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "JPS - Kovitz Matter Closer";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).EndInit();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.statusGroupBox.ResumeLayout(false);
            this.statusGroupBox.PerformLayout();
            this.TabControl1.ResumeLayout(false);
            this.ClientTab.ResumeLayout(false);
            this.ClientTab.PerformLayout();
            this.MatterTab.ResumeLayout(false);
            this.MatterTab.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox JurisLogoImageBox;
        private System.Windows.Forms.PictureBox LexisNexisLogoPictureBox;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ListBox listBoxCompanies;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.GroupBox statusGroupBox;
        private System.Windows.Forms.Label labelCurrentStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label labelPercentComplete;
        private System.Windows.Forms.OpenFileDialog OpenFileDialogOpen;
        private System.Windows.Forms.Button buttonReport;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabControl TabControl1;
        private System.Windows.Forms.TabPage ClientTab;
        private System.Windows.Forms.TabPage MatterTab;
        private System.Windows.Forms.Button buttonClientExcel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxClientText;
        private System.Windows.Forms.Button buttonMatterExcel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxMatterText;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxClientMatter;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBoxCliName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxCliBranch;
        private System.Windows.Forms.RichTextBox richTextBoxCliRemarks;
        private System.Windows.Forms.RichTextBox richTextBoxRemarksMatter;
        private System.Windows.Forms.Label label9;
    }
}

