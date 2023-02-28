namespace BorcSorgulama
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnDosyaYolu = new System.Windows.Forms.Button();
            this.txtDosyaYolu = new System.Windows.Forms.TextBox();
            this.btnSayfa = new System.Windows.Forms.Button();
            this.txtSayfaIsmi = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 87);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(894, 546);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(886, 518);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(880, 512);
            this.dataGridView1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(886, 518);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnDosyaYolu
            // 
            this.btnDosyaYolu.Location = new System.Drawing.Point(19, 12);
            this.btnDosyaYolu.Name = "btnDosyaYolu";
            this.btnDosyaYolu.Size = new System.Drawing.Size(105, 27);
            this.btnDosyaYolu.TabIndex = 1;
            this.btnDosyaYolu.Text = "Dosya Seç";
            this.btnDosyaYolu.UseVisualStyleBackColor = true;
            this.btnDosyaYolu.Click += new System.EventHandler(this.btnDosyaYolu_Click);
            // 
            // txtDosyaYolu
            // 
            this.txtDosyaYolu.Location = new System.Drawing.Point(130, 16);
            this.txtDosyaYolu.Name = "txtDosyaYolu";
            this.txtDosyaYolu.ReadOnly = true;
            this.txtDosyaYolu.Size = new System.Drawing.Size(373, 23);
            this.txtDosyaYolu.TabIndex = 2;
            // 
            // btnSayfa
            // 
            this.btnSayfa.Location = new System.Drawing.Point(19, 45);
            this.btnSayfa.Name = "btnSayfa";
            this.btnSayfa.Size = new System.Drawing.Size(105, 27);
            this.btnSayfa.TabIndex = 3;
            this.btnSayfa.Text = "Sayfa Seç";
            this.btnSayfa.UseVisualStyleBackColor = true;
            this.btnSayfa.Click += new System.EventHandler(this.btnSayfa_Click);
            // 
            // txtSayfaIsmi
            // 
            this.txtSayfaIsmi.Location = new System.Drawing.Point(130, 45);
            this.txtSayfaIsmi.Name = "txtSayfaIsmi";
            this.txtSayfaIsmi.ReadOnly = true;
            this.txtSayfaIsmi.Size = new System.Drawing.Size(373, 23);
            this.txtSayfaIsmi.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1242, 645);
            this.Controls.Add(this.txtSayfaIsmi);
            this.Controls.Add(this.btnSayfa);
            this.Controls.Add(this.txtDosyaYolu);
            this.Controls.Add(this.btnDosyaYolu);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage1;
        private DataGridView dataGridView1;
        private TabPage tabPage2;
        private Button btnDosyaYolu;
        private TextBox txtDosyaYolu;
        private Button btnSayfa;
        private TextBox txtSayfaIsmi;
    }
}