namespace Adres_Defteri
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.pageSetupDialog1 = new System.Windows.Forms.PageSetupDialog();
            this.btnSayfaYapisi = new System.Windows.Forms.Button();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.btnBaskiOnizleme = new System.Windows.Forms.Button();
            this.btnYaz = new System.Windows.Forms.Button();
            this.btnCikis = new System.Windows.Forms.Button();
            this.btnSil = new System.Windows.Forms.Button();
            this.btnGuncelle = new System.Windows.Forms.Button();
            this.btnEkle = new System.Windows.Forms.Button();
            this.lblAdres = new System.Windows.Forms.Label();
            this.lbleMail = new System.Windows.Forms.Label();
            this.lblTelefon = new System.Windows.Forms.Label();
            this.lblSoyad = new System.Windows.Forms.Label();
            this.lblAd = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtAranan = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbGruplar = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblGrup = new System.Windows.Forms.Label();
            this.lblID = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.gridKisiler = new System.Windows.Forms.DataGridView();
            this.IDKisi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.soyad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.telefon = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.seçiliKaydıSilToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seçiliKaydıDüzenleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridKisiler)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // btnSayfaYapisi
            // 
            this.btnSayfaYapisi.Location = new System.Drawing.Point(444, 445);
            this.btnSayfaYapisi.Name = "btnSayfaYapisi";
            this.btnSayfaYapisi.Size = new System.Drawing.Size(95, 47);
            this.btnSayfaYapisi.TabIndex = 21;
            this.btnSayfaYapisi.Text = "SAYFA YAPISI";
            this.btnSayfaYapisi.UseVisualStyleBackColor = true;
            this.btnSayfaYapisi.Click += new System.EventHandler(this.btnSayfaYapisi_Click);
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // btnBaskiOnizleme
            // 
            this.btnBaskiOnizleme.Location = new System.Drawing.Point(551, 445);
            this.btnBaskiOnizleme.Name = "btnBaskiOnizleme";
            this.btnBaskiOnizleme.Size = new System.Drawing.Size(95, 47);
            this.btnBaskiOnizleme.TabIndex = 22;
            this.btnBaskiOnizleme.Text = "BASKI ÖNİZLEME";
            this.btnBaskiOnizleme.UseVisualStyleBackColor = true;
            this.btnBaskiOnizleme.Click += new System.EventHandler(this.btnBaskiOnizleme_Click);
            // 
            // btnYaz
            // 
            this.btnYaz.Location = new System.Drawing.Point(337, 454);
            this.btnYaz.Name = "btnYaz";
            this.btnYaz.Size = new System.Drawing.Size(95, 29);
            this.btnYaz.TabIndex = 20;
            this.btnYaz.Text = "YAZDIR";
            this.btnYaz.UseVisualStyleBackColor = true;
            this.btnYaz.Click += new System.EventHandler(this.btnYaz_Click);
            // 
            // btnCikis
            // 
            this.btnCikis.Location = new System.Drawing.Point(658, 451);
            this.btnCikis.Name = "btnCikis";
            this.btnCikis.Size = new System.Drawing.Size(95, 29);
            this.btnCikis.TabIndex = 19;
            this.btnCikis.Text = "ÇIKIŞ";
            this.btnCikis.UseVisualStyleBackColor = true;
            this.btnCikis.Click += new System.EventHandler(this.btnCikis_Click);
            // 
            // btnSil
            // 
            this.btnSil.Location = new System.Drawing.Point(230, 454);
            this.btnSil.Name = "btnSil";
            this.btnSil.Size = new System.Drawing.Size(95, 29);
            this.btnSil.TabIndex = 18;
            this.btnSil.Text = "SİL";
            this.btnSil.UseVisualStyleBackColor = true;
            this.btnSil.Click += new System.EventHandler(this.btnSil_Click);
            // 
            // btnGuncelle
            // 
            this.btnGuncelle.Location = new System.Drawing.Point(123, 454);
            this.btnGuncelle.Name = "btnGuncelle";
            this.btnGuncelle.Size = new System.Drawing.Size(95, 29);
            this.btnGuncelle.TabIndex = 17;
            this.btnGuncelle.Text = "GÜNCELLE";
            this.btnGuncelle.UseVisualStyleBackColor = true;
            this.btnGuncelle.Click += new System.EventHandler(this.btnGuncelle_Click);
            // 
            // btnEkle
            // 
            this.btnEkle.Location = new System.Drawing.Point(16, 454);
            this.btnEkle.Name = "btnEkle";
            this.btnEkle.Size = new System.Drawing.Size(95, 29);
            this.btnEkle.TabIndex = 16;
            this.btnEkle.Text = "EKLE";
            this.btnEkle.UseVisualStyleBackColor = true;
            this.btnEkle.Click += new System.EventHandler(this.btnEkle_Click);
            // 
            // lblAdres
            // 
            this.lblAdres.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblAdres.Location = new System.Drawing.Point(77, 277);
            this.lblAdres.Name = "lblAdres";
            this.lblAdres.Size = new System.Drawing.Size(221, 82);
            this.lblAdres.TabIndex = 13;
            this.lblAdres.Text = resources.GetString("lblAdres.Text");
            // 
            // lbleMail
            // 
            this.lbleMail.AutoSize = true;
            this.lbleMail.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lbleMail.Location = new System.Drawing.Point(77, 236);
            this.lbleMail.Name = "lbleMail";
            this.lbleMail.Size = new System.Drawing.Size(216, 16);
            this.lbleMail.TabIndex = 12;
            this.lbleMail.Text = "....................................................";
            // 
            // lblTelefon
            // 
            this.lblTelefon.AutoSize = true;
            this.lblTelefon.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblTelefon.Location = new System.Drawing.Point(77, 195);
            this.lblTelefon.Name = "lblTelefon";
            this.lblTelefon.Size = new System.Drawing.Size(216, 16);
            this.lblTelefon.TabIndex = 11;
            this.lblTelefon.Text = "....................................................";
            // 
            // lblSoyad
            // 
            this.lblSoyad.AutoSize = true;
            this.lblSoyad.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblSoyad.Location = new System.Drawing.Point(77, 154);
            this.lblSoyad.Name = "lblSoyad";
            this.lblSoyad.Size = new System.Drawing.Size(216, 16);
            this.lblSoyad.TabIndex = 10;
            this.lblSoyad.Text = "....................................................";
            // 
            // lblAd
            // 
            this.lblAd.AutoSize = true;
            this.lblAd.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblAd.Location = new System.Drawing.Point(77, 113);
            this.lblAd.Name = "lblAd";
            this.lblAd.Size = new System.Drawing.Size(216, 16);
            this.lblAd.TabIndex = 9;
            this.lblAd.Text = "....................................................";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtAranan);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(13, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(430, 55);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Kayıt Arama";
            // 
            // txtAranan
            // 
            this.txtAranan.Location = new System.Drawing.Point(190, 20);
            this.txtAranan.Name = "txtAranan";
            this.txtAranan.Size = new System.Drawing.Size(228, 23);
            this.txtAranan.TabIndex = 2;
            this.txtAranan.TextChanged += new System.EventHandler(this.txtAranan_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(176, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Aranacak Kelimeyi Yazın:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbGruplar);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(449, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(304, 55);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Grup Seçimi";
            // 
            // cmbGruplar
            // 
            this.cmbGruplar.FormattingEnabled = true;
            this.cmbGruplar.Location = new System.Drawing.Point(127, 20);
            this.cmbGruplar.Name = "cmbGruplar";
            this.cmbGruplar.Size = new System.Drawing.Size(160, 24);
            this.cmbGruplar.TabIndex = 2;
            this.cmbGruplar.SelectedIndexChanged += new System.EventHandler(this.cmbGruplar_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Grubu Seçin:";
            // 
            // lblGrup
            // 
            this.lblGrup.AutoSize = true;
            this.lblGrup.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblGrup.Location = new System.Drawing.Point(77, 72);
            this.lblGrup.Name = "lblGrup";
            this.lblGrup.Size = new System.Drawing.Size(216, 16);
            this.lblGrup.TabIndex = 8;
            this.lblGrup.Text = "....................................................";
            // 
            // lblID
            // 
            this.lblID.AutoSize = true;
            this.lblID.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.lblID.Location = new System.Drawing.Point(77, 31);
            this.lblID.Name = "lblID";
            this.lblID.Size = new System.Drawing.Size(216, 16);
            this.lblID.TabIndex = 7;
            this.lblID.Text = "....................................................";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(21, 277);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(50, 16);
            this.label9.TabIndex = 6;
            this.label9.Text = "Adres:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 236);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 16);
            this.label8.TabIndex = 5;
            this.label8.Text = "e-Mail:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(21, 195);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 16);
            this.label7.TabIndex = 4;
            this.label7.Text = "Telfon:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(18, 154);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 16);
            this.label6.TabIndex = 3;
            this.label6.Text = "Soyad:";
            // 
            // gridKisiler
            // 
            this.gridKisiler.AllowUserToAddRows = false;
            this.gridKisiler.AllowUserToDeleteRows = false;
            this.gridKisiler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridKisiler.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.IDKisi,
            this.ad,
            this.soyad,
            this.telefon});
            this.gridKisiler.ContextMenuStrip = this.contextMenuStrip1;
            this.gridKisiler.Location = new System.Drawing.Point(13, 73);
            this.gridKisiler.Name = "gridKisiler";
            this.gridKisiler.ReadOnly = true;
            this.gridKisiler.RowHeadersVisible = false;
            this.gridKisiler.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.gridKisiler.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.gridKisiler.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridKisiler.Size = new System.Drawing.Size(430, 360);
            this.gridKisiler.TabIndex = 14;
            this.gridKisiler.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridKisiler_RowEnter);
            this.gridKisiler.DoubleClick += new System.EventHandler(this.gridKisiler_DoubleClick);
            this.gridKisiler.MouseDown += new System.Windows.Forms.MouseEventHandler(this.gridKisiler_MouseDown);
            // 
            // IDKisi
            // 
            this.IDKisi.DataPropertyName = "IDKisi";
            this.IDKisi.HeaderText = "ID";
            this.IDKisi.Name = "IDKisi";
            this.IDKisi.ReadOnly = true;
            this.IDKisi.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.IDKisi.Width = 50;
            // 
            // ad
            // 
            this.ad.DataPropertyName = "ad";
            this.ad.HeaderText = "Ad";
            this.ad.Name = "ad";
            this.ad.ReadOnly = true;
            this.ad.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ad.Width = 130;
            // 
            // soyad
            // 
            this.soyad.DataPropertyName = "soyad";
            this.soyad.HeaderText = "Soyad";
            this.soyad.Name = "soyad";
            this.soyad.ReadOnly = true;
            this.soyad.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.soyad.Width = 130;
            // 
            // telefon
            // 
            this.telefon.DataPropertyName = "telefon";
            this.telefon.HeaderText = "Telefon";
            this.telefon.Name = "telefon";
            this.telefon.ReadOnly = true;
            this.telefon.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.telefon.Width = 130;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.contextMenuStrip1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.seçiliKaydıSilToolStripMenuItem,
            this.seçiliKaydıDüzenleToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(181, 48);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(40, 113);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 16);
            this.label5.TabIndex = 2;
            this.label5.Text = "Ad:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 16);
            this.label4.TabIndex = 1;
            this.label4.Text = "Grup:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(46, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 16);
            this.label3.TabIndex = 0;
            this.label3.Text = "ID:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblAdres);
            this.groupBox3.Controls.Add(this.lbleMail);
            this.groupBox3.Controls.Add(this.lblTelefon);
            this.groupBox3.Controls.Add(this.lblSoyad);
            this.groupBox3.Controls.Add(this.lblAd);
            this.groupBox3.Controls.Add(this.lblGrup);
            this.groupBox3.Controls.Add(this.lblID);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(449, 64);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(304, 369);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Kişi Bilgileri";
            // 
            // seçiliKaydıSilToolStripMenuItem
            // 
            this.seçiliKaydıSilToolStripMenuItem.Image = global::Adres_Defteri.Properties.Resources.Actions_window_close_icon;
            this.seçiliKaydıSilToolStripMenuItem.Name = "seçiliKaydıSilToolStripMenuItem";
            this.seçiliKaydıSilToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.seçiliKaydıSilToolStripMenuItem.Text = "Seçili Kaydı Sil";
            this.seçiliKaydıSilToolStripMenuItem.Click += new System.EventHandler(this.seçiliKaydıSilToolStripMenuItem_Click);
            // 
            // seçiliKaydıDüzenleToolStripMenuItem
            // 
            this.seçiliKaydıDüzenleToolStripMenuItem.Image = global::Adres_Defteri.Properties.Resources.Text_Edit_icon;
            this.seçiliKaydıDüzenleToolStripMenuItem.Name = "seçiliKaydıDüzenleToolStripMenuItem";
            this.seçiliKaydıDüzenleToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.seçiliKaydıDüzenleToolStripMenuItem.Text = "Seçili Kaydı Düzenle";
            this.seçiliKaydıDüzenleToolStripMenuItem.Click += new System.EventHandler(this.seçiliKaydıDüzenleToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 495);
            this.ControlBox = false;
            this.Controls.Add(this.btnSayfaYapisi);
            this.Controls.Add(this.btnBaskiOnizleme);
            this.Controls.Add(this.btnYaz);
            this.Controls.Add(this.btnCikis);
            this.Controls.Add(this.btnSil);
            this.Controls.Add(this.btnGuncelle);
            this.Controls.Add(this.btnEkle);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.gridKisiler);
            this.Controls.Add(this.groupBox3);
            this.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ADRES DEFTERİ by Nazmi ACAR";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridKisiler)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.PageSetupDialog pageSetupDialog1;
        private System.Windows.Forms.Button btnSayfaYapisi;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.Button btnBaskiOnizleme;
        private System.Windows.Forms.Button btnYaz;
        private System.Windows.Forms.Button btnCikis;
        private System.Windows.Forms.Button btnSil;
        private System.Windows.Forms.Button btnGuncelle;
        private System.Windows.Forms.Button btnEkle;
        private System.Windows.Forms.Label lblAdres;
        private System.Windows.Forms.Label lbleMail;
        private System.Windows.Forms.Label lblTelefon;
        private System.Windows.Forms.Label lblSoyad;
        private System.Windows.Forms.Label lblAd;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtAranan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cmbGruplar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblGrup;
        private System.Windows.Forms.Label lblID;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView gridKisiler;
        private System.Windows.Forms.DataGridViewTextBoxColumn IDKisi;
        private System.Windows.Forms.DataGridViewTextBoxColumn ad;
        private System.Windows.Forms.DataGridViewTextBoxColumn soyad;
        private System.Windows.Forms.DataGridViewTextBoxColumn telefon;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem seçiliKaydıSilToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seçiliKaydıDüzenleToolStripMenuItem;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox3;
    }
}

