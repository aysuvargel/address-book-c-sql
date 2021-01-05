using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;//ekle
using System.Drawing.Printing;//ekle

namespace Adres_Defteri
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection baglanti;
        OleDbCommand komut;

        OleDbDataAdapter dtAdaptorGrup, dtAdaptorKisi;
        DataSet dsGrup, dsKisi;

        public static string yol;
        public static int kullaniciID;//İlgilikaydı silmek için
        int grupID;//ComboBoxta seçili olan üyeye ait id noya ihtiyaç olduğundan

        void grupDoldur()
        {
            string sqlGrup = "SELECT * FROM Grup";
            dtAdaptorGrup = new OleDbDataAdapter(sqlGrup, baglanti);
            dsGrup = new DataSet();
            baglanti.Open();
            dtAdaptorGrup.Fill(dsGrup, "Grup");
            baglanti.Close();
            cmbGruplar.DataSource = dsGrup.Tables["Grup"];
            cmbGruplar.DisplayMember = "GrupAdi";
            cmbGruplar.ValueMember = "IDGrup";
            cmbGruplar.SelectedIndex = -1;
        }

        void gridDoldur(string sqlMetni)
        {
            dtAdaptorKisi = new OleDbDataAdapter(sqlMetni, baglanti);
            dsKisi = new DataSet();
            baglanti.Open();
            dtAdaptorKisi.Fill(dsKisi, "Kisi");
            baglanti.Close();
            gridKisiler.DataSource = dsKisi.Tables["Kisi"];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            yol = "Provider=Microsoft.jet.oledb.4.0; data source=adresdefteri.mdb";
            baglanti = new OleDbConnection(yol);
            grupDoldur();
            string sqlKisi = "SELECT * FROM Kisi";
            gridDoldur(sqlKisi);
        }

        int seciliSatir;
        private void gridKisiler_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            seciliSatir = e.RowIndex;
            lblID.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["IDKisi"].ToString();
            int geciciGrupID = Convert.ToInt32(dsKisi.Tables["Kisi"].Rows[seciliSatir]["IDGrup"]);
            lblGrup.Text = grupAdi(geciciGrupID);
            lblAd.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["ad"].ToString();
            lblSoyad.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["soyad"].ToString();
            lblTelefon.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["telefon"].ToString();
            lbleMail.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["email"].ToString();
            lblAdres.Text = dsKisi.Tables["Kisi"].Rows[seciliSatir]["adres"].ToString();

            kullaniciID = Convert.ToInt32(lblID.Text);//kayıt silme işleminde kullanılacak.
        }

        public string grupAdi(int geciciGrupID)
        {
            string sqlGrupAdi = "SELECT GrupAdi FROM Grup WHERE IDGrup=" + geciciGrupID;
            dtAdaptorGrup = new OleDbDataAdapter(sqlGrupAdi, baglanti);
            dsGrup = new DataSet();
            baglanti.Open();
            dtAdaptorGrup.Fill(dsGrup, "Grup");
            baglanti.Close();
            string grupAdi = dsGrup.Tables["Grup"].Rows[0]["GrupAdi"].ToString();
            return grupAdi;
        }

        private void txtAranan_TextChanged(object sender, EventArgs e)
        {
            string sqlAranan = "SELECT * FROM Kisi WHERE ad LIKE '" + txtAranan.Text + "%'";
            gridDoldur(sqlAranan);
        }

        private void cmbGruplar_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                grupID = Convert.ToInt32(cmbGruplar.SelectedValue);
                string sqlGrupID = "SELECT * FROM Kisi WHERE IDGrup=" + grupID;
                gridDoldur(sqlGrupID);
            }
            catch
            {

            }
        }

        private void btnCikis_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #region SİLME
        private void btnSil_Click(object sender, EventArgs e)
        {
            DialogResult cevap = MessageBox.Show(kullaniciID+" Nolu Kayıt Silinecek!", "DİKKAT!",MessageBoxButtons.OKCancel);
            if (cevap==DialogResult.OK)
            {
                string silinecekKayit = "DELETE FROM Kisi WHERE IDKisi="+kullaniciID;
                komut = new OleDbCommand(silinecekKayit, baglanti);
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
                gridDoldur("SELECT * FROM Kisi");
            }
        }
        #endregion

        private void btnEkle_Click(object sender, EventArgs e)
        {
            YeniKayit frmYenikayit = new YeniKayit();
            frmYenikayit.Show();
            this.Visible = false;
        }

        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            KayitGuncelle kguncelle = new KayitGuncelle();
            kguncelle.Show();
            this.Visible = false;
        }

        private void seçiliKaydıSilToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSil_Click(sender, e);
        }

        private void seçiliKaydıDüzenleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnGuncelle_Click(sender, e);
        }

        private void gridKisiler_MouseDown(object sender, MouseEventArgs e)
        {//grid üzerinde fare ile kliklendiğinde x ve y koordinatlarını belirler
            if (e.Button==MouseButtons.Right)
            {
                var hti = gridKisiler.HitTest(e.X, e.Y);
                gridKisiler.ClearSelection();
                gridKisiler.Rows[hti.RowIndex].Selected = true;
                kullaniciID = Convert.ToInt32(gridKisiler.Rows[hti.RowIndex].Cells[0].Value);
            }
        }

        private void gridKisiler_DoubleClick(object sender, EventArgs e)
        {
            btnGuncelle_Click(sender, e);
        }

        #region YAZDIRMA
        int i = 0;//satırlar kullanılırken gerekli
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font baslik = new Font("Arial", 10, FontStyle.Bold);
            Font altbaslik = new Font("Arial", 8, FontStyle.Regular);
            PageSettings sayfa = printDocument1.DefaultPageSettings;
            int y = 135, satir = gridKisiler.Rows.Count;

            e.Graphics.DrawLine(new Pen(Color.Black, 1), sayfa.Margins.Left - 75, 100, sayfa.PaperSize.Width - sayfa.Margins.Right + 70, 100);
            e.Graphics.DrawString("SNO", baslik, Brushes.Black, 25, 100);
            e.Graphics.DrawString("AD", baslik, Brushes.Black, 60, 100);
            e.Graphics.DrawString("SOYAD", baslik, Brushes.Black, 160, 100);
            e.Graphics.DrawString("TELEFON", baslik, Brushes.Black, 260, 100);
            e.Graphics.DrawString("EMAİL", baslik, Brushes.Black, 335, 100);
            e.Graphics.DrawString("ADRES", baslik, Brushes.Black, 485, 100);
            e.Graphics.DrawLine(new Pen(Color.Black, 1), sayfa.Margins.Left - 75, 120, sayfa.PaperSize.Width - sayfa.Margins.Right + 70, 120);

            while (i<satir)
            {
                e.Graphics.DrawString((i+1).ToString(), altbaslik, Brushes.Black, 25, y);
                e.Graphics.DrawString(gridKisiler.Rows[i].Cells[2].Value.ToString(), altbaslik, Brushes.Black, 60, y);
                e.Graphics.DrawString(gridKisiler.Rows[i].Cells[3].Value.ToString(), altbaslik, Brushes.Black, 160, y);
                e.Graphics.DrawString(gridKisiler.Rows[i].Cells[4].Value.ToString(), altbaslik, Brushes.Black, 260, y);
                e.Graphics.DrawString(gridKisiler.Rows[i].Cells[5].Value.ToString(), altbaslik, Brushes.Black, 335, y);
                e.Graphics.DrawString(gridKisiler.Rows[i].Cells[6].Value.ToString(), altbaslik, Brushes.Black, 485, y);
                i++; y += 25;

                if (y + 155 > sayfa.PaperSize.Height + 80 - sayfa.Margins.Bottom + 80)
                {
                    e.HasMorePages = true;
                    break;
                }
            }

            if (i>=satir)
            {
                e.HasMorePages = false;
                i = 0;
            }

        }

        private void btnBaskiOnizleme_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();       
        }
        #endregion

        private void btnYaz_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void btnSayfaYapisi_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.PageSettings = printDocument1.DefaultPageSettings;
            if (pageSetupDialog1.ShowDialog()==DialogResult.OK)
            {
                printDocument1.DefaultPageSettings = pageSetupDialog1.PageSettings;
            }
        }


    }
}
