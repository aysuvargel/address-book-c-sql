using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Adres_Defteri
{
    public partial class KayitGuncelle : Form
    {
        public KayitGuncelle()
        {
            InitializeComponent();
        }

        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter dtAdaptorGrup, dtAdaptorKisi;
        DataSet dsGrup, dsKisi;

        int Kull_ID, grupID;

        private void KayitGuncelle_Load(object sender, EventArgs e)
        {
            baglanti = new OleDbConnection(Form1.yol);
            grupDoldur();
            Kull_ID = Form1.kullaniciID;
            kisiBilgileriniGetir();
        }

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

        void kisiBilgileriniGetir()
        {
            string sqlKisiler = "SELECT * FROM Kisi WHERE IDKisi="+Kull_ID;
            dtAdaptorKisi = new OleDbDataAdapter(sqlKisiler, baglanti);
            dsKisi = new DataSet();
            baglanti.Open();
            dtAdaptorKisi.Fill(dsKisi, "Kisi");
            baglanti.Close();
            lblID.Text = Kull_ID.ToString();
            txtAd.Text = dsKisi.Tables[0].Rows[0]["ad"].ToString();
            txtSoyad.Text = dsKisi.Tables[0].Rows[0]["soyad"].ToString();
            txtTelefon.Text = dsKisi.Tables[0].Rows[0]["telefon"].ToString();
            txtemail.Text = dsKisi.Tables[0].Rows[0]["email"].ToString();
            txtAdres.Text = dsKisi.Tables[0].Rows[0]["adres"].ToString();
            grupID = Convert.ToInt32(dsKisi.Tables[0].Rows[0]["IDGrup"]);
            cmbGruplar.SelectedValue = grupID;
        }

        void temizle()
        {
            txtAd.Clear();
            txtSoyad.Clear();
            txtTelefon.Clear();
            txtemail.Clear();
            txtAdres.Clear();
            cmbGruplar.Text = "";
            lblID.Text="";
        }

        private void btnGeriDon_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            frm1.Show();
            this.Close();
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {
            if (txtAd.Text == "" || txtSoyad.Text == "" || txtTelefon.Text == "" || txtemail.Text == "" || txtAdres.Text == "" || cmbGruplar.Text == "")
            {
                txtAd.Focus();
            }
            else
            {
                int IDgrup = Convert.ToInt32(cmbGruplar.SelectedValue);
                string sqlGuncelle = "UPDATE Kisi SET IDGrup=@IDgrup, ad=@txtAd, soyad=@txtSoyad, telefon=@txtTelefon, email=@txtemail, adres=@txtAdres WHERE IDKisi="+Kull_ID;
                komut = new OleDbCommand(sqlGuncelle, baglanti);
                baglanti.Open();
                komut.Parameters.Add("@IDgrup", OleDbType.Integer).Value = IDgrup;
                komut.Parameters.Add("@txtAd", OleDbType.Char, 50).Value = txtAd.Text;
                komut.Parameters.Add("@txtSoyad", OleDbType.Char, 50).Value = txtSoyad.Text;
                komut.Parameters.Add("@txtTelefon", OleDbType.Char, 50).Value = txtTelefon.Text;
                komut.Parameters.Add("@txtemail", OleDbType.Char, 50).Value = txtemail.Text;
                komut.Parameters.Add("@txtAdres", OleDbType.Char, 255).Value = txtAdres.Text;
                DialogResult cevap = MessageBox.Show("KAYIT GÜNCELLENİYOR...", "DİKKAT!", MessageBoxButtons.YesNo);
                if (cevap==DialogResult.Yes)
                {
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                    temizle();
                }
                else
                {
                    baglanti.Close();
                }
            }
        }
    }
}
