using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Globalization;

namespace KastamonuElektrik
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=KElektrikDb1.accdb");
        DateTime labeltarih;
        bool islem = true;
        bool cari = false;
        decimal toplamtutar = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            OleDbDataAdapter adp = new OleDbDataAdapter("select * from Firma order by FirmaAd", baglanti);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            dgvFirma.DataSource = dt;
            dgvFirma.Columns["FirmaId"].Visible = dgvFirma.Columns["FirmaTel1"].Visible = dgvFirma.Columns["FirmaTel2"].Visible = dgvFirma.Columns["FirmaJenNo"].Visible = dgvFirma.Columns["FirmaMotorNo"].Visible = dgvFirma.Columns["FirmaTarih"].Visible = false;
            dgvFirma.Columns["FirmaAd"].Width = dgvFirma.Width - 20;

            CmbIslemTip.SelectedIndex = 4;
            CmbCariListele.SelectedIndex = 4;
            //CmbUrun.SelectedIndex = 0;
            TxtIslemDiger.Visible = false;
            LblIslemDiger.Visible = false;
            GrpDuzenCari.Visible = false;

            string listele2 = "select UrunId, UrunKod from Urun";
            OleDbDataAdapter adp2 = new OleDbDataAdapter(listele2, baglanti);
            DataSet ds = new DataSet();
            adp2.Fill(ds);
            CmbUrunKod.DataSource = ds.Tables[0];
            CmbUrunKod.DisplayMember = "UrunKod";
            CmbUrunKod.ValueMember = "UrunId";

            OleDbDataAdapter adpu = new OleDbDataAdapter("select UrunId, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Fiyat from Urun order by UrunKod", baglanti);
            DataTable dtu = new DataTable();
            adpu.Fill(dtu);
            DgvUrunler.DataSource = dtu;
            DgvUrunler.Columns["UrunId"].Visible = false;
            DgvUrunler.Columns["Fiyat"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            if (DgvUrunler.Rows.Count > 0)
            {
                for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                {
                    if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                    {
                        LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                        LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                    }
                }
            }
            else if (DgvUrunler.Rows.Count < 1)
            {
                LblUrunAd.Text = "";
                LblUrunFiyat.Text = "";
            }
        }

        private void Listele_Firma()
        {
            OleDbDataAdapter adp = new OleDbDataAdapter("select * from Firma order by FirmaAd", baglanti);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            dgvFirma.DataSource = dt;
            dgvFirma.Columns["FirmaId"].Visible = dgvFirma.Columns["FirmaTel1"].Visible = dgvFirma.Columns["FirmaTel2"].Visible = dgvFirma.Columns["FirmaJenNo"].Visible = dgvFirma.Columns["FirmaMotorNo"].Visible = dgvFirma.Columns["FirmaTarih"].Visible = false;
            dgvFirma.Columns["FirmaAd"].Width = dgvFirma.Width - 20;

            CmbIslemTip.SelectedIndex = 4;
            CmbCariListele.SelectedIndex = 4;
            TxtIslemDiger.Visible = false;
            LblIslemDiger.Visible = false;
            GrpDuzenCari.Visible = false;
        }

        private void BtnEkleFirma_Click(object sender, EventArgs e)
        {
            if (TxtFirmaAd.Text == null || TxtFirmaAd.Text == "")
            {
                MessageBox.Show("Lütfen Firma Adı 'nı girelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Ekle
                OleDbCommand cmd = new OleDbCommand("insert into Firma(FirmaAd,FirmaTel1,FirmaTel2,FirmaJenNo,FirmaMotorNo,[FirmaTarih]) values(@firmaad,@tel1,@tel2,@jenno,@motorno,@firmatarih)", baglanti);

                cmd.Parameters.AddWithValue("@firmaad", TxtFirmaAd.Text);
                cmd.Parameters.AddWithValue("@tel1", MskTxtTel1.Text);
                cmd.Parameters.AddWithValue("@tel2", MskTxtTel2.Text);
                cmd.Parameters.AddWithValue("@jenno", TxtJen.Text);
                cmd.Parameters.AddWithValue("@motorno", TxtMotor.Text);
                cmd.Parameters.AddWithValue("@firmatarih", DtpFirmaTarih.Text);

                baglanti.Open();
                cmd.ExecuteNonQuery();
                baglanti.Close();

                //Listele
                Listele_Firma();

                //Temizle
                TxtFirmaAd.Text = null;
                MskTxtTel1.Text = null;
                MskTxtTel2.Text = null;
                TxtJen.Text = null;
                TxtMotor.Text = null;
                DtpFirmaTarih.Text = DateTime.Now.ToString("d/M/yyyy");
            }
        }

        private void BtnTemizleFirma_Click(object sender, EventArgs e)
        {
            TxtFirmaAd.Text = null;
            MskTxtTel1.Text = null;
            MskTxtTel2.Text = null;
            TxtJen.Text = null;
            TxtMotor.Text = null;
            DtpFirmaTarih.Text = DateTime.Now.ToString("d/M/yyyy");
        }

        private void BtnFirmaGuncelle_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count < 1)
            {
                MessageBox.Show("Listede firma yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (dgvFirma.SelectedCells[1].Value == null || TxtFirmaAd.Text == "")
            {
                MessageBox.Show("Firma seçili değil!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Açıklamalar   
                int aId2 = dgvFirma.SelectedRows[0].Index;
                string aId = "No";
                if (dgvFirma.SelectedCells[1].Value == null)
                {
                    aId = aId2.ToString();
                }
                else if (dgvFirma.SelectedCells[1].Value != null)
                {
                    aId = (aId2 + 1).ToString();
                }

                string aTarih = "Gün/Ay/Yıl";
                if (dgvFirma.SelectedCells[6].Value != null)
                {
                    DateTime aTarih2 = (DateTime)dgvFirma.SelectedCells[6].Value;
                    aTarih = aTarih2.ToString("dd/MM/yyyy");
                }

                string aAd = "Ad";
                if (dgvFirma.SelectedCells[1].Value != null)
                {
                    aAd = dgvFirma.SelectedCells[1].Value.ToString();
                }

                string aTel1 = "Tel - 1";
                if (dgvFirma.SelectedCells[2].Value != null)
                {
                    aTel1 = dgvFirma.SelectedCells[2].Value.ToString();
                }

                string aTel2 = "Tel - 2";
                if (dgvFirma.SelectedCells[3].Value != null)
                {
                    aTel2 = dgvFirma.SelectedCells[3].Value.ToString();
                }

                string aJenno = "Jeneratör Seri No";
                if (dgvFirma.SelectedCells[4].Value != null)
                {
                    aJenno = dgvFirma.SelectedCells[4].Value.ToString();
                }

                string aMotor = "Motor Seri No";
                if (dgvFirma.SelectedCells[5].Value != null)
                {
                    aMotor = dgvFirma.SelectedCells[5].Value.ToString();
                }


                if (TxtFirmaAd.Text == null || TxtFirmaAd.Text == "")
                {
                    MessageBox.Show("Firma Adı 'nı girelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    DialogResult sonuc = MessageBox.Show("Sıralanan Firmalardan Satır No: " + aId + "\n" + "Tarih: " + aTarih + "\n" + "Firma Adı: " + aAd + "\n" + "Tel-1: " + aTel1 + "\n" + "Tel-2: " + aTel2 + "\n" + "Jeneratör Seri No: " + aJenno + "\n" + "Motor Seri No: " + aMotor + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + "Tarih: " + DtpFirmaTarih.Text + "\n" + "Firma Adı: " + TxtFirmaAd.Text + "\n" + "Tel-1: " + MskTxtTel1.Text + "\n" + "Tel-2: " + MskTxtTel2.Text + "\n" + "Jeneratör Seri No: " + TxtJen.Text + "\n" + "Motor Seri No: " + TxtMotor.Text + "\n" + "\n" + "Güncellensin mi?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (sonuc == DialogResult.Yes)
                    {
                        OleDbCommand cmd = new OleDbCommand("Update Firma set FirmaAd=@firmaad,FirmaTel1=@tel1,FirmaTel2=@tel2,FirmaJenNo=@jenno,FirmaMotorNo=@motorno,[FirmaTarih]=@firmatarih where FirmaId=@fid", baglanti);

                        cmd.Parameters.AddWithValue("@firmaad", TxtFirmaAd.Text);
                        cmd.Parameters.AddWithValue("@tel1", MskTxtTel1.Text);
                        cmd.Parameters.AddWithValue("@tel2", MskTxtTel2.Text);
                        cmd.Parameters.AddWithValue("@jenno", TxtJen.Text);
                        cmd.Parameters.AddWithValue("@motorno", TxtMotor.Text);
                        cmd.Parameters.AddWithValue("@firmatarih", DtpFirmaTarih.Text);
                        cmd.Parameters.AddWithValue("@fid", (int)TxtFirmaAd.Tag);

                        baglanti.Open();
                        cmd.ExecuteNonQuery();
                        baglanti.Close();

                        //Listele
                        Listele_Firma();

                        //Temizle                        
                        TxtFirmaAd.Text = null;
                        MskTxtTel1.Text = null;
                        MskTxtTel2.Text = null;
                        TxtJen.Text = null;
                        TxtMotor.Text = null;
                        DtpFirmaTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                    }
                }
            }
        }

        private void BtnSil_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count < 1)
            {
                MessageBox.Show("Listede firma yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (TxtFirmaAd.Text == "" || TxtFirmaAd.Text == null)
            {
                MessageBox.Show("Firma seçili değil!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                int id = (int)TxtFirmaAd.Tag;
                //Açıklamalar   
                int aId2 = dgvFirma.SelectedRows[0].Index;
                string aId = (aId2 + 1).ToString();

                DateTime aTarih2 = (DateTime)dgvFirma.SelectedCells[6].Value;
                string aTarih = aTarih2.ToString("dd/MM/yyyy");

                DialogResult sonuc = MessageBox.Show("Sıralanan Firmalardan Satır No: " + aId + "\n" + "Tarih: " + aTarih + "\n" + "Firma Adı: " + GrpBoxFirma.Text + "\n" + "Tel-1: " + LblTel1.Text + "\n" + "Tel-2: " + LblTel2.Text + "\n" + "Jeneratör Seri No: " + LblJenNo.Text + "\n" + "Motor Seri No: " + LblMotorNo.Text + "\n" + "\n" + "Silinsin mi?", "Silme", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (sonuc == DialogResult.Yes)
                {
                    OleDbCommand cmd2 = new OleDbCommand("delete from Islem where FirmaId=@fid2", baglanti);
                    OleDbCommand cmd = new OleDbCommand("delete from Firma where FirmaId=@fid", baglanti);
                    cmd2.Parameters.AddWithValue("@fid2", id);
                    cmd.Parameters.AddWithValue("@fid", id);
                    
                    baglanti.Open();
                    cmd2.ExecuteNonQuery();
                    cmd.ExecuteNonQuery();                  
                    baglanti.Close();

                    //Listele
                    Listele_Firma();

                    //Temizle                        
                    TxtFirmaAd.Text = null;
                    MskTxtTel1.Text = null;
                    MskTxtTel2.Text = null;
                    TxtJen.Text = null;
                    TxtMotor.Text = null;
                    DtpFirmaTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                }
            }
        }

        private void dgvFirma_SelectionChanged(object sender, EventArgs e)
        {
            if (islem)
            {
                if (dgvFirma.Rows.Count < 1)
                {
                    GrpBoxFirma.Text = "Firma Bilgileri";
                    LblFirmaTarih.Text = "Eklenme Tarihi";
                    LblTel1.Text = "İletişim 1";
                    LblTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblJenNo.Text = "";
                    LblMotorNo.Text = "";
                }
                //Soldaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
                DataGridViewRow row = dgvFirma.CurrentRow;
                if (row != null)
                {
                    GrpBoxFirma.Text = row.Cells["FirmaAd"].Value.ToString();
                    TxtFirmaAd.Tag = row.Cells["FirmaId"].Value;
                    TxtFirmaAd.Text = row.Cells["FirmaAd"].Value.ToString();
                    LblTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    MskTxtTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    LblTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    MskTxtTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    LblJenNo.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    TxtJen.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    LblMotorNo.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    TxtMotor.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    labeltarih = (DateTime)row.Cells["FirmaTarih"].Value;
                    LblFirmaTarih.Text = labeltarih.ToString("d/M/yyyy");
                    DtpFirmaTarih.Text = row.Cells["FirmaTarih"].Value.ToString();
                    CmbIslemTip.SelectedIndex = 4;

                    //oledbdata tablolar birleştirme
                    //string listele = "select IslemId, Tarih, IslemTipAciklama as [İşlem Tipi], UrunAd as [Ürün Adı], UrunKod as [Ürün Kodu], Adet, Birim from Islem INNER JOIN IslemTip ON Islem.IslemTipId = IslemTip.IslemTipId where FirmaId = " + TxtFirmaAd.Tag.ToString() + " order by Tarih desc";

                    Listele_Urun();
                    /*
                    if (dgvUrun.Rows.Count<1)
                    {
                        Temizle_Urun();
                    }
                    */
                }               
            }
            else if (cari)
            {
                if (dgvFirma.Rows.Count<1)
                {
                    LblCariTel1.Text = "İletişim 1";
                    LblCariTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblCariAdres.Text = "Adres";
                    LblAlacak.Text = "0,00";
                    LblBorc.Text = "0,00";
                }
                DataGridViewRow row = dgvFirma.CurrentRow;
                if (row != null)
                {
                    LblAlacak.Text = "0,00";
                    LblBorc.Text = "0,00";
                    GrpKisi.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Tag = row.Cells["KisiId"].Value;
                    LblCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    MskCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    LblCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    MskCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    LblCariAdres.Text = row.Cells["Adres"].Value.ToString();
                    TxtAdres.Text = row.Cells["Adres"].Value.ToString();
                    CmbCariListele.SelectedIndex = 4;

                    Cari_Listele();

                    if (DgvCari.Rows.Count < 1)
                    {
                        Temizle_Urun();
                        CmbCari.SelectedIndex = 0;
                        TxtCariTutar.Text = "0";
                        DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                        TxtAciklama.Text = "";
                    }
                }              
            }
        }

        private void TxtAraFirma_TextChanged(object sender, EventArgs e)
        {
            if (islem)
            {
                if (TxtAraFirma.Text.Length > 1)
                {
                    TxtAraFirma.ForeColor = Color.Blue;
                    OleDbDataAdapter adp = new OleDbDataAdapter("select * from Firma where FirmaAd like '%" + TxtAraFirma.Text + "%' order by FirmaAd", baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvFirma.DataSource = dt;
                    dgvFirma.Columns["FirmaId"].Visible = dgvFirma.Columns["FirmaTel1"].Visible = dgvFirma.Columns["FirmaTel2"].Visible = dgvFirma.Columns["FirmaJenNo"].Visible = dgvFirma.Columns["FirmaMotorNo"].Visible = dgvFirma.Columns["FirmaTarih"].Visible = false;
                    dgvFirma.Columns["FirmaAd"].Width = dgvFirma.Width - 20;
                }
                else
                {
                    TxtAraFirma.ForeColor = Color.Red;
                    OleDbDataAdapter adp = new OleDbDataAdapter("select * from Firma order by FirmaAd", baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvFirma.DataSource = dt;
                    dgvFirma.Columns["FirmaId"].Visible = dgvFirma.Columns["FirmaTel1"].Visible = dgvFirma.Columns["FirmaTel2"].Visible = dgvFirma.Columns["FirmaJenNo"].Visible = dgvFirma.Columns["FirmaMotorNo"].Visible = dgvFirma.Columns["FirmaTarih"].Visible = false;
                    dgvFirma.Columns["FirmaAd"].Width = dgvFirma.Width - 20;
                }
            }
            else if (cari)
            {
                if (TxtAraFirma.Text.Length > 1)
                {
                    TxtAraFirma.ForeColor = Color.Blue;
                    OleDbDataAdapter adp = new OleDbDataAdapter("select * from Kisi where Ad like '%" + TxtAraFirma.Text + "%' order by Ad", baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvFirma.DataSource = dt;
                    dgvFirma.Columns["KisiId"].Visible = dgvFirma.Columns["Tel1"].Visible = dgvFirma.Columns["Tel2"].Visible = dgvFirma.Columns["Adres"].Visible = false;
                    dgvFirma.Columns["Ad"].Width = dgvFirma.Width - 20;
                }
                else
                {
                    TxtAraFirma.ForeColor = Color.Red;
                    Listele_Kisi();
                }
            }
        }

        private void dgvUrun_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.dgvUrun.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.dgvUrun.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void dgvUrun_SelectionChanged(object sender, EventArgs e)
        {
            //Alttaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
            DataGridViewRow row = dgvUrun.CurrentRow;
            if (row != null)
            {
                CmbUrun.Tag = row.Cells["IslemId"].Value;
                //LblUrunAd.Text = row.Cells["Ürün Adı"].Value.ToString();
                //LblUrunAd.Tag = row.Cells["UrunID"].Value;
                //LblUrunFiyat.Text = row.Cells["Fiyat"].Value.ToString();
                NudUrunAdet.Value = (decimal)row.Cells["Adet"].Value;
                DtpIslemTarih.Text = row.Cells["Tarih"].Value.ToString();

                if (row.Cells["İşlem Tipi"].Value.ToString() == "Periyodik Kontrol")
                {
                    CmbUrun.SelectedIndex = 0;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Genel Bakım")
                {
                    CmbUrun.SelectedIndex = 1;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Arıza")
                {
                    CmbUrun.SelectedIndex = 2;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Diğer" || row.Cells["İşlem Tipi"].Value.ToString() != "Arıza" || row.Cells["İşlem Tipi"].Value.ToString() != "Genel Bakım" || row.Cells["İşlem Tipi"].Value.ToString() != "Periyodik Kontrol")
                {
                    CmbUrun.SelectedIndex = 3;
                    TxtIslemDiger.Visible = true;
                    LblIslemDiger.Visible = true;
                    TxtIslemDiger.Text = row.Cells["İşlem Tipi"].Value.ToString();
                }
            }
            else
            {
                CmbUrun.Tag = null;
            }
        }

        private void CmbUrun_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (CmbUrun.SelectedIndex == 3)
            {
                TxtIslemDiger.Visible = true;
                LblIslemDiger.Visible = true;
            }
            else
            {
                TxtIslemDiger.Visible = false;
                LblIslemDiger.Visible = false;
            }
        }

        private void BtnUrunEkle_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0)
            {
                if (CmbUrunKod.SelectedIndex.ToString() == null || CmbUrunKod.SelectedIndex.ToString() == "" || CmbUrunKod.SelectedIndex.ToString() == "-1")
                {
                    MessageBox.Show("Lütfen Ürün Kodu nu seçelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (CmbUrun.SelectedIndex == -1)
                {
                    MessageBox.Show("Lütfen İşlem Tipini Seçelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //Ekle
                    OleDbCommand cmd = new OleDbCommand("insert into Islem(FirmaId,IslemTip,[Tarih],UrunId,Adet) values(@fid,@islem,@tarih,@urunid,@adet)", baglanti);

                    cmd.Parameters.AddWithValue("@fid", (int)TxtFirmaAd.Tag);
                    if (CmbUrun.SelectedIndex == 0)
                    {
                        cmd.Parameters.AddWithValue("@islem", "Periyodik Kontrol");
                    }
                    else if (CmbUrun.SelectedIndex == 1)
                    {
                        cmd.Parameters.AddWithValue("@islem", "Genel Bakım");
                    }
                    else if (CmbUrun.SelectedIndex == 2)
                    {
                        cmd.Parameters.AddWithValue("@islem", "Arıza");
                    }
                    else if (CmbUrun.SelectedIndex == 3)
                    {
                        if (TxtIslemDiger.Text == "" || TxtIslemDiger.Text == null)
                        {
                            cmd.Parameters.AddWithValue("@islem", "Diğer");
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@islem", TxtIslemDiger.Text);
                        }
                    }
                    cmd.Parameters.AddWithValue("@tarih", DtpIslemTarih.Text);
                    cmd.Parameters.AddWithValue("@urunid", CmbUrunKod.SelectedValue);
                    cmd.Parameters.AddWithValue("@adet", NudUrunAdet.Text);

                    baglanti.Open();
                    cmd.ExecuteNonQuery();
                    baglanti.Close();

                    //Listele
                    Listele_Urun();

                    //
                    if (DgvUrunler.Rows.Count > 0)
                    {
                        for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                        {
                            if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                            {
                                LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                                LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                            }
                        }
                    }

                    //Temizle
                    Temizle_Urun();
                }
            }
        }

        private void BtnUrunTemizle_Click(object sender, EventArgs e)
        {
            Temizle_Urun();
        }

        private void CmbIslemTip_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0)
            {
                if (CmbIslemTip.SelectedIndex != 4 && CmbIslemTip.SelectedIndex != 3)
                {
                    string listele = "select IslemId, IslemTip as [İşlem Tipi], Tarih, Islem.UrunId as [UrunId], UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Adet, Fiyat from Islem INNER JOIN Urun ON Islem.UrunId = Urun.UrunId where FirmaId = " + TxtFirmaAd.Tag.ToString() + " and IslemTip = '" + CmbIslemTip.SelectedItem.ToString() + "' order by Tarih desc";
                    OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvUrun.DataSource = dt;
                    dgvUrun.Columns["IslemId"].Visible = dgvUrun.Columns["UrunId"].Visible = false;
                }
                else if (CmbIslemTip.SelectedIndex == 3)
                {
                    string listele = "select IslemId, IslemTip as [İşlem Tipi], Tarih, Islem.UrunId as [UrunId], UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Adet, Fiyat, [Adet] * [Fiyat] as [Tutar] from Islem INNER JOIN Urun ON Islem.UrunId = Urun.UrunId where FirmaId = " + TxtFirmaAd.Tag.ToString() + " and IslemTip <> 'Periyodik Kontrol' and IslemTip <> 'Genel Bakım' and IslemTip <> 'Arıza' order by Tarih desc";
                    OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvUrun.DataSource = dt;
                    dgvUrun.Columns["IslemId"].Visible = dgvUrun.Columns["UrunId"].Visible = false;
                    dgvUrun.Columns["Adet"].DefaultCellStyle.Alignment = dgvUrun.Columns["Fiyat"].DefaultCellStyle.Alignment = dgvUrun.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    string listele = "select IslemId, IslemTip as [İşlem Tipi], Tarih, Islem.UrunId as [UrunId], UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Adet, Fiyat, [Adet] * [Fiyat] as [Tutar] from Islem INNER JOIN Urun ON Islem.UrunId = Urun.UrunId where FirmaId = " + TxtFirmaAd.Tag.ToString() + " order by Tarih desc";
                    OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    dgvUrun.DataSource = dt;
                    dgvUrun.Columns["IslemId"].Visible = dgvUrun.Columns["UrunId"].Visible = false;
                    dgvUrun.Columns["Adet"].DefaultCellStyle.Alignment = dgvUrun.Columns["Fiyat"].DefaultCellStyle.Alignment = dgvUrun.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }

        private void Listele_Urun()
        {
            string listele = "select IslemId, Islem.UrunId as [UrunId], IslemTip as [İşlem Tipi], Tarih, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Adet, Fiyat, [Adet] * [Fiyat] as [Tutar] from Islem INNER JOIN Urun ON Islem.UrunId = Urun.UrunId where FirmaId = " + TxtFirmaAd.Tag.ToString() + " order by Tarih desc";

            OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            dgvUrun.DataSource = dt;
            dgvUrun.Columns["IslemId"].Visible = dgvUrun.Columns["UrunId"].Visible = false;
            dgvUrun.Columns["Adet"].DefaultCellStyle.Alignment = dgvUrun.Columns["Fiyat"].DefaultCellStyle.Alignment = dgvUrun.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            toplamtutar = 0;
            for (int i = 0; i < dgvUrun.Rows.Count; i++)
            {
                toplamtutar += (decimal)dgvUrun.Rows[i].Cells["Tutar"].Value;
            }
            LblToplamTutar.Text = String.Format("{0:N}\n", toplamtutar);
        }

        private void Temizle_Urun()
        {
            //CmbUrun.SelectedIndex = 0;
            DtpIslemTarih.Text = DateTime.Now.ToString("d/M/yyyy");
            NudUrunAdet.Value = 1;
            TxtIslemDiger.Text = null;
            TxtIslemDiger.Visible = false;
            LblIslemDiger.Visible = false;
            CmbUrun.Tag = null;
        }

        private void BtnUrunSil_Click(object sender, EventArgs e)
        {
            if (dgvUrun.Rows.Count > 0)
            {
                if (LblUrunAd.Text == null || LblUrunAd.Text == "")
                {
                    MessageBox.Show("Ürün seçili değil..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    int id = (int)dgvUrun.CurrentRow.Cells["IslemId"].Value;
                    //Açıklamalar   
                    int aId2 = dgvUrun.SelectedRows[0].Index;
                    string aId = "No";
                    aId = (aId2 + 1).ToString();

                    string aIslemTip = dgvUrun.SelectedCells[2].Value.ToString();

                    DateTime aTarih2 = (DateTime)dgvUrun.SelectedCells[3].Value;
                    string aTarih = aTarih2.ToString("dd/MM/yyyy");

                    DialogResult sonuc = MessageBox.Show("Sıralanan İşlemlerden Satır No: " + aId + "\n" + "İşlem Tipi: " + aIslemTip + "\n" + "Tarih: " + aTarih  + "\n" + "Ürün Kodu: " + dgvUrun.CurrentRow.Cells["Ürün Kodu"].Value.ToString() + "\n" + "Ürün Adı: " + dgvUrun.CurrentRow.Cells["Ürün Adı"].Value.ToString() + "\n" + "Adet: " + dgvUrun.CurrentRow.Cells["Adet"].Value.ToString() + "\n" + "Birim Fiyat: " + dgvUrun.CurrentRow.Cells["Fiyat"].Value.ToString() + "\n" + "\n" + "Silinsin mi?", "Sil?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (sonuc == DialogResult.Yes)
                    {
                        OleDbCommand cmd = new OleDbCommand("delete from Islem where IslemId=@iid", baglanti);

                        cmd.Parameters.AddWithValue("@iid", id);

                        baglanti.Open();
                        cmd.ExecuteNonQuery();
                        baglanti.Close();

                        //Listele
                        Listele_Urun();

                        //
                        if (DgvUrunler.Rows.Count > 0)
                        {
                            for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                            {
                                if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                                {
                                    LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                                    LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                                }
                            }
                        }

                        //Temizle                        
                        Temizle_Urun();
                    }
                }
            }
        }

        private void BtnUrunGuncelle_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0 && dgvUrun.Rows.Count > 0)
            {
                if (dgvUrun.SelectedCells[0].Value == null || CmbUrun.Tag == null)
                {
                    MessageBox.Show("İşlem seçili değil!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //Açıklamalar   
                    int aId2 = dgvUrun.SelectedRows[0].Index;
                    string aId = "No";
                    if (dgvUrun.SelectedCells[0].Value != null)
                    {
                        aId = (aId2 + 1).ToString();
                    }

                    string aIslemTip = "İşlem Tipi";
                    if (dgvUrun.SelectedCells[2].Value != null)
                    {
                        aIslemTip = dgvUrun.SelectedCells[2].Value.ToString();
                    }

                    string bIslemTip = "Yeni İşlem Tipi";
                    if (CmbUrun.SelectedIndex == 3)
                    {
                        if (TxtIslemDiger.Text == "" || TxtIslemDiger == null)
                        {
                            bIslemTip = "Diğer";
                        }
                        else
                        {
                            bIslemTip = TxtIslemDiger.Text;
                        }
                    }
                    else
                    {
                        bIslemTip = CmbUrun.SelectedItem.ToString();
                    }

                    DateTime aTarih2 = (DateTime)dgvUrun.SelectedCells[3].Value;
                    string aTarih = aTarih2.ToString("dd/MM/yyyy");

                    string aUrunAd = dgvUrun.SelectedCells[5].Value.ToString();
                    string aAdet = dgvUrun.SelectedCells[6].Value.ToString();
                    string aBirim = dgvUrun.SelectedCells[7].Value.ToString();

                    if (NudUrunAdet.Value == 0)
                    {
                        MessageBox.Show("Lütfen Adet girelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "İşlem Tipi: " + aIslemTip + "\n" + "Tarih: " + aTarih  + "\n" +"Ürün Adı: " + aUrunAd + "\n" + "Adet: " + aAdet + "\n" + "Birim Fiyat: " + aBirim + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + "İşlem Tipi: " + bIslemTip + "\n" + "Tarih: " + DtpIslemTarih.Text + "\n" + "Ürün Adı: " + LblUrunAd.Text + "\n" + "Adet: " + NudUrunAdet.Text + "\n" + "Birim Fiyat: " + LblUrunFiyat.Text + "\n" + "\n" + "Güncellensin mi?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (sonuc == DialogResult.Yes)
                        {
                            OleDbCommand cmd = new OleDbCommand("Update Islem set FirmaId=@fid,IslemTip=@islem,[Tarih]=@tarih,UrunId=@urunid,Adet=@adet where IslemId=@aid3", baglanti);

                            cmd.Parameters.AddWithValue("@fid", (int)TxtFirmaAd.Tag);
                            cmd.Parameters.AddWithValue("@islem", bIslemTip);
                            cmd.Parameters.AddWithValue("@tarih", DtpIslemTarih.Text);
                            cmd.Parameters.AddWithValue("@urunid", CmbUrunKod.SelectedValue);
                            cmd.Parameters.AddWithValue("@adet", NudUrunAdet.Text);
                            cmd.Parameters.AddWithValue("@aid3", (int)CmbUrun.Tag);

                            baglanti.Open();
                            cmd.ExecuteNonQuery();
                            baglanti.Close();

                            //Listele
                            Listele_Urun();

                            //
                            if (DgvUrunler.Rows.Count > 0)
                            {
                                for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                                {
                                    if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                                    {
                                        LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                                        LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                                    }
                                }
                            }

                            //Temizle                        
                            Temizle_Urun();
                        }
                    }
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtAraFirma.Text = null;
            if (tabControl1.SelectedIndex == 0)
            {
                cari = false;
                Listele_Firma();                
                GrpDuzenCari.Visible = false;
                GrpDuzenIslem.Visible = true;
                if (dgvFirma.Rows.Count > 0)
                {
                    DataGridViewRow row = dgvFirma.SelectedRows[0];
                    GrpBoxFirma.Text = row.Cells["FirmaAd"].Value.ToString();
                    TxtFirmaAd.Tag = row.Cells["FirmaId"].Value;
                    TxtFirmaAd.Text = row.Cells["FirmaAd"].Value.ToString();
                    LblTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    MskTxtTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    LblTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    MskTxtTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    LblJenNo.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    TxtJen.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    LblMotorNo.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    TxtMotor.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    labeltarih = (DateTime)row.Cells["FirmaTarih"].Value;
                    LblFirmaTarih.Text = labeltarih.ToString("d/M/yyyy");
                    DtpFirmaTarih.Text = row.Cells["FirmaTarih"].Value.ToString();
                    CmbIslemTip.SelectedIndex = 4;
                    Listele_Urun();
                }
                else if (dgvFirma.Rows.Count < 1)
                {
                    GrpBoxFirma.Text = "Firma Bilgileri";
                    LblFirmaTarih.Text = "Eklenme Tarihi";
                    LblTel1.Text = "İletişim 1";
                    LblTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblJenNo.Text = "";
                    LblMotorNo.Text = "";
                }

                if (DgvUrunler.Rows.Count > 0)
                {
                    for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                    {
                        if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                        {
                            LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                            LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                        }
                    }
                }
                else if (DgvUrunler.Rows.Count < 1)
                {
                    LblUrunAd.Text = "";
                    LblUrunFiyat.Text = "";
                }

                islem = true;
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                LblAlacak.Text = "0,00";
                LblBorc.Text = "0,00";
                islem = false;
                Listele_Kisi();                
                GrpDuzenCari.Visible = true;
                GrpDuzenCari.Location = new Point (3, 6);
                GrpDuzenIslem.Visible = false;
                if (dgvFirma.Rows.Count > 0)
                {
                    DataGridViewRow row = dgvFirma.SelectedRows[0];
                    GrpKisi.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Tag = row.Cells["KisiId"].Value;
                    LblCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    MskCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    LblCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    MskCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    LblCariAdres.Text = row.Cells["Adres"].Value.ToString();
                    TxtAdres.Text = row.Cells["Adres"].Value.ToString();
                    Cari_Listele();
                }
                else if (dgvFirma.Rows.Count < 1)
                {
                    LblCariTel1.Text = "İletişim 1";
                    LblCariTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblCariAdres.Text = "Adres";
                }
                cari = true;
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                OleDbDataAdapter adpu = new OleDbDataAdapter("select UrunId, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Fiyat from Urun order by UrunKod", baglanti);
                DataTable dtu = new DataTable();
                adpu.Fill(dtu);
                DgvUrunler.DataSource = dtu;
                DgvUrunler.Columns["UrunId"].Visible = false;
                DgvUrunler.Columns["Fiyat"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                //Firma Listesini ve İlgili Yerleri Getir
                cari = false;
                Listele_Firma();
                GrpDuzenCari.Visible = false;
                GrpDuzenIslem.Visible = true;
                if (dgvFirma.Rows.Count > 0)
                {
                    DataGridViewRow row = dgvFirma.SelectedRows[0];
                    GrpBoxFirma.Text = row.Cells["FirmaAd"].Value.ToString();
                    TxtFirmaAd.Tag = row.Cells["FirmaId"].Value;
                    TxtFirmaAd.Text = row.Cells["FirmaAd"].Value.ToString();
                    LblTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    MskTxtTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    LblTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    MskTxtTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    LblJenNo.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    TxtJen.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    LblMotorNo.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    TxtMotor.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    labeltarih = (DateTime)row.Cells["FirmaTarih"].Value;
                    LblFirmaTarih.Text = labeltarih.ToString("d/M/yyyy");
                    DtpFirmaTarih.Text = row.Cells["FirmaTarih"].Value.ToString();
                    CmbIslemTip.SelectedIndex = 4;
                    Listele_Urun();
                }
                else if (dgvFirma.Rows.Count < 1)
                {
                    GrpBoxFirma.Text = "Firma Bilgileri";
                    LblFirmaTarih.Text = "Eklenme Tarihi";
                    LblTel1.Text = "İletişim 1";
                    LblTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblJenNo.Text = "";
                    LblMotorNo.Text = "";
                }

                if (DgvUrunler.Rows.Count > 0)
                {
                    for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                    {
                        if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                        {
                            LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                            LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                        }
                    }
                }
                else if (DgvUrunler.Rows.Count < 1)
                {
                    LblUrunAd.Text = "";
                    LblUrunFiyat.Text = "";
                }

                islem = true;

                //Kontrol
                if (dgvFirma.Rows.Count > 0)
                {
                    OleDbDataAdapter adpk = new OleDbDataAdapter("select FirmaAd as [Firma], IslemTip as [İşlem Tipi], Tarih from Firma INNER JOIN Islem ON Firma.FirmaId = Islem.FirmaId where (IslemTip = 'Periyodik Kontrol' or IslemTip = 'Genel Bakım') order by  FirmaAd ASC, Tarih DESC", baglanti);
                    DataTable dtk = new DataTable();
                    adpk.Fill(dtk);
                    DgvKontrol.DataSource = dtk;

                    for (int i = 0; i < dgvFirma.Rows.Count; i++)
                    {
                        bool bulunduPer = false;
                        bool kontrolPer = false;
                        bool bulunduGen = false;
                        bool kontrolGen = false;
                        foreach (DataGridViewRow dGVRow in this.DgvKontrol.Rows)
                        {
                            if (dgvFirma.Rows[i].Cells[1].Value.ToString() == dGVRow.Cells[0].Value.ToString())
                            {
                                //Periyodik Kontrol
                                if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 < (double)NumPer.Value * 30) && bulunduPer == false)
                                {
                                    bulunduPer = true;
                                    kontrolPer = true;
                                    //Satır Gizlemek için ufak bir ayar lazım
                                    CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                    currencyManager1.SuspendBinding();
                                    dGVRow.Visible = false;
                                    currencyManager1.ResumeBinding();
                                }
                                else if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 >= ((double)NumPer.Value * 30)) && bulunduPer == false)
                                {
                                    bulunduPer = true;
                                    dGVRow.DefaultCellStyle.BackColor = Color.DarkOrange;
                                }
                                else if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && (kontrolPer == true || bulunduPer == true))
                                {
                                    //Satır Gizlemek için ufak bir ayar lazım
                                    CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                    currencyManager1.SuspendBinding();
                                    dGVRow.Visible = false;
                                    currencyManager1.ResumeBinding();
                                }

                                //Genel Bakım
                                if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 < (double)NumGen.Value * 30) && bulunduGen == false)
                                {
                                    bulunduGen = true;
                                    kontrolGen = true;
                                    //Satır Gizlemek için ufak bir ayar lazım
                                    CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                    currencyManager1.SuspendBinding();
                                    dGVRow.Visible = false;
                                    currencyManager1.ResumeBinding();
                                }
                                else if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 >= ((double)NumGen.Value * 30)) && bulunduGen == false)
                                {
                                    bulunduGen = true;
                                    dGVRow.DefaultCellStyle.BackColor = Color.DarkRed;
                                    dGVRow.DefaultCellStyle.ForeColor = Color.White;
                                }
                                else if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && (kontrolGen == true || bulunduGen == true))
                                {
                                    //Satır Gizlemek için ufak bir ayar lazım
                                    CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                    currencyManager1.SuspendBinding();
                                    dGVRow.Visible = false;
                                    currencyManager1.ResumeBinding();
                                }
                            }
                        }
                       
                    }
                }
                else
                {
                    DgvKontrol.DataSource = null;
                }
            }
        }

        private void Listele_Kisi()
        {
            OleDbDataAdapter adp = new OleDbDataAdapter("select * from Kisi order by Ad", baglanti);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            dgvFirma.DataSource = dt;
            dgvFirma.Columns["KisiId"].Visible = dgvFirma.Columns["Tel1"].Visible = dgvFirma.Columns["Tel2"].Visible = dgvFirma.Columns["Adres"].Visible = false;
            dgvFirma.Columns["Ad"].Width = dgvFirma.Width - 20;
        }

        private void BtnEkleKisi_Click(object sender, EventArgs e)
        {
            if (TxtCariAd.Text == null || TxtCariAd.Text == "")
            {
                MessageBox.Show("Lütfen Ad girelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Ekle
                OleDbCommand cmd = new OleDbCommand("insert into Kisi(Ad,Tel1,Tel2,Adres) values(@cariad,@caritel1,@caritel2,@cariadres)", baglanti);

                cmd.Parameters.AddWithValue("@cariad", TxtCariAd.Text);
                cmd.Parameters.AddWithValue("@caritel1", MskCariTel1.Text);
                cmd.Parameters.AddWithValue("@caritel2", MskCariTel2.Text);
                cmd.Parameters.AddWithValue("@cariadres", TxtAdres.Text);

                baglanti.Open();
                cmd.ExecuteNonQuery();
                baglanti.Close();

                //Listele
                Listele_Kisi();

                //Temizle
                TxtCariAd.Text = null;
                MskCariTel1.Text = null;
                MskCariTel2.Text = null;
                TxtAdres.Text = null;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TxtCariAd.Text = null;
            MskCariTel1.Text = null;
            MskCariTel2.Text = null;
            TxtAdres.Text = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (TxtCariAd.Text == "" || TxtCariAd.Text == null)
            {
                MessageBox.Show("Seçim yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                int id = (int)TxtCariAd.Tag;
                //Açıklamalar   
                int aId2 = dgvFirma.SelectedRows[0].Index;
                string aId = (aId2 + 1).ToString();

                DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Ad: " + GrpKisi.Text + "\n" + "\n" + "Tel-1: " + LblCariTel1.Text + "\n" + "Tel-2: " + LblCariTel2.Text + "\n" + "Adres: " + LblCariAdres.Text + "\n" + "\n" + "Silinsin mi?", "Sil?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (sonuc == DialogResult.Yes)
                {
                    OleDbCommand cmd = new OleDbCommand("delete from Kisi where KisiId=@kid", baglanti);
                    OleDbCommand cmd2 = new OleDbCommand("delete from Cari where KisiId=@kid2", baglanti);

                    cmd.Parameters.AddWithValue("@kid", id);
                    cmd2.Parameters.AddWithValue("@kid2", id);

                    baglanti.Open();
                    cmd.ExecuteNonQuery();
                    cmd2.ExecuteNonQuery();
                    baglanti.Close();

                    //Listele
                    Listele_Kisi();

                    //Temizle                        
                    TxtCariAd.Text = null;
                    MskCariTel1.Text = null;
                    MskCariTel2.Text = null;
                    TxtAdres.Text = null;
                }
            }
        }

        private void BtnKisiGuncelle_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count < 1)
            {
                MessageBox.Show("Liste boş!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (TxtCariAd.Text == "")
            {
                MessageBox.Show("Seçili veri yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Açıklamalar   
                int aId2 = dgvFirma.SelectedRows[0].Index;
                string aId = "No";
                if (dgvFirma.SelectedCells[1].Value == null)
                {
                    aId = aId2.ToString();
                }
                else if (dgvFirma.SelectedCells[1].Value != null)
                {
                    aId = (aId2 + 1).ToString();
                }

                string aAd = dgvFirma.SelectedCells[1].Value.ToString();

                string aTel1 = "";
                if (dgvFirma.SelectedCells[2].Value != null)
                {
                    aTel1 = dgvFirma.SelectedCells[2].Value.ToString();
                }

                string aTel2 = "";
                if (dgvFirma.SelectedCells[3].Value != null)
                {
                    aTel2 = dgvFirma.SelectedCells[3].Value.ToString();
                }

                string aAdres = "";
                if (dgvFirma.SelectedCells[4].Value != null)
                {
                    aAdres = dgvFirma.SelectedCells[4].Value.ToString();
                }


                if (TxtCariAd.Text == null || TxtCariAd.Text == "")
                {
                    MessageBox.Show("Ad girelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Ad: " + aAd + "\n" + "Tel-1: " + aTel1 + "\n" + "Tel-2: " + aTel2 + "\n" + "Adres: " + aAdres + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + "Ad: " + TxtCariAd.Text + "\n" + "Tel-1: " + MskCariTel1.Text + "\n" + "Tel-2: " + MskCariTel2.Text + "\n" + "Adres: " + TxtAdres.Text + "\n" + "\n" + "Güncellensin mi?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (sonuc == DialogResult.Yes)
                    {
                        OleDbCommand cmd = new OleDbCommand("Update Kisi set Ad=@cariad,Tel1=@caritel1,Tel2=@caritel2,Adres=@cariadres where KisiId=@kid", baglanti);

                        cmd.Parameters.AddWithValue("@cariad", TxtCariAd.Text);
                        cmd.Parameters.AddWithValue("@caritel1", MskCariTel1.Text);
                        cmd.Parameters.AddWithValue("@caritel2", MskCariTel2.Text);
                        cmd.Parameters.AddWithValue("@cariadres", TxtAdres.Text);
                        cmd.Parameters.AddWithValue("@kid", (int)TxtCariAd.Tag);

                        baglanti.Open();
                        cmd.ExecuteNonQuery();
                        baglanti.Close();

                        //Listele
                        Listele_Kisi();

                        //Temizle
                        TxtCariAd.Text = null;
                        MskCariTel1.Text = null;
                        MskCariTel2.Text = null;
                        TxtAdres.Text = null;
                    }
                }
            }
        }

        private void TxtCariTutar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // decimal için virgül
            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }

        private void DgvCari_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.DgvCari.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.DgvCari.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void Cari_Listele()
        {
            string listele = "select CariId, Tarih, Durum, Tutar, Aciklama as [Açıklama] from Cari where KisiId = " + TxtCariAd.Tag.ToString() + " order by Tarih desc";
            OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            DgvCari.DataSource = dt;
            DgvCari.Columns["CariId"].Visible = false;
            DgvCari.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            if (DgvCari.Rows.Count > 0)
            {
                decimal alacak = 0;
                decimal alindi = 0;
                decimal borc = 0;
                decimal odendi = 0;
                for (int i = 0; i < DgvCari.Rows.Count; i++)
                {
                    if ((string)DgvCari.Rows[i].Cells[2].Value == "Alacak")
                    {
                        alacak += (decimal)DgvCari.Rows[i].Cells[3].Value;
                    }
                    else if ((string)DgvCari.Rows[i].Cells[2].Value == "Alındı")
                    {
                        alindi += (decimal)DgvCari.Rows[i].Cells[3].Value;
                    }
                    else if ((string)DgvCari.Rows[i].Cells[2].Value == "Borç")
                    {
                        borc += (decimal)DgvCari.Rows[i].Cells[3].Value;
                    }
                    else if ((string)DgvCari.Rows[i].Cells[2].Value == "Ödendi")
                    {
                        odendi += (decimal)DgvCari.Rows[i].Cells[3].Value;
                    }
                }
                LblAlacak.Text = String.Format("{0:N}\n", alacak - alindi);
                LblBorc.Text = String.Format("{0:N}\n", borc - odendi);
            }
        }

        private void CmbCariListele_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0)
            {
                if (CmbCariListele.SelectedIndex != 4)
                {
                    var d = TxtCariAd.Tag.ToString();
                    string listele = "select CariId, Tarih, Durum, Tutar, Aciklama as [Açıklama] from Cari where KisiId = " + TxtCariAd.Tag.ToString() + " and Durum = '" + CmbCariListele.SelectedItem.ToString() + "' order by Tarih desc";
                    OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    DgvCari.DataSource = dt;
                    DgvCari.Columns["CariId"].Visible = false;
                    DgvCari.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    var d = TxtCariAd.Tag.ToString();
                    string listele = "select CariId, Tarih, Durum, Tutar, Aciklama as [Açıklama] from Cari where KisiId = " + TxtCariAd.Tag.ToString() + " order by Tarih desc";
                    OleDbDataAdapter adp = new OleDbDataAdapter(listele, baglanti);
                    DataTable dt = new DataTable();
                    adp.Fill(dt);
                    DgvCari.DataSource = dt;
                    DgvCari.Columns["CariId"].Visible = false;
                    DgvCari.Columns["Tutar"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                } 
            }
        }

        private void DgvCari_SelectionChanged(object sender, EventArgs e)
        {
            //Alttaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
            DataGridViewRow row = DgvCari.CurrentRow;
            if (row != null)
            {
                TxtCariTutar.Tag = row.Cells["CariId"].Value;
                DtpCariTarih.Text = row.Cells["Tarih"].Value.ToString();
                TxtCariTutar.Text = row.Cells["Tutar"].Value.ToString();
                TxtAciklama.Text = row.Cells["Açıklama"].Value.ToString();

                if (row.Cells["Durum"].Value.ToString() == "Alacak")
                {
                    CmbCari.SelectedIndex = 0;
                }
                else if (row.Cells["Durum"].Value.ToString() == "Alındı")
                {
                    CmbCari.SelectedIndex = 1;
                }
                else if (row.Cells["Durum"].Value.ToString() == "Borç")
                {
                    CmbCari.SelectedIndex = 2;
                }
                else
                {
                    CmbCari.SelectedIndex = 3;
                }
            }
            else
            {
                TxtCariTutar.Tag = null;
            }
        }

        private void BtnCariEkle_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0)
            {
                if (CmbCari.SelectedIndex == -1)
                {
                    MessageBox.Show("Lütfen Durum u seçiniz..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //Ekle
                    OleDbCommand cmd = new OleDbCommand("insert into Cari(KisiId,[Tarih],Durum,Tutar,Aciklama) values(@kid,@tarih,@durum,@tutar,@aciklama)", baglanti);

                    cmd.Parameters.AddWithValue("@kid", (int)TxtCariAd.Tag);
                    cmd.Parameters.AddWithValue("@tarih", DtpCariTarih.Text);
                    if (CmbCari.SelectedIndex == 0)
                    {
                        cmd.Parameters.AddWithValue("@durum", "Alacak");
                    }
                    else if (CmbCari.SelectedIndex == 1)
                    {
                        cmd.Parameters.AddWithValue("@durum", "Alındı");
                    }
                    else if (CmbCari.SelectedIndex == 2)
                    {
                        cmd.Parameters.AddWithValue("@durum", "Borç");
                    }
                    else if (CmbCari.SelectedIndex == 3)
                    {
                        cmd.Parameters.AddWithValue("@durum", "Ödendi");
                    }

                    if (TxtCariTutar.Text == null || TxtCariTutar.Text == "")
                    {
                        cmd.Parameters.AddWithValue("@tutar", 0);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@tutar", TxtCariTutar.Text);
                    }

                    cmd.Parameters.AddWithValue("@aciklama", TxtAciklama.Text);

                    baglanti.Open();
                    cmd.ExecuteNonQuery();
                    baglanti.Close();

                    //Listele
                    Cari_Listele();

                    //Temizle
                    CmbCari.SelectedIndex = 0;
                    TxtCariTutar.Text = "0";
                    DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                    TxtAciklama.Text = "";
                }
            }
        }

        private void BtnCariSil_Click(object sender, EventArgs e)
        {
            if (DgvCari.Rows.Count < 1 || TxtCariTutar.Tag == null)
            {
                MessageBox.Show("Seçim yok..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                int id = (int)TxtCariTutar.Tag;
                //Açıklamalar   
                int aId2 = DgvCari.SelectedRows[0].Index;
                string aId = (aId2 + 1).ToString();

                DateTime aTarih2 = (DateTime)DgvCari.SelectedCells[1].Value;
                string aTarih = aTarih2.ToString("dd/MM/yyyy");

                DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Tarih: " + aTarih + "\n" + "Durum: " + DgvCari.SelectedCells[2].Value.ToString() + "\n" + "Tutar: " + DgvCari.SelectedCells[3].Value.ToString() + "\n" + "Açıklama: " + DgvCari.SelectedCells[4].Value.ToString() + "\n" + "\n" + "Silinsin mi ? ", "Sil ? ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (sonuc == DialogResult.Yes)
                {
                    OleDbCommand cmd = new OleDbCommand("delete from Cari where CariId=@cid", baglanti);

                    cmd.Parameters.AddWithValue("@cid", id);

                    baglanti.Open();
                    cmd.ExecuteNonQuery();
                    baglanti.Close();

                    //Listele
                    Cari_Listele();

                    //Temizle
                    CmbCari.SelectedIndex = 0;
                    TxtCariTutar.Text = "0";
                    DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                    TxtAciklama.Text = "";
                }
            }
        }

        private void BtnCariTemizle_Click(object sender, EventArgs e)
        {
            CmbCari.SelectedIndex = 0;
            TxtCariTutar.Text = "0";
            DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
            TxtAciklama.Text = "";
        }

        private void BtnCariGuncelle_Click(object sender, EventArgs e)
        {
            if (DgvCari.Rows.Count < 1)
            {
                MessageBox.Show("Seçim yok!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                int id = (int)TxtCariTutar.Tag;
                //Açıklamalar   
                int aId2 = DgvCari.SelectedRows[0].Index;
                string aId = (aId2 + 1).ToString(); ;

                DateTime aTarih2 = (DateTime)DgvCari.SelectedCells[1].Value;
                string aTarih = aTarih2.ToString("dd/MM/yyyy");

                string aDurum = DgvCari.SelectedCells[2].Value.ToString();
                string aTutar = DgvCari.SelectedCells[3].Value.ToString();
                string aAciklama = DgvCari.SelectedCells[4].Value.ToString();

                DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Tarih: " + aTarih + "\n" + "Durum: " + aDurum + "\n" + "Tutar: " + aTutar + "\n" + "Açıklama: " + aAciklama + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + "Tarih: " + DtpCariTarih.Text + "\n" + "Durum: " + CmbCari.SelectedIndex.ToString() + "\n" + "Tutar: " + TxtCariTutar.Text + "\n" + "Açıklama: " + TxtAciklama.Text + "\n" + "\n" + "Güncellensin mi?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (sonuc == DialogResult.Yes)
                {
                    OleDbCommand cmd = new OleDbCommand("Update Cari set KisiId=@kid, [Tarih]=@tarih, Durum=@durum, Tutar=@tutar, Aciklama=@aciklama where CariId=@cid", baglanti);

                    cmd.Parameters.AddWithValue("@kid", TxtCariAd.Tag);
                    cmd.Parameters.AddWithValue("@tarih", DtpCariTarih.Text);
                    cmd.Parameters.AddWithValue("@durum", CmbCari.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@tutar", TxtCariTutar.Text);
                    cmd.Parameters.AddWithValue("@aciklama", TxtAciklama.Text);

                    cmd.Parameters.AddWithValue("@cid", (int)TxtCariTutar.Tag);

                    baglanti.Open();
                    cmd.ExecuteNonQuery();
                    baglanti.Close();

                    //Listele
                    Cari_Listele();

                    //Temizle                        
                    CmbCari.SelectedIndex = 0;
                    TxtCariTutar.Text = "0";
                    DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                    TxtAciklama.Text = "";
                }
            }
        
        }

        private void dgvFirma_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (islem)
            {
                if (dgvFirma.Rows.Count < 1)
                {
                    GrpBoxFirma.Text = "Firma Bilgileri";
                    LblFirmaTarih.Text = "Eklenme Tarihi";
                    LblTel1.Text = "İletişim 1";
                    LblTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblJenNo.Text = "";
                    LblMotorNo.Text = "";
                }
                //Soldaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
                DataGridViewRow row = dgvFirma.CurrentRow;
                if (row != null)
                {
                    GrpBoxFirma.Text = row.Cells["FirmaAd"].Value.ToString();
                    TxtFirmaAd.Tag = row.Cells["FirmaId"].Value;
                    TxtFirmaAd.Text = row.Cells["FirmaAd"].Value.ToString();
                    LblTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    MskTxtTel1.Text = row.Cells["FirmaTel1"].Value.ToString();
                    LblTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    MskTxtTel2.Text = row.Cells["FirmaTel2"].Value.ToString();
                    LblJenNo.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    TxtJen.Text = row.Cells["FirmaJenNo"].Value.ToString();
                    LblMotorNo.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    TxtMotor.Text = row.Cells["FirmaMotorNo"].Value.ToString();
                    labeltarih = (DateTime)row.Cells["FirmaTarih"].Value;
                    LblFirmaTarih.Text = labeltarih.ToString("d/M/yyyy");
                    DtpFirmaTarih.Text = row.Cells["FirmaTarih"].Value.ToString();
                    CmbIslemTip.SelectedIndex = 4;

                    Listele_Urun();
                    /*
                    if (dgvUrun.Rows.Count < 1)
                    {
                        Temizle_Urun();
                    }
                    */
                }
            }
            else if (cari)
            {
                if (dgvFirma.Rows.Count < 1)
                {
                    LblCariTel1.Text = "İletişim 1";
                    LblCariTel2.Text = "İletişim 2";
                    GrpKisi.Text = "Ad";
                    LblCariAdres.Text = "Adres";
                }
                DataGridViewRow row = dgvFirma.CurrentRow;
                if (row != null)
                {
                    GrpKisi.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Text = row.Cells["Ad"].Value.ToString();
                    TxtCariAd.Tag = row.Cells["KisiId"].Value;
                    LblCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    MskCariTel1.Text = row.Cells["Tel1"].Value.ToString();
                    LblCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    MskCariTel2.Text = row.Cells["Tel2"].Value.ToString();
                    LblCariAdres.Text = row.Cells["Adres"].Value.ToString();
                    TxtAdres.Text = row.Cells["Adres"].Value.ToString();
                    CmbCariListele.SelectedIndex = 4;

                    Cari_Listele();

                    if (DgvCari.Rows.Count < 1)
                    {
                        Temizle_Urun();
                        CmbCari.SelectedIndex = 0;
                        TxtCariTutar.Text = "0";
                        DtpCariTarih.Text = DateTime.Now.ToString("d/M/yyyy");
                        TxtAciklama.Text = "";
                    }
                }
            }
        }

        private void dgvUrun_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //Alttaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
            DataGridViewRow row = dgvUrun.CurrentRow;
            if (row != null)
            {
                CmbUrun.Tag = row.Cells["IslemId"].Value;
                //LblUrunAd.Tag = row.Cells["UrunID"].Value;
                TxtUrunAd.Text = row.Cells["Ürün Adı"].Value.ToString();
                TxtUrunKod.Text = row.Cells["Ürün Kodu"].Value.ToString();
                //LblUrunFiyat.Text = row.Cells["Fiyat"].Value.ToString();
                NudUrunAdet.Value = (decimal)row.Cells["Adet"].Value;
                DtpIslemTarih.Text = row.Cells["Tarih"].Value.ToString();
                CmbUrunKod.SelectedValue = row.Cells["UrunId"].Value;

                if (row.Cells["İşlem Tipi"].Value.ToString() == "Periyodik Kontrol")
                {
                    CmbUrun.SelectedIndex = 0;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Genel Bakım")
                {
                    CmbUrun.SelectedIndex = 1;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Arıza")
                {
                    CmbUrun.SelectedIndex = 2;
                    TxtIslemDiger.Visible = false;
                    LblIslemDiger.Visible = false;
                    TxtIslemDiger.Text = "";
                }
                else if (row.Cells["İşlem Tipi"].Value.ToString() == "Diğer" || row.Cells["İşlem Tipi"].Value.ToString() != "Arıza" || row.Cells["İşlem Tipi"].Value.ToString() != "Genel Bakım" || row.Cells["İşlem Tipi"].Value.ToString() != "Periyodik Kontrol")
                {
                    CmbUrun.SelectedIndex = 3;
                    TxtIslemDiger.Visible = true;
                    LblIslemDiger.Visible = true;
                    TxtIslemDiger.Text = row.Cells["İşlem Tipi"].Value.ToString();
                }
            }
            else
            {
                CmbUrun.Tag = null;
            }
        }

        private void DgvCari_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //Alttaki listede Firma seçildiğinde ilgili yerlere değer atıyoruz.
            DataGridViewRow row = DgvCari.CurrentRow;
            if (row != null)
            {
                TxtCariTutar.Tag = row.Cells["CariId"].Value;
                DtpCariTarih.Text = row.Cells["Tarih"].Value.ToString();
                TxtCariTutar.Text = row.Cells["Tutar"].Value.ToString();
                TxtAciklama.Text = row.Cells["Açıklama"].Value.ToString();

                if (row.Cells["Durum"].Value.ToString() == "Alacak")
                {
                    CmbCari.SelectedIndex = 0;
                }
                else if (row.Cells["Durum"].Value.ToString() == "Alındı")
                {
                    CmbCari.SelectedIndex = 1;
                }
                else if (row.Cells["Durum"].Value.ToString() == "Borç")
                {
                    CmbCari.SelectedIndex = 2;
                }
                else
                {
                    CmbCari.SelectedIndex = 3;
                }
            }
            else
            {
                TxtCariTutar.Tag = null;
            }
        }

        private void BtnIslemExcel_Click(object sender, EventArgs e)
        {
            if (dgvUrun.Rows.Count > 0)
            {
                //ClosedXml - Version - 0.93.1
                //Dosya ismi
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    FileName = GrpBoxFirma.Text + " - İşlemler (" + DateTime.Now.ToString("dd-MM-yyyy") + ")"
                };

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable dt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };
                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";

                    //Diğer Tüm Sütunları Ekle
                    foreach (DataGridViewColumn column in dgvUrun.Columns)
                    {
                        dt.Columns.Add(column.HeaderText, column.ValueType);
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in dgvUrun.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //IslemiId, UrunId Sütunları Kaldır
                    dt.Columns.RemoveAt(1);
                    dt.Columns.RemoveAt(1);

                    //Toplam Satırı Ekle
                    dt.Rows.Add();
                    DataRow rowToplam = dt.NewRow();
                    dt.Rows.Add(rowToplam);
                    rowToplam[7] = toplamtutar;

                    //Excel Sayfasına ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, GrpBoxFirma.Text);
                    }

                    //Access Veritabanında "Tutar" Sayı:Ondalık, Ölçek:2, Ondalık Basamaklar:2 - . ve , leri getiriyoruz "Tutar" a
                    workbook.Worksheet(GrpBoxFirma.Text).Column(7).Style.NumberFormat.NumberFormatId = workbook.Worksheet(GrpBoxFirma.Text).Column(8).Style.NumberFormat.NumberFormatId = 4;

                    //Fazla No yazanları siliyoruz
                    workbook.Worksheet(GrpBoxFirma.Text).Cell(dgvUrun.Rows.Count + 2, 1).Value = "";
                    workbook.Worksheet(GrpBoxFirma.Text).Cell(dgvUrun.Rows.Count + 3, 1).Value = "";

                    //Sütun Genişliklerini ayarla
                    workbook.Worksheet(GrpBoxFirma.Text).Columns().AdjustToContents();
                    workbook.Worksheet(GrpBoxFirma.Text).Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet(GrpBoxFirma.Text).Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //Tarih, Başlıklar ekle
                    workbook.Worksheet(GrpBoxFirma.Text).PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet(GrpBoxFirma.Text).PageSetup.Footer.Right.AddText("kastamonuelektrik.com");
                    workbook.Worksheet(GrpBoxFirma.Text).PageSetup.Header.Left.AddText(GrpBoxFirma.Text + " | " + LblTel1.Text + " | " + LblTel2.Text);
                    workbook.Worksheet(GrpBoxFirma.Text).PageSetup.Footer.Left.AddText("Jeneratör: " + LblJenNo.Text + " | Motor: " + LblMotorNo.Text);

                    //1 Sayfaya Sığmazsa
                    if ((workbook.Worksheet(GrpBoxFirma.Text).Column(2).Width + workbook.Worksheet(GrpBoxFirma.Text).Column(4).Width + workbook.Worksheet(GrpBoxFirma.Text).Column(5).Width + workbook.Worksheet(GrpBoxFirma.Text).Column(7).Width + workbook.Worksheet(GrpBoxFirma.Text).Column(8).Width) > 58)
                    {
                        workbook.Worksheet(GrpBoxFirma.Text).PageSetup.FitToPages(1, 2);
                    }

                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            break;
                        }
                    } while (true);
                }
            }
        }

        private void BtnCariExcel_Click(object sender, EventArgs e)
        {
            if (DgvCari.Rows.Count > 0)
            {
                //ClosedXml - Version - 0.93.1
                string fileName;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    Title = "To Excel",
                    FileName = GrpKisi.Text + " - Cari (" + DateTime.Now.ToString("dd-MM-yyyy") + ")"
                };

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();

                    //DataTable Oluştur
                    DataTable dt = new DataTable();

                    //No Sütunu Ekle
                    DataColumn columno = new DataColumn
                    {
                        DataType = System.Type.GetType("System.Int32"),
                        AutoIncrement = true,
                        AutoIncrementSeed = 1,
                        AutoIncrementStep = 1
                    };
                    dt.Columns.Add(columno);
                    dt.Columns["Column1"].ColumnName = "No";

                    //Diğer Tüm Sütunları Ekle
                    foreach (DataGridViewColumn column in DgvCari.Columns)
                    {
                        dt.Columns.Add(column.HeaderText, column.ValueType);
                    }

                    //Satırları Ekle
                    foreach (DataGridViewRow row in DgvCari.Rows)
                    {
                        dt.Rows.Add();
                        //Hücreleri Ekle
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex + 1] = cell.Value;
                            }
                        }
                    }

                    //CariId Sütunu Kaldır
                    dt.Columns.RemoveAt(1);

                    //Toplamlar Ekle

                    //Alacak Toplam Hesapla

                    decimal alacak = 0;
                    for (int j = 0; j < DgvCari.Rows.Count + 1; ++j)
                    {
                        if (j < DgvCari.Rows.Count && DgvCari[2, j].Value.ToString() == "Alacak")
                        {
                            alacak += Convert.ToDecimal(DgvCari.Rows[j].Cells[3].Value);
                        }
                    }
                    decimal alindi = 0;
                    for (int j = 0; j < DgvCari.Rows.Count + 1; ++j)
                    {
                        if (j < DgvCari.Rows.Count && DgvCari[2, j].Value.ToString() == "Alındı")
                        {
                            alindi += Convert.ToDecimal(DgvCari.Rows[j].Cells[3].Value);
                        }
                    }
                    decimal borc = 0;
                    for (int j = 0; j < DgvCari.Rows.Count + 1; ++j)
                    {
                        if (j < DgvCari.Rows.Count && DgvCari[2, j].Value.ToString() == "Borç")
                        {
                            borc += Convert.ToDecimal(DgvCari.Rows[j].Cells[3].Value);
                        }
                    }
                    decimal odendi = 0;
                    for (int j = 0; j < DgvCari.Rows.Count + 1; ++j)
                    {
                        if (j < DgvCari.Rows.Count && DgvCari[2, j].Value.ToString() == "Ödendi")
                        {
                            odendi += Convert.ToDecimal(DgvCari.Rows[j].Cells[3].Value);
                        }
                    }
                    int satir = 1;
                    dt.Rows.Add();
                    if (alacak != 0)
                    {
                        DataRow rowAlacak = dt.NewRow();
                        dt.Rows.Add(rowAlacak);
                        rowAlacak[2] = "Alacaklar :";
                        rowAlacak[3] = alacak;
                        satir += 1;
                    }
                    if (alindi != 0)
                    {
                        DataRow rowAlindi = dt.NewRow();
                        dt.Rows.Add(rowAlindi);
                        rowAlindi[2] = "Alındılar :";
                        rowAlindi[3] = alindi;
                        satir += 1;
                    }
                    if (alacak != 0 || alindi != 0)
                    {
                        DataRow rowKalana1 = dt.NewRow();
                        dt.Rows.Add(rowKalana1);
                        rowKalana1[2] = "Kalan :";
                        rowKalana1[3] = alacak - alindi;
                        dt.Rows.Add();
                        satir += 1;
                    }
                    if (borc != 0)
                    {
                        DataRow rowBorc = dt.NewRow();
                        dt.Rows.Add(rowBorc);
                        rowBorc[2] = "Borçlar :";
                        rowBorc[3] = borc;
                        satir += 1;
                    }
                    if (odendi != 0)
                    {
                        DataRow rowOdendi = dt.NewRow();
                        dt.Rows.Add(rowOdendi);
                        rowOdendi[2] = "Ödemeler :";
                        rowOdendi[3] = odendi;
                        satir += 1;
                    }
                    if (borc != 0 || odendi != 0)
                    {
                        DataRow rowKalan2 = dt.NewRow();
                        dt.Rows.Add(rowKalan2);
                        rowKalan2[2] = "Kalan :";
                        rowKalan2[3] = borc - odendi;
                        satir += 1;
                    }

                    //Excel Sayfasına ekle.
                    using (workbook)
                    {
                        workbook.Worksheets.Add(dt, GrpKisi.Text);
                    }

                    //Access Veritabanında "Tutar" Sayı:Ondalık, Ölçek:2, Ondalık Basamaklar:2 - . ve , leri getiriyoruz "Tutar" a
                    workbook.Worksheet(GrpKisi.Text).Column(4).Style.NumberFormat.NumberFormatId = 4;

                    //No Sütununda fazladan yazılan değerleri siliyor, Eklenen Satırları "Bold" yapıyor, ve alacak vs sağ hizala
                    for (int i = 0; i < satir + 1; i++)
                    {
                        workbook.Worksheet(GrpKisi.Text).Cell(DgvCari.Rows.Count + i + 2, 1).Value = "";
                        workbook.Worksheet(GrpKisi.Text).Cell(DgvCari.Rows.Count + i + 2, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        workbook.Worksheet(GrpKisi.Text).Row(DgvCari.Rows.Count + i + 2).Style.Font.SetBold();
                    }

                    //Sütun Genişliklerini ayarla
                    workbook.Worksheet(GrpKisi.Text).Columns().AdjustToContents();
                    workbook.Worksheet(GrpKisi.Text).Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Worksheet(GrpKisi.Text).Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    //Tarih, Başlıklar ekle
                    workbook.Worksheet(GrpKisi.Text).PageSetup.Header.Right.AddText(string.Format("{0:dd/MM/yyyy}", DateTime.Now));
                    workbook.Worksheet(GrpKisi.Text).PageSetup.Header.Left.AddText(GrpKisi.Text + " | " + LblCariTel1.Text + " | " + LblCariTel2.Text);
                    workbook.Worksheet(GrpKisi.Text).PageSetup.Footer.Left.AddText("Adres: " + LblCariAdres.Text);
                    
                    //1 Sayfaya Sığmazsa
                    if (workbook.Worksheet(GrpKisi.Text).Column(4).Width + workbook.Worksheet(GrpKisi.Text).Column(5).Width > 58)
                    {
                        workbook.Worksheet(GrpKisi.Text).PageSetup.FitToPages(1, 2);
                    }
                    
                    //Kaydet
                    do
                    {
                        try
                        {
                            workbook.SaveAs(fileName);
                            MessageBox.Show("Excel dosyası kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        }
                        catch (System.IO.IOException)
                        {
                            MessageBox.Show("Kayıt yapılamadı! Kaydetmeye çalıştığınız dosya açık olabilir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            break;
                        }
                    } while (true);
                }
            }
        }

        private void DgvUrunler_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = DgvUrunler.CurrentRow;
            if (row != null)
            {
                TxtUrunKod.Tag = row.Cells["UrunId"].Value;
                TxtUrunKod.Text = row.Cells["Ürün Kodu"].Value.ToString();
                TxtUrunAd.Text = row.Cells["Ürün Adı"].Value.ToString();
                TxtUrunFiyat.Text = row.Cells["Fiyat"].Value.ToString();
            }
        }

        private void DgvUrunler_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.DgvUrunler.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.DgvUrunler.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void DgvUrunler_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow row = DgvUrunler.CurrentRow;
            if (row != null)
            {
                TxtUrunKod.Tag = row.Cells["UrunId"].Value;
                TxtUrunKod.Text = row.Cells["Ürün Kodu"].Value.ToString();
                TxtUrunAd.Text = row.Cells["Ürün Adı"].Value.ToString();
                TxtUrunFiyat.Text = row.Cells["Fiyat"].Value.ToString();
            }
        }

        private void TxtUrunFiyat_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // decimal için virgül
            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }

        private void BtnUTemizle_Click(object sender, EventArgs e)
        {
            TxtUrunKod.Tag = null;
            TxtUrunKod.Text = null;
            TxtUrunAd.Text = null;
            TxtUrunFiyat.Text = "0";
        }

        private void BtnUEkle_Click(object sender, EventArgs e)
        {
            if (TxtUrunKod.Text == null || TxtUrunKod.Text == "" || TxtUrunAd.Text == null || TxtUrunAd.Text == "")
            {
                MessageBox.Show("Lütfen ilgili alanları dolduralım..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Ekle
                OleDbCommand cmd = new OleDbCommand("insert into Urun(UrunKod,UrunAd,Fiyat) values(@ukod,@uad,@ufiyat)", baglanti);
                cmd.Parameters.AddWithValue("@ukod", TxtUrunKod.Text);
                cmd.Parameters.AddWithValue("@uad", TxtUrunAd.Text);
                cmd.Parameters.AddWithValue("@ufiyat", TxtUrunFiyat.Text);

                baglanti.Open();
                cmd.ExecuteNonQuery();
                baglanti.Close();

                //Listele
                OleDbDataAdapter adpu = new OleDbDataAdapter("select UrunId, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Fiyat from Urun order by UrunKod", baglanti);
                DataTable dtu = new DataTable();
                adpu.Fill(dtu);
                DgvUrunler.DataSource = dtu;
                DgvUrunler.Columns["UrunId"].Visible = false;
                DgvUrunler.Columns["Fiyat"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                //Combo Box Listele
                string listele2 = "select UrunId, UrunKod from Urun";
                OleDbDataAdapter adp2 = new OleDbDataAdapter(listele2, baglanti);
                DataSet ds = new DataSet();
                adp2.Fill(ds);
                CmbUrunKod.DataSource = ds.Tables[0];
                CmbUrunKod.DisplayMember = "UrunKod";
                CmbUrunKod.ValueMember = "UrunId";

                //Temizle
                TxtUrunKod.Tag = null;
                TxtUrunKod.Text = null;
                TxtUrunAd.Text = null;
                TxtUrunFiyat.Text = "0";
            }            
        }

        private void BtnUSil_Click(object sender, EventArgs e)
        {
            if (DgvUrunler.Rows.Count > 0)
            {
                if (TxtUrunKod.Tag == null)
                {
                    MessageBox.Show("Lütfen ürünü seçelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    int id = (int)TxtUrunKod.Tag;
                    //Açıklamalar
                    int aId2 = DgvUrunler.SelectedRows[0].Index;
                    
                    string aId = (aId2 + 1).ToString();
                    
                    DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Ürün Kodu: " + TxtUrunKod.Text + "\n" + "Ürün Adı: " + TxtUrunAd.Text + "\n" + "Fiyat: " + TxtUrunFiyat.Text + "\n" + "\n" + "Bu Ürün ve Ürünü içeren işlemlerin hepsi silinsin mi?", "Sil?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (sonuc == DialogResult.Yes)
                    {
                        OleDbCommand cmd = new OleDbCommand("delete from Urun where UrunId=@uid", baglanti);
                        OleDbCommand cmd2 = new OleDbCommand("delete from Islem where UrunId=@uid2", baglanti);

                        cmd.Parameters.AddWithValue("@uid", id);
                        cmd2.Parameters.AddWithValue("@uid2", id);

                        baglanti.Open();
                        cmd.ExecuteNonQuery();
                        cmd2.ExecuteNonQuery();
                        baglanti.Close();

                        //Listele
                        OleDbDataAdapter adpu = new OleDbDataAdapter("select UrunId, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Fiyat from Urun order by UrunKod", baglanti);
                        DataTable dtu = new DataTable();
                        adpu.Fill(dtu);
                        DgvUrunler.DataSource = dtu;
                        DgvUrunler.Columns["UrunId"].Visible = false;
                        DgvUrunler.Columns["Fiyat"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        //Combo Box Listele
                        string listele2 = "select UrunId, UrunKod from Urun";
                        OleDbDataAdapter adp2 = new OleDbDataAdapter(listele2, baglanti);
                        DataSet ds = new DataSet();
                        adp2.Fill(ds);
                        CmbUrunKod.DataSource = ds.Tables[0];
                        CmbUrunKod.DisplayMember = "UrunKod";
                        CmbUrunKod.ValueMember = "UrunId";

                        //Temizle
                        TxtUrunKod.Tag = null;
                        TxtUrunKod.Text = null;
                        TxtUrunAd.Text = null;
                        TxtUrunFiyat.Text = "0";
                    }
                }
            }
        }

        private void BtnUGuncelle_Click(object sender, EventArgs e)
        {
            if (DgvUrunler.Rows.Count > 0)
            {
                if (TxtUrunKod.Tag == null)
                {
                    MessageBox.Show("Lütfen ürünü seçelim..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //Açıklamalar   
                    int aId2 = DgvUrunler.SelectedRows[0].Index;
                    string aId = (aId2 + 1).ToString();

                    string uKod = DgvUrunler.SelectedCells[1].Value.ToString();
                    string uAd = DgvUrunler.SelectedCells[2].Value.ToString();
                    string uFiyat = DgvUrunler.SelectedCells[3].Value.ToString();

                    if (TxtUrunKod.Text == null || TxtUrunKod.Text == "" || TxtUrunAd.Text == null || TxtUrunAd.Text == "")
                    {
                        MessageBox.Show("Lütfen ilgili alanları dolduralım..", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        DialogResult sonuc = MessageBox.Show("Satır No: " + aId + "\n" + "Ürün Kodu: " + uKod + "\n" + "Ürün Adı: " + uAd + "\n" + "Fiyat: " + uFiyat + "\n" + "\n" + "Aşağıdaki yeni veri ile," + "\n" + "\n" + "Ürün Kodu: " + TxtUrunKod.Text + "\n" + "Ürün Adı: " + TxtUrunAd.Text + "\n" + "Fiyat: " + TxtUrunFiyat.Text + "\n" + "\n" + "Güncellensin mi?", "Güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (sonuc == DialogResult.Yes)
                        {
                            OleDbCommand cmd = new OleDbCommand("Update Urun set UrunKod=@ukod,UrunAd=@uad,Fiyat=@ufiyat where UrunId=@uid", baglanti);

                            cmd.Parameters.AddWithValue("@ukod", TxtUrunKod.Text);
                            cmd.Parameters.AddWithValue("@uad", TxtUrunAd.Text);
                            cmd.Parameters.AddWithValue("@ufiyat", TxtUrunFiyat.Text);
                            cmd.Parameters.AddWithValue("@uid", (int)TxtUrunKod.Tag);

                            baglanti.Open();
                            cmd.ExecuteNonQuery();
                            baglanti.Close();

                            //Listele
                            OleDbDataAdapter adpu = new OleDbDataAdapter("select UrunId, UrunKod as [Ürün Kodu], UrunAd as [Ürün Adı], Fiyat from Urun order by UrunKod", baglanti);
                            DataTable dtu = new DataTable();
                            adpu.Fill(dtu);
                            DgvUrunler.DataSource = dtu;
                            DgvUrunler.Columns["UrunId"].Visible = false;
                            DgvUrunler.Columns["Fiyat"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                            //Combo Box Listele
                            string listele2 = "select UrunId, UrunKod from Urun";
                            OleDbDataAdapter adp2 = new OleDbDataAdapter(listele2, baglanti);
                            DataSet ds = new DataSet();
                            adp2.Fill(ds);
                            CmbUrunKod.DataSource = ds.Tables[0];
                            CmbUrunKod.DisplayMember = "UrunKod";
                            CmbUrunKod.ValueMember = "UrunId";

                            //Temizle
                            TxtUrunKod.Tag = null;
                            TxtUrunKod.Text = null;
                            TxtUrunAd.Text = null;
                            TxtUrunFiyat.Text = "0";
                        }
                    }
                }
            }
        }

        private void DgvKontrol_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow dGVRow in this.DgvKontrol.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }
            //Genişlik Ayarı
            this.DgvKontrol.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void BtnKontrol1_Click(object sender, EventArgs e)
        {
            if (dgvFirma.Rows.Count > 0)
            {
                OleDbDataAdapter adpk = new OleDbDataAdapter("select FirmaAd as [Firma], IslemTip as [İşlem Tipi], Tarih from Firma INNER JOIN Islem ON Firma.FirmaId = Islem.FirmaId where (IslemTip = 'Periyodik Kontrol' or IslemTip = 'Genel Bakım') order by  FirmaAd ASC, Tarih DESC", baglanti);
                DataTable dtk = new DataTable();
                adpk.Fill(dtk);
                DgvKontrol.DataSource = dtk;

                for (int i = 0; i < dgvFirma.Rows.Count; i++)
                {
                    bool bulunduPer = false;
                    bool kontrolPer = false;
                    bool bulunduGen = false;
                    bool kontrolGen = false;
                    foreach (DataGridViewRow dGVRow in this.DgvKontrol.Rows)
                    {
                        if (dgvFirma.Rows[i].Cells[1].Value.ToString() == dGVRow.Cells[0].Value.ToString())
                        {
                            if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 < (double)NumPer.Value * 30) && bulunduPer == false)
                            {
                                bulunduPer = true;
                                kontrolPer = true;
                                //Satır Gizlemek için ufak bir ayar lazım
                                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                currencyManager1.SuspendBinding();
                                dGVRow.Visible = false;
                                currencyManager1.ResumeBinding();
                            }
                            else if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 >= ((double)NumPer.Value * 30)) && bulunduPer == false)
                            {
                                bulunduPer = true;
                                dGVRow.DefaultCellStyle.BackColor = Color.DarkOrange;
                            }
                            else if ((string)dGVRow.Cells[1].Value == "Periyodik Kontrol" && (kontrolPer == true || bulunduPer == true))
                            {
                                //Satır Gizlemek için ufak bir ayar lazım
                                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                currencyManager1.SuspendBinding();
                                dGVRow.Visible = false;
                                currencyManager1.ResumeBinding();
                            }

                            //Genel Bakım
                            if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 < (double)NumGen.Value * 30) && bulunduGen == false)
                            {
                                bulunduGen = true;
                                kontrolGen = true;
                                //Satır Gizlemek için ufak bir ayar lazım
                                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                currencyManager1.SuspendBinding();
                                dGVRow.Visible = false;
                                currencyManager1.ResumeBinding();
                            }
                            else if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && ((DateTime.Today - (DateTime)dGVRow.Cells["Tarih"].Value).TotalDays + 15 >= ((double)NumGen.Value * 30)) && bulunduGen == false)
                            {
                                bulunduGen = true;
                                dGVRow.DefaultCellStyle.BackColor = Color.DarkRed;
                                dGVRow.DefaultCellStyle.ForeColor = Color.White;
                            }
                            else if ((string)dGVRow.Cells[1].Value == "Genel Bakım" && (kontrolGen == true || bulunduGen == true))
                            {
                                //Satır Gizlemek için ufak bir ayar lazım
                                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[DgvKontrol.DataSource];
                                currencyManager1.SuspendBinding();
                                dGVRow.Visible = false;
                                currencyManager1.ResumeBinding();
                            }
                        }
                    }

                }
            }
        }

        private void CmbUrunKod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DgvUrunler.Rows.Count > 0)
            {
                for (int i = 0; i < DgvUrunler.Rows.Count; i++)
                {
                    if ((int)CmbUrunKod.SelectedValue == (int)DgvUrunler.Rows[i].Cells[0].Value)
                    {
                        LblUrunAd.Text = DgvUrunler.Rows[i].Cells[2].Value.ToString();
                        LblUrunFiyat.Text = DgvUrunler.Rows[i].Cells[3].Value.ToString();
                    }
                }
            }
        }
    }
}
