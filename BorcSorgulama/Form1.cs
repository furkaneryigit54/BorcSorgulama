using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Control = System.Windows.Forms.Control;
using Font = System.Drawing.Font;


namespace BorcSorgulama
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnDosyaYolu_Click(object sender, EventArgs e)
        {

            //Malik tablosu i�in excel'in dosya yolunu se�me ekran�n� a�ma
            cmbMalik.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyas� Se�iniz..";
            file.ShowDialog();
            txtDosyaYolu.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu �a��r�p sayfa isimlerini alma
            if (txtDosyaYolu.Text!="")
            {
                List<DataRow> SayfalarDR = ExcelSayfaAdlariGetir(txtDosyaYolu.Text).ToList();
                List<string> Sayfalar = new List<string>();
                string kesmeisareti = "'";
                foreach (DataRow dr in SayfalarDR)
                {
                    Sayfalar.Add(dr["TABLE_NAME"].ToString().Trim('$', kesmeisareti[0]));
                }
                //�ekilen sayfa isimlerini combobox'a atama
                foreach (string sayfa in Sayfalar)
                {
                    cmbMalik.Items.Add(sayfa);
                }
                cmbMalik.Enabled = true;
            }

            if (cmbMalik.Items.Count>0)
            {
                cmbMalik.SelectedIndex = 0;
                btnSayfa.PerformClick();
            }
        }
        private void btnSayfa_Click(object sender, EventArgs e)
        {
            //Se�ilen dosyadaki verileri datagridviewe y�kleme
            clbSutunlar.Items.Clear();
            try
            {
                String constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu.Text +
                                "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand cmd =
                    new OleDbCommand("Select * From [" + cmbMalik.SelectedItem.ToString() + "$] where Kimlik is not NULL order by Kimlik asc ",
                        con);
                con.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();
            }
            catch (Exception ex)
            {
                dataGridView1.DataSource = null;
                MessageBox.Show("Ge�erli Bir Sayfa �smi Se�in!","MAL�K");
            }
            //Malik sayfas�n�n s�tun isimlerini al�p combobox'a atama
            string[] malikSutunlar = new string[dataGridView1.Columns.Count];
            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                malikSutunlar[i] = (dataGridView1.Columns[i].Name);
            }
            foreach (var sutun in malikSutunlar)
            {
                clbSutunlar.Items.Add(sutun);
            }
            //Datagridview'e veri eklenince devre d��� olan ara�lar� aktif hale getirme
            if (dataGridView1.Rows.Count > 0)
            {
                clbOturum.Items.Clear();
                clbKiraci.Items.Clear();
                clbSatti.Items.Clear();
                clbTelefon.Items.Clear();
                string[] dogruYanlis = new[] { "DO�RU", "YANLI�" };
                foreach (string s in dogruYanlis)
                {
                    clbOturum.Items.Add(s);
                    clbKiraci.Items.Add(s);
                    clbSatti.Items.Add(s);
                }
                string[] varYok = new[] { "VAR", "YOK" };
                foreach (string s in varYok)
                {
                    clbTelefon.Items.Add(s);
                }
                clbSutunlar.Enabled = true;
                clbOturum.Enabled = true;
                clbKiraci.Enabled = true;
                clbSatti.Enabled = true;
                clbTelefon.Enabled = true;
                btnFiltre.Enabled = true;
                if (txtMalikExcel.Text!="")
                {
                    btnMalikExport.Enabled = true;
                }
            }
            else
            {
                clbSutunlar.Enabled = false;
                clbSutunlar.Items.Clear();
                clbOturum.Enabled = false;
                clbOturum.Items.Clear();
                clbKiraci.Enabled = false;
                clbKiraci.Items.Clear();
                clbSatti.Enabled = false;
                clbSatti.Items.Clear();
                clbTelefon.Enabled = false;
                clbTelefon.Items.Clear();
                btnFiltre.Enabled = false;
                btnMalikExport.Enabled = false;
            }

        }
        public List<DataRow> ExcelSayfaAdlariGetir(string DosyaYolu)
        {
            //Excel dosyas�ndaki sayfa isimlerini �ekme
            try
            {

                string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + DosyaYolu +
                                "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
                OleDbConnection con = new OleDbConnection(constr);
                con.Open();
                List<DataRow> SayfaIsimleriList = con.GetSchema("Tables").AsEnumerable().ToList<DataRow>();
                con.Close();
                return SayfaIsimleriList;
            }
            finally
            {
            }
        }
        private void btnDosyaYolu2_Click(object sender, EventArgs e)
        {
            //Logo tablosu i�in excel'in dosya yolunu se�me ekran�n� a�ma
            cmbLogo.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyas� Se�iniz..";
            file.ShowDialog();
            txtDosyaYolu2.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu �a��r�p sayfa isimlerini alma
            if (txtDosyaYolu2.Text!="")
            {
                List<DataRow> SayfalarDR = ExcelSayfaAdlariGetir(txtDosyaYolu2.Text).ToList();
                List<string> Sayfalar = new List<string>();
                string kesmeisareti = "'";
                foreach (DataRow dr in SayfalarDR)
                {
                    Sayfalar.Add(dr["TABLE_NAME"].ToString().Trim('$', kesmeisareti[0]));
                }
                //�ekilen sayfa isimlerini combobox'a atama
                foreach (string sayfa in Sayfalar)
                {
                    cmbLogo.Items.Add(sayfa);
                }
                //Veri giri�i yap�ld���nda pasif ara�lar� aktif hale getirme
                cmbLogo.Enabled = true;
                cmbBolge.Enabled = true;
                cmbBolge.Text = cmbBolge.Items[0].ToString();
            }

            if (cmbLogo.Items.Count>0)
            {
                int logoind = 0;
                for (int i = 0; i < cmbLogo.Items.Count; i++)
                {
                    if (cmbLogo.Items[i].ToString().ToLower() == "logo")
                    {
                        logoind = i;
                    }
                }
                cmbLogo.SelectedIndex = logoind;
                if (cmbMalik.Items.Count>0)
                {
                    string secilen = cmbMalik.SelectedItem.ToString().ToLower();
                    if (secilen == "v�llalar")
                    {
                        cmbBolge.SelectedIndex = 1;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "cars�_evler�")
                    {
                        cmbBolge.SelectedIndex = 2;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "lbloklar�")
                    {
                        cmbBolge.SelectedIndex = 3;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "acarblu")
                    {
                        cmbBolge.SelectedIndex = 4;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "acar_vad�")
                    {
                        cmbBolge.SelectedIndex = 5;
                        btnSayfa2.PerformClick();
                    }
                }
                else
                {
                    cmbBolge.SelectedIndex = 0;
                    btnSayfa2.PerformClick();
                }
            }
        }
        private void btnSayfa2_Click(object sender, EventArgs e)
        {
            //Se�ilen dosyadaki verileri datagridviewe y�kleme
            clbSutunlarLogo.Items.Clear();
            clbOdemeTuruLogo.Items.Clear();
            clbKatLogo.Items.Clear();
            string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu2.Text +
                            "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
            OleDbConnection con = new OleDbConnection(constr);
            //B�lge se�imi i�in filtreleme
            string bolge = "";
            if (cmbBolge.SelectedIndex == 1)
            {
                bolge = " where AKRO_KODU like 'ACRA%' OR AKRO_KODU like 'ACRB%' OR AKRO_KODU like 'ACRC%' OR AKRO_KODU like 'ACRD%' OR AKRO_KODU like 'ACRI%' " +
                        "OR AKRO_KODU like 'ACRK%' OR AKRO_KODU like 'ACRO%' OR AKRO_KODU like 'ACRT%' ";
            }else if (cmbBolge.SelectedIndex==2)
            {
                bolge = "where AKRO_KODU like 'ACRS%'";
            }else if (cmbBolge.SelectedIndex==3)
            {
                bolge = "where AKRO_KODU like 'ACRL%'";
            }
            else if (cmbBolge.SelectedIndex == 4)
            {
                bolge = "where AKRO_KODU like 'BLU%'";
            }
            else if (cmbBolge.SelectedIndex == 5)
            {
                bolge = "where AKRO_KODU like 'VDI%'";
            }
            else if (cmbBolge.SelectedIndex == 6)
            {
                bolge = "where AKRO_KODU L�KE 'VRD%'";
            }
            else if (cmbBolge.SelectedIndex == 0)
            {
                bolge = "";
            }
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select * From [" + cmbLogo.SelectedItem.ToString() + "$] " + bolge + " order by AKRO_KODU asc ", con);
                con.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                con.Close();
            }
            catch (Exception ex)
            {
                dataGridView2.DataSource = null;
                MessageBox.Show("Ge�erli Bir Sayfa �smi Se�in!","LOGO");
            }
            try
            {
                //Logo dosyas�n�n sutun isimlerini, filtreleme i�in kat ve odemeturu s�tunlar�n� �ekip combobox'a atama
                string[] logoSutunlar = new string[dataGridView2.Columns.Count];
                string[] katlar = new string[dataGridView2.Rows.Count - 1];
                string[] odemeTuru = new string[dataGridView2.Rows.Count - 1];
                //Kat ve �deme t�r� s�tunundaki de�erleri dizilere atama
                for (int i = 0; i <= dataGridView2.Columns.Count - 1; i++)
                {
                    logoSutunlar[i] = (dataGridView2.Columns[i].Name);
                    if (logoSutunlar[i] == "KAT")
                    {
                        for (int j = 0; j < dataGridView2.Rows.Count - 1; j++)
                        {
                            katlar[j] = dataGridView2.Rows[j].Cells[i].Value.ToString();
                        }
                    }

                    if (logoSutunlar[i] == "ODEMETUR")
                    {
                        for (int j = 0; j < dataGridView2.Rows.Count - 1; j++)
                        {
                            odemeTuru[j] = dataGridView2.Rows[j].Cells[i].Value.ToString();
                        }
                    }
                }
                //S�tun isimlerini combobox'a atama
                foreach (var sutun in logoSutunlar)
                {
                    clbSutunlarLogo.Items.Add(sutun);
                }
                List<string> katList = new List<string>();
                for (int i = 0; i < katlar.Length; i++)
                {
                    katList.Add(katlar[i]);
                }
                List<string> odemeTuruList = new List<string>();
                for (int i = 0; i < odemeTuru.Length; i++)
                {
                    odemeTuruList.Add(odemeTuru[i]);
                }
                //Kat listesindeki tekrar eden de�erleri silme ve s�ralama
                katList = katList.Distinct().ToList();
                katList.Sort();
                for (int i = 0; i < katList.Count - 1; i++)
                {
                    if (katList[i] == "")
                    {
                        katList.RemoveAt(i);
                    }
                }
                //�deme t�r� listesindeki tekrar eden de�erleri silme ve s�ralama
                odemeTuruList = odemeTuruList.Distinct().ToList();
                odemeTuruList.Sort();
                for (int i = 0; i < odemeTuruList.Count - 1; i++)
                {
                    if (odemeTuruList[i] == "")
                    {
                        odemeTuruList.RemoveAt(i);
                    }
                }
                //Kat ve �deme t�r� bilgilerini combobox'lara atama
                for (int i = 0; i < katList.Count; i++)
                {
                    clbKatLogo.Items.Add(katList[i]);
                }
                foreach (string odemeturu in odemeTuruList)
                {
                    clbOdemeTuruLogo.Items.Add(odemeturu);
                }
                //Bakiye de�erlerini tekrar hesaplay�p bakiye s�tununa atama
                int bakiyeIndex = 0;
                int borcIndex = 0;
                int alacakIndex = 0;
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    if (dataGridView2.Columns[j].Name == "BAKIYE")
                    {
                        bakiyeIndex = j;
                    }
                    if (dataGridView2.Columns[j].Name == "BORC")
                    {
                        borcIndex = j;
                    }
                    if (dataGridView2.Columns[j].Name == "ALACAK")
                    {
                        alacakIndex = j;
                    }
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    Decimal bakiye = 0;
                    bakiye =
                        Convert.ToDecimal(dataGridView2.Rows[i].Cells[borcIndex].Value) -
                        Convert.ToDecimal(dataGridView2.Rows[i].Cells[alacakIndex].Value);
                    dataGridView2.Rows[i].Cells[bakiyeIndex].Value = String.Format("{0:0.00}", bakiye.ToString());
                }
            }
            catch (Exception exception)
            {
            }
            //S�tun filtresi combobox'�na veri giri�i yap�ld���nda pasif haldeki ara�lar� aktif hale getirme
            if (dataGridView2.Rows.Count > 0)
            {
                clbSutunlarLogo.Enabled = true;
                clbKatLogo.Enabled = true;
                clbBakiyeLogo.Enabled = true;
                clbOdemeTuruLogo.Enabled = true;
                btnFiltreLogo.Enabled = true;
                clbBakiyeIsPositive.Enabled = true;
                clbBakiyeLogo.Items.Clear();
                clbBakiyeIsPositive.Items.Clear();
                string[] clbbakiyeStrings = new[] { "B�y�kt�r","K���kt�r","E�ittir","B�y�k E�ittir","K���k E�ittir" };
                string[] clbbakiyeDurumuStrings = new[] { "VAR", "YOK", "ALACAK VAR" };
                foreach (string s in clbbakiyeDurumuStrings)
                {
                    clbBakiyeIsPositive.Items.Add(s);
                }
                foreach (string s in clbbakiyeStrings)
                {
                    clbBakiyeLogo.Items.Add(s);
                }
                txtBakiyeLogo.Enabled = true;
            }
            else
            {
                clbSutunlarLogo.Items.Clear();
                clbSutunlarLogo.Enabled = false;
                clbKatLogo.Enabled = false;
                clbKatLogo.Items.Clear();
                clbBakiyeLogo.Enabled = false;
                clbBakiyeLogo.Items.Clear();
                clbOdemeTuruLogo.Enabled = false;
                clbOdemeTuruLogo.Items.Clear();
                btnFiltreLogo.Enabled = false;
                txtBakiyeLogo.Enabled = false;
                clbBakiyeIsPositive.Items.Clear();
                clbBakiyeIsPositive.Enabled=false;
            }
            //Datagridview'daki b�t�n s�tunlar� g�r�n�r hale getirme
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].Visible = true;
            }
        }
        private void btnAyir_Click(object sender, EventArgs e)
        {
       
            if (dgvEslesmeyenler.Rows.Count>0)
            {
                dgvEslesmeyenler.Rows.Clear();
                dgvEslesmeyenler.Refresh();
            }

            
            int sira = 1;
            //Datagridview'a gerekli s�tun isimlerini ekleme
            dataGridView3.Columns.Clear();
            dataGridView3.ColumnCount = 4;
            dataGridView3.Columns[0].Name = "S�ra No";
            dataGridView3.Columns[1].Name = "AD";
            dataGridView3.Columns[2].Name = "TUTAR";
            dataGridView3.Columns[3].Name = "TELEFON";
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView3);
            //��kt� dosyas� i�in gerekli olan 2 dosyayadaki verileri e�le�tirecek de�i�kenlerin indexlerinin tan�m�
            //Logo dosyas� index de�i�kenleri
            int NoLogoind = 0;
            int KatLogoind = 0;
            int aciklamaLogoind = 0;
            int borcLogoind = 0;
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                if (dataGridView2.Columns[i].Name == "NO")
                {
                    NoLogoind = dataGridView2.Columns[i].Index;
                }
                else if (dataGridView2.Columns[i].Name == "KAT")
                {
                    KatLogoind = dataGridView2.Columns[i].Index;
                }
                else if (dataGridView2.Columns[i].Name == "ACIKLAMA")
                {
                    aciklamaLogoind = dataGridView2.Columns[i].Index;

                }
                else if (dataGridView2.Columns[i].Name == "BAKIYE")
                {
                    borcLogoind = dataGridView2.Columns[i].Index;
                }
            }
            //Malik dosyas� index de�i�kenleri
            int kimlikMalikind = 0;
            int telefonMalikind = 0;
            int adiMalikind = 0;
            int soyadMalikind = 0;
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].Name == "Kimlik")
                {
                    kimlikMalikind = dataGridView1.Columns[i].Index;
                }
                else if (dataGridView1.Columns[i].Name == "CepTelefonu")
                {
                    telefonMalikind = dataGridView1.Columns[i].Index;
                }
                else if (dataGridView1.Columns[i].Name == "Ad�")
                {
                    adiMalikind = dataGridView1.Columns[i].Index;
                }
                else if (dataGridView1.Columns[i].Name == "Soyad")
                {
                    soyadMalikind = dataGridView1.Columns[i].Index;
                }
            }
            //�ki dosyadaki verileri e�le�tirip ��kt� dosyas� i�in datagridview'e verileri ekleme
            List<string> eslesmeyenler = new List<string>();
            dgvEslesmeyenler.ColumnCount = dataGridView2.ColumnCount;
            for (int j = 0; j < dataGridView2.Columns.Count; j++)
            {
                dgvEslesmeyenler.Columns[j].Name = dataGridView2.Columns[j].Name;
            }
            for (int i = 0; i < dataGridView2.RowCount ; i++)
            {
                int eslesti = 1;
                string borc = dataGridView2.Rows[i].Cells[borcLogoind].Value.ToString().Replace(',', '.');
                if (Convert.ToDouble(dataGridView2.Rows[i].Cells[borcLogoind].Value.ToString()) > 0)
                {
                    //Logo dosyas�ndan no ve kat de�erlerini string de�i�kenlere tan�mlama
                    string noLogo = dataGridView2.Rows[i].Cells[NoLogoind].Value.ToString();
                    string katLogo = dataGridView2.Rows[i].Cells[KatLogoind].Value.ToString();
                    //E�le�tirme i�in logo dosyas�ndaki kat s�tunundan �ekilen verilerden 0 say�l�r�n� silme
                    if (cmbBolge.SelectedItem != "�AR�I EVLER�" & cmbBolge.SelectedItem != "L BLOKLARI")
                    {
                        string[] katlist = katLogo.Split('0');
                        katLogo = "";
                        foreach (string de�er in katlist)
                        {
                            katLogo += de�er;
                        }
                    }
                    //Logo tablosundan no ve kat bilgilerini birle�tirerek �nite de�i�kenini tan�mlama
                    string uniteLogo = noLogo + katLogo;
                    //Logo tablosundan bor� bilgisini de�i�kene tan�mlama
                    string borclogo = dataGridView2.Rows[i].Cells[borcLogoind].Value.ToString();
                    for (int j = 0; j < dataGridView1.RowCount ; j++)
                    {
                        //Malik tablosundan �nite bilgisini de�i�kene atay�p e�le�tirmek i�in nokta varsa silme
                        string unite = dataGridView1.Rows[j].Cells[kimlikMalikind].Value.ToString();
                        string[] uniteList = unite.Split('.');
                        //�nite de�i�kenindeki A ve B de�erlerini e�le�tirme i�in say�ya �evirme
                        foreach (string deger in uniteList)
                        {
                            if (deger == "A")
                            {
                                uniteList[Array.IndexOf(uniteList, "A")] = "1";
                            }
                            else if (deger == "B")
                            {
                                uniteList[Array.IndexOf(uniteList, "B")] = "2";
                            }
                        }
                        //Yap�lan i�lemlerden sonra �nite de�ikenini olu�turulan listeden tekrar birle�tirme
                        unite = "";
                        foreach (var de�er in uniteList)
                        {
                            unite += de�er;
                        }
                        //�nite bilgileri iki tablodada uyu�an ki�ilerin bilgilerini ��kt� tablosuna ekleme ve d�ng�y� sonland�rma
                        if (unite == uniteLogo)
                        {
                            string telefon = dataGridView1.Rows[j].Cells[telefonMalikind].Value.ToString();
                            dataGridView3.Rows.Add(sira,dataGridView1.Rows[j].Cells[kimlikMalikind].Value.ToString() + " " + dataGridView1.Rows[j].Cells[adiMalikind].Value.ToString() + " " + dataGridView1.Rows[j].Cells[soyadMalikind].Value.ToString(), borclogo, telefon);
                            sira++;
                            eslesti++;
                            goto don;
                        }
                        
                    }
                    don:;
                }

                if (eslesti<=1)
                {
                    dgvEslesmeyenler.Rows.Add(dataGridView2.Rows[i].Cells[0].Value, dataGridView2.Rows[i].Cells[1].Value, dataGridView2.Rows[i].Cells[2].Value, dataGridView2.Rows[i].Cells[3].Value, dataGridView2.Rows[i].Cells[4].Value, dataGridView2.Rows[i].Cells[5].Value, dataGridView2.Rows[i].Cells[6].Value, dataGridView2.Rows[i].Cells[7].Value, dataGridView2.Rows[i].Cells[8].Value, dataGridView2.Rows[i].Cells[9].Value, dataGridView2.Rows[i].Cells[10].Value);
                    eslesmeyenler.Add(dataGridView2.Rows[i].Cells[aciklamaLogoind].Value.ToString());
                }
            }

            
            eklenemeyenler frm = new eklenemeyenler();
            frm.listeDoldur(eslesmeyenler);
            //Excel'e verilecek isim textboxuna giri� yap�ld���nda ve ��kt� datagridview'�na veri giri�i yap�ld���nda dosyay� kaydetme tu�unu aktif hale getirme
            if (txtExcelisim.Text != String.Empty & dataGridView3.Rows.Count > 0)
            {
                btnExcelAktar.Enabled = true;
            }
            else
            {
                btnExcelAktar.Enabled = false;
            }

            if (dataGridView3.Rows.Count>0)
            {
                btnIsimleEslestir.Enabled = true;
            }
            else
            {
                btnIsimleEslestir.Enabled = false;
            }
        }
        private void btnFiltre_Click(object sender, EventArgs e)
        {
            //Malik tablosunda oturum bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string oturumFiltre = "";
            if (clbOturum.CheckedItems.Count > 0)
            {
                if (clbOturum.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbOturum.CheckedItems[i].ToString() == "DO�RU")
                        {
                            oturumFiltre = "AND Oturum=True";
                        }
                        else
                        {
                            oturumFiltre = "AND Oturum=False";
                        }
                    }
                }

            }
            //Malik tablosunda satt� bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string sattiFiltre = "";
            if (clbSatti.CheckedItems.Count > 0)
            {
                if (clbSatti.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbSatti.CheckedItems[i].ToString() == "DO�RU")
                        {
                            sattiFiltre = "and Satt�=True";
                        }
                        else
                        {
                            sattiFiltre = "and Satt�=False";
                        }
                    }
                }
            }
            // Malik tablosunda kirac� bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string kiraciFiltre = "";
            if (clbKiraci.CheckedItems.Count > 0)
            {
                if (clbKiraci.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbKiraci.CheckedItems[i].ToString() == "DO�RU")
                        {
                            kiraciFiltre = "and Kirac�=True";
                        }
                        else
                        {
                            kiraciFiltre = "and Kirac�=False";
                        }
                    }
                }
            }
            //Malik tablosunda telefon bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string telefonFiltre = "";
            if (clbTelefon.CheckedItems.Count > 0)
            {
                if (clbTelefon.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbTelefon.CheckedItems[i].ToString() == "VAR")
                        {
                            telefonFiltre = "and CepTelefonu <>''";
                        }
                        else
                        {
                            telefonFiltre = "and CepTelefonu IS NULL";
                        }
                    }
                }
            }
            //Malik tablosunun oldu�u datagridview'i temizleyip filtrelerle olu�turulan yeni sorguyu �al��t�rma
            dataGridView1.DataSource = null;
            string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu.Text +
                            "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
            OleDbConnection conn = new OleDbConnection(constr);
            OleDbCommand cmd = new OleDbCommand("Select * From [" + cmbMalik.SelectedItem.ToString() + "$] where Kimlik is not null " + oturumFiltre + " " + sattiFiltre + " " + kiraciFiltre + " " + telefonFiltre + " order by Kimlik asc ", conn);
            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                //T�m s�tunlar� g�r�nmez hale getirme
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].Visible = false;
                }
                //Filtrelenen s�tunlar� g�r�n�r hale getirme
                if (clbSutunlar.CheckedItems.Count > 0)
                {
                    for (int i = 0; i < clbSutunlar.Items.Count; i++)
                    {
                        for (int j = 0; j < clbSutunlar.CheckedItems.Count; j++)
                        {
                            if (clbSutunlar.Items[i] == clbSutunlar.CheckedItems[j])
                            {
                                dataGridView1.Columns[i].Visible = true;
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        dataGridView1.Columns[i].Visible = true;
                    }
                }
            }
            catch
            {
                dataGridView1.DataSource = null;
                MessageBox.Show("Ge�erli Bir Sayfa Se�in!","MAL�K");
                if (dataGridView1.Rows.Count > 0)
                {
                    clbOturum.Items.Clear();
                    clbKiraci.Items.Clear();
                    clbSatti.Items.Clear();
                    clbTelefon.Items.Clear();
                    string[] dogruYanlis = new[] { "DO�RU", "YANLI�" };
                    foreach (string s in dogruYanlis)
                    {
                        clbOturum.Items.Add(s);
                        clbKiraci.Items.Add(s);
                        clbSatti.Items.Add(s);
                    }
                    string[] varYok = new[] { "VAR", "YOK" };
                    foreach (string s in varYok)
                    {
                        clbTelefon.Items.Add(s);
                    }
                    clbSutunlar.Enabled = true;
                    clbOturum.Enabled = true;
                    clbKiraci.Enabled = true;
                    clbSatti.Enabled = true;
                    clbTelefon.Enabled = true;
                    btnFiltre.Enabled = true;
                }
                else
                {
                    clbSutunlar.Enabled = false;
                    clbSutunlar.Items.Clear();
                    clbOturum.Enabled = false;
                    clbOturum.Items.Clear();
                    clbKiraci.Enabled = false;
                    clbKiraci.Items.Clear();
                    clbSatti.Enabled = false;
                    clbSatti.Items.Clear();
                    clbTelefon.Enabled = false;
                    clbTelefon.Items.Clear();
                    btnFiltre.Enabled = false;
                }
            }

            if (dataGridView1.Rows.Count>0&txtMalikExcel.Text!="")
            {
                btnMalikExport.Enabled = true;
            }
            conn.Close();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            tabPage1.Text = "Malik Kay�tlar�";
            tabPage2.Text = "LOGO Kay�tlar�";
            tabPage3.Text = "��LENM�� TABLO";
            tabPage4.Text = "Malik Filtreler";
            tabPage5.Text = "LOGO Filtreler";
            cmbBolge.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLogo.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbMalik.DropDownStyle = ComboBoxStyle.DropDownList;
            this.MaximizeBox = false;
        }
        private void btnFiltreLogo_Click(object sender, EventArgs e)
        {
            // Logo tablosunda kat bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string kat = "";
            if (clbKatLogo.CheckedItems.Count > 0)
            {
                if (clbKatLogo.CheckedItems.Count == 1)
                {
                    kat = "and Kat='" + clbKatLogo.CheckedItems[0].ToString() + "'";
                }
                else if (clbKatLogo.CheckedItems.Count > 1)
                {
                    string filtreler = "";
                    for (int i = 0; i < clbKatLogo.CheckedItems.Count; i++)
                    {
                        filtreler = filtreler + "'" + clbKatLogo.CheckedItems[i] + "',";
                    }
                    kat = "and KAT in (" + filtreler + ")";
                }
            }
            // Logo tablosunda �deme t�r� bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string odemeturu = "";
            if (clbOdemeTuruLogo.CheckedItems.Count > 0)
            {
                if (clbOdemeTuruLogo.CheckedItems.Count == 1)
                {
                    odemeturu = "and ODEMETUR='" + clbOdemeTuruLogo.CheckedItems[0].ToString() + "'";
                }
                else if (clbOdemeTuruLogo.CheckedItems.Count > 1)
                {
                    string filtreler = "";
                    for (int i = 0; i < clbOdemeTuruLogo.CheckedItems.Count; i++)
                    {
                        filtreler = filtreler + "'" + clbOdemeTuruLogo.CheckedItems[i] + "',";
                    }
                    odemeturu = "and ODEMETUR in (" + filtreler + ")";
                }
            }
            // Logo tablosunda bakiye bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string bakiyeDurumu = "";
            if (clbBakiyeIsPositive.CheckedItems.Count > 0)
            {
                if (clbBakiyeIsPositive.CheckedItems.Count == 2)
                {
                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {
                        if (clbBakiyeIsPositive.CheckedItems[i] == "VAR")
                        {
                            bakiyeDurumu = "and BORC>ALACAK";
                        }
                        else if (clbBakiyeIsPositive.CheckedItems[i] == "YOK")
                        {
                            bakiyeDurumu = "and BORC=ALACAK";
                        }
                        else if (clbBakiyeIsPositive.CheckedItems[i] == "ALACAK VAR")
                        {
                            bakiyeDurumu = "and BORC<ALACAK";
                        }

                    }
                }
            }
            string bakiye = "";
            if (clbBakiyeLogo.CheckedItems.Count>0&txtBakiyeLogo.Text!="")
            {
                if (clbBakiyeLogo.CheckedItems.Count > 1)
                {
                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {
                        if (clbBakiyeLogo.CheckedItems[i] == "B�y�kt�r")
                        {
                            bakiye = "and BAKIYE>" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "K���kt�r")
                        {
                            bakiye = "and BAKIYE<" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "E�ittir")
                        {
                            bakiye = "and BAKIYE=" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "B�y�k E�ittir")
                        {
                            bakiye = "and BAKIYE>=" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "K���k E�ittir")
                        {
                            bakiye = "and BAKIYE<=" + txtBakiyeLogo.Text + "";
                        }


                    }
                }
            }
            // Logo tablosunda b�lge bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string bolge = "";
            if (cmbBolge.SelectedIndex == 1)
            {
                bolge = " and (AKRO_KODU like 'ACRA%' OR AKRO_KODU like 'ACRB%' OR AKRO_KODU like 'ACRC%' OR AKRO_KODU like 'ACRD%' OR AKRO_KODU like 'ACRI%' " +
                        "OR AKRO_KODU like 'ACRK%' OR AKRO_KODU like 'ACRO%' OR AKRO_KODU like 'ACRT%') ";
            }
            else if (cmbBolge.SelectedIndex == 2)
            {
                bolge = "and AKRO_KODU like 'ACRS%'";
            }
            else if (cmbBolge.SelectedIndex == 3)
            {
                bolge = "and AKRO_KODU like 'ACRL%'";
            }
            else if (cmbBolge.SelectedIndex == 4)
            {
                bolge = "and AKRO_KODU like 'BLU%'";
            }
            else if (cmbBolge.SelectedIndex == 5)
            {
                bolge = "and AKRO_KODU like 'VDI%'";
            }
            else if (cmbBolge.SelectedIndex == 6)
            {
                bolge = "and AKRO_KODU L�KE 'VRD%'";
            }
            else if (cmbBolge.SelectedIndex == 0)
            {
                bolge = "";
            }
            //Logo tablosunundaki verileri temizleyip olu�turulan yeni sorguyu �al��t�rma
            dataGridView2.DataSource = null;
            
            try
            {
                string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu2.Text +
                                "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
                OleDbConnection conn = new OleDbConnection(constr);
                OleDbCommand cmd = new OleDbCommand("Select *From [" + cmbLogo.SelectedItem.ToString() + "$] " +
                                                    "where ACIKLAMA is not null " + kat + " " + odemeturu + " " + bakiye + " " + bolge + " " + bakiyeDurumu + " order by AKRO_KODU ", conn);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                conn.Close();
            }
            catch
            {
                dataGridView2.DataSource = null;
                MessageBox.Show("Ge�erli Bir Sayfa Se�in!","LOGO");
            }
            //Bakiye s�tunundaki verileri hesaplay�p yeni de�erler ile de�i�tirme
            int bakiyeIndex = 0;
            int borcIndex = 0;
            int alacakIndex = 0;
            for (int j = 0; j < dataGridView2.Columns.Count; j++)
            {
                if (dataGridView2.Columns[j].Name == "BAKIYE")
                {
                    bakiyeIndex = j;
                }

                if (dataGridView2.Columns[j].Name == "BORC")
                {
                    borcIndex = j;
                }

                if (dataGridView2.Columns[j].Name == "ALACAK")
                {
                    alacakIndex = j;
                }
            }
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                decimal bakiyeHesaplanan = 0;
                bakiyeHesaplanan =
                    Convert.ToDecimal(dataGridView2.Rows[i].Cells[borcIndex].Value) -
                    Convert.ToDecimal(dataGridView2.Rows[i].Cells[alacakIndex].Value);
                dataGridView2.Rows[i].Cells[bakiyeIndex].Value = String.Format("{0:0.00}", bakiyeHesaplanan.ToString());
            }
            
            //T�m s�tunlar� g�r�nmez hale getirme
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].Visible = false;
            }
            //Filtrelenen s�tunlar� g�r�n�r hale getirme
            try
            {
                if (clbSutunlarLogo.CheckedItems.Count > 0)
                {
                    for (int i = 0; i < clbSutunlarLogo.Items.Count; i++)
                    {
                        for (int j = 0; j < clbSutunlarLogo.CheckedItems.Count; j++)
                        {
                            if (clbSutunlarLogo.Items[i] == clbSutunlarLogo.CheckedItems[j])
                            {
                                dataGridView2.Columns[i].Visible = true;
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        dataGridView2.Columns[i].Visible = true;
                    }
                }
                //Logo tablosu datagridview'�na veri giri�i yap�ld���nda pasif haldeki ara�lar� aktif hale getirme
                //Veri yoksa eski verileri temizlemeyip pasif hale getirme
                if (dataGridView2.Rows.Count > 0)
                {
                    clbSutunlarLogo.Enabled = true;
                    clbKatLogo.Enabled = true;
                    clbBakiyeLogo.Enabled = true;
                    clbOdemeTuruLogo.Enabled = true;
                    btnFiltreLogo.Enabled = true;
                    txtBakiyeLogo.Enabled=true;
                    clbBakiyeIsPositive.Enabled = true;
                   
                }
                else
                {
                    clbSutunlarLogo.Items.Clear();
                    clbSutunlarLogo.Enabled = false;
                    clbKatLogo.Enabled = false;
                    clbKatLogo.Items.Clear();
                    clbBakiyeLogo.Enabled = false;
                    clbBakiyeLogo.Items.Clear();
                    clbOdemeTuruLogo.Enabled = false;
                    clbOdemeTuruLogo.Items.Clear();
                    btnFiltreLogo.Enabled = false;
                    txtBakiyeLogo.Enabled=false;
                    clbBakiyeIsPositive.Items.Clear();
                    clbBakiyeIsPositive.Enabled=false;
                }
            }
            finally
            {
            }
        }
        private void cmbMalik_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Malik dosyas�ndan sayfa ismi se�ildi�ine verileri getirecek butonu aktif hale getirme
            if (cmbMalik.SelectedItem != "")
            {
                btnSayfa.Enabled = true;
            }
            else
            {
                btnSayfa.Enabled = false;
            }

            string secilen = cmbMalik.SelectedItem.ToString().ToLower();
            if (secilen == "v�llalar")
            {
                cmbBolge.SelectedIndex = 1;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "cars�_evler�")
            {
                cmbBolge.SelectedIndex = 2;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "lbloklar�")
            {
                cmbBolge.SelectedIndex = 3;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "acarblu")
            {
                cmbBolge.SelectedIndex = 4;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "acar_vad�")
            {
                cmbBolge.SelectedIndex = 5;
                btnSayfa2.PerformClick();
            }
            btnSayfa.PerformClick();
        }
        private void cmbLogo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Logo dosyas�ndan sayfa ismi se�ildi�ine verileri getirecek butonu aktif hale getirme
            if (cmbLogo.SelectedItem != "")
            {
                btnSayfa2.Enabled = true;

            }
            else
            {
                btnSayfa2.Enabled = false;
            }
            btnSayfa2.PerformClick();
        }
        private void btnExcelAktar_Click(object sender, EventArgs e)
        {
            //��kt� dosyas�n� excel' d�n��t�r�p C s�r�c�s� i�erinde excel klas�r� olu�turup girilen isimle kaydetme
            if (dataGridView3.Rows.Count > 0 & txtExcelisim.Text != string.Empty)
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn sutun in dataGridView3.Columns)
                {
                    dt.Columns.Add(sutun.HeaderText);
                }

                foreach (DataGridViewRow satir in dataGridView3.Rows)
                {
                    dt.Rows.Add();
                    foreach (DataGridViewCell hucre in satir.Cells)
                    {
                        dt.Rows[dt.Rows.Count - 1][hucre.ColumnIndex] = hucre.Value.ToString();
                    }
                }
                string dosyaYolu = "C:\\Excel\\";
                if (!Directory.Exists(dosyaYolu))
                {
                    Directory.CreateDirectory(dosyaYolu);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt, "SMS Tablosu");
                    wb.SaveAs(dosyaYolu + txtExcelisim.Text + ".xlsx");
                }
            }
        }
        private void txtExcelisim_TextChanged(object sender, EventArgs e)
        {
            //Excel'e verilecek isim textboxuna giri� yap�ld���nda ve ��kt� datagridview'�na veri giri�i yap�ld���nda dosyay� kaydetme tu�unu aktif hale getirme
            if (txtExcelisim.Text!=String.Empty& dataGridView3.Rows.Count > 0)
            {
                btnExcelAktar.Enabled = true;
            }
            else
            {
                btnExcelAktar.Enabled = false;
            }
        }

        private void txtMalikExcel_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count>0 & txtMalikExcel.Text!=String.Empty)        
            {
                btnMalikExport.Enabled = true;
            }
            else
            {
                btnMalikExport.Enabled = false;
            }
        }

        private void btnMalikExport_Click(object sender, EventArgs e)
        {
            string oturumFiltre = "";
            if (clbOturum.CheckedItems.Count > 0)
            {
                if (clbOturum.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbOturum.CheckedItems[i].ToString() == "DO�RU")
                        {
                            oturumFiltre = "AND Oturum=True";
                        }
                        else
                        {
                            oturumFiltre = "AND Oturum=False";
                        }
                    }
                }

            }
            //Se�ilen s�tunlar� filtreleme
            string sutunlar = "*";
            if (clbSutunlar.CheckedItems.Count>0)
            {
                if (clbSutunlar.CheckedItems.Count==1)
                {
                    sutunlar = "";
                    foreach (string s in clbSutunlar.CheckedItems)
                    {
                        sutunlar += s;
                    }
                }
                else
                {
                    sutunlar = clbSutunlar.CheckedItems[0].ToString();
                    for (int i = 1; i < clbSutunlar.CheckedItems.Count; i++)
                    {
                        sutunlar = sutunlar + "," + clbSutunlar.CheckedItems[i].ToString();
                    }
                }
            }
            //Malik tablosunda satt� bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string sattiFiltre = "";
            if (clbSatti.CheckedItems.Count > 0)
            {
                if (clbSatti.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbSatti.CheckedItems[i].ToString() == "DO�RU")
                        {
                            sattiFiltre = "and Satt�=True";
                        }
                        else
                        {
                            sattiFiltre = "and Satt�=False";
                        }
                    }
                }
            }
            // Malik tablosunda kirac� bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string kiraciFiltre = "";
            if (clbKiraci.CheckedItems.Count > 0)
            {
                if (clbKiraci.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbKiraci.CheckedItems[i].ToString() == "DO�RU")
                        {
                            kiraciFiltre = "and Kirac�=True";
                        }
                        else
                        {
                            kiraciFiltre = "and Kirac�=False";
                        }
                    }
                }
            }
            //Malik tablosunda telefon bilgisini filtrelemek i�in ko�ul olu�turup yeni sorguya ekleme
            string telefonFiltre = "";
            if (clbTelefon.CheckedItems.Count > 0)
            {
                if (clbTelefon.CheckedItems.Count == 2)
                {

                }
                else
                {
                    for (int i = 0; i < 1; i++)
                    {

                        if (clbTelefon.CheckedItems[i].ToString() == "VAR")
                        {
                            telefonFiltre = "and CepTelefonu <>''";
                        }
                        else
                        {
                            telefonFiltre = "and CepTelefonu IS NULL";
                        }
                    }
                }
            }
            //Malik tablosunun oldu�u datagridview'i temizleyip filtrelerle olu�turulan yeni sorguyu �al��t�rma
            dgvMalik.DataSource = null;
            string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu.Text +
                            "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
            OleDbConnection conn = new OleDbConnection(constr);
            OleDbCommand cmd = new OleDbCommand("Select "+sutunlar+" From [" + cmbMalik.SelectedItem.ToString() + "$] where CepTelefonu is not null " + oturumFiltre + " " + sattiFiltre + " " + kiraciFiltre + " " + telefonFiltre + " ", conn);
            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                dgvMalik.DataSource = dt;
            }
            catch
            {
            }
            //��kt� dosyas�n� excel' d�n��t�r�p C s�r�c�s� i�erinde excel klas�r� olu�turup girilen isimle kaydetme
            if (dgvMalik.Rows.Count>0 & txtMalikExcel.Text != "")
            {
                DataTable dte = new DataTable();
                foreach (DataGridViewColumn sutun in dgvMalik.Columns)
                {
                    dte.Columns.Add(sutun.HeaderText);
                }

                foreach (DataGridViewRow satir in dgvMalik.Rows)
                {
                    dte.Rows.Add();
                    foreach (DataGridViewCell hucre in satir.Cells)
                    {
                        dte.Rows[dte.Rows.Count - 1][hucre.ColumnIndex] = hucre.Value.ToString();
                    }
                }
                string dosyaYolu = "C:\\Excel\\Malik Kay�tlar\\";
                if (!Directory.Exists(dosyaYolu))
                {
                    Directory.CreateDirectory(dosyaYolu);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dte, "Malik Kay�tlar");
                    wb.SaveAs(dosyaYolu + txtMalikExcel.Text + ".xlsx");
                }
            }
        }

        private void btnIsimleEslestir_Click(object sender, EventArgs e)
        {
            int AciklamaEslesmeyenInd=0;
            int BakiyeEslesmeyenInd= 0;
            int NoEslesmeyenInd = 0;
            List<string> AciklamaList = new List<string>();
            List<string> TelefonList = new List<string>();
            List<string> BakiyeList = new List<string>();
            if (dgvEslesmeyenler.Columns.Count>0)
            {
                for (int i = 0; i < dgvEslesmeyenler.Columns.Count; i++)
                {
                    if (dgvEslesmeyenler.Columns[i].Name== "ACIKLAMA")
                    {
                        AciklamaEslesmeyenInd=dgvEslesmeyenler.Columns[i].Index;
                    }else if (dgvEslesmeyenler.Columns[i].Name== "BAKIYE")
                    {
                        BakiyeEslesmeyenInd = dgvEslesmeyenler.Columns[i].Index;
                    }else if (dgvEslesmeyenler.Columns[i].Name == "NO")
                    {
                        NoEslesmeyenInd=dgvEslesmeyenler.Columns[i].Index;
                    }
                }
            }

            int AdMalikInd = 0;
            int SoyadMalikInd = 0;
            int CepTelefonuMalikInd = 0;
            int KimlikMalikInd = 0;
            if (dataGridView1.Columns.Count>0)
            {
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (dataGridView1.Columns[i].Name=="Ad�")
                    {
                        AdMalikInd=dataGridView1.Columns[i].Index;
                    }else if (dataGridView1.Columns[i].Name=="Soyad")
                    {
                        SoyadMalikInd = dataGridView1.Columns[i].Index;
                    }else if (dataGridView1.Columns[i].Name== "CepTelefonu")
                    {
                        CepTelefonuMalikInd= dataGridView1.Columns[i].Index;
                    }else if (dataGridView1.Columns[i].Name == "Kimlik")
                    {
                        KimlikMalikInd= dataGridView1.Columns[i].Index;
                    }
                }
            }
            List<string> eklenemeyenler = new List<string>();
            if (dgvEslesmeyenler.Rows.Count>0)
            {
                for (int i = 0; i < dgvEslesmeyenler.Rows.Count; i++)
                {
                    int eslesti = 0;
                    string aciklama = dgvEslesmeyenler.Rows[i].Cells[AciklamaEslesmeyenInd].Value.ToString();
                    string No = dgvEslesmeyenler.Rows[i].Cells[NoEslesmeyenInd].Value.ToString();
                    string[] aciklamaArray = aciklama.Split(" ");
                    aciklama = "";
                    for (int j = 1; j < aciklamaArray.Length; j++)
                    {
                        aciklama += aciklamaArray[j];
                    }

                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    {
                        string adMalik = dataGridView1.Rows[j].Cells[AdMalikInd].Value.ToString();
                        adMalik = adMalik.Trim();
                        string soyadMalik = dataGridView1.Rows[j].Cells[SoyadMalikInd].Value.ToString();
                        soyadMalik = soyadMalik.Trim();
                        string CepTelefonuMalik = dataGridView1.Rows[j].Cells[CepTelefonuMalikInd].Value.ToString();
                        if (aciklama.Contains(adMalik+soyadMalik))
                        {
                            AciklamaList.Add(dgvEslesmeyenler.Rows[i].Cells[AciklamaEslesmeyenInd].Value.ToString());
                            TelefonList.Add(CepTelefonuMalik);
                            BakiyeList.Add(dgvEslesmeyenler.Rows[i].Cells[BakiyeEslesmeyenInd].Value.ToString());
                            eslesti++;
                            goto don;
                        }
                       
                    }
                    don: ;
                    if (eslesti < 1&dataGridView1.Rows.Count>0)
                    {
                        aciklama= dgvEslesmeyenler.Rows[i].Cells[AciklamaEslesmeyenInd].Value.ToString();
                        for (int j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            if (dataGridView1.Rows[j].Cells[KimlikMalikInd].Value.ToString().ToLower().Contains(No.ToLower()))
                            {
                                int isimleEslesti = 0;
                                string adMalik = dataGridView1.Rows[j].Cells[AdMalikInd].Value.ToString();
                                string[] adStrings = adMalik.Split(" ");
                                string soyadMalik = dataGridView1.Rows[j].Cells[SoyadMalikInd].Value.ToString();
                                string[] soyadStrings = soyadMalik.Split(" ");
                                string CepTelefonuMalik = dataGridView1.Rows[j].Cells[CepTelefonuMalikInd].Value.ToString();
                                foreach (string adString in adStrings)
                                {
                                    if (aciklama.ToLower().Contains(adString.ToLower()))
                                    {
                                        isimleEslesti++;
                                    }
                                }

                                foreach (string soyadString in soyadStrings)
                                {
                                    if (aciklama.ToLower().Contains(soyadString.ToLower()))
                                    {
                                        isimleEslesti++;
                                    }
                                }
                                if (isimleEslesti>0)
                                {
                                    AciklamaList.Add(dgvEslesmeyenler.Rows[i].Cells[AciklamaEslesmeyenInd].Value.ToString());
                                    TelefonList.Add(CepTelefonuMalik);
                                    BakiyeList.Add(dgvEslesmeyenler.Rows[i].Cells[BakiyeEslesmeyenInd].Value.ToString());
                                    eslesti++;
                                    goto don2;
                                }
                            }
                        }
                        don2: ;
                    }

                    if (eslesti<1)
                    {
                        eklenemeyenler.Add(dgvEslesmeyenler.Rows[i].Cells[AciklamaEslesmeyenInd].Value.ToString());
                    }
                }
            }

            isimleEslestir iE = new isimleEslestir();
            iE.listeDoldur(AciklamaList,BakiyeList,TelefonList);
            eklenemeyenler frmeEklenemeyenler = new eklenemeyenler();
            frmeEklenemeyenler.listeDoldur(eklenemeyenler);
        }

        private void btnLogoDonustur_Click(object sender, EventArgs e)
        {
            LogoDonustur frmlLogoDonustur = new LogoDonustur();
            frmlLogoDonustur.ShowDialog();
        }

        private void cmbBolge_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSayfa2.PerformClick();
        }
    }
}