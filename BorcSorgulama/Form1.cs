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

            //Malik tablosu için excel'in dosya yolunu seçme ekranýný açma
            cmbMalik.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyasý Seçiniz..";
            file.ShowDialog();
            txtDosyaYolu.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu çaðýrýp sayfa isimlerini alma
            if (txtDosyaYolu.Text!="")
            {
                List<DataRow> SayfalarDR = ExcelSayfaAdlariGetir(txtDosyaYolu.Text).ToList();
                List<string> Sayfalar = new List<string>();
                string kesmeisareti = "'";
                foreach (DataRow dr in SayfalarDR)
                {
                    Sayfalar.Add(dr["TABLE_NAME"].ToString().Trim('$', kesmeisareti[0]));
                }
                //Çekilen sayfa isimlerini combobox'a atama
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
            //Seçilen dosyadaki verileri datagridviewe yükleme
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
                MessageBox.Show("Geçerli Bir Sayfa Ýsmi Seçin!","MALÝK");
            }
            //Malik sayfasýnýn sütun isimlerini alýp combobox'a atama
            string[] malikSutunlar = new string[dataGridView1.Columns.Count];
            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                malikSutunlar[i] = (dataGridView1.Columns[i].Name);
            }
            foreach (var sutun in malikSutunlar)
            {
                clbSutunlar.Items.Add(sutun);
            }
            //Datagridview'e veri eklenince devre dýþý olan araçlarý aktif hale getirme
            if (dataGridView1.Rows.Count > 0)
            {
                clbOturum.Items.Clear();
                clbKiraci.Items.Clear();
                clbSatti.Items.Clear();
                clbTelefon.Items.Clear();
                string[] dogruYanlis = new[] { "DOÐRU", "YANLIÞ" };
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
            //Excel dosyasýndaki sayfa isimlerini çekme
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
            //Logo tablosu için excel'in dosya yolunu seçme ekranýný açma
            cmbLogo.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyasý Seçiniz..";
            file.ShowDialog();
            txtDosyaYolu2.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu çaðýrýp sayfa isimlerini alma
            if (txtDosyaYolu2.Text!="")
            {
                List<DataRow> SayfalarDR = ExcelSayfaAdlariGetir(txtDosyaYolu2.Text).ToList();
                List<string> Sayfalar = new List<string>();
                string kesmeisareti = "'";
                foreach (DataRow dr in SayfalarDR)
                {
                    Sayfalar.Add(dr["TABLE_NAME"].ToString().Trim('$', kesmeisareti[0]));
                }
                //Çekilen sayfa isimlerini combobox'a atama
                foreach (string sayfa in Sayfalar)
                {
                    cmbLogo.Items.Add(sayfa);
                }
                //Veri giriþi yapýldýðýnda pasif araçlarý aktif hale getirme
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
                    if (secilen == "výllalar")
                    {
                        cmbBolge.SelectedIndex = 1;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "carsý_evlerý")
                    {
                        cmbBolge.SelectedIndex = 2;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "lbloklarý")
                    {
                        cmbBolge.SelectedIndex = 3;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "acarblu")
                    {
                        cmbBolge.SelectedIndex = 4;
                        btnSayfa2.PerformClick();
                    }
                    else if (secilen == "acar_vadý")
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
            //Seçilen dosyadaki verileri datagridviewe yükleme
            clbSutunlarLogo.Items.Clear();
            clbOdemeTuruLogo.Items.Clear();
            clbKatLogo.Items.Clear();
            string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu2.Text +
                            "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
            OleDbConnection con = new OleDbConnection(constr);
            //Bölge seçimi için filtreleme
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
                bolge = "where AKRO_KODU LÝKE 'VRD%'";
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
                MessageBox.Show("Geçerli Bir Sayfa Ýsmi Seçin!","LOGO");
            }
            try
            {
                //Logo dosyasýnýn sutun isimlerini, filtreleme için kat ve odemeturu sütunlarýný çekip combobox'a atama
                string[] logoSutunlar = new string[dataGridView2.Columns.Count];
                string[] katlar = new string[dataGridView2.Rows.Count - 1];
                string[] odemeTuru = new string[dataGridView2.Rows.Count - 1];
                //Kat ve ödeme türü sütunundaki deðerleri dizilere atama
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
                //Sütun isimlerini combobox'a atama
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
                //Kat listesindeki tekrar eden deðerleri silme ve sýralama
                katList = katList.Distinct().ToList();
                katList.Sort();
                for (int i = 0; i < katList.Count - 1; i++)
                {
                    if (katList[i] == "")
                    {
                        katList.RemoveAt(i);
                    }
                }
                //Ödeme türü listesindeki tekrar eden deðerleri silme ve sýralama
                odemeTuruList = odemeTuruList.Distinct().ToList();
                odemeTuruList.Sort();
                for (int i = 0; i < odemeTuruList.Count - 1; i++)
                {
                    if (odemeTuruList[i] == "")
                    {
                        odemeTuruList.RemoveAt(i);
                    }
                }
                //Kat ve ödeme türü bilgilerini combobox'lara atama
                for (int i = 0; i < katList.Count; i++)
                {
                    clbKatLogo.Items.Add(katList[i]);
                }
                foreach (string odemeturu in odemeTuruList)
                {
                    clbOdemeTuruLogo.Items.Add(odemeturu);
                }
                //Bakiye deðerlerini tekrar hesaplayýp bakiye sütununa atama
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
            //Sütun filtresi combobox'ýna veri giriþi yapýldýðýnda pasif haldeki araçlarý aktif hale getirme
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
                string[] clbbakiyeStrings = new[] { "Büyüktür","Küçüktür","Eþittir","Büyük Eþittir","Küçük Eþittir" };
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
            //Datagridview'daki bütün sütunlarý görünür hale getirme
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
            //Datagridview'a gerekli sütun isimlerini ekleme
            dataGridView3.Columns.Clear();
            dataGridView3.ColumnCount = 4;
            dataGridView3.Columns[0].Name = "Sýra No";
            dataGridView3.Columns[1].Name = "AD";
            dataGridView3.Columns[2].Name = "TUTAR";
            dataGridView3.Columns[3].Name = "TELEFON";
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView3);
            //Çýktý dosyasý için gerekli olan 2 dosyayadaki verileri eþleþtirecek deðiþkenlerin indexlerinin tanýmý
            //Logo dosyasý index deðiþkenleri
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
            //Malik dosyasý index deðiþkenleri
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
                else if (dataGridView1.Columns[i].Name == "Adý")
                {
                    adiMalikind = dataGridView1.Columns[i].Index;
                }
                else if (dataGridView1.Columns[i].Name == "Soyad")
                {
                    soyadMalikind = dataGridView1.Columns[i].Index;
                }
            }
            //Ýki dosyadaki verileri eþleþtirip çýktý dosyasý için datagridview'e verileri ekleme
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
                    //Logo dosyasýndan no ve kat deðerlerini string deðiþkenlere tanýmlama
                    string noLogo = dataGridView2.Rows[i].Cells[NoLogoind].Value.ToString();
                    string katLogo = dataGridView2.Rows[i].Cells[KatLogoind].Value.ToString();
                    //Eþleþtirme için logo dosyasýndaki kat sütunundan çekilen verilerden 0 sayýlýrýný silme
                    if (cmbBolge.SelectedItem != "ÇARÞI EVLERÝ" & cmbBolge.SelectedItem != "L BLOKLARI")
                    {
                        string[] katlist = katLogo.Split('0');
                        katLogo = "";
                        foreach (string deðer in katlist)
                        {
                            katLogo += deðer;
                        }
                    }
                    //Logo tablosundan no ve kat bilgilerini birleþtirerek ünite deðiþkenini tanýmlama
                    string uniteLogo = noLogo + katLogo;
                    //Logo tablosundan borç bilgisini deðiþkene tanýmlama
                    string borclogo = dataGridView2.Rows[i].Cells[borcLogoind].Value.ToString();
                    for (int j = 0; j < dataGridView1.RowCount ; j++)
                    {
                        //Malik tablosundan ünite bilgisini deðiþkene atayýp eþleþtirmek için nokta varsa silme
                        string unite = dataGridView1.Rows[j].Cells[kimlikMalikind].Value.ToString();
                        string[] uniteList = unite.Split('.');
                        //Ünite deðiþkenindeki A ve B deðerlerini eþleþtirme için sayýya çevirme
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
                        //Yapýlan iþlemlerden sonra ünite deðikenini oluþturulan listeden tekrar birleþtirme
                        unite = "";
                        foreach (var deðer in uniteList)
                        {
                            unite += deðer;
                        }
                        //Ünite bilgileri iki tablodada uyuþan kiþilerin bilgilerini çýktý tablosuna ekleme ve döngüyü sonlandýrma
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
            //Excel'e verilecek isim textboxuna giriþ yapýldýðýnda ve çýktý datagridview'ýna veri giriþi yapýldýðýnda dosyayý kaydetme tuþunu aktif hale getirme
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
            //Malik tablosunda oturum bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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

                        if (clbOturum.CheckedItems[i].ToString() == "DOÐRU")
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
            //Malik tablosunda sattý bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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

                        if (clbSatti.CheckedItems[i].ToString() == "DOÐRU")
                        {
                            sattiFiltre = "and Sattý=True";
                        }
                        else
                        {
                            sattiFiltre = "and Sattý=False";
                        }
                    }
                }
            }
            // Malik tablosunda kiracý bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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

                        if (clbKiraci.CheckedItems[i].ToString() == "DOÐRU")
                        {
                            kiraciFiltre = "and Kiracý=True";
                        }
                        else
                        {
                            kiraciFiltre = "and Kiracý=False";
                        }
                    }
                }
            }
            //Malik tablosunda telefon bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
            //Malik tablosunun olduðu datagridview'i temizleyip filtrelerle oluþturulan yeni sorguyu çalýþtýrma
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
                //Tüm sütunlarý görünmez hale getirme
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].Visible = false;
                }
                //Filtrelenen sütunlarý görünür hale getirme
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
                MessageBox.Show("Geçerli Bir Sayfa Seçin!","MALÝK");
                if (dataGridView1.Rows.Count > 0)
                {
                    clbOturum.Items.Clear();
                    clbKiraci.Items.Clear();
                    clbSatti.Items.Clear();
                    clbTelefon.Items.Clear();
                    string[] dogruYanlis = new[] { "DOÐRU", "YANLIÞ" };
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
            tabPage1.Text = "Malik Kayýtlarý";
            tabPage2.Text = "LOGO Kayýtlarý";
            tabPage3.Text = "ÝÞLENMÝÞ TABLO";
            tabPage4.Text = "Malik Filtreler";
            tabPage5.Text = "LOGO Filtreler";
            cmbBolge.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLogo.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbMalik.DropDownStyle = ComboBoxStyle.DropDownList;
            this.MaximizeBox = false;
        }
        private void btnFiltreLogo_Click(object sender, EventArgs e)
        {
            // Logo tablosunda kat bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
            // Logo tablosunda ödeme türü bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
            // Logo tablosunda bakiye bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
                        if (clbBakiyeLogo.CheckedItems[i] == "Büyüktür")
                        {
                            bakiye = "and BAKIYE>" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "Küçüktür")
                        {
                            bakiye = "and BAKIYE<" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "Eþittir")
                        {
                            bakiye = "and BAKIYE=" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "Büyük Eþittir")
                        {
                            bakiye = "and BAKIYE>=" + txtBakiyeLogo.Text + "";
                        }
                        else if (clbBakiyeLogo.CheckedItems[i] == "Küçük Eþittir")
                        {
                            bakiye = "and BAKIYE<=" + txtBakiyeLogo.Text + "";
                        }


                    }
                }
            }
            // Logo tablosunda bölge bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
                bolge = "and AKRO_KODU LÝKE 'VRD%'";
            }
            else if (cmbBolge.SelectedIndex == 0)
            {
                bolge = "";
            }
            //Logo tablosunundaki verileri temizleyip oluþturulan yeni sorguyu çalýþtýrma
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
                MessageBox.Show("Geçerli Bir Sayfa Seçin!","LOGO");
            }
            //Bakiye sütunundaki verileri hesaplayýp yeni deðerler ile deðiþtirme
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
            
            //Tüm sütunlarý görünmez hale getirme
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].Visible = false;
            }
            //Filtrelenen sütunlarý görünür hale getirme
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
                //Logo tablosu datagridview'ýna veri giriþi yapýldýðýnda pasif haldeki araçlarý aktif hale getirme
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
            //Malik dosyasýndan sayfa ismi seçildiðine verileri getirecek butonu aktif hale getirme
            if (cmbMalik.SelectedItem != "")
            {
                btnSayfa.Enabled = true;
            }
            else
            {
                btnSayfa.Enabled = false;
            }

            string secilen = cmbMalik.SelectedItem.ToString().ToLower();
            if (secilen == "výllalar")
            {
                cmbBolge.SelectedIndex = 1;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "carsý_evlerý")
            {
                cmbBolge.SelectedIndex = 2;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "lbloklarý")
            {
                cmbBolge.SelectedIndex = 3;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "acarblu")
            {
                cmbBolge.SelectedIndex = 4;
                btnSayfa2.PerformClick();
            }
            else if (secilen == "acar_vadý")
            {
                cmbBolge.SelectedIndex = 5;
                btnSayfa2.PerformClick();
            }
            btnSayfa.PerformClick();
        }
        private void cmbLogo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Logo dosyasýndan sayfa ismi seçildiðine verileri getirecek butonu aktif hale getirme
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
            //Çýktý dosyasýný excel' dönüþtürüp C sürücüsü içerinde excel klasörü oluþturup girilen isimle kaydetme
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
            //Excel'e verilecek isim textboxuna giriþ yapýldýðýnda ve çýktý datagridview'ýna veri giriþi yapýldýðýnda dosyayý kaydetme tuþunu aktif hale getirme
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

                        if (clbOturum.CheckedItems[i].ToString() == "DOÐRU")
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
            //Seçilen sütunlarý filtreleme
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
            //Malik tablosunda sattý bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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

                        if (clbSatti.CheckedItems[i].ToString() == "DOÐRU")
                        {
                            sattiFiltre = "and Sattý=True";
                        }
                        else
                        {
                            sattiFiltre = "and Sattý=False";
                        }
                    }
                }
            }
            // Malik tablosunda kiracý bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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

                        if (clbKiraci.CheckedItems[i].ToString() == "DOÐRU")
                        {
                            kiraciFiltre = "and Kiracý=True";
                        }
                        else
                        {
                            kiraciFiltre = "and Kiracý=False";
                        }
                    }
                }
            }
            //Malik tablosunda telefon bilgisini filtrelemek için koþul oluþturup yeni sorguya ekleme
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
            //Malik tablosunun olduðu datagridview'i temizleyip filtrelerle oluþturulan yeni sorguyu çalýþtýrma
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
            //Çýktý dosyasýný excel' dönüþtürüp C sürücüsü içerinde excel klasörü oluþturup girilen isimle kaydetme
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
                string dosyaYolu = "C:\\Excel\\Malik Kayýtlar\\";
                if (!Directory.Exists(dosyaYolu))
                {
                    Directory.CreateDirectory(dosyaYolu);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dte, "Malik Kayýtlar");
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
                    if (dataGridView1.Columns[i].Name=="Adý")
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