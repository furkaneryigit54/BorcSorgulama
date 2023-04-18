using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BorcSorgulama
{
    public partial class LogoDonustur : Form
    {
        public LogoDonustur()
        {
            InitializeComponent();
        }

        private void btnDosyaYolu_Click(object sender, EventArgs e)
        {
            //Malik tablosu için excel'in dosya yolunu seçme ekranını açma
            cmbMalik.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyası Seçiniz..";
            file.ShowDialog();
            txtDosyaYolu.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu çağırıp sayfa isimlerini alma
            if (txtDosyaYolu.Text != "")
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
               
                if (cmbMalik.Items.Count>0)
                {
                    cmbMalik.Enabled = true;
                    cmbMalik.SelectedIndex = 0;
                    btnSayfa.Enabled=true;
                    btnSayfa.PerformClick();
                }
                else
                {
                    cmbMalik.Enabled = false;
                    btnSayfa.Enabled = false;
                }
            }
        }
        public List<DataRow> ExcelSayfaAdlariGetir(string DosyaYolu)
        {
            //Excel dosyasındaki sayfa isimlerini çekme
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

        private void btnSayfa_Click(object sender, EventArgs e)
        {
            //Seçilen dosyadaki verileri datagridviewe yükleme
            try
            {
                String constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu.Text +
                                "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand cmd =
                    new OleDbCommand("Select * From [" + cmbMalik.SelectedItem.ToString() + "$] ",
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
                MessageBox.Show("Geçerli Bir Sayfa İsmi Seçin!", "MALİK");
            }
        }

        private void btnDosyaYolu2_Click(object sender, EventArgs e)
        {
            //Logo tablosu için excel'in dosya yolunu seçme ekranını açma
            cmbLogo.Items.Clear();
            OpenFileDialog file = new OpenFileDialog();
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyası Seçiniz..";
            file.ShowDialog();
            txtDosyaYolu2.Text = file.FileName;
            //Excel sayfa isimlerini alan metodu çağırıp sayfa isimlerini alma
            if (txtDosyaYolu2.Text != "")
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
                //Veri girişi yapıldığında pasif araçları aktif hale getirme
                
                if (cmbLogo.Items.Count > 0)
                {
                    cmbLogo.Enabled = true;
                    btnSayfa2.Enabled = true;
                    int logoind = 0;
                    for (int i = 0; i < cmbLogo.Items.Count; i++)
                    {
                        if (cmbLogo.Items[i].ToString().ToLower() == "logo")
                        {
                            logoind = i;
                        }
                    }
                    cmbLogo.SelectedIndex = logoind;
                    btnSayfa2.PerformClick();
                }
                else
                {
                    cmbLogo.Enabled = false;
                    btnSayfa2.Enabled = false;
                }
            }
        }

        private void btnSayfa2_Click(object sender, EventArgs e)
        {
            //Seçilen dosyadaki verileri datagridviewe yükleme
            try
            {
                String constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu2.Text +
                                "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
                OleDbConnection con = new OleDbConnection(constr);
                OleDbCommand cmd =
                    new OleDbCommand("Select * From [" + cmbLogo.SelectedItem.ToString() + "$] ",
                        con);
                con.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                con.Close();
            }
            catch (Exception ex)
            {
                dataGridView1.DataSource = null;
                MessageBox.Show("Geçerli Bir Sayfa İsmi Seçin!", "LOGO");
            }
        }

        private void excelIsim_TextChanged(object sender, EventArgs e)
        {
            if (excelIsim.Text!="")
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = null;
            dataGridView3.ColumnCount = dataGridView2.ColumnCount;
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView3.Columns[i].Name = dataGridView2.Columns[i].Name;
            }
            if (dataGridView1.Rows.Count>0&dataGridView2.Rows.Count>0)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    string CariHesapKodu = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {
                        if (dataGridView2.Rows[j].Cells[0].Value.ToString()==CariHesapKodu)
                        {
                            dataGridView3.Rows.Add(dataGridView2.Rows[j].Cells[0].Value,
                                dataGridView2.Rows[j].Cells[1].Value, dataGridView2.Rows[j].Cells[2].Value,
                                dataGridView2.Rows[j].Cells[3].Value, dataGridView1.Rows[i].Cells[2].Value,
                                0, dataGridView2.Rows[j].Cells[6].Value,
                                dataGridView2.Rows[j].Cells[7].Value, dataGridView2.Rows[j].Cells[8].Value,
                                dataGridView2.Rows[j].Cells[9].Value, dataGridView1.Rows[i].Cells[2].Value);
                        }
                    }
                    
                }
            }
            if (dataGridView3.Rows.Count > 0 & excelIsim.Text != string.Empty)
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn sutun in dataGridView2.Columns)
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
                string dosyaYolu = "C:\\Excel\\Dönüştürülen Excel Dosyaları\\";
                if (!Directory.Exists(dosyaYolu))
                {
                    Directory.CreateDirectory(dosyaYolu);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt, "LOGO");
                    wb.SaveAs(dosyaYolu + excelIsim.Text + ".xlsx");
                }
            }

        }

        private void LogoDonustur_Load(object sender, EventArgs e)
        {
            tabPage1.Text = "Dönüştürülecek Dosya";
            tabPage2.Text = "Logo Dosyası";
        }

        private void cmbMalik_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSayfa.PerformClick();
        }

        private void cmbLogo_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSayfa2.PerformClick();
        }
    }
}
