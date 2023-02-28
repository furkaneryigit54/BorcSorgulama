using System.Data;
using System.Data.OleDb;

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
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = "Excel Dosyasý |*.xlsx| Excel Dosyasý|*.xls";  
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyasý Seçiniz..";
            file.ShowDialog();

            //string DosyaYolu = file.FileName;
            //string DosyaAdi = file.SafeFileName;
            txtDosyaYolu.Text = file.FileName;
        }

        private void btnSayfa_Click(object sender, EventArgs e)
        {
            string constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + txtDosyaYolu.Text +
                            "; Extended Properties =\"Excel 8.0; HDR = Yes;\";";
            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + txtSayfaIsmi.Text + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;
        }

    }

    
}