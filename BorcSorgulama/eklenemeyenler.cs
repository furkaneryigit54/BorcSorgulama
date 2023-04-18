using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BorcSorgulama
{
    public partial class eklenemeyenler : Form
    {
        public eklenemeyenler()
        {
            InitializeComponent();
        }

        private void eklenemeyenler_Load(object sender, EventArgs e)
        {
            tabPage1.Text = "Eşleşmeyenler";
        }

        public void listeDoldur(List<string> degerler)
        {
            dataGridView1.Columns.Clear();
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "Sıra";
            dataGridView1.Columns[1].Name = "Ünite Bilgileri";
            int sira = 1;
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView1);
            for (int i = 0; i < degerler.Count; i++)
            {
                dataGridView1.Rows.Add(sira,degerler[i]);
                sira++;
            }
            this.Show();
        }

        public void excelAktar(string dosyaYolu)
        {
            //Çıktı dosyasını excel' dönüştürüp C sürücüsü içerinde excel klasörü oluşturup girilen isimle kaydetme
            if (dataGridView1.Rows.Count > 0 & dosyaYolu != string.Empty)
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn sutun in dataGridView1.Columns)
                {
                    dt.Columns.Add(sutun.HeaderText);
                }

                foreach (DataGridViewRow satir in dataGridView1.Rows)
                {
                    dt.Rows.Add();
                    foreach (DataGridViewCell hucre in satir.Cells)
                    {
                        dt.Rows[dt.Rows.Count - 1][hucre.ColumnIndex] = hucre.Value.ToString();
                    }
                }
                string klasorYolu = "C:\\Excel\\Eşleşmeyenler\\";
                try
                {
                    if (!Directory.Exists(klasorYolu))
                    {
                        Directory.CreateDirectory(klasorYolu);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(dt, "Eşleşmeyen Kayıtlar");
                        wb.SaveAs(klasorYolu + dosyaYolu + ".xlsx");
                    }
                }
                finally{}
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelAktar(textBox1.Text);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count>0 & textBox1.Text!="")
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }
    }
}
