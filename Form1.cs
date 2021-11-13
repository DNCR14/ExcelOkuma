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

namespace ExcelOkuma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnDosyaSec_Click(object sender, EventArgs e)
        {
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                // seçtiğimiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel bağlantı adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //bağlantı 
                OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);

                //tüm verileri seçmek için select sorgumuz. Sayfa1 olan kısmı Excel'de hangi sayfayı açmak istiyosanız orayı yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + "Sayfa1" + "$]", baglanti);
                
                //bağlantıyı açıyoruz.
                baglanti.Open(); 

                //Gelen verileri DataAdapter'e atıyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluşturuyoruz.
                DataTable data = new DataTable();

                //DataAdapter'da ki verileri data adındaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynağını oluşturduğumuz DataTable ile dolduruyoruz.
                dataGridView1.DataSource = data;
            }
            catch (Exception ex)
            {
                // Hata alırsak ekrana bastırıyoruz.
                MessageBox.Show(ex.Message);
            }
        }
    }
}
