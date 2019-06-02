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

namespace frmexcelexample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C://Users/burak_000/Desktop/burak.xlsx; Extended Properties='Excel 12.0 Xml;HDR=YES'");


     
                xlsxbaglanti.Open();
                OleDbCommand komutt = new OleDbCommand("SELECT * FROM [Sayfa1$]", xlsxbaglanti);
                OleDbDataReader oku = komutt.ExecuteReader();
                while (oku.Read())
                {

                    string adSoyad = oku["ad"].ToString();
                    string Cinsiyet = oku["numara"].ToString();

                    //baglan.Open();
                    //SqlCommand komut = new SqlCommand("insert into kisi (AdSoyad,cisiyet,yas) values('" + adSoyad.ToString() + "','" + Cinsiyet.ToString() + "') ", baglan);
                    //komut.ExecuteNonQuery();
                    //baglan.Close();




                }
                xlsxbaglanti.Close();
            }
    }
}
