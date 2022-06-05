using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace stok_programi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public SqlConnection bag = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=data;Integrated Security=True");

        public DataTable tablo = new DataTable();/
        public SqlDataAdapter adtr = new SqlDataAdapter();
        public SqlCommand kmt = new SqlCommand();
        string DosyaYolu, DosyaAdi = "";
       int id;
        public void listele()
        {
            tablo.Clear();
            bag.Open();
            SqlDataAdapter adtr = new SqlDataAdapter("select stokAdi,stokModeli,stokSeriNo,stokAdedi,stokTarih,kayitYapan From stokbil", bag);
            adtr.Fill(tablo);
            dataGridView1.DataSource = tablo;
            adtr.Dispose();
            bag.Close();
            try
            {
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                //datagridview1'deki tüm satırı seç              
                dataGridView1.Columns[0].HeaderText = "STOK ADI";
                //sütunlardaki textleri değiştirme
                dataGridView1.Columns[1].HeaderText = "STOK MODELİ";
                dataGridView1.Columns[2].HeaderText = "STOK SERİNO";
                dataGridView1.Columns[3].HeaderText = "STOK ADEDİ";
                dataGridView1.Columns[4].HeaderText = "STOK TARİH";
                dataGridView1.Columns[5].HeaderText = "KAYIT YAPAN";
                dataGridView1.Columns[0].Width = 120;
                //sütunların genişliğini belirleme
                dataGridView1.Columns[1].Width = 120;
                dataGridView1.Columns[2].Width = 120;
                dataGridView1.Columns[3].Width = 80;
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[5].Width = 120;               
            }
            catch
            {
                ;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            listele();
                      
        }
    
        private void btnStokEkle_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text.Trim() == "") errorProvider1.SetError(textBox1, "Boş geçilmez");
                else errorProvider1.SetError(textBox1, "");
                if (textBox2.Text.Trim() == "") errorProvider1.SetError(textBox2, "Boş geçilmez");
                else errorProvider1.SetError(textBox2, "");
                if (textBox3.Text.Trim() == "") errorProvider1.SetError(textBox3, "Boş geçilmez");
                else errorProvider1.SetError(textBox3, "");
                if (textBox4.Text.Trim() == "") errorProvider1.SetError(textBox4, "Boş geçilmez");
                else errorProvider1.SetError(textBox4, "");
                if (textBox5.Text.Trim() == "") errorProvider1.SetError(textBox5, "Boş geçilmez");
                else errorProvider1.SetError(textBox5, "");

                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "" && textBox3.Text.Trim() != "" && textBox4.Text.Trim() != "" && textBox5.Text.Trim() != "")
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "INSERT INTO stokbil(stokAdi,stokModeli,stokSeriNo,stokAdedi,stokTarih,kayitYapan,dosyaAdi) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + dateTimePicker1.Text + "','" + textBox5.Text + "','" + DosyaAdi + "') ";
                    kmt.ExecuteNonQuery();
                    kmt.Dispose();
                    bag.Close();
                    for (int i = 0; i < this.Controls.Count; i++)
                    {
                        if (this.Controls[i] is TextBox) this.Controls[i].Text = "";
                    }
                    listele();
                    
                    MessageBox.Show("Kayıt İşlemi Tamamlandı ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
              
            }
            catch 
            {
                MessageBox.Show("Kayıtlı Seri No !");
                bag.Close();              
            }           
        }        
        private void btnResimEkle_Click(object sender, EventArgs e)
        {

            if (DosyaAc.ShowDialog() == DialogResult.OK)
            {
                foreach (string i in DosyaAc.FileName.Split('\\'))
                {
                    if (i.Contains(".jpg")) { DosyaAdi = i; }
                    else if (i.Contains(".bmp")) { DosyaAdi = i; }
                    else if (i.Contains(".png")) { DosyaAdi = i; }
                    else if (i.Contains(".gif")) { DosyaAdi = i; }
                    else { DosyaYolu += i + "\\"; }
                }
                pictureBox1.ImageLocation = DosyaAc.FileName;
                //cmd = new OleDbCommand("inser into tablom (ResimAdi,DosyaYolu,DosyaAdi) values ('TEST',"+
                //MessageBox.Show(DosyaAc.FileName.Split(@"\"));
            }
            else
            {
                MessageBox.Show("Herhangibir Resim Girilmedi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            try
            {
                kmt = new SqlCommand("select * from stokbil where stokSeriNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'", bag);
                bag.Open();
                SqlDataReader oku = kmt.ExecuteReader();
                oku.Read();
                if (oku.HasRows)
                {
                    pictureBox1.ImageLocation = oku[7].ToString();
                    id=Convert.ToInt32(oku[0].ToString());
                }
                bag.Close();
            }
            catch
            {
                bag.Close();
            }
            DosyaAdi = pictureBox1.ImageLocation;
            
        }
        private void btnStokAdiAra_Click(object sender, EventArgs e)
        {
            SqlDataAdapter adtr = new SqlDataAdapter("select * From stokbil", bag);
            if (textBox6.Text.Trim()== "")
            {
                tablo.Clear();
                kmt.Connection = bag;
                kmt.CommandText = "Select * from stokbil";
                adtr.SelectCommand = kmt;
                adtr.Fill(tablo);               
            }
            if (Convert.ToBoolean(bag.State) == false)
            {
                bag.Open();
            }
            if (textBox6.Text.Trim() != "")
            {
                adtr.SelectCommand.CommandText = " Select * From stokbil" +
                     " where(stokAdi='" + textBox6.Text + "' )";
                tablo.Clear();
                adtr.Fill(tablo);
                bag.Close();
            }
        }
        private void btnStokModelAra_Click(object sender, EventArgs e)
        {
            SqlDataAdapter adtr = new SqlDataAdapter("select * From stokbil", bag);
            if (textBox7.Text.Trim() == "")
            {
                tablo.Clear();
                kmt.Connection = bag;
                kmt.CommandText = "Select * from stokbil";
                adtr.SelectCommand = kmt;
                adtr.Fill(tablo);            
            }
            if (Convert.ToBoolean(bag.State) == false)
            {
                bag.Open();
            }
            if (textBox7.Text.Trim() != "")
            {
                adtr.SelectCommand.CommandText = " Select * From stokbil" +
                     " where(stokModeli='" + textBox7.Text + "' )";
                tablo.Clear();
                adtr.Fill(tablo);
                bag.Close();
            }
        }
        private void btnResimSil_Click(object sender, EventArgs e)
        {
            pictureBox1.ImageLocation = "";
            DosyaAdi = "";
        }
        private void btnStokSil_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kaydı silmek istediğinizden eminmisiniz", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes && dataGridView1.CurrentRow.Cells[0].Value.ToString().Trim() != "")
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "DELETE from stokbil WHERE stokSeriNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "' ";
                    kmt.ExecuteNonQuery();
                    kmt.Dispose();
                    bag.Close();
                    listele();
                }
            }
            catch
            {
                ;
            }
        }      
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Sayı Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);               
            }            
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Harf Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }        
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Harf Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }      
        }
        private void btnCikis_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Harf Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }      
        }
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Harf Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }      
        }
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && e.KeyChar != (char)44)
            {
                e.Handled = true;
                MessageBox.Show("Sadece Harf Girişi Yapabilirsiniz ! ", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }      
        }
        private void btnStokGuncelle_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "" && textBox3.Text.Trim() != "" && textBox4.Text.Trim() != "" && textBox5.Text.Trim() != "")
                {
                   
                    //OleDbCommand kmt=new OleDbCommand("UPDATE stokbil SET stokAdi='" + textBox1.Text + "',stokModeli='" + textBox2.Text + "',stokSeriNo='" + textBox3.Text + "',stokAdedi='" + textBox4.Text + "',stokTarih='" + dateTimePicker1.Text + "',kayitYapan='" + textBox5.Text + "',dosyaAdi='" + DosyaAdi + "' WHERE id="+id,bag);
                    string sorgu = "UPDATE stokbil SET stokAdi='" + textBox1.Text + "',stokModeli='" + textBox2.Text + "',stokSeriNo='" + textBox3.Text + "',stokAdedi='" + textBox4.Text + "',stokTarih='" + dateTimePicker1.Text + "',kayitYapan='" + textBox5.Text + "',dosyaAdi='" + DosyaAdi + "' WHERE id="+id;
                    SqlCommand kmt = new SqlCommand(sorgu,bag);
                    bag.Open();
                    kmt.ExecuteNonQuery();
                    kmt.Dispose();
                    bag.Close();
                    listele();
                    
                    MessageBox.Show("Güncelleme İşlemi Tamamlandı !","Bilgi",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Boş Alan Bırakmayınız !");
                }
            //}
            //catch(Exception h)
            //{ MessageBox.Show(h.Message); }
        }        
    }
}
