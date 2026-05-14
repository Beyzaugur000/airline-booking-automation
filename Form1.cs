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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
namespace finalproje2_beyza_ugur_21907047
{
    public partial class Form1 : Form
    {
       
        
        private object webBrowser1;

        public Form1()
        {
            InitializeComponent();
        }

      

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }
     
        private string TcDogruMu(string tcNo)
        {
            string durum = "";
            try
            {
                if (tcNo != "")
                {
                    if (tcNo.Length == 11)
                    {
                        char[] rakamlar = tcNo.ToCharArray();
                        int kural1 = 0, hane11 = rakamlar[10], hane10 = rakamlar[9];

                        for (int i = 0; i < 10; i++)
                        {
                            kural1 += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural1 = kural1.ToString().ToCharArray();

                        int kural2tek = 0, kural2cift = 0;

                        for (int i = 0; i < 10; i += 2)
                        {
                            kural2tek += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        for (int i = 1; i < 9; i += 2)
                        {
                            kural2cift += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural2 = ((7 * kural2tek) + (9 * kural2cift)).ToString().ToCharArray();

                        int kural3 = 0;

                        for (int i = 0; i < 10; i += 2)
                        {
                            kural3 += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural3 = (8 * kural3).ToString().ToCharArray();

                        if ((birlerbasamagikural1[birlerbasamagikural1.Length - 1] == hane11) && (birlerbasamagikural2[birlerbasamagikural2.Length - 1] == hane10) && (birlerbasamagikural3[birlerbasamagikural3.Length - 1] == hane11))
                        {
                            durum = "Kimlik Numarası Geçerli";
                        }
                        else
                        {
                            durum = "Kimlik Numarası Geçerli Değildir";
                        }

                        textBox1.Focus();
                    }
                    else
                    {
                        durum = "TC Kimlik Numaranızı Eksik Girdiniz\n Lütfen Kontrol Ediniz!!!";
                    }
                }
                else
                {
                    durum = "Lütfen TC Kimlik Numaranızı Giriniz!!!";
                }
            }
            catch (Exception ex)
            {
                durum = ex.Message;
            }
            return durum;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
             {
                label12.Text = (TcDogruMu( textBox1.Text));
   
             }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser2.Navigate("https://www.thy.com");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns.Add(label2.Text, label2.Text);
            dataGridView1.Columns.Add(label3.Text, label3.Text);
            dataGridView1.Columns.Add(label4.Text, label4.Text);
            dataGridView1.Columns.Add(label5.Text, label5.Text);
            dataGridView1.Columns.Add(label6.Text, label6.Text);
            dataGridView1.Columns.Add(label7.Text, label7.Text);
            dataGridView1.Columns.Add(label8.Text, label8.Text);
            dataGridView1.Columns.Add(label9.Text, label9.Text);
            dataGridView1.Columns.Add(label10.Text, label10.Text);
            dataGridView1.Columns.Add(label11.Text, label11.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbCommand komut = new OleDbCommand();

            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange =
                    (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange =
                            (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        myRange.Select();
                    }
                    catch
                    {
                        ;
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, comboBox1.Text, comboBox2.Text, comboBox3.Text, comboBox4.Text, textBox5.Text, textBox6.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true) { textBox5.Text = textBox5.Text + checkBox1.Text; }
            if (checkBox2.Checked == true) { textBox5.Text = textBox5.Text + checkBox2.Text; }
            if (checkBox3.Checked == true) { textBox5.Text = textBox5.Text + checkBox3.Text; }
            if (checkBox4.Checked == true) { textBox5.Text = textBox5.Text + checkBox4.Text; }
            if (checkBox5.Checked == true) { textBox5.Text = textBox5.Text + checkBox5.Text; }
            if (checkBox6.Checked == true) { textBox5.Text = textBox5.Text + checkBox6.Text; }
            if (checkBox7.Checked == true) { textBox5.Text = textBox5.Text + checkBox7.Text; }
            if (checkBox8.Checked == true) { textBox5.Text = textBox5.Text + checkBox8.Text; }
            if (checkBox9.Checked == true) { textBox5.Text = textBox5.Text + checkBox9.Text; }
            if (checkBox10.Checked == true) { textBox5.Text = textBox5.Text + checkBox10.Text; }
            if (checkBox11.Checked == true) { textBox5.Text = textBox5.Text + checkBox11.Text; }
            if (checkBox12.Checked == true) { textBox5.Text = textBox5.Text + checkBox12.Text; }
            if (checkBox13.Checked == true) { textBox5.Text = textBox5.Text + checkBox13.Text; }
            if (checkBox14.Checked == true) { textBox5.Text = textBox5.Text + checkBox14.Text; }
            if (checkBox15.Checked == true) { textBox5.Text = textBox5.Text + checkBox15.Text; }
            if (checkBox16.Checked == true) { textBox5.Text = textBox5.Text + checkBox16.Text; }
            if (checkBox17.Checked == true) { textBox5.Text = textBox5.Text + checkBox17.Text; }
            if (checkBox18.Checked == true) { textBox5.Text = textBox5.Text + checkBox18.Text; }
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }
    }
  
}
