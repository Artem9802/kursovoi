using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Win32;
using word = Microsoft.Office.Interop.Word;

namespace KPSushiBar
{
    public partial class Заказ_оборудования : Form
    {
        private readonly string TemplateFileName = @"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\ZakazOborud.docx";

        public Заказ_оборудования()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox3.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView1.Columns.Count; c++)
                {
                    if (dataGridView1[c, i].Value.ToString() == textBox2.Text)
                    {
                        dataGridView1.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView2.Columns.Count; c++)
                {
                    if (dataGridView2[c, i].Value.ToString() == textBox4.Text)
                    {
                        dataGridView2.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void оборудованиеПоставщиковToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button7.Enabled = true;
            button6.Enabled = true;
            textBox6.Enabled = true;
            textBox5.Enabled = true;
            textBox7.Enabled = true;
            textBox8.Enabled = true;
            textBox9.Enabled = true;
            textBox10.Enabled = true;
            textBox11.Enabled = true;
            tabControl1.Enabled = true;
            textBox13.Enabled = false;
            textBox12.Enabled = false;
            textBox14.Enabled = false;
            textBox15.Enabled = false;

            textBox3.Enabled = true;
            textBox4.Enabled = true;
            label1.Text = "Оборудование";
            label2.Text = "Количество";
            label3.Text = "Марка";
            label4.Text = "Продукты";
            label5.Text = "Тип продуктов";
            label6.Text = "Цена оборудования";
            label7.Text = "Цена продуктов";

            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Tovar_post] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView2.DataSource = table;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[8].Visible = false;
            con.Close();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void Заказ_оборудования_Load(object sender, EventArgs e)
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
           
            tabControl1.Enabled = false;

            label1.Text = "";
            label2.Text = "";
            label3.Text = "";
            label4.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";

            label8.Text = "Поиск";
            label9.Text = "Фильтрация";

        }

        private void страныToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Text = "";
            label2.Text = "";
            label3.Text = "";
            label4.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";


            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
            textBox10.Enabled = false;
            textBox11.Enabled = false;
            button5.Enabled = false;
            textBox14.Enabled = false;
            textBox15.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;


            tabControl1.Enabled = true;
            textBox13.Enabled = true;
            textBox12.Enabled = true;


            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Stran] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView2.DataSource = table;
            dataGridView2.Columns[0].Visible = false;
            con.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Главное_меню glavMenu = new Главное_меню();
            glavMenu.Show();
            this.Close();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            textBox5.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            textBox6.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
            textBox7.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            textBox8.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
            textBox9.Text = dataGridView2.CurrentRow.Cells[5].Value.ToString();
            textBox10.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
            textBox11.Text = dataGridView2.CurrentRow.Cells[7].Value.ToString();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {
            dataGridView2.CurrentCell = null;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView2.Columns.Count; c++)
                {
                    if (dataGridView2[c, i].Value.ToString() == textBox4.Text)
                    {
                        dataGridView2.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox13.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            dataGridView2.CurrentCell = null;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView2.Columns.Count; c++)
                {
                    if (dataGridView2[c, i].Value.ToString() == textBox12.Text)
                    {
                        dataGridView2.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void городаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Text = "";
            label2.Text = "";
            label3.Text = "";
            label4.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";

            tabControl1.Enabled = true;
            textBox13.Enabled = false;
            textBox12.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
            textBox10.Enabled = false;
            textBox11.Enabled = false;
            button5.Enabled = false;

            textBox14.Enabled = true;
            textBox15.Enabled = true;

            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Gor] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView2.DataSource = table;
            dataGridView2.Columns[0].Visible = false;
            con.Close();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox14.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            dataGridView2.CurrentCell = null;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView2.Columns.Count; c++)
                {
                    if (dataGridView2[c, i].Value.ToString() == textBox15.Text)
                    {
                        dataGridView2.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox3.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox3_Click_1(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void textBox4_Click_1(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void textBox13_Click(object sender, EventArgs e)
        {
            textBox13.Text = "";
        }

        private void textBox12_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";

        }

        private void textBox14_Click(object sender, EventArgs e)
        {
            textBox14.Text = "";

        }

        private void textBox15_Click(object sender, EventArgs e)
        {
            textBox15.Text = "";
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void оборудованиеНаСкладеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl2.Enabled = true;
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Oborud_na_sklad] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[3].Visible = false;

            con.Close();
        }

        private void продуктыНаСкладеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Produkt_na_sklad] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;


            con.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var oborud = textBox5.Text;
            var kolich = textBox6.Text;
            var marka = textBox7.Text;
            var produkt = textBox8.Text;
            var cenaoborud = textBox10.Text;
            var cenaprodukt = textBox11.Text;

            var wordApp = new word.Application();
            wordApp.Visible = false;

            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWord("{oborud}", oborud, wordDocument);
                ReplaceWord("{marka}", marka, wordDocument);
                ReplaceWord("{kolich}", kolich, wordDocument);
                ReplaceWord("{cenaoborud}", cenaoborud, wordDocument);
                ReplaceWord("{produkt}", produkt, wordDocument);
                ReplaceWord("{cenaprodukt}", cenaprodukt, wordDocument);

                wordDocument.SaveAs(@"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\result.docx");
                wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");
            }

        }

        private void ReplaceWord(string stubToReplace, string text, word.Document worddocument)
        {
            var range = worddocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox16.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView1.Columns.Count; c++)
                {
                    if (dataGridView1[c, i].Value.ToString() == textBox17.Text)
                    {
                        dataGridView1.Rows[i].Visible = true;
                        break;
                    }
                }

            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (textBox6.Text != "" && (textBox10.Text != "" && (textBox11.Text != "")))
            {
                label19.Text = (Convert.ToDecimal(textBox6.Text) * Convert.ToDecimal(textBox10.Text) + Convert.ToDecimal(textBox11.Text)).ToString();
            }
        }
    }
}
