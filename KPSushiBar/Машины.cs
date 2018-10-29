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
    public partial class Машины : Form
    {
        private readonly string TemplateFileName = @"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\Машины.docx";
        public Машины()
        {
            InitializeComponent();
        }
        public void zagr()
        {
            SqlConnection connection = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.1beta;Persist Security Info = True; User ID = sa; Password = 281199");
            connection.Open();
            SqlCommand upd = new SqlCommand("Spisok_mashin_edit", connection);
            upd.CommandType = CommandType.StoredProcedure;
            upd.Parameters.AddWithValue("@id_spisok_mashin", textBox5.Text);
            upd.Parameters.AddWithValue("@Marka", textBox3.Text);
            upd.Parameters.AddWithValue("@Kolich_mashin", textBox4.Text);
            upd.Parameters.AddWithValue("@id_sotr", textBox6.Text);

            upd.ExecuteNonQuery();
            connection.Close();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Главное_меню glavMenu = new Главное_меню();
            glavMenu.Show();
            this.Close();
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

        private void Машины_Load(object sender, EventArgs e)
        {
            textBox5.Visible = true;

            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Spisok_mashin] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            //dataGridView1.Columns[0].Visible = false;
            // dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[0].Visible = false;
             dataGridView1.Columns[3].Visible = false;
            con.Close();

            textBox5.Visible = false;
            textBox6.Visible = false;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
                SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
                con.Open();
                SqlCommand StrPrc = new SqlCommand("Spisok_mashin_add", con);
                StrPrc.CommandType = CommandType.StoredProcedure;
               //StrPrc.Parameters.AddWithValue("@id_sotr", comboBox1.Text);
                StrPrc.Parameters.AddWithValue("@Marka", textBox3.Text);
                StrPrc.Parameters.AddWithValue("@Kolich_mashin", textBox4.Text);
                StrPrc.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Добавление прошло успешно !");
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Spisok_mashin] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            zagr();
            MessageBox.Show("Изменение произошло успешно");
            zagr();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Spisok_mashin] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            con.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand mashdel = new SqlCommand("Spisok_mashin_delete", con);
            mashdel.CommandType = CommandType.StoredProcedure;
            mashdel.Parameters.AddWithValue("@id_Spisok_mashin", Convert.ToInt32(textBox5.Text));
            mashdel .ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Удаление произошло успешно!");
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
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
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Marka = textBox3.Text;
            var Kolvo = textBox4.Text;

            var wordApp = new word.Application();
            wordApp.Visible = false;
            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWord("{Marka}", Marka, wordDocument);
                ReplaceWord("{Kolvo}", Kolvo, wordDocument);
               

                wordDocument.SaveAs(@"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\Машины.docx");
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
    }
}
