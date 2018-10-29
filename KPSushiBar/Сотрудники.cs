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
    public partial class Сотрудники : Form
    {
        private readonly string TemplateFileName = @"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\Информация о сотруднике.docx";
        public Сотрудники()
        {
            InitializeComponent();
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

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox3.Text =="" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "" || textBox10.Text == "" || textBox11.Text == "")
            {
                MessageBox.Show("Не все поля заполнены!", "Внимние!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
                con.Open();
                SqlCommand stadd = new SqlCommand("Sotr_add", con);
                stadd.CommandType = CommandType.StoredProcedure;
                stadd.Parameters.AddWithValue("@F_sotr", textBox3.Text);
                stadd.Parameters.AddWithValue("@N_sotr", textBox4.Text);
                stadd.Parameters.AddWithValue("@O_sotr", textBox5.Text);
                stadd.Parameters.AddWithValue("@Data_rojd", textBox6.Text);
                stadd.Parameters.AddWithValue("@Numb_sotr", textBox7.Text);
                stadd.Parameters.AddWithValue("@Staj_rab", textBox8.Text);
                stadd.Parameters.AddWithValue("@Dolj", textBox8.Text);
                stadd.Parameters.AddWithValue("@Log_sotr", textBox9.Text);
                stadd.Parameters.AddWithValue("@Pass_sotr", textBox10.Text);
                stadd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Сотрудник успешно добавлен");
                
                    
            }
        }

        public void get_oboryd()
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД 
            con.Open();
            SqlCommand get_ob_name = new SqlCommand("select id_oborud as \"ido\",Naim_oborud as \"nameo\" from Oborud", con);
            SqlDataReader dr = get_ob_name.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "nameo";
            comboBox1.ValueMember = "ido";
            con.Close();
        }

        private void Сотрудники_Load(object sender, EventArgs e)
        {
            get_oboryd();
            // TODO: данная строка кода позволяет загрузить данные в таблицу "_kp_v_3_1betaDataSet.Sotr". При необходимости она может быть перемещена или удалена.
            comboBox1.Visible = true;
            this.sotrTableAdapter.Fill(this._kp_v_3_1betaDataSet.Sotr);
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Sotr] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[10].Visible = false;

            con.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox7.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            textBox9.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox10.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            textBox11.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            textBox12.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox13.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();


        }

      

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand edit = new SqlCommand("Sotr_edit " , con);
            edit.CommandType = CommandType.StoredProcedure;
            edit.Parameters.AddWithValue("@id_oborud", comboBox1.SelectedValue);
           edit.Parameters.AddWithValue("@id_sotr", textBox12.Text);
            edit.Parameters.AddWithValue("@F_sotr",textBox3.Text);
            edit.Parameters.AddWithValue("@N_sotr", textBox4.Text);
            edit.Parameters.AddWithValue("@O_sotr", textBox5.Text);
            edit.Parameters.AddWithValue("@Data_rojd", textBox6.Text);
            edit.Parameters.AddWithValue("@Numb_sotr", textBox7.Text);
            edit.Parameters.AddWithValue("@Staj_rab", textBox8.Text);
            edit.Parameters.AddWithValue("@Dolj", textBox9.Text);
            edit.Parameters.AddWithValue("@Log_sotr", textBox10.Text);
            edit.Parameters.AddWithValue("@Pass_sotr", textBox11.Text);
            edit.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Изменение произошло успешно!");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand dlstr = new SqlCommand("Sotr_delete",con);
            dlstr.CommandType = CommandType.StoredProcedure;
            dlstr.Parameters.AddWithValue("@id_sotr", Convert.ToInt32(textBox12.Text));
            dlstr.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Удаление произошло успешно!");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var Fam = textBox3.Text;
            var Im = textBox4.Text;
            var Otch = textBox5.Text;
            var Datarojd = textBox6.Text;
            var NumbTel = textBox7.Text;
            var StajRab = textBox8.Text;
            var Dolj = textBox9.Text;

            var wordApp = new word.Application();
            wordApp.Visible = false;

            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWord("{Fam}", Fam, wordDocument);
                ReplaceWord("{Im}", Im, wordDocument);
                ReplaceWord("{Otch}", Otch, wordDocument);
                ReplaceWord("{Datarojd}", Datarojd, wordDocument);
                ReplaceWord("{NumbTel}", NumbTel, wordDocument);
                ReplaceWord("{StajRab}", StajRab, wordDocument);
                ReplaceWord("{ Dolj}", Dolj, wordDocument);

                wordDocument.SaveAs(@"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\Информация о сотруднике.docx");
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

        private void button6_Click(object sender, EventArgs e)
        {
            var Fam = textBox3.Text;
            var Im = textBox4.Text;
            var Otch = textBox5.Text;
            var Datarojd = textBox6.Text;
            var NumbTel = textBox7.Text;
            var StajRab = textBox8.Text;
            var Dolj = textBox9.Text;

            var wordApp = new word.Application();
            wordApp.Visible = false;

            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);
                ReplaceWord("{Fam}", Fam, wordDocument);
                ReplaceWord("{Im}", Im, wordDocument);
                ReplaceWord("{Otch}", Otch, wordDocument);
                ReplaceWord("{Datarojd}", Datarojd, wordDocument);
                ReplaceWord("{NumbTel}", NumbTel, wordDocument);
                ReplaceWord("{StajRab}", StajRab, wordDocument);
                ReplaceWord("{Dolj}", Dolj, wordDocument);

                wordDocument.SaveAs(@"C:\Users\KostF\OneDrive\Рабочий стол\МПТ\РВиАПООН\КП\Информация о сотруднике.docx");
                wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

