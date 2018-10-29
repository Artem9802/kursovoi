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

namespace KPSushiBar
{
    public partial class Заказы_клиентов : Form
    {
        public Заказы_клиентов()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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

        private void Заказы_клиентов_Load(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("select * from [dbo].[Jurnal_zakaz] ", con);
            SqlDataReader dr = sc.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(dr);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[3].Visible = false;

            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
