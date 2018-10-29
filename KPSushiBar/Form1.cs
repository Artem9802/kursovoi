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
    public partial class Form1 : Form
    {
        private RegInfo _RI;
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _RI = new RegInfo();
            _RI.Register_set(comboBox1.Text, textBox2.Text, textBox3.Text, comboBox2.Text);
            Авторизация avtoriz = new Авторизация();
            avtoriz.Show();
            this.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            SqlDataSourceEnumerator sse = SqlDataSourceEnumerator.Instance;
            DataTable dt = sse.GetDataSources();
            foreach (DataRow r in dt.Rows)
            {
                comboBox1.Items.Add(r[0] + "\\" + r[1]); // выгрузка списка серверов
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ds = comboBox1.Text; //переменная для вывода имени сервера
            string log = textBox2.Text; //переменная в которой будет выводится логин базы
            string pas = textBox3.Text;//переменная в которой будет пароль логин базы

            SqlConnection connection = new SqlConnection("Data Source =" + ds + ";Initial Catalog = master; Persist Security Info = True; User ID = " + log +
            ";Password = \"" + pas + "\"");
            connection.Open();
            SqlCommand cmd = new SqlCommand("select name from sys.databases", connection);
            DataTable dt = new DataTable();
            SqlDataReader rd = cmd.ExecuteReader();
            dt.Load(rd);
            connection.Close();
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "name";
            comboBox2.Enabled = true;
            button2.Enabled = true;
            
          
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            if ((textBox2.Text == "") || (textBox3.Text == "") || (comboBox1.Text == ""))
            {
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox1.Text = "Скрыть пароль";
                textBox3.PasswordChar = (char)0;
            }
            else
            {
                checkBox1.Text = "Показать пароль";
                textBox3.PasswordChar = '*';
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox1.Text = "";
        }
    }
}
