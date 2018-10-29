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
    public partial class Авторизация : Form
    {
        private RegInfo _RI;
        public Авторизация()
        {
            InitializeComponent();
        }
        //public string id_sotr = "select [dbo].[Sotr].[id_sotr] from [dbo].[Sotr] inner join [dbo].[Prava_dostyp] on [dbo].[Prava_dostyp].[id_prava_dostyp]"

        private void button1_Click(object sender, EventArgs e)
        {
            
           
            SqlConnection con = new SqlConnection("Data Source = DESKTOP-AVKS5AE\\FAILBDS;Initial Catalog=kp v.3.2beta;Persist Security Info = True; User ID = sa; Password = 281199 "); //подключение к БД
            con.Open();
            SqlCommand sc = new SqlCommand("Select * from Sotr where[Log_sotr] = '" + textBox1.Text + "' and[Pass_sotr] = '" + textBox2.Text + "'", con); // выбока данных из таблицы сотрцдников и проверка логина и пароля при введени их в textbox
           
            SqlDataReader dr;
            dr = sc.ExecuteReader();
            int count = 0; // переменная для нахождения логина и пароля
            while (dr.Read())   
            {
                count += 1;
            }
            dr.Close();

            if (count == 1)
            {
               
                Главное_меню glavMenu = new Главное_меню();
                glavMenu.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Вас нет в системе");
            }
            
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if((textBox1.Text == "") || (textBox2.Text ==""))
            {
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           DialogResult result = MessageBox.Show("Вы уверены что хотите выйти", "Внимание!", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if(result == DialogResult.OK)
            {
                Application.Exit();
            }
            else
            {
               
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox1.Text = "Скрыть пароль";
                textBox2.PasswordChar = (char)0;
            }
            else
            {
                checkBox1.Text = "Показать пароль";
                textBox2.PasswordChar = '*';
            }
        }

        private void Авторизация_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
