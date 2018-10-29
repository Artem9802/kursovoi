using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSushiBar
{
    public partial class Главное_меню : Form
    {
        public Главное_меню()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Сотрудники sotr = new Сотрудники();
            sotr.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Машины masin = new Машины();
            masin.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Заказ_оборудования zakazOborud = new Заказ_оборудования();
            zakazOborud.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Заказы_клиентов zakazKlient = new Заказы_клиентов();
            zakazKlient.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
   
            Авторизация avtoriz = new Авторизация();
            avtoriz.Show();
            this.Close();
        }

        private void Главное_меню_Load(object sender, EventArgs e)
        {
      
        }
    }
}
