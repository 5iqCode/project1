using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Курсовая_2
{
    public partial class Авторизация : Form
    {
        public int user = 0;
        public Авторизация()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "1111")
            {
                user = 1;
                Form1 form1 = new Form1();
                Close();
            }
            else
            {
                label2.Visible = true;
                Task.Delay(1000).GetAwaiter().GetResult();
                label2.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Form1 form1 = new Form1();
            Close();
        }
    }
}
