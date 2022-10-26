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
    public partial class Найти_по_дате : Form
    {
        public Найти_по_дате()
        {
            InitializeComponent();
        }
        public int cl=0;
        private void button1_Click(object sender, EventArgs e)
        {
            cl = 1;
            Close();
        }
    }
}
