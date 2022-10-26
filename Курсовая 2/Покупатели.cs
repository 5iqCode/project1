using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Курсовая_2
{
    public partial class Покупатели : Form
    {
        public Покупатели()
        {
            InitializeComponent();
        }

        private void Покупатели_Load(object sender, EventArgs e)
        {
            listView1.GridLines = true;

            listView1.FullRowSelect = true;

            listView1.View = View.Details;

            listView1.Columns.Add("№ п/п", Width = 75);
            listView1.Columns.Add("ФИО", Width = 250);
            listView1.Columns.Add("Количество покупок", Width = 200);
            listView1.Columns.Add("Общая сумма", Width = 150);

            LoadBD();
        }

        private void LoadBD()
        {
            

        }

        private void сохранитьdocxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile1 = new SaveFileDialog();
            saveFile1.DefaultExt = "*.txt";
            saveFile1.Filter = "Text files|*.txt";
            if (saveFile1.ShowDialog() == System.Windows.Forms.DialogResult.OK &&
                saveFile1.FileName.Length > 0)
            {
                using (StreamWriter sw = new StreamWriter(saveFile1.FileName, true, Encoding.UTF8))
                {
                    sw.WriteLine("Id                ФИО             Кол-во покупок  Общая сумма");
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        for (int j = 0; j < listView1.Columns.Count; j++)
                        {
                            sw.Write(listView1.Items[i].SubItems[j].Text);
                            sw.Write("        ");
                        }
                        sw.WriteLine();
                    }
                    sw.Close();
                }
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

