using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Курсовая_2
{
    public partial class Укажите_путь_БД : Form
    {
        public Укажите_путь_БД()
        {
            InitializeComponent();
        }
        public int cls = 0;
        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == ""|| textBox1.Text ==" ")
            {
                MessageBox.Show("Введите путь к базе данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("CREATE TABLE [dbo].[Продажи] (" + "\r\n" +
 "   [Id]                   INT           IDENTITY(1, 1) NOT NULL, " + "\r\n" +
 "   [Дата_продажи]         DATETIME      NULL, " + "\r\n" +
 "   [ФИО]                  NVARCHAR(50) NULL, " + "\r\n" +
 "   [Категория]            NVARCHAR(50) NULL, " + "\r\n" +
 "   [Наименование]         NVARCHAR(50) NULL, " + "\r\n" +
 "   [Цена_с_учётом_скидки] FLOAT(53)    NULL, " + "\r\n" +
 "   [Доставка]             FLOAT(53)    NULL, " + "\r\n" +
 "   [Общая_стоимость]      FLOAT(53)    NULL, " + "\r\n" +
 "   PRIMARY KEY CLUSTERED([Id] ASC)" + "\r\n" +
"); ");
            label6.Visible = true;
            Task.Delay(1000).GetAwaiter().GetResult();
            label6.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            cls = 1;
            Close();
        }
    }
}
