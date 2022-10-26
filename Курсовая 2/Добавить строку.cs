using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Курсовая_2
{
    public partial class Добавить_строку : Form
    {
        private SqlConnection SqlConnection = null;
        public Добавить_строку(SqlConnection connection)
        {
            InitializeComponent();

            SqlConnection = connection;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
            SqlCommand insertBDCommand = new SqlCommand("INSERT INTO [Продажи] (Дата_продажи, ФИО, Категория, Наименование, Цена_с_учётом_скидки, Доставка, Общая_стоимость) VALUES (@Дата_продажи, @ФИО, @Категория, @Наименование, @Цена_с_учётом_скидки, @Доставка, @Общая_стоимость) ", SqlConnection);
                insertBDCommand.Parameters.AddWithValue("Дата_продажи", Convert.ToDateTime(textBox1.Text));
                insertBDCommand.Parameters.AddWithValue("ФИО", textBox2.Text);
                insertBDCommand.Parameters.AddWithValue("Категория", textBox3.Text);
                insertBDCommand.Parameters.AddWithValue("Наименование", textBox6.Text);
                float price = float.Parse(textBox5.Text), price2 = float.Parse(textBox4.Text);
                insertBDCommand.Parameters.AddWithValue("Цена_с_учётом_скидки", price);
                insertBDCommand.Parameters.AddWithValue("Доставка", price2);
                insertBDCommand.Parameters.AddWithValue("Общая_стоимость", price+price2);

 
                await insertBDCommand.ExecuteNonQueryAsync();

                Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
