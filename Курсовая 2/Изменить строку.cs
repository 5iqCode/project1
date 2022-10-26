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
    public partial class Изменить_строку : Form
    {
        private SqlConnection SqlConnection = null;

        private int id;

        public Изменить_строку(SqlConnection connection,int id)
        {
            InitializeComponent();

            SqlConnection = connection;

            this.id = id;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SqlCommand updateBDCommand = new SqlCommand("UPDATE [Продажи] SET [Дата_продажи]=@Дата_продажи, [ФИО]=@ФИО, [Категория]=@Категория, [Наименование]=@Наименование, [Цена_с_учётом_скидки]=@Цена_с_учётом_скидки, [Доставка]=@Доставка, [Общая_стоимость]=@Общая_стоимость WHERE [id]=@id", SqlConnection);
            try
            {
                updateBDCommand.Parameters.AddWithValue("Дата_продажи", Convert.ToDateTime(textBox1.Text));
                updateBDCommand.Parameters.AddWithValue("ФИО", textBox2.Text);
                updateBDCommand.Parameters.AddWithValue("Категория", textBox3.Text);
                updateBDCommand.Parameters.AddWithValue("Наименование", textBox6.Text);
                float price = float.Parse(textBox5.Text), price2 = float.Parse(textBox4.Text);
                updateBDCommand.Parameters.AddWithValue("Цена_с_учётом_скидки", price);
                updateBDCommand.Parameters.AddWithValue("Доставка", price2);
                updateBDCommand.Parameters.AddWithValue("Общая_стоимость", price + price2);
                updateBDCommand.Parameters.AddWithValue("id", id);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                await updateBDCommand.ExecuteNonQueryAsync();

                Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void Изменить_строку_Load(object sender, EventArgs e)
        {
            SqlCommand getBDInfoCommand = new SqlCommand("SELECT [Дата_продажи],[ФИО],[Категория],[Наименование],[Цена_с_учётом_скидки],[Доставка] FROM [Продажи] WHERE [id]=@id",SqlConnection);

            getBDInfoCommand.Parameters.AddWithValue("id", id);

            SqlDataReader sqlReader = null;

            try
            {
                sqlReader = await getBDInfoCommand.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = Convert.ToString(sqlReader["Дата_продажи"]);

                    textBox2.Text = Convert.ToString(sqlReader["ФИО"]);

                    textBox3.Text = Convert.ToString(sqlReader["Категория"]);

                    textBox6.Text = Convert.ToString(sqlReader["Наименование"]);

                    textBox5.Text = Convert.ToString(sqlReader["Цена_с_учётом_скидки"]);

                    textBox4.Text = Convert.ToString(sqlReader["Доставка"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null && !sqlReader.IsClosed)
                {
                    sqlReader.Close();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
