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
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_2
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
        }
        public string pt;
        public int user;
        private void Form1_Load(object sender, EventArgs e)
        {
            Авторизация авторизация = new Авторизация();

            авторизация.ShowDialog();
            int user = авторизация.user;
            if (user == 0)
            {
                toolStripButton1.Visible = false;
                toolStripButton2.Visible = false;
                toolStripButton3.Visible = false;
            }
            listView1.GridLines = true;

            listView1.FullRowSelect = true;

            listView1.View = View.Details;

            listView1.Columns.Add("№ п/п", Width = 75);
            listView1.Columns.Add("Дата продажи", Width = 150);
            listView1.Columns.Add("ФИО", Width = 250);
            listView1.Columns.Add("Категория", Width = 250);
            listView1.Columns.Add("Наименование", Width = 250);
            listView1.Columns.Add("Цена с учётом скидки", Width = 200);
            listView1.Columns.Add("Доставка", Width = 150);
            listView1.Columns.Add("Общая стоимость", Width = 150);
            string pt;
            Укажите_путь_БД укажите_Путь_БД = new Укажите_путь_БД();
            if (укажите_Путь_БД.cls == 0)
            {
                if (sqlConnection == null || sqlConnection.State == ConnectionState.Closed)
                {
                    укажите_Путь_БД.ShowDialog();

                    pt = укажите_Путь_БД.textBox1.Text;

                    string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + pt;

                    sqlConnection = new SqlConnection(connectionString);

                    try
                    {
                        sqlConnection.Open();
                    }
                    catch
                    {
                        MessageBox.Show("Не удалось подключиться к базе данных повторите попытку или измените путь (Файл --> Изменить путь БД)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        укажите_Путь_БД.cls = 1;
                    }

                }
                if (sqlConnection.State == ConnectionState.Open)
                {
                    LoadBDAsync();
                }
            }
        }
        private async Task LoadBDAsync() //Select
        {
            SqlDataReader sqlReader = null;

            SqlCommand getProdashiCommand = new SqlCommand("SELECT * FROM [Продажи]", sqlConnection);


            try
            {
                sqlReader = await getProdashiCommand.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    ListViewItem item = new ListViewItem(new string[] {

                        Convert.ToString(sqlReader["id"]),
                        Convert.ToString(String.Format("{0:d}",sqlReader["Дата_продажи"])),
                        Convert.ToString(sqlReader["ФИО"]),
                        Convert.ToString(sqlReader["Категория"]),
                        Convert.ToString(sqlReader["Наименование"]),
                        Convert.ToString(String.Format("{0:0.00}",sqlReader["Цена_с_учётом_скидки"])),
                        Convert.ToString(String.Format("{0:0.00}",sqlReader["Доставка"])),
                        Convert.ToString(String.Format("{0:0.00}", sqlReader["Общая_стоимость"]))
                    });

                    listView1.Items.Add(item);

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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private async void toolStripButton4_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            await LoadBDAsync();
        }
        private async void toolStripButton1_Click(object sender, EventArgs e)
        {
            Добавить_строку insert = new Добавить_строку(sqlConnection);
            insert.Show();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Изменить_строку update = new Изменить_строку(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                update.Show();
            }
            else
            {
                MessageBox.Show("Выберите строку которую хотите изменить", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private async void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult res = MessageBox.Show("Вы действительно хотите удалить эту строку?", "Удаление строки", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

                switch (res)
                {
                    case DialogResult.OK:

                        SqlCommand deleteStrocCommand = new SqlCommand("DELETE FROM [Продажи] WHERE [id]=@id", sqlConnection);

                        deleteStrocCommand.Parameters.AddWithValue("id", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));

                        try
                        {
                            await deleteStrocCommand.ExecuteNonQueryAsync();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        listView1.Items.Clear();

                        await LoadBDAsync();

                        break;

                }
            }

            else
            {
                MessageBox.Show("Выберите строку которую хотите удалить", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void изменитьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pt;
            sqlConnection = null;
            Укажите_путь_БД укажите_Путь_БД = new Укажите_путь_БД();
            if (укажите_Путь_БД.cls == 0)
            {
                if (sqlConnection == null || sqlConnection.State == ConnectionState.Closed)
                {
                    укажите_Путь_БД.ShowDialog();

                    pt = укажите_Путь_БД.textBox1.Text;

                    string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + pt;

                    sqlConnection = new SqlConnection(connectionString);

                    try
                    {
                        sqlConnection.Open();
                    }
                    catch
                    {
                        MessageBox.Show("Не удалось подключиться к базе данных повторите попытку или измените путь (Файл --> Изменить путь БД)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        укажите_Путь_БД.cls = 1;
                    }

                }
                if (sqlConnection.State == ConnectionState.Open)
                {
                    listView1.Items.Clear();
                    LoadBDAsync();
                }
            }
        }
            private void Sort(SqlCommand getSortCommand)
        {
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = getSortCommand.ExecuteReader();

                while (sqlReader.Read())
                {
                    ListViewItem item = new ListViewItem(new string[] {

                        Convert.ToString(sqlReader["id"]),
                        Convert.ToString(String.Format("{0:d}",sqlReader["Дата_продажи"])),
                        Convert.ToString(sqlReader["ФИО"]),
                        Convert.ToString(sqlReader["Категория"]),
                        Convert.ToString(sqlReader["Наименование"]),
                        Convert.ToString(String.Format("{0:0.00}",sqlReader["Цена_с_учётом_скидки"])),
                        Convert.ToString(String.Format("{0:0.00}",sqlReader["Доставка"])),
                        Convert.ToString(String.Format("{0:0.00}", sqlReader["Общая_стоимость"]))
                    });

                    listView1.Items.Add(item);

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
        private void отАДоЯToolStripMenuItem_Click(object sender, EventArgs e)
        {

            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Категория ASC", sqlConnection);

            Sort(getSortCommand);

        }

        private void отЯДоАToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Категория DESC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отАДоЯToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Наименование ASC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отЯДоАToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Наименование DESC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отПозднихДоНовыхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Дата_продажи ASC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отНовыхДоПозднихToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Дата_продажи DESC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отАДоЯToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY ФИО ASC", sqlConnection);

            Sort(getSortCommand);
        }

        private void отЯДоАToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY ФИО DESC", sqlConnection);

            Sort(getSortCommand);
        }

        private void фИОToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Найти найти = new Найти();
            найти.label1.Text = "Введите ФИО покупателя";
            найти.label2.Text = "которое хотите найти";
            найти.ShowDialog();
            while (найти.cl == 1)
            {
                SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] WHERE ФИО LIKE @ФИО", sqlConnection);
                getSortCommand.Parameters.AddWithValue("ФИО", найти.textBox1.Text);
                listView1.Items.Clear();
                Sort(getSortCommand);
                найти.cl = 0;
                if (listView1.Items.Count == 0)
                {
                    MessageBox.Show("Не найдено ни одной строки, подходящей под заданные параметры", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadBDAsync();
                    найти.ShowDialog();
                }
            }
        }

        private void категорияСтройматериаловToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Найти найти = new Найти();
            найти.label1.Text = "Введите категорию товара";
            найти.label2.Text = "которую хотите найти";
            найти.ShowDialog();
            while (найти.cl == 1)
            {
                SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] WHERE Категория LIKE @Категория", sqlConnection);
                getSortCommand.Parameters.AddWithValue("Категория", найти.textBox1.Text);
                listView1.Items.Clear();
                Sort(getSortCommand);
                найти.cl = 0;
                if (listView1.Items.Count == 0)
                {
                    MessageBox.Show("Не найдено ни одной строки, подходящей под заданные параметры", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadBDAsync();
                    найти.ShowDialog();
                }
            }
        }

        private void наименованиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Найти найти = new Найти();
            найти.label1.Text = "Введите наименование товара";
            найти.label2.Text = "которое хотите найти";
            найти.ShowDialog();
            while (найти.cl == 1)
            {
                SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] WHERE Наименование LIKE @Наименование", sqlConnection);
                getSortCommand.Parameters.AddWithValue("Наименование", найти.textBox1.Text);
                listView1.Items.Clear();
                Sort(getSortCommand);
                найти.cl = 0;
                if (listView1.Items.Count == 0)
                {
                    MessageBox.Show("Не найдено ни одной строки, подходящей под заданные параметры", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadBDAsync();
                    найти.ShowDialog();
                }
            }
        }

        private void датаПродажиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Найти_по_дате найти = new Найти_по_дате();
            найти.ShowDialog();

            while (найти.cl == 1)
            {
                if (Convert.ToDateTime(найти.textBox2.Text) < Convert.ToDateTime(найти.textBox1.Text))
                {
                    try
                    {
                        SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] WHERE @Дата_продажи1 < Дата_продажи AND @Дата_продажи2 > Дата_продажи", sqlConnection);
                        getSortCommand.Parameters.AddWithValue("Дата_продажи1", Convert.ToDateTime(найти.textBox2.Text));
                        getSortCommand.Parameters.AddWithValue("Дата_продажи2", Convert.ToDateTime(найти.textBox1.Text));
                        listView1.Items.Clear();
                        Sort(getSortCommand);
                        найти.cl = 0;
                        if (listView1.Items.Count == 0)
                        {
                            MessageBox.Show("Не найдено ни одной строки, подходящей под заданные параметры", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadBDAsync();
                            найти.ShowDialog();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Введите существующую дату!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        найти.textBox1.Text = "12.12.2012";
                        найти.textBox2.Text = "11.11.2011 ";
                        найти.cl = 0;
                        найти.ShowDialog();
                    }
                }
                else {
                    MessageBox.Show("Начальная дата должна быть меньше конечной!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    найти.textBox1.Text = "12.12.2012";
                    найти.textBox2.Text = "11.11.2011 ";
                    найти.cl = 0;
                    найти.ShowDialog();
                }
            }

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
                    sw.WriteLine("Id      Дата продажи                    ФИО               Категория         Наименование         Цена          Доставка     Общая стоимость");
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

        private void покупателиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY ФИО ASC", sqlConnection);

            Sort(getSortCommand);

            Покупатели покупатели = new Покупатели();
            int i = 1, kolvo = 1;
            ListViewItem item = new ListViewItem(new string[] {


                        Convert.ToString(i),
                        Convert.ToString(listView1.Items[i-1].SubItems[2].Text),
                        Convert.ToString(kolvo),
                        Convert.ToString(String.Format("{0:0.00}", listView1.Items[i-1].SubItems[7].Text))
                    });

            покупатели.listView1.Items.Add(item);
            for (int j = 1; j < listView1.Items.Count; j++)
            {
                if (listView1.Items[j].SubItems[2].Text == покупатели.listView1.Items[i - 1].SubItems[1].Text)
                {
                    float c = float.Parse(покупатели.listView1.Items[i - 1].SubItems[3].Text);
                    float c1 = float.Parse(listView1.Items[j].SubItems[7].Text);
                    покупатели.listView1.Items[i - 1].SubItems[3].Text = Convert.ToString(c + c1);
                    kolvo++; c = 0; c1 = 0;
                    покупатели.listView1.Items[i - 1].SubItems[2].Text = Convert.ToString(kolvo);
                }
                else
                {
                    kolvo = 1; i += 1;
                    ListViewItem itemnew = new ListViewItem(new string[] {


                        Convert.ToString(i),
                        Convert.ToString(listView1.Items[j].SubItems[2].Text),
                        Convert.ToString(kolvo),
                        Convert.ToString(String.Format("{0:0.00}", listView1.Items[j].SubItems[7].Text))
                    });
                    покупатели.listView1.Items.Add(itemnew);
                }
            }
            покупатели.Show();
        }

        private void общаяСтоимостьПродажToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Общая_стоимость_продаж общая_Стоимость_Продаж = new Общая_стоимость_продаж();
            общая_Стоимость_Продаж.Show();


            int kolvo = 1, poc = 0, allpoc = 0;
            float stoim = 0, allstoim = 0;
            listView1.Items.Clear();

            SqlCommand getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Дата_продажи ASC", sqlConnection);

            Sort(getSortCommand);
            string dat = "";
            string cat = "";

            for (int j = 0; j < listView1.Items.Count; j++)
            {
                if (dat != Convert.ToString(Convert.ToDateTime(listView1.Items[j].SubItems[1].Text).ToString("MMMM yyyy")))
                {

                    dat = Convert.ToString(Convert.ToDateTime(listView1.Items[j].SubItems[1].Text).ToString("MMMM yyyy"));
                    ListViewItem datenew = new ListViewItem(new string[] {
                    dat,
                    });
                    общая_Стоимость_Продаж.listView1.Items.Add(datenew);



                    getSortCommand = new SqlCommand("SELECT * FROM [Продажи] WHERE MONTH(Дата_продажи) = @MOUNTH  and YEAR(Дата_продажи) = @YEAR ORDER BY Категория", sqlConnection);
                    getSortCommand.Parameters.AddWithValue("MOUNTH", Convert.ToDateTime(listView1.Items[j].SubItems[1].Text).ToString("MM"));
                    getSortCommand.Parameters.AddWithValue("YEAR", Convert.ToDateTime(listView1.Items[j].SubItems[1].Text).ToString("yyyy"));
                    listView1.Items.Clear();
                    Sort(getSortCommand);
                    for (int j1 = 0; j1 < listView1.Items.Count; j1++)
                    {
                        if (cat != Convert.ToString(listView1.Items[j1].SubItems[3].Text))
                        {
                            ListViewItem kategnew = new ListViewItem(new string[] {
                        Convert.ToString(""),
                        Convert.ToString(listView1.Items[j1].SubItems[3].Text),
                        });
                            cat = Convert.ToString(listView1.Items[j1].SubItems[3].Text);
                            общая_Стоимость_Продаж.listView1.Items.Add(kategnew);
                            for (int i = 0; i < listView1.Items.Count; i++)
                            {
                                if ((dat == Convert.ToString(Convert.ToDateTime(listView1.Items[i].SubItems[1].Text).ToString("MMMM yyyy"))) && (cat == Convert.ToString(listView1.Items[i].SubItems[3].Text)))
                                {
                                    ListViewItem itemnew = new ListViewItem(new string[] {
                        Convert.ToString(""),
                        Convert.ToString(""),
                        Convert.ToString(listView1.Items[i].SubItems[0].Text),
                        Convert.ToString(listView1.Items[i].SubItems[4].Text),
                        Convert.ToString(kolvo),
                        Convert.ToString(String.Format("{0:0.00}", listView1.Items[i].SubItems[7].Text))
                    });
                                    общая_Стоимость_Продаж.listView1.Items.Add(itemnew);
                                    poc++;
                                    stoim += float.Parse(listView1.Items[i].SubItems[7].Text);
                                }
                            }
                        }
                    }
                    cat = " ";
                    listView1.Items.Clear();

                    getSortCommand = new SqlCommand("SELECT * FROM [Продажи] ORDER BY Дата_продажи ASC", sqlConnection);

                    Sort(getSortCommand);

                    ListViewItem itog = new ListViewItem(new string[] {
                        Convert.ToString(""),
                        Convert.ToString(""),
                        Convert.ToString(""),
                        Convert.ToString("Итого"),
                        Convert.ToString(poc),
                        Convert.ToString(String.Format("{0:0.00}",stoim))
                    });

                    allstoim += stoim; allpoc += poc;
                    stoim = 0; poc = 0;

                    общая_Стоимость_Продаж.listView1.Items.Add(itog);
                    int color = общая_Стоимость_Продаж.listView1.Items.Count;
                    общая_Стоимость_Продаж.listView1.Items[color - 1].BackColor = Color.LightYellow;
                }

            }

            ListViewItem allitog = new ListViewItem(new string[] {
                        Convert.ToString(""),
                        Convert.ToString(""),
                        Convert.ToString(""),
                        Convert.ToString("Общий итого"),
                        Convert.ToString(allpoc),
                        Convert.ToString(String.Format("{0:0.00}",allstoim))
                    });
            общая_Стоимость_Продаж.listView1.Items.Add(allitog);
            int color2 = общая_Стоимость_Продаж.listView1.Items.Count;
            общая_Стоимость_Продаж.listView1.Items[color2 - 1].BackColor = Color.LightSkyBlue;

        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}





