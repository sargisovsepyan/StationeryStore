using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace PR_5_6
{
    public partial class Form3 : Form
    {
        public int selectrow = -1;
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sql = "SELECT * FROM Услуги";
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ДОБАВЛЕНИЕ ЗАПИСИ В БД
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;

            string KodVida = textBox1.Text.ToString();
            string Format = textBox2.Text.ToString();
            string Kolichestvo = textBox3.Text.ToString();
            string EdIzmereniya = textBox4.Text.ToString();
            string Stoimost = textBox5.Text.ToString();

            sql = "INSERT INTO Услуги" + "(Код_вида, Формат, Количество, Единицы_измерения, Стоимость) " + " VALUES ( " + KodVida + ", " + "'" + Format + "', " + Kolichestvo + ", " + "'" + EdIzmereniya + "', " + Stoimost + " )";
            //ВЫПОЛНЕНИЯ ЗАПРОСА
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Услуги";
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            groupBox1.Visible = false;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";

            selectrow = dataGridView1.CurrentCell.RowIndex;

            if (selectrow < (dataGridView1.RowCount - 1)){
                textBox6.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox7.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox8.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox9.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[4, selectrow].Value.ToString();
                textBox11.Text = dataGridView1[5, selectrow].Value.ToString();
                textBox12.Text = dataGridView1[0, selectrow].Value.ToString();
            }



        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            selectrow = dataGridView1.CurrentCell.RowIndex;

            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox6.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox7.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox8.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox9.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[4, selectrow].Value.ToString();
                textBox11.Text = dataGridView1[5, selectrow].Value.ToString();
                textBox12.Text = dataGridView1[0, selectrow].Value.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //РЕДАКТИРОВАНИЕ ЗАПИСИ В БД

            //ЕСЛИ ПОЛЬЗОВАТЕЛЬ НЕ ВЫДЕЛИЛ СТРОКУ ДЛЯ ИЗМЕНЕНИЯ
            //ИЛИ ВЫДЕЛИЛ В КОНЦЕ СЕТКИ ПУСТУЮ СТРОКУ,
            //ТО ОТМЕНА ДЕЙСТВИЯ "РЕКДАКТИРВОАНИЯ" И ВЫВОДП ПРЕДУПРЕЖДЕНИЯ

            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1)
            {

                MessageBox.Show("Выделите в сетке строку для редактирования");
                return;
            }

            //ИНАЧЕ ВЫПОЛНЕНИЕ ДАЛЬНЕЙШЕГО ХОДА РЕДАКТИРОВАНИЯ
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;
            //считывание данных с полей ввода
            string KodUslugi = textBox6.Text.ToString();
            string KodVida = textBox7.Text.ToString();
            string Format = textBox8.Text.ToString();
            string Kolichestvo = textBox9.Text.ToString();
            string EdIzmereniya  = textBox10.Text.ToString();
            string Stoimost = textBox11.Text.ToString();


            //ФОРМИРОВАНИЕ ЗАПРОСА НА РЕДАКТИРОВАНИЕ
            sql = " UPDATE Услуги SET " + " Код_вида = " + KodVida + ", " + " Формат = '" + Format + "'" + "," + " Количество = " + Kolichestvo + ", " + " Единицы_измерения = '" + EdIzmereniya + "'" + " WHERE Код_Услуги = " + KodUslugi;  

            //выполнение запроса
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Услуги";
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox2.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";

            groupBox1.Visible = false;
            groupBox2.Visible = true;
            groupBox3.Visible = false;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //УДАЛЕНИЕ ЗАПИСИ В БД

            //ЕСЛИ ПОЛЬЗОВАТЕЛЬ НЕ ВЫДЕЛИЛ СТРОКУ ДЛЯ ИЗМЕНЕНИЯ
            //ИЛИ ВЫДЕЛИЛ В КОНЦЕ СЕТКИ ПУСТУЮ СТРОКУ,
            //ТО ОТМЕНА ДЕЙСТВИЯ "УДАЛЕНИЕ" И ВЫВОДП ПРЕДУПРЕЖДЕНИЯ

            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1)
            {

                MessageBox.Show("Выделите в сетке строку для удаления");
                return;
            }

            //ИНАЧЕ ВЫПОЛНЕНИЕ ДАЛЬНЕЙШЕГО ХОДА УДАЛЕНИЯ
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;
            string KodUslugi = textBox12.Text.ToString();
            //формирование запроса на удаление
            sql = "DELETE * FROM Услуги WHERE Код_услуги = " + KodUslugi;

            //выполнение запроса
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Услуги";
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox3.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //если пользователь не выбрал критерий поиска из выпадающего списка,
            //то вывод предупреждения
            if (comboBox1.SelectedIndex < 0){
                MessageBox.Show("Выбирите критерий поиска");
                return;
            }

            //ИНАЧЕ ПОИСК ДАННЫХ В БД
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;
            //считывание критерия поиска из выпадающего списка
            string kriteriy = comboBox1.SelectedItem.ToString();
            //считывание данных с поля ввода 
            string znachenie = textBox13.Text.ToString();
            sql = "SELECT * FROM Услуги WHERE " + kriteriy + " Like " + " '" + znachenie + "' ";

            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;
            sql = "SELECT * FROM Услуги";

            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
        }

        private void поискДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            textBox13.Text = "";
            groupBox4.Visible = true;

        }
    }
}
