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
    public partial class Form2 : Form
    {
        public int selectrow = -1;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=center.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sql = "SELECT * FROM Виды_услуг";
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds,"Результат");
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
            string NazvanieVida = textBox1.Text.ToString();
            sql = "INSERT INTO Виды_услуг ( Название_вида ) " + " VALUES (" + "'" + NazvanieVida + "' " + ")";
            //ВЫПОЛНЕНИЕ ЗАПРОСА
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Виды_услуг";
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
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            //определить номер строки, на которую нажал пользователь
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox2.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox3.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            //определить номер строки, на которую нажал пользователь
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox2.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox3.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //РЕДАКТИРОВАНИЕ ЗАПИСИ В БД

            //ЕСЛИ ПОЛЬЗОВАТЕЛЬ НЕ ВЫДЕЛИЛ СТРОКУ ДЛЯ ИЗМЕНЕНИЯ
            //ИЛИ ВЫДЕЛИЛ В КОНЦЕ СЕТКИ ПУСТУЮ СТРОКУ,
            //ТО ОТМЕНА ДЕЙСТВИЯ "РЕКДАКТИРВОАНИЯ" И ВЫВОДП ПРЕДУПРЕЖДЕНИЯ

            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1){

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
            string KodVida = textBox2.Text.ToString();
            string NazvaniyeVida = textBox3.Text.ToString();

            //ФОРМИРОВАНИЕ ЗАПРОСА НА РЕДАКТИРОВАНИЕ
            sql = " UPDATE Виды_услуг SET " + " Название_вида = '" + NazvaniyeVida + "'" + " Where Код_Вида = "+ KodVida;

            //выполнение запроса
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Виды_услуг";
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
            textBox2.Text = "";
            textBox3.Text = "";
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
            string KodVida = textBox4.Text.ToString();
            //формирование запроса на удаление
            sql = "DELETE * FROM Виды_услуг WHERE Код_вида = "+KodVida;

            //выполнение запроса
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            //ОБНОВЛЕНИЕ ИНФОРМАЦИИ НА ФОРМЕ ПОСЛЕ ДОБАВЛЕНИЯ ДАННЫХ
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT * FROM Виды_услуг";
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
            textBox4.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //если пользователь не выбрал критерий поиска из выпадающего списка,
            //то вывод предупреждения
            if (comboBox1.SelectedIndex < 0)
            {
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
            string znachenie = textBox5.Text.ToString();
            sql = "SELECT * FROM Виды_услуг WHERE " + kriteriy + " Like " + " '" + znachenie + "' ";

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
            sql = "SELECT * FROM Виды_услуг";

            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
        }

        private void поискДаннызToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            textBox5.Text = "";
            groupBox4.Visible = true;
        }
    }
}
