using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    internal class DbManeger
    {
        // Строка подключения к базе данных
        static string connect = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=TRPObd;Integrated Security=True;Connect Timeout=30;Encrypt=False;";

        SqlConnection sqlCon = new SqlConnection(connect); // Создаем соединение с базой данных

        // Метод для открытия или закрытия соединения с базой данных
        void Connect()
        {
            // Если соединение открыто, закрываем его
            if (sqlCon.State == System.Data.ConnectionState.Open)
            {
                sqlCon.Close();
            }
            else
            {
                // Иначе открываем соединение
                sqlCon.Open();
            }
        }

        // Метод для выполнения SQL-запроса, который возвращает одно скалярное значение
        public static object SelectScalar(string query)
        {
            using (SqlConnection connection = new SqlConnection(connect))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open(); // Открываем соединение
                return command.ExecuteScalar(); // Возвращает первое значение первой строки
            }
        }

        // Метод для выполнения SQL-запроса и заполнения DataGridView данными
        public void Select(string query, DataGridView dataGridView)
        {
            Connect(); // Устанавливаем соединение с базой данных
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(query, sqlCon); // Создаем адаптер для выполнения запроса
            DataTable dataTable = new DataTable(); // Создаем DataTable для хранения результатов
            sqlDataAdapter.Fill(dataTable); // Заполняем DataTable данными из базы
            Connect(); // Устанавливаем соединение с базой данных
            dataGridView.DataSource = dataTable; // Устанавливаем источник данных для DataGridView
        }

        // Метод для выполнения SQL-запроса и заполнения ComboBox данными
        public void SelectComb(string query, string dispMember, string valueMemb, ComboBox combox)
        {
            Connect(); // Устанавливаем соединение с базой данных
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(query, sqlCon); // Создаем адаптер для выполнения запроса
            DataTable dataTable = new DataTable(); // Создаем DataTable для хранения результатов
            sqlDataAdapter.Fill(dataTable); // Заполняем DataTable данными из базы
            Connect(); // Устанавливаем соединение с базой данных
            combox.DataSource = dataTable; // Устанавливаем источник данных для ComboBox
            combox.DisplayMember = dispMember; // Устанавливаем поле, отображаемое в ComboBox
            combox.ValueMember = valueMemb; // Устанавливаем поле, представляющее значение в ComboBox
        }

        // Метод для выполнения SQL-запросов, которые не возвращают данные (INSERT, UPDATE, DELETE)
        public void Action(string query)
        {
            Connect(); // Устанавливаем соединение с базой данных
            SqlCommand sqlCommand = new SqlCommand(query, sqlCon); // Создаем команду с SQL-запросом
            sqlCommand.ExecuteNonQuery(); // Выполняем команду
            Connect(); // Устанавливаем соединение с базой данных
        }
    }
}