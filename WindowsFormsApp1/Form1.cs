using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;



namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.MaxLength = 4;
            textBox1.KeyPress += textBox1_KeyPress;
            textBox2.KeyPress += textBox2_KeyPress;
            textBox4.KeyPress += textBox4_KeyPress;
            textBox6.KeyPress += textBox6_KeyPress;
            textBox7.KeyPress += textBox7_KeyPress;
            textBox8.KeyPress += textBox8_KeyPress;
            textBox9.KeyPress += textBox9_KeyPress;
            textBox10.KeyPress += textBox10_KeyPress;
            textBox11.KeyPress += textBox11_KeyPress;

            comboBox1.KeyPress += comboBox1_KeyPress;
            comboBox2.KeyPress += comboBox2_KeyPress;
            comboBox3.KeyPress += comboBox3_KeyPress;
            comboBox4.KeyPress += comboBox4_KeyPress;
            comboBox5.KeyPress += comboBox5_KeyPress;
            comboBox6.KeyPress += comboBox6_KeyPress;
            comboBox7.KeyPress += comboBox7_KeyPress;
            comboBox9.KeyPress += comboBox9_KeyPress;

            //запрет на ввод вручную 
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
            dataGridView3.ReadOnly = true;
            dataGridView4.ReadOnly = true;

            // Установка размера формы
            this.Width = 1070;   // Ширина формы
            this.Height = 600;  // Высота формы

            // Создаем экземпляр ToolTip
            System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();

            // Устанавливаем параметры для ToolTip (необязательно)
            toolTip.AutoPopDelay = 5000; // Время отображения (в миллисекундах)
            toolTip.InitialDelay = 10; // Время задержки перед появлением
            toolTip.ReshowDelay = 1000; // Время, через какое появится снова
            toolTip.ShowAlways = true; // Всегда показывать подсказки

            toolTip.SetToolTip(textBox2, "Введите ФИО студента.");
            toolTip.SetToolTip(dateTimePicker1, "Выберите дату рождения студента. Не младше 15 и не старше 21.");
            toolTip.SetToolTip(textBox3, "Введите номер студента.");
            toolTip.SetToolTip(textBox4, "Введите адрес студента.");
            toolTip.SetToolTip(dateTimePicker5, "Фильтрация осуществляется от выбранной даты до текущей.");

            toolTip.SetToolTip(comboBox1, "Выберите ФИО студента.");
            toolTip.SetToolTip(textBox1, "Введите средний балл студента (от 1 до 10).");
            toolTip.SetToolTip(dateTimePicker2, "Выберите дату составления ведомости. Дата не может быть в будущем.");
            toolTip.SetToolTip(textBox10, "Фильтрация осуществляется для введенного диапазона.");
            toolTip.SetToolTip(comboBox6, "Отобразятся только ведомости выбранного студента.");

            toolTip.SetToolTip(comboBox2, "Выберите ФИО студента.");
            toolTip.SetToolTip(comboBox3, "Выберите причину доплаты.");
            toolTip.SetToolTip(dateTimePicker2, "Выберите дату составления приказа. Дата не может быть в прошлом.");
            toolTip.SetToolTip(comboBox9, "Отобразятся только приказы с выбранной причиной.");


            toolTip.SetToolTip(comboBox4, "Выберите ФИО студента.");
            toolTip.SetToolTip(comboBox5, "Выберите тип стипендии.");
            toolTip.SetToolTip(textBox14, "Введите сумму доплаты по приказу от 10 до 99. Максимум 5 символов.");
            toolTip.SetToolTip(dateTimePicker4, "Выберите дату выдачи стипендии. Дата не может быть в будущем или прошлом.");
            toolTip.SetToolTip(dateTimePicker6, "Фильтрация осуществляется от выбранной даты до текущей.");
            toolTip.SetToolTip(comboBox9, "Отобразятся только стипендии с выбранным типом.");

            toolTip.SetToolTip(textBox6, "Введите значение которое хотите найти.");
            toolTip.SetToolTip(textBox7, "Введите значение которое хотите найти.");
            toolTip.SetToolTip(textBox8, "Введите значение которое хотите найти.");
            toolTip.SetToolTip(textBox9, "Введите значение которое хотите найти.");
        }

        DbManeger DbManeger = new DbManeger();

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void Form1_Load(object sender, EventArgs e)
        {
            DbManeger.Select("Select id_student, fio as 'Студент', " +
                              "birthday as 'Дата рождения', " +
                              "numder as 'Номер телефона', " +
                              "adress as 'Адрес' from student", dataGridView1);


            DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                                "gpa AS 'Средний балл', " +
                                "cof AS 'Коэффициент', " +
                                "date_cr AS 'Дата составления' FROM grade INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

            DbManeger.Select("Select id_orders, fio as 'Студент', " +
                             "cause as 'Причина', " +
                             "date_cr as 'Дата составления' from orders inner join student on  orders.id_student=student.id_student", dataGridView3);

            DbManeger.Select("Select id_schoolarship, fio as 'Студент'," +
                             "type_s as 'Тип',    " +
                             "schoolarship.date_cr as 'Дата выдачи' ," +
                             "cause as 'Причина надбавки', " +
                             "summ as 'Сумма по приказу',  " +
                             "gpa as 'Средний балл', " +
                             "cof as 'Коэффициент', " +
                             "cof*118+summ as 'Сумма стипендии'" +
                             "from((schoolarship inner join student on  schoolarship.id_student = student.id_student) " +
                             "inner join orders on orders.id_student = student.id_student) inner join grade on grade.id_student = student.id_student", dataGridView4);

            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox1);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox2);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox4);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox6);
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //обновление комбобоксов
        private void UpdateComboBoxes()
        {
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox1);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox2);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox4);
            DbManeger.SelectComb("SELECT id_student, fio FROM student", "fio", "id_student", comboBox6);
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //ввод цифр в textBox СРЕДНИЙ БАЛЛ
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && textBox1.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        //ввод букв в textBox ФИО
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //цифры и + НОМЕР ТЕЛЕФОНА
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '+' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //буквы и точка АДРЕС
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Разрешаем ввод букв, цифр, точки, пробела и клавиши Backspace
            if (!char.IsLetterOrDigit(e.KeyChar) &&
                e.KeyChar != ' ' &&
                e.KeyChar != ',' &&
                e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Отменяем ввод
            }
        }

        //ввод цифр в textBox ФИЛЬТРАЦИЯ ОТ СУММА
        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && textBox1.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        //ввод цифр в textBox ФИЛЬТРАЦИЯ ДО СУММА
        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && textBox1.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        //ввод цифр в textBox ФИЛЬТРАЦИЯ ОТ СРЕДНИЙ БАЛЛ
        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && textBox1.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        //ввод цифр в textBox ФИЛЬТРАЦИЯ ДО СРЕДНИЙ БАЛЛ
        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && textBox11.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //ввоод только букв в comboBox
        private void comboBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //поиск первое
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                e.KeyChar != '.' &&
                e.KeyChar != ',' &&
                e.KeyChar != ' ' &&
                e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //поиск второе
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                e.KeyChar != '.' &&
                e.KeyChar != ',' &&
                e.KeyChar != ' ' &&
                e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //поиск третье
        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                e.KeyChar != '.' &&
                e.KeyChar != ',' &&
                e.KeyChar != ' ' &&
                e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //поиск четвертое
        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) &&
                !char.IsDigit(e.KeyChar) &&
                e.KeyChar != '.' &&
                e.KeyChar != ',' &&
                e.KeyChar != ' ' &&
                e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //добавление записи студента
        private void button1_Click(object sender, EventArgs e)
        {
            // Проверяем, что все текстовые поля заполнены
            if (string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox3.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем формат номера телефона
            string input = textBox3.Text;
            if (!Regex.IsMatch(input, @"^\+375(29|25|44)\d{7}$"))
            {
                MessageBox.Show("Некорректный код номера телефона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем дату из dateTimePicker
            DateTime birthday = dateTimePicker1.Value.Date;
            DateTime currentDate = DateTime.Now;

            // Проверяем, что дата рождения не больше 21 года назад и не меньше 15 лет назад
            if (birthday < currentDate.AddYears(-21) || birthday > currentDate.AddYears(-15))
            {
                MessageBox.Show("Студент не может быть младше 15 или старше 21.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем, существует ли студент с уже существующими данными
            string checkQuery = $"SELECT COUNT(*) FROM student WHERE fio = N'{textBox2.Text}' AND birthday = '{birthday.ToString("yyyy-MM-dd")}' " +
                                $"AND numder = N'{textBox3.Text}' AND adress = N'{textBox4.Text}'";
            int count = (int)DbManeger.SelectScalar(checkQuery); // Предполагается, что SelectScalar возвращает одно значение.

            // Если студент уже существует
            if (count > 0)
            {
                MessageBox.Show("Студент с такими данными уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Если все проверки пройдены, выполняем действие
            DbManeger.Action($"INSERT INTO student (fio, birthday, numder, adress) VALUES (N'{textBox2.Text}', '{birthday.ToString("yyyy-MM-dd")}', N'{textBox3.Text}', N'{textBox4.Text}')");

            // Обновляем DataGridView
            DbManeger.Select("SELECT id_student, fio AS 'Студент', " +
                             "birthday AS 'Дата рождения', " +
                             "numder AS 'Номер телефона', " +
                             "adress AS 'Адрес' FROM student", dataGridView1);

            // Очищаем текстовые поля и устанавливаем дату на текущую
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            dateTimePicker1.Value = DateTime.Now; // Устанавливаем дату на текущую

            UpdateComboBoxes();
        }

        //добавление записи о успеваемости
        private void button2_Click(object sender, EventArgs e)
        {
            float a;

            // Проверяем, что все поля заполнены
            if (string.IsNullOrWhiteSpace(textBox1.Text) || comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, заполните все поля и выберите студента.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверка на наличие более одной точки
            if (textBox1.Text.Count(c => c == '.') > 1)
            {
                MessageBox.Show("Ошибка: в поле GPA может быть только одна точка.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем корректность ввода GPA и диапазон от 1 до 10
            if (float.TryParse(textBox1.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out a))
            {
                // Проверяем, что GPA находится в диапазоне от 1 до 10
                if (a < 1 || a > 10)
                {
                    MessageBox.Show("Ошибка: средний балл (GPA) должен быть в диапазоне от 1 до 10.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string coefficient = "0";

                // Устанавливаем коэффициент в зависимости от значения GPA
                if (a <= 5)
                {
                    coefficient = "0";
                }
                else if (a > 5 && a < 8)
                {
                    coefficient = "1";
                }
                else if (a >= 8 && a <= 10)
                {
                    coefficient = "1.5";
                }

                // Проверяем дату
                DateTime date = dateTimePicker2.Value.Date;
                DateTime currentDate = DateTime.Now.Date;

                if (date > currentDate || date < currentDate.AddMonths(-1))
                {
                    MessageBox.Show("Дата составления должна быть не в будущем и не больше месяца назад.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Проверяем, существует ли уже запись с такой же датой и ФИО
                string checkQuery = $"SELECT COUNT(*) FROM grade WHERE id_student = N'{comboBox1.SelectedValue}' AND date_cr = '{date.ToString("yyyy-MM-dd")}'";
                int count = (int)DbManeger.SelectScalar(checkQuery); // Предполагается, что SelectScalar возвращает одно значение.

                // Если запись уже существует
                if (count > 0)
                {
                    MessageBox.Show("Запись с данной датой и ФИО уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Вставка данных в базу данных
                DbManeger.Action($"INSERT INTO grade (id_student, gpa, cof, date_cr) VALUES (N'{comboBox1.SelectedValue}', N'{textBox1.Text}', '{coefficient}', '{date.ToString("yyyy-MM-dd")}')");

                // Обновляем DataGridView
                DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                                 "gpa AS 'Средний балл', " +
                                 "cof AS 'Коэффициент', " +
                                 "date_cr AS 'Дата составления' FROM grade INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

                // Очищаем поля ввода
                textBox1.Text = "";
                comboBox1.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Некорректный формат GPA. Убедитесь, что вы ввели число.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        //добавление записи о приказе
        private void button3_Click(object sender, EventArgs e)
        {
            // Проверяем, что все поля заполнены
            if (comboBox2.SelectedIndex == -1 || string.IsNullOrWhiteSpace(comboBox3.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем дату из dateTimePicker
            DateTime selectedDate = dateTimePicker3.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата не прошедшая и не более чем на месяц вперед
            if (selectedDate < currentDate)
            {
                MessageBox.Show("Дата не может быть в прошлом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (selectedDate > currentDate.AddMonths(1))
            {
                MessageBox.Show("Дата не может быть более чем на месяц вперед.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем, существует ли уже запись с такими же данными, кроме id
            string checkQuery = $"SELECT COUNT(*) FROM orders WHERE id_student = N'{comboBox2.SelectedValue}' " +
                                $"AND cause = N'{comboBox3.Text}' AND date_cr = '{selectedDate:yyyy-MM-dd}'";
            int count = (int)DbManeger.SelectScalar(checkQuery); // Предполагается, что SelectScalar возвращает одно значение.

            // Если запись с такими данными уже существует
            if (count > 0)
            {
                MessageBox.Show("Такой приказ уже есть.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Если все проверки пройдены, выполняем действие добавления в базу
            DbManeger.Action($"INSERT INTO orders (id_student, cause, date_cr) VALUES (N'{comboBox2.SelectedValue}', " +
                             $"N'{comboBox3.Text}', '{selectedDate:yyyy-MM-dd}')");

            // Обновляем DataGridView
            DbManeger.Select("SELECT id_orders, fio AS 'Студент', " +
                             "cause AS 'Причина', " +
                             "date_cr AS 'Дата составления' FROM orders INNER JOIN student ON orders.id_student = student.id_student", dataGridView3);

            // Очищаем поля ввода
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }

        //добавление записи о стипендии
        private void button4_Click(object sender, EventArgs e)
        {
            // Проверка корректности суммы доплаты
            if (!decimal.TryParse(textBox14.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal summ) ||
                summ < 0 || summ > 150 ||
                !Regex.IsMatch(textBox14.Text, @"^(0|[1-9]\d*)(\.\d{1,2})?$"))
            {
                MessageBox.Show("Ошибка: сумма доплаты должна быть от 0 до 150 и иметь максимум 2 знака " +
                                "после запятой.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем выбранную дату
            DateTime selectedDate = dateTimePicker4.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата не может быть в будущем или в прошлом
            if (selectedDate < currentDate)
            {
                MessageBox.Show("Дата не может быть в прошлом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем, существует ли уже запись с такими же данными (по дате и имени студента)
            string checkQuery = $"SELECT COUNT(*) FROM schoolarship WHERE id_student = " +
                                $"N'{comboBox4.SelectedValue}' AND date_cr = '{selectedDate:yyyy-MM-dd}'";
            int count = (int)DbManeger.SelectScalar(checkQuery); // Предполагается, что SelectScalar возвращает одно значение.

            // Если запись с такой датой и студентом уже существует
            if (count > 0)
            {
                MessageBox.Show("Стипендия для этого студента уже есть.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Если все проверки пройдены, выполняем вставку
            DbManeger.Action($"INSERT INTO schoolarship (id_student, type_s, date_cr, summ) VALUES (N'{comboBox4.SelectedValue}', " +
                             $"N'{comboBox5.Text}', '{selectedDate:yyyy-MM-dd}', N'{textBox14.Text}')");

            // Обновляем DataGridView
            DbManeger.Select("SELECT id_schoolarship, fio AS 'Студент'," +
                             "type_s AS 'Тип', " +
                             "schoolarship.date_cr AS 'Дата выдачи', " +
                             "cause AS 'Причина надбавки', " +
                             "summ AS 'Сумма по приказу', " +
                             "gpa AS 'Средний балл', " +
                             "cof AS 'Коэффициент', " +
                             "cof * 118 + summ AS 'Сумма стипендии' " +
                             "FROM (schoolarship INNER JOIN student ON schoolarship.id_student = student.id_student) " +
                             "INNER JOIN orders ON orders.id_student = student.id_student " +
                             "INNER JOIN grade ON grade.id_student = student.id_student", dataGridView4);

            // Очищаем выборы
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            textBox14.Text = "";
            dateTimePicker4.Value = DateTime.Now; // Установить дату на текущую
        }


        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //удаление студента
        private void button8_Click(object sender, EventArgs e)
        {
            // Запрос на подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить данного студента?",
                                                   "Подтверждение удаления",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем удаление
            if (result == DialogResult.Yes)
            {
                // Проверяем, выбрана ли строка в dataGridView1
                if (dataGridView1.CurrentRow != null)
                {
                    try
                    {
                        // Получаем ID студента для удаления
                        int studentId = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);

                        // Выполнение операции удаления
                        DbManeger.Action($"DELETE FROM student WHERE id_student = {studentId}");

                        // Обновляем данные в DataGridView
                        DbManeger.Select("SELECT id_student, fio AS 'Студент', " +
                                         "birthday AS 'Дата рождения', " +
                                         "numder AS 'Номер телефона', " +
                                         "adress AS 'Адрес' FROM student", dataGridView1);

                        // Сообщение об успешном удалении
                        MessageBox.Show("Студент успешно удален.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Нельзя удалить данного студента: ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Для удаления выберите строку таблицы, нажмите на нее и нажмите кнопку " +
                    "'Удалить'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        //удаление успеваемости
        private void button7_Click(object sender, EventArgs e)
        {
            // Проверка, есть ли выделенная строка
            if (dataGridView2.CurrentRow == null)
            {
                MessageBox.Show("Для удаления выберите строку таблицы, нажмите на нее и нажмите кнопку " +
                    "'Удалить'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Запрос на подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить данную ведомость?",
                                                   "Подтверждение удаления",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем удаление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Получаем id_grade и проверяем, что преобразование успешно
                    if (int.TryParse(dataGridView2.CurrentRow.Cells[0].Value.ToString(), out int idGrade))
                    {
                        // Выполнение операции удаления
                        DbManeger.Action($"DELETE FROM grade WHERE id_grade={idGrade}");

                        // Обновляем данные в DataGridView
                        DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                                         "gpa AS 'Средний балл', " +
                                         "cof AS 'Коэффициент', " +
                                         "date_cr AS 'Дата составления' " +
                                         "FROM grade " +
                                         "INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

                        MessageBox.Show("Ведомость успешно удалена.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information); // сообщение об успешном удалении
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при получении ID ведомости.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Нельзя удалить данную ведомость: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //удаление приказа
        private void button6_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView3.CurrentRow == null)
            {
                MessageBox.Show("Для удаления выберите строку таблицы, нажмите на нее и нажмите кнопку " +
                    "'Удалить'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Запрос на подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить данный приказ?",
                                                   "Подтверждение удаления",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем удаление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции удаления
                    DbManeger.Action($"DELETE FROM orders WHERE id_orders={dataGridView3.CurrentRow.Cells[0].Value.ToString()}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("SELECT id_orders, fio AS 'Студент', " +
                                     "cause AS 'Причина', " +
                                     "date_cr AS 'Дата составления' FROM orders INNER JOIN student ON orders.id_student=student.id_student", dataGridView3);
                }
                catch (Exception)
                {
                    MessageBox.Show("Нельзя удалить данный приказ.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //удаление стипендии
        private void button5_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView4.CurrentRow == null)
            {
                MessageBox.Show("Для удаления выберите строку таблицы, нажмите на нее и нажмите кнопку " +
                    "'Удалить'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Запрос на подтверждение удаления
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить данную информацию о стипендии?",
                                                   "Подтверждение удаления",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем удаление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции удаления
                    DbManeger.Action($"DELETE FROM schoolarship WHERE id_schoolarship={dataGridView4.CurrentRow.Cells[0].Value.ToString()}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("Select id_schoolarship, fio as 'Студент'," +
                             "type_s as 'Тип',    " +
                             "schoolarship.date_cr as 'Дата выдачи' ," +
                             "cause as 'Причина надбавки', " +
                             "summ as 'Сумма по приказу',  " +
                             "gpa as 'Средний балл', " +
                             "cof as 'Коэффициент', " +
                             "cof*118+summ as 'Сумма стипендии'" +
                             "from((schoolarship inner join student on  schoolarship.id_student = student.id_student) " +
                             "inner join orders on orders.id_student = student.id_student) inner join grade on " +
                             "grade.id_student = student.id_student", dataGridView4);
                }
                catch (Exception)
                {
                    MessageBox.Show("Нельзя удалить данную информацию о стипендии.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //изменение студента
        private void button9_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Для изменения записи о студенте выберите строку таблицы, нажмите на нее, затем" +
                                " замените значения и нажмите кнопку  'Изменить'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Проверяем, что все текстовые поля заполнены
            if (string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox3.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем формат номера телефона
            if (!Regex.IsMatch(textBox3.Text, @"^\+375(29|25|44)\d{7}$"))
            {
                MessageBox.Show("Некорректный формат номера телефона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем дату рождения
            DateTime birthday = dateTimePicker1.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата рождения соответствует критериям
            if (birthday < currentDate.AddYears(-21) || birthday > currentDate.AddYears(-15))
            {
                MessageBox.Show("Дата рождения должна быть не старше 21 года и не младше 15 лет.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Запрос на подтверждение изменения
            DialogResult confirmationResult = MessageBox.Show("Вы уверены, что хотите внести изменения в данные студента?",
                                              "Подтверждение изменения",
                                              MessageBoxButtons.YesNo,
                                              MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем обновление
            if (confirmationResult == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции обновления
                    DbManeger.Action($"UPDATE student SET " +
                                     $"fio=N'{textBox2.Text}', " +
                                     $"birthday='{birthday:yyyy-MM-dd}', " +
                                     $"numder=N'{textBox3.Text}', " +
                                     $"adress=N'{textBox4.Text}' " +
                                     $"WHERE id_student={dataGridView1.CurrentRow.Cells[0].Value}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("SELECT id_student, fio AS 'Студент', " +
                                     "birthday AS 'Дата рождения', " +
                                     "numder AS 'Номер телефона', " +
                                     "adress AS 'Адрес' FROM student", dataGridView1);

                    // Очищаем текстовые поля
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";

                    // Убедитесь, что первый столбец виден
                    dataGridView1.Columns[0].Visible = true; // или просто уберите это, если столбец должен быть всегда видимым
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //изменение успеваемости
        private void button10_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView2.CurrentRow == null)
            {
                MessageBox.Show("Для изменения записи о студенте выберите строку таблицы, нажмите на нее, затем" +
                    " замените значения и нажмите кнопку  'Изменить'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Проверяем, что все поля заполнены
            if (string.IsNullOrWhiteSpace(textBox1.Text) || comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем корректность ввода GPA
            if (float.TryParse(textBox1.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out float gpa))
            {
                // Проверяем, что GPA в диапазоне от 1 до 10
                if (gpa < 1 || gpa > 10)
                {
                    MessageBox.Show("Ошибка: средний балл (GPA) должен быть в диапазоне от 1 до 10.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Некорректный формат GPA. Убедитесь, что вы ввели число.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Получаем дату из dateTimePicker
            DateTime selectedDate = dateTimePicker2.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата не может быть в будущем и не старше месяца назад
            if (selectedDate > currentDate)
            {
                MessageBox.Show("Дата не может быть в будущем.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (selectedDate < currentDate.AddMonths(-1))
            {
                MessageBox.Show("Дата не может быть более чем на месяц назад.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем, существует ли уже запись с такой же датой для выбранного студента
            string checkQuery = $"SELECT COUNT(*) FROM grade WHERE id_student = N'{comboBox1.SelectedValue}' " +
                                $"AND date_cr = '{selectedDate:yyyy-MM-dd}' AND id_grade != {dataGridView2.CurrentRow.Cells[0].Value}";
            int count = (int)DbManeger.SelectScalar(checkQuery);

            // Если запись с такой датой уже существует
            if (count > 0)
            {
                MessageBox.Show("Запись с данной датой для этого студента уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Выполнение операции обновления
            try
            {
                DbManeger.Action($"UPDATE grade SET gpa = N'{textBox1.Text}', date_cr = '{selectedDate:yyyy-MM-dd}' " +
                                 $"WHERE id_grade = {dataGridView2.CurrentRow.Cells[0].Value}");

                // Обновляем данные в DataGridView
                DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                                 "gpa AS 'Средний балл', " +
                                 "cof AS 'Коэффициент', " +
                                 "date_cr AS 'Дата составления' FROM grade INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

                // Очищаем поля ввода
                textBox1.Text = "";
                comboBox1.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            // Запрос на подтверждение изменения
            DialogResult result = MessageBox.Show("Вы уверены, что хотите внести изменения в данные о средней оценке?",
                                                       "Подтверждение изменения",
                                                       MessageBoxButtons.YesNo,
                                                       MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем обновление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции обновления
                    DbManeger.Action($"UPDATE grade SET " +
                                     $"id_student={comboBox1.SelectedValue}, " +
                                     $"gpa=N'{textBox1.Text}', " +
                                     $"date_cr='{selectedDate:yyyy-MM-dd}' " +
                                     $"WHERE id_grade={dataGridView2.CurrentRow.Cells[0].Value.ToString()}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                                "gpa AS 'Средний балл', " +
                                "cof AS 'Коэффициент', " +
                                "date_cr AS 'Дата составления' FROM grade INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

                    // Очищаем текстовые поля
                    textBox1.Text = "";
                    comboBox1.SelectedIndex = -1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //изменение приказа
        private void button11_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView3.CurrentRow == null)
            {
                MessageBox.Show("Для изменения записи о студенте выберите строку таблицы, нажмите на нее, затем" +
                    " замените значения и нажмите кнопку  'Изменить'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Проверяем, что все поля заполнены
            if (comboBox2.SelectedIndex == -1 || comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем дату из dateTimePicker
            DateTime selectedDate = dateTimePicker3.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата не может быть в прошлом и не больше месяца вперед
            if (selectedDate < currentDate)
            {
                MessageBox.Show("Дата не может быть в прошлом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (selectedDate > currentDate.AddMonths(1))
            {
                MessageBox.Show("Дата не может быть более чем на месяц вперед.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Запрос на подтверждение изменения
            DialogResult result = MessageBox.Show("Вы уверены, что хотите внести изменения в данные о приказе?",
                                                   "Подтверждение изменения",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем обновление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции обновления
                    DbManeger.Action($"UPDATE orders SET " +
                                     $"id_student={comboBox2.SelectedValue}, " +
                                     $"cause=N'{comboBox3.SelectedValue}', " +
                                     $"date_cr='{selectedDate:yyyy-MM-dd}' " +
                                     $"WHERE id_orders={dataGridView3.CurrentRow.Cells[0].Value}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("SELECT id_orders, fio AS 'Студент', " +
                                     "cause AS 'Причина', " +
                                     "date_cr AS 'Дата составления' FROM orders INNER JOIN student ON orders.id_student = student.id_student", dataGridView3);

                    // Очищаем выборы
                    comboBox2.SelectedIndex = -1;
                    comboBox3.SelectedIndex = -1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //изменение стипендии
        private void button12_Click(object sender, EventArgs e)
        {
            // Проверка, выбрана ли какая-либо строка
            if (dataGridView4.CurrentRow == null)
            {
                MessageBox.Show("Для изменения записи о студенте выберите строку таблицы, нажмите на нее, затем" +
                    " замените значения и нажмите кнопку  'Изменить'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Проверяем, что все поля заполнены
            if (comboBox4.SelectedIndex == -1 || comboBox5.SelectedIndex == -1 || string.IsNullOrWhiteSpace(textBox14.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем дату из dateTimePicker
            DateTime selectedDate = dateTimePicker4.Value.Date;
            DateTime currentDate = DateTime.Now.Date;

            // Проверяем, что дата не может быть в будущем и не в прошлом
            if (selectedDate < currentDate)
            {
                MessageBox.Show("Дата не может быть в прошлом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (selectedDate > currentDate)
            {
                MessageBox.Show("Дата не может быть в будущем.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Запрос на подтверждение изменения
            DialogResult result = MessageBox.Show("Вы уверены, что хотите внести изменения в данные о стипендии?",
                                                   "Подтверждение изменения",
                                                   MessageBoxButtons.YesNo,
                                                   MessageBoxIcon.Question);

            // Если пользователь нажал "Да", выполняем обновление
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Выполнение операции обновления
                    DbManeger.Action($"UPDATE schoolarship SET " +
                                     $"id_student={comboBox4.SelectedValue}, " +
                                     $"type_s=N'{comboBox5.Text}', " +
                                     $"date_cr='{selectedDate:yyyy-MM-dd}', " +
                                     $"summ=N'{textBox14.Text}' " +
                                     $"WHERE id_schoolarship={dataGridView4.CurrentRow.Cells[0].Value}");

                    // Обновляем данные в DataGridView
                    DbManeger.Select("Select id_schoolarship, fio as 'Студент'," +
                             "type_s as 'Тип',    " +
                             "schoolarship.date_cr as 'Дата выдачи' ," +
                             "cause as 'Причина надбавки', " +
                             "summ as 'Сумма по приказу',  " +
                             "gpa as 'Средний балл', " +
                             "cof as 'Коэффициент', " +
                             "cof*118+summ as 'Сумма стипендии'" +
                             "from((schoolarship inner join student on  schoolarship.id_student = student.id_student) " +
                             "inner join orders on orders.id_student = student.id_student) inner join grade on grade.id_student = student.id_student", dataGridView4);

                    // Очищаем выборы
                    comboBox4.SelectedIndex = -1;
                    comboBox5.SelectedIndex = -1;
                    dateTimePicker4.Value = DateTime.Now; // Устанавливаем текущую дату
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при обновлении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //поиск студента
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            string searchText = textBox6.Text; // Получаем текст для поиска из текстового поля
            bool found = false; // Флаг для отслеживания найденных совпадений

            // Проверяем, является ли текст поиска пустым
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Если текст поиска пуст, сбрасываем цвет фона всех ячеек в DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = dataGridView1.DefaultCellStyle.BackColor; // Сбрасываем цвет к стандартному
                    }
                }
            }
            else
            {
                // Если введен текст для поиска, начинаем искать совпадения
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        // Проверяем, содержит ли ячейка текст поиска (игнорируя регистр)
                        if (cell.Value != null && cell.Value.ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Если совпадение найдено, изменяем цвет фона ячейки
                            cell.Style.BackColor = Color.CadetBlue;
                            found = true; // Устанавливаем флаг найденного совпадения
                        }
                        else
                        {
                            // Если совпадение не найдено, сбрасываем цвет фона ячейки
                            cell.Style.BackColor = dataGridView1.DefaultCellStyle.BackColor; // Сбрасываем цвет к стандартному
                        }
                    }
                }

                // Если совпадений не найдено, отображаем сообщение
                if (!found)
                {
                    MessageBox.Show("Ничего не найдено. Проверьте введенный текст.", "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        //поиск успеваемости
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string searchText = textBox7.Text; // Получаем текст для поиска из текстового поля
            bool found = false; // Флаг для отслеживания найденных совпадений

            // Проверяем, является ли текст поиска пустым
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Если текст поиска пуст, сбрасываем цвет фона всех ячеек в DataGridView
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = dataGridView2.DefaultCellStyle.BackColor; // Сбрасываем цвет к стандартному
                    }
                }
            }
            else
            {
                // Если введен текст для поиска, начинаем искать совпадения
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        // Проверяем, содержит ли ячейка текст поиска (игнорируя регистр)
                        if (cell.Value != null && cell.Value.ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Если совпадение найдено, изменяем цвет фона ячейки
                            cell.Style.BackColor = Color.CadetBlue;
                            found = true; // Устанавливаем флаг найденного совпадения
                        }
                        else
                        {
                            // Если совпадение не найдено, сбрасываем цвет фона ячейки
                            cell.Style.BackColor = dataGridView2.DefaultCellStyle.BackColor; // Сбрасываем цвет к стандартному
                        }
                    }
                }

                // Если совпадений не найдено, отображаем сообщение
                if (!found)
                {
                    MessageBox.Show("Ничего не найдено. Проверьте введенный текст.", "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        //поиск приказа
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            string searchText = textBox8.Text; // Получаем текст для поиска из текстового поля
            bool found = false; // Флаг для отслеживания найденных совпадений

            // Проверяем, является ли текст поиска пустым
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Если текст поиска пуст, сбрасываем цвет фона всех ячеек в DataGridView
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = dataGridView3.DefaultCellStyle.BackColor;
                    }
                }
            }
            else
            {
                // Если введен текст для поиска, начинаем искать совпадения
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        // Проверяем, содержит ли ячейка текст поиска (игнорируя регистр)
                        if (cell.Value != null && cell.Value.ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Если совпадение найдено, изменяем цвет фона ячейки
                            cell.Style.BackColor = Color.CadetBlue;
                            found = true; // Устанавливаем флаг найденного совпадения
                        }
                        else
                        {
                            // Если совпадение не найдено, сбрасываем цвет фона ячейки
                            cell.Style.BackColor = dataGridView3.DefaultCellStyle.BackColor;
                        }
                    }
                }

                // Если совпадений не найдено, отображаем сообщение
                if (!found)
                {
                    MessageBox.Show("Ничего не найдено. Проверьте введенный текст.", "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        //поиск студента
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            string searchText = textBox9.Text; // Получаем текст для поиска из текстового поля
            bool found = false; // Флаг для отслеживания найденных совпадений

            // Проверяем, является ли текст поиска пустым
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Если текст поиска пуст, сбрасываем цвет фона всех ячеек в DataGridView
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = dataGridView4.DefaultCellStyle.BackColor;
                    }
                }
            }
            else
            {
                // Если введен текст для поиска, начинаем искать совпадения
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        // Проверяем, содержит ли ячейка текст поиска (игнорируя регистр)
                        if (cell.Value != null && cell.Value.ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Если совпадение найдено, изменяем цвет фона ячейки
                            cell.Style.BackColor = Color.CadetBlue;
                            found = true; // Устанавливаем флаг найденного совпадения
                        }
                        else
                        {
                            // Если совпадение не найдено, сбрасываем цвет фона ячейки
                            cell.Style.BackColor = dataGridView4.DefaultCellStyle.BackColor;
                        }
                    }
                }

                // Если совпадений не найдено, отображаем сообщение
                if (!found)
                {
                    MessageBox.Show("Ничего не найдено. Проверьте введенный текст.", "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //для отображения при нажатии на ячейку СТУДЕНТ
        private void dataGridView1_Click_1(object sender, EventArgs e)
        {
            try
            {
                textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();

                //поиск индекса для отчета студента
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    int selectedRowIndex = dataGridView1.SelectedRows[0].Index;
                    textBox5.Text = selectedRowIndex.ToString();
                }
                else
                {
                    textBox5.Text = "";
                }
            }
            catch (Exception) { }
        }

        //для отображения при нажатии на ячейку УСПЕВАЕМОСТЬ
        private void dataGridView2_Click_1(object sender, EventArgs e)
        {
            try
            {
                comboBox1.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
                textBox1.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                dateTimePicker2.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();

                //поиск индекса для отчета успеваемости
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int selectedRowIndex = dataGridView2.SelectedRows[0].Index;
                    textBox15.Text = selectedRowIndex.ToString();
                }
                else
                {
                    textBox15.Text = "";
                }
            }
            catch (Exception) { }
        }

        //для отображения при нажатии на ячейку ПРИКАЗ
        private void dataGridView3_Click_1(object sender, EventArgs e)
        {
            try
            {
                comboBox2.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
                comboBox3.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
                dateTimePicker3.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();

                //поиск индекса для отчета приказа
                if (dataGridView3.SelectedRows.Count > 0)
                {
                    int selectedRowIndex = dataGridView3.SelectedRows[0].Index;
                    textBox16.Text = selectedRowIndex.ToString();
                }
                else
                {
                    textBox16.Text = "";
                }
            }

            catch (Exception) { }
        }

        //для отображения при нажатии на ячейку СТИПЕНДИЯ
        private void dataGridView4_Click(object sender, EventArgs e)
        {
            try
            {
                comboBox4.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();
                comboBox5.Text = dataGridView4.CurrentRow.Cells[2].Value.ToString();
                dateTimePicker4.Text = dataGridView4.CurrentRow.Cells[3].Value.ToString();
                textBox14.Text = dataGridView4.CurrentRow.Cells[5].Value.ToString();

                //поиск индекса для отчета стипендии
                if (dataGridView4.SelectedRows.Count > 0)
                {
                    int selectedRowIndex = dataGridView4.SelectedRows[0].Index;
                    textBox17.Text = selectedRowIndex.ToString();
                }
                else
                {
                    textBox17.Text = "";
                }
            }
            catch (Exception) { }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //фильтрация рождения СТУДЕНТА
        private void button16_Click(object sender, EventArgs e)
        {
            DateTime d;
            // Проверка на правильный формат даты из dateTimePicker5
            if (DateTime.TryParse(dateTimePicker5.Value.Date.ToString("yyyy-MM-dd"), out d))
            {
                string date = d.ToString("yyyy-MM-dd");
                string columnName = "Дата рождения"; //имя столбца свое где нужно фильтровать
                string form = string.Format("[{0}] >= #{1}# AND [{0}] <= #{2}#", columnName, date, DateTime.Now.ToString("MM/dd/yyyy"));

                // Убираем неоднозначность с DataTable
                var defaultView = ((System.Data.DataTable)dataGridView1.DataSource).DefaultView;
                defaultView.RowFilter = form;

                // Проверка, есть ли записи в отфильтрованном представлении
                if (defaultView.Count == 0)
                {
                    MessageBox.Show("Записи не найдены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Некорректная дата.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //фильтрация среднего балла в УСПЕВАЕМОСТИ  
        private void button13_Click(object sender, EventArgs e)
        {
            // Обработка нажатия кнопки или другого триггера для фильтрации
            try
            {
                // Проверяем, что текстовые поля не пустые
                if (string.IsNullOrWhiteSpace(textBox10.Text) || string.IsNullOrWhiteSpace(textBox11.Text))
                {
                    MessageBox.Show("Введите значения для диапазона.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Проверяем, что введенные значения являются корректными числами
                if (int.TryParse(textBox10.Text, out int minGpa) && int.TryParse(textBox11.Text, out int maxGpa))
                {
                    if (minGpa > maxGpa)
                    {
                        MessageBox.Show("Минимальный балл не может быть больше максимального балла.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Выполнение SQL-запроса
                    var query = $"SELECT id_grade, fio AS 'Студент', " +
                                        "gpa AS 'Средний балл', " +
                                        "cof AS 'Коэффициент', " +
                                        "date_cr AS 'Дата составления' FROM grade " +
                                        "INNER JOIN student ON grade.id_student = student.id_student " +
                                        $"WHERE grade.gpa BETWEEN {minGpa} AND {maxGpa}";

                    DbManeger.Select(query, dataGridView2);

                    // Проверка наличия записей
                    if (dataGridView2.Rows.Count == 0)
                    {
                        MessageBox.Show("Записи не найдены для указанного диапазона.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Введите корректные значения для диапазона.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтрация фио в УСПЕВАЕМОСТИ
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что элемент выбран в comboBox6
                if (comboBox6.SelectedValue != null)
                {
                    // Получаем выбранное значение как строку. В зависимости от того, как у вас организован ComboBox, может потребоваться преобразование типа.
                    var selectedStudentId = comboBox6.SelectedValue;

                    // Выполняем SQL-запрос с использованием выбранного значения
                    var query = $"SELECT id_grade, fio AS 'Студент', " +
                                "gpa AS 'Средний балл', " +
                                "cof AS 'Коэффициент', " +
                                "date_cr AS 'Дата составления' FROM grade " +
                                "INNER JOIN student ON grade.id_student = student.id_student " +
                                $"WHERE grade.id_student = {selectedStudentId}";

                    DbManeger.Select(query, dataGridView2);

                    // Проверка наличия записей после выполнения запроса
                    if (dataGridView2.Rows.Count == 0)
                    {
                        MessageBox.Show("Записи не найдены для выбранного студента.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Пожалуйста, выберите студента из списка.", "Ошибка выбора", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтрация причины ПРИКАЗ
        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(comboBox9.Text))
                {
                    string cause = comboBox9.Text;  // Получение текста причины из ComboBox

                    // Выполнение SQL-запроса
                    DbManeger.Select($"SELECT id_orders, fio AS 'Студент', " +
                                     $"cause AS 'Причина', " +
                                     $"date_cr AS 'Дата составления' " +
                                     $"FROM orders INNER JOIN student ON orders.id_student = student.id_student " +
                                     $"WHERE orders.cause = N'{cause}'", dataGridView3);

                    // Проверка наличия записей после выполнения запроса
                    if (dataGridView3.Rows.Count == 0)
                    {
                        MessageBox.Show("Записи не найдены для выбранной причины.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Пожалуйста, выберите причину из списка.", "Ошибка выбора", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтрация дата выдачи СТИПЕНДИЯ
        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime d;
                if (DateTime.TryParse(dateTimePicker5.Value.Date.ToString("yyyy-MM-dd"), out d))
                {
                    string date = d.ToString("yyyy-MM-dd");
                    string columnName = "Дата выдачи";
                    string form = string.Format("[{0}] >= #{1}# AND [{0}] <= #{2}#", columnName, date, DateTime.Now.ToString("yyyy-MM-dd"));

                    var defaultView = ((System.Data.DataTable)dataGridView4.DataSource).DefaultView;
                    defaultView.RowFilter = form;

                    if (defaultView.Count == 0)
                    {
                        MessageBox.Show("Записи не найдены.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Неверный формат даты.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при фильтрации записей: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтрация тип СТИПЕНДИЯ
        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(comboBox7.Text))
                {
                    string typeScholarship = comboBox7.Text;  // Получение текста типа стипендии из ComboBox

                    // Добавляем условие WHERE для фильтрации по типу стипендии
                    string query = $"SELECT id_schoolarship, fio AS 'Студент', " +
                                   "type_s AS 'Тип', " +
                                   "schoolarship.date_cr AS 'Дата выдачи', " +
                                   "cause AS 'Причина надбавки', " +
                                   "summ AS 'Сумма по приказу', " +
                                   "gpa AS 'Средний балл', " +
                                   "cof AS 'Коэффициент', " +
                                   "(cof * 118 + summ) AS 'Сумма стипендии' " +
                                   "FROM ((schoolarship " +
                                   "INNER JOIN student ON schoolarship.id_student = student.id_student) " +
                                   "INNER JOIN orders ON orders.id_student = student.id_student) " +
                                   "INNER JOIN grade ON grade.id_student = student.id_student " +
                                   $"WHERE type_s = N'{typeScholarship}'"; // Фильтрация по типу стипендии

                    DbManeger.Select(query, dataGridView4);
                }
                else
                {
                    MessageBox.Show("Пожалуйста, выберите тип стипендии из списка.", "Ошибка выбора", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //вернуть все в таблицу успеваемости
        private void button15_Click(object sender, EventArgs e)
        {
            DbManeger.Select("SELECT id_grade, fio AS 'Студент', " +
                             "gpa AS 'Средний балл', " +
                             "date_cr AS 'Дата составления' " +
                             "FROM grade " +
                             "INNER JOIN student ON grade.id_student = student.id_student", dataGridView2);

            MessageBox.Show("Все данные в таблице.", "Возврат данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //вернуть все в таблицу студент
        private void button17_Click(object sender, EventArgs e)
        {
            DbManeger.Select("SELECT id_student, fio AS 'Студент', " +
                 "birthday AS 'Дата рождения', " +
                 "numder AS 'Номер телефона', " +
                 "adress AS 'Адрес' FROM student", dataGridView1);

            MessageBox.Show("Все данные в таблице.", "Возврат данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //вернуть все в таблицу приказ
        private void button21_Click(object sender, EventArgs e)
        {
            DbManeger.Select("Select id_orders, fio as 'Студент', " +
                             "cause as 'Причина', " +
                             "date_cr as 'Дата составления' from orders inner join student on  orders.id_student=student.id_student", dataGridView3);

            MessageBox.Show("Все данные в таблице.", "Возврат данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //вернуть все в таблицу стипендия
        private void button23_Click(object sender, EventArgs e)
        {
            DbManeger.Select("Select id_schoolarship, fio as 'Студент'," +
                             "type_s as 'Тип',    " +
                             "schoolarship.date_cr as 'Дата выдачи' ," +
                             "cause as 'Причина надбавки', " +
                             "summ as 'Сумма по приказу',  " +
                             "gpa as 'Средний балл', " +
                             "cof as 'Коэффициент', " +
                             "cof*118+summ as 'Сумма стипендии'" +
                             "from((schoolarship inner join student on  schoolarship.id_student = student.id_student) " +
                             "inner join orders on orders.id_student = student.id_student) inner join grade on grade.id_student " +
                             "= student.id_student", dataGridView4);

            MessageBox.Show("Все данные в таблице.", "Возврат данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //Excel студенты
        private void button24_Click(object sender, EventArgs e)
        {
            // Создаем новый экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Добавляем новую книгу
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Устанавливаем ширину колонок
            ExcelApp.Columns.ColumnWidth = 15;

            // Получаем активный лист
            Microsoft.Office.Interop.Excel.Worksheet wsh = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;

            // Объединяем ячейки в первой строке для заголовка
            wsh.Cells.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGridView1.ColumnCount]].Merge();

            // Центрируем текст в заголовке
            wsh.Columns.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Задаем цвет фона заголовка
            ExcelApp.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            // Устанавливаем текст заголовка и названия колонок
            ExcelApp.Cells[1, 1] = "Студент";
            ExcelApp.Cells[2, 1] = "id студента";
            ExcelApp.Cells[2, 2] = "ФИО";
            ExcelApp.Cells[2, 3] = "Дата рождения";
            ExcelApp.Cells[2, 4] = "Номер телефона";
            ExcelApp.Cells[2, 5] = "Адрес";

            // Копируем данные из DataGridView в Excel
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount - 1; j++)
                {
                    // Заполняем ячейки Excel значениями из DataGridView
                    ExcelApp.Cells[j + 3, i + 1] = dataGridView1[i, j].Value.ToString();
                }
            }

            // Получаем диапазон всех использованных ячеек
            Microsoft.Office.Interop.Excel.Range eRange = wsh.UsedRange;

            // Автоматически подгоняем высоту строк и ширину столбцов
            eRange.EntireRow.AutoFit();
            eRange.EntireColumn.AutoFit();

            // Добавляем границы ко всем ячейкам в диапазоне
            Microsoft.Office.Interop.Excel.Borders border = eRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Добавляем строку для подписи
            ExcelApp.Cells[dataGridView1.RowCount + 3, dataGridView1.ColumnCount - 1] = "подпись";

            // Делаем приложение Excel видимым для пользователя
            ExcelApp.Visible = true;
        }

        //Excel успеваемость
        private void button25_Click(object sender, EventArgs e)
        {
            // Создаем новый экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Добавляем новую книгу
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Устанавливаем ширину колонок
            ExcelApp.Columns.ColumnWidth = 15;

            // Получаем активный лист
            Microsoft.Office.Interop.Excel.Worksheet wsh = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;

            // Объединяем ячейки в первой строке для заголовка
            wsh.Cells.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGridView2.ColumnCount]].Merge();

            // Центрируем текст в заголовке
            wsh.Columns.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Задаем цвет фона заголовка
            ExcelApp.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            // Устанавливаем текст заголовка и названия колонок
            ExcelApp.Cells[1, 1] = "Успеваемость";
            ExcelApp.Cells[2, 1] = "id успеваемости";
            ExcelApp.Cells[2, 2] = "ФИО";
            ExcelApp.Cells[2, 3] = "Средний балл";
            ExcelApp.Cells[2, 4] = "Коэффициент";
            ExcelApp.Cells[2, 5] = "Дата составления";

            // Копируем данные из DataGridView в Excel
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView2.RowCount - 1; j++)
                {
                    // Заполняем ячейки Excel значениями из DataGridView
                    ExcelApp.Cells[j + 3, i + 1] = dataGridView2[i, j].Value.ToString();
                }
            }

            // Получаем диапазон всех использованных ячеек
            Microsoft.Office.Interop.Excel.Range eRange = wsh.UsedRange;

            // Автоматически подгоняем высоту строк и ширину столбцов
            eRange.EntireRow.AutoFit();
            eRange.EntireColumn.AutoFit();

            // Добавляем границы ко всем ячейкам в диапазоне
            Microsoft.Office.Interop.Excel.Borders border = eRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Добавляем строку для подписи
            ExcelApp.Cells[dataGridView2.RowCount + 3, dataGridView2.ColumnCount - 1] = "подпись";

            // Делаем приложение Excel видимым для пользователя
            ExcelApp.Visible = true;
        }

        //Excel приказ
        private void button26_Click(object sender, EventArgs e)
        {
            // Создаем новый экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Добавляем новую книгу
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Устанавливаем ширину колонок
            ExcelApp.Columns.ColumnWidth = 15;

            // Получаем активный лист
            Microsoft.Office.Interop.Excel.Worksheet wsh = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;

            // Объединяем ячейки в первой строке для заголовка
            wsh.Cells.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGridView3.ColumnCount]].Merge();

            // Центрируем текст в заголовке
            wsh.Columns.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Задаем цвет фона заголовка
            ExcelApp.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            // Устанавливаем текст заголовка и названия колонок
            ExcelApp.Cells[1, 1] = "Приказ о доплате";
            ExcelApp.Cells[2, 1] = "id приказа";
            ExcelApp.Cells[2, 2] = "ФИО";
            ExcelApp.Cells[2, 3] = "Причина";
            ExcelApp.Cells[2, 4] = "Дата составления";

            // Копируем данные из DataGridView в Excel
            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView3.RowCount - 1; j++)
                {
                    // Заполняем ячейки Excel значениями из DataGridView
                    ExcelApp.Cells[j + 3, i + 1] = dataGridView3[i, j].Value.ToString();
                }
            }

            // Получаем диапазон всех использованных ячеек
            Microsoft.Office.Interop.Excel.Range eRange = wsh.UsedRange;

            // Автоматически подгоняем высоту строк и ширину столбцов
            eRange.EntireRow.AutoFit();
            eRange.EntireColumn.AutoFit();

            // Добавляем границы ко всем ячейкам в диапазоне
            Microsoft.Office.Interop.Excel.Borders border = eRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Добавляем строку для подписи
            ExcelApp.Cells[dataGridView3.RowCount + 3, dataGridView3.ColumnCount - 1] = "подпись";

            // Делаем приложение Excel видимым для пользователя
            ExcelApp.Visible = true;
        }

        //Excel стипендия
        private void button27_Click(object sender, EventArgs e)
        {
            // Создаем новый экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Добавляем новую книгу
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Устанавливаем ширину колонок
            ExcelApp.Columns.ColumnWidth = 15;

            // Получаем активный лист
            Microsoft.Office.Interop.Excel.Worksheet wsh = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveSheet;

            // Объединяем ячейки в первой строке для заголовка
            wsh.Cells.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGridView3.ColumnCount]].Merge();

            // Центрируем текст в заголовке
            wsh.Columns.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Задаем цвет фона заголовка
            ExcelApp.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            // Устанавливаем текст заголовка и названия колонок
            ExcelApp.Cells[1, 1] = "Стипендия";
            ExcelApp.Cells[2, 1] = "id стипендии";
            ExcelApp.Cells[2, 2] = "ФИО";
            ExcelApp.Cells[2, 3] = "Тип";
            ExcelApp.Cells[2, 4] = "Дата выдачи";
            ExcelApp.Cells[2, 5] = "Причина доплаты";
            ExcelApp.Cells[2, 6] = "Сумма доплаты";
            ExcelApp.Cells[2, 7] = "Средний балл";
            ExcelApp.Cells[2, 8] = "Коэффициент";
            ExcelApp.Cells[2, 9] = "Сумма стипендии";

            // Копируем данные из DataGridView в Excel
            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView3.RowCount - 1; j++)
                {
                    // Заполняем ячейки Excel значениями из DataGridView
                    ExcelApp.Cells[j + 3, i + 1] = dataGridView3[i, j].Value.ToString();
                }
            }

            // Получаем диапазон всех использованных ячеек
            Microsoft.Office.Interop.Excel.Range eRange = wsh.UsedRange;

            // Автоматически подгоняем высоту строк и ширину столбцов
            eRange.EntireRow.AutoFit();
            eRange.EntireColumn.AutoFit();

            // Добавляем границы ко всем ячейкам в диапазоне
            Microsoft.Office.Interop.Excel.Borders border = eRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Добавляем строку для подписи
            ExcelApp.Cells[dataGridView3.RowCount + 3, dataGridView3.ColumnCount - 1] = "подпись";

            // Делаем приложение Excel видимым для пользователя
            ExcelApp.Visible = true;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //отчет студенты 
        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что введенное значение в textBox5 можно преобразовать в целое число
                if (int.TryParse(textBox5.Text, out int selectedRowIndex))
                {
                    // Проверяем, что индекс находится в пределах допустимого диапазона
                    if (selectedRowIndex >= 0 && selectedRowIndex < dataGridView1.Rows.Count)
                    {
                        // Проверяем, что строка не пустая
                        bool isRowEmpty = true;
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (dataGridView1.Rows[selectedRowIndex].Cells[i].Value != null &&
                                !string.IsNullOrWhiteSpace(dataGridView1.Rows[selectedRowIndex].Cells[i].Value.ToString()))
                            {
                                isRowEmpty = false; // Строка не пустая
                                break;
                            }
                        }

                        // Если строка пустая, показываем сообщение об ошибке
                        if (isRowEmpty)
                        {
                            MessageBox.Show("Выбранная строка пуста. Пожалуйста, выберите строку с данными.",
                                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Создаем приложение Word
                        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.Visible = true; // Делаем приложение видимым
                        Document wordDoc = wordApp.Documents.Add(); // Создаем новый документ

                        string fontName = "Times New Roman"; // Устанавливаем шрифт
                        int fontSize = 16; // Устанавливаем размер шрифта

                        // Добавляем пустой абзац
                        Paragraph emptyParagraph = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph.Range.Text = "";
                        emptyParagraph.Range.Font.Name = fontName;
                        emptyParagraph.Range.Font.Size = fontSize;
                        emptyParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Добавление заголовка
                        Paragraph titleParagraph = wordDoc.Content.Paragraphs.Add();
                        titleParagraph.Range.Text = "CТУДЕНТ\n";
                        titleParagraph.Range.Font.Name = fontName;
                        titleParagraph.Range.Font.Size = fontSize;
                        titleParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Добавляем еще один пустой абзац
                        Paragraph anotherEmptyParagraph = wordDoc.Content.Paragraphs.Add();
                        anotherEmptyParagraph.Range.Text = "";
                        anotherEmptyParagraph.Range.Font.Name = fontName;
                        anotherEmptyParagraph.Range.Font.Size = fontSize;
                        anotherEmptyParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Добавляем информацию о студенте
                        // Номер студента
                        Paragraph studentIdParagraph = wordDoc.Content.Paragraphs.Add();
                        studentIdParagraph.Range.Text = $"Номер студента: {dataGridView1.Rows[selectedRowIndex].Cells[0].Value.ToString()}\n";
                        studentIdParagraph.Range.Font.Name = fontName;
                        studentIdParagraph.Range.Font.Size = fontSize;

                        // ФИО
                        Paragraph fullNameParagraph = wordDoc.Content.Paragraphs.Add();
                        fullNameParagraph.Range.Text = $"ФИО: {dataGridView1.Rows[selectedRowIndex].Cells[1].Value.ToString()}\n";
                        fullNameParagraph.Range.Font.Name = fontName;
                        fullNameParagraph.Range.Font.Size = fontSize;

                        // Дата рождения
                        Paragraph birthDateParagraph = wordDoc.Content.Paragraphs.Add();
                        birthDateParagraph.Range.Text = $"Дата рождения: {dataGridView1.Rows[selectedRowIndex].Cells[2].Value.ToString()}\n";
                        birthDateParagraph.Range.Font.Name = fontName;
                        birthDateParagraph.Range.Font.Size = fontSize;

                        // Номер телефона
                        Paragraph phoneNumberParagraph = wordDoc.Content.Paragraphs.Add();
                        phoneNumberParagraph.Range.Text = $"Номер телефона: {dataGridView1.Rows[selectedRowIndex].Cells[3].Value.ToString()}\n";
                        phoneNumberParagraph.Range.Font.Name = fontName;
                        phoneNumberParagraph.Range.Font.Size = fontSize;

                        // Адрес
                        Paragraph addressParagraph = wordDoc.Content.Paragraphs.Add();
                        addressParagraph.Range.Text = $"Адрес: {dataGridView1.Rows[selectedRowIndex].Cells[4].Value.ToString()}\n";
                        addressParagraph.Range.Font.Name = fontName;
                        addressParagraph.Range.Font.Size = fontSize;
                        addressParagraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    else
                    {
                        // Если индекс не в пределах допустимого диапазона
                        MessageBox.Show("Для создания отчета Word выберите строку, нажмите на нее и нажмите кнопку " +
                                        "'Отчет Word'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //отчет успеваемость
        private void button29_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что введенное значение в textBox15 можно преобразовать в целое число
                if (int.TryParse(textBox15.Text, out int selectedRowIndex))
                {
                    // Проверяем, что индекс находится в пределах допустимого диапазона
                    if (selectedRowIndex >= 0 && selectedRowIndex < dataGridView2.Rows.Count)
                    {
                        // Проверяем, что строка не пустая
                        bool isRowEmpty = true;
                        for (int i = 0; i < dataGridView2.Columns.Count; i++)
                        {
                            if (dataGridView2.Rows[selectedRowIndex].Cells[i].Value != null &&
                                !string.IsNullOrWhiteSpace(dataGridView2.Rows[selectedRowIndex].Cells[i].Value.ToString()))
                            {
                                isRowEmpty = false; // Строка не пустая
                                break;
                            }
                        }

                        // Если строка пуста, показываем сообщение об ошибке
                        if (isRowEmpty)
                        {
                            MessageBox.Show("Выбранная строка пуста. Пожалуйста, выберите строку с данными.",
                                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Создаем приложение Word
                        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.Visible = true; // Делаем приложение видимым
                        Document wordDoc = wordApp.Documents.Add(); // Создаем новый документ

                        string fontName = "Times New Roman"; // Устанавливаем шрифт
                        int fontSize = 16; // Устанавливаем размер шрифта

                        // Добавление заголовка
                        Paragraph emptyParagraph1 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph1.Range.Text = ""; // Пустой абзац для отступа
                        emptyParagraph1.Range.Font.Name = fontName;
                        emptyParagraph1.Range.Font.Size = fontSize;
                        emptyParagraph1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Заголовок "УСПЕВАЕМОСТЬ"
                        Paragraph titleParagraph = wordDoc.Content.Paragraphs.Add();
                        titleParagraph.Range.Text = "УСПЕВАЕМОСТЬ\n";
                        titleParagraph.Range.Font.Name = fontName;
                        titleParagraph.Range.Font.Size = fontSize;
                        titleParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Добавление пустого абзаца для отступа
                        Paragraph emptyParagraph2 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph2.Range.Text = ""; // Пустой абзац
                        emptyParagraph2.Range.Font.Name = fontName;
                        emptyParagraph2.Range.Font.Size = fontSize;
                        emptyParagraph2.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Добавляем информацию о студенте
                        // Номер ведомости
                        Paragraph recordNumberParagraph = wordDoc.Content.Paragraphs.Add();
                        recordNumberParagraph.Range.Text = $"Номер ведомости: {dataGridView2.Rows[selectedRowIndex].Cells[0].Value.ToString()}\n";
                        recordNumberParagraph.Range.Font.Name = fontName;
                        recordNumberParagraph.Range.Font.Size = fontSize;

                        // ФИО
                        Paragraph fullNameParagraph = wordDoc.Content.Paragraphs.Add();
                        fullNameParagraph.Range.Text = $"ФИО: {dataGridView2.Rows[selectedRowIndex].Cells[1].Value.ToString()}\n";
                        fullNameParagraph.Range.Font.Name = fontName;
                        fullNameParagraph.Range.Font.Size = fontSize;

                        // Средний балл
                        Paragraph averageScoreParagraph = wordDoc.Content.Paragraphs.Add();
                        averageScoreParagraph.Range.Text = $"Средний балл: {dataGridView2.Rows[selectedRowIndex].Cells[2].Value.ToString()}\n";
                        averageScoreParagraph.Range.Font.Name = fontName;
                        averageScoreParagraph.Range.Font.Size = fontSize;

                        // Коэффициент
                        Paragraph coefficientParagraph = wordDoc.Content.Paragraphs.Add();
                        coefficientParagraph.Range.Text = $"Коэффициент: {dataGridView2.Rows[selectedRowIndex].Cells[3].Value.ToString()}\n";
                        coefficientParagraph.Range.Font.Name = fontName;
                        coefficientParagraph.Range.Font.Size = fontSize;

                        // Дата составления
                        Paragraph dateParagraph = wordDoc.Content.Paragraphs.Add();
                        dateParagraph.Range.Text = $"Дата составления: {dataGridView2.Rows[selectedRowIndex].Cells[4].Value.ToString()}\n";
                        dateParagraph.Range.Font.Name = fontName;
                        dateParagraph.Range.Font.Size = fontSize;
                        dateParagraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        // Если индекс не в пределах допустимого диапазона
                        MessageBox.Show("Для создания отчета Word выберите строку, нажмите на нее и нажмите кнопку " +
                                        "'Отчет Word'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //отчет приказ
        private void button30_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что введенное значение в textBox16 можно преобразовать в целое число
                if (int.TryParse(textBox16.Text, out int selectedRowIndex))
                {
                    // Проверяем, что индекс находится в пределах допустимого диапазона
                    if (selectedRowIndex >= 0 && selectedRowIndex < dataGridView3.Rows.Count)
                    {
                        // Проверяем, что строка не пустая
                        bool isRowEmpty = true;
                        for (int i = 0; i < dataGridView3.Columns.Count; i++)
                        {
                            if (dataGridView3.Rows[selectedRowIndex].Cells[i].Value != null &&
                                !string.IsNullOrWhiteSpace(dataGridView3.Rows[selectedRowIndex].Cells[i].Value.ToString()))
                            {
                                isRowEmpty = false; // Строка не пустая
                                break;
                            }
                        }

                        // Если строка пуста, показываем сообщение об ошибке
                        if (isRowEmpty)
                        {
                            MessageBox.Show("Выбранная строка пуста. Пожалуйста, выберите строку с данными.",
                                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Создаем приложение Word
                        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.Visible = true; // Делаем приложение видимым
                        Document wordDoc = wordApp.Documents.Add(); // Создаем новый документ

                        string fontName = "Times New Roman"; // Устанавливаем шрифт
                        int fontSize = 16; // Устанавливаем размер шрифта
                        int fontSiz = 20; // Размер шрифта для заголовка

                        // Добавление пустого параграфа для отступа
                        Paragraph emptyParagraph1 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph1.Range.Text = "";
                        emptyParagraph1.Range.Font.Name = fontName;
                        emptyParagraph1.Range.Font.Size = fontSiz;
                        emptyParagraph1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Заголовок "ПРИКАЗ О ДОПЛАТЕ"
                        Paragraph titleParagraph = wordDoc.Content.Paragraphs.Add();
                        titleParagraph.Range.Text = "ПРИКАЗ О ДОПЛАТЕ\n";
                        titleParagraph.Range.Font.Name = fontName;
                        titleParagraph.Range.Font.Size = fontSize;
                        titleParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Добавление пустого параграфа для отступа
                        Paragraph emptyParagraph2 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph2.Range.Text = ""; // Пустой абзац
                        emptyParagraph2.Range.Font.Name = fontName;
                        emptyParagraph2.Range.Font.Size = fontSize;
                        emptyParagraph2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Добавляем информацию о приказе
                        // Номер приказа
                        Paragraph orderNumberParagraph = wordDoc.Content.Paragraphs.Add();
                        orderNumberParagraph.Range.Text = $"Номер приказа: {dataGridView3.Rows[selectedRowIndex].Cells[0].Value.ToString()}\n";
                        orderNumberParagraph.Range.Font.Name = fontName;
                        orderNumberParagraph.Range.Font.Size = fontSize;

                        // ФИО
                        Paragraph fullNameParagraph = wordDoc.Content.Paragraphs.Add();
                        fullNameParagraph.Range.Text = $"ФИО: {dataGridView3.Rows[selectedRowIndex].Cells[1].Value.ToString()}\n";
                        fullNameParagraph.Range.Font.Name = fontName;
                        fullNameParagraph.Range.Font.Size = fontSize;

                        // Причина доплаты
                        Paragraph reasonParagraph = wordDoc.Content.Paragraphs.Add();
                        reasonParagraph.Range.Text = $"Причина: {dataGridView3.Rows[selectedRowIndex].Cells[2].Value.ToString()}\n";
                        reasonParagraph.Range.Font.Name = fontName;
                        reasonParagraph.Range.Font.Size = fontSize;

                        // Дата составления
                        Paragraph dateParagraph = wordDoc.Content.Paragraphs.Add();
                        dateParagraph.Range.Text = $"Дата составления: {dataGridView3.Rows[selectedRowIndex].Cells[3].Value.ToString()}\n";
                        dateParagraph.Range.Font.Name = fontName;
                        dateParagraph.Range.Font.Size = fontSize;
                        dateParagraph.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        // Если индекс не в пределах допустимого диапазона
                        MessageBox.Show("Для создания отчета Word выберите строку, нажмите на нее и нажмите " +
                                        "кнопку 'Отчет Word'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //отчет стипендия
        private void button31_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что введенное значение в textBox17 можно преобразовать в целое число
                if (int.TryParse(textBox17.Text, out int selectedRowIndex))
                {
                    // Проверяем, что индекс находится в пределах допустимого диапазона
                    if (selectedRowIndex >= 0 && selectedRowIndex < dataGridView4.Rows.Count)
                    {
                        // Проверяем, что строка не пустая
                        bool isRowEmpty = true;
                        for (int i = 0; i < dataGridView4.Columns.Count; i++)
                        {
                            // Если хотя бы одна ячейка не пустая
                            if (dataGridView4.Rows[selectedRowIndex].Cells[i].Value != null &&
                                !string.IsNullOrWhiteSpace(dataGridView4.Rows[selectedRowIndex].Cells[i].Value.ToString()))
                            {
                                isRowEmpty = false; // Строка не пустая
                                break;
                            }
                        }

                        // Если строка пуста, показываем сообщение об ошибке
                        if (isRowEmpty)
                        {
                            MessageBox.Show("Выбранная строка пуста. Пожалуйста, выберите строку с " +
                                            "данными.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Создаем приложение Word
                        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.Visible = true; // Делаем приложение видимым
                        Document wordDoc = wordApp.Documents.Add(); // Создаем новый документ

                        string fontName = "Times New Roman"; // Устанавливаем шрифт
                        int fontSize = 16; // Устанавливаем размер шрифта
                        int fontSiz = 20; // Размер шрифта для заголовка

                        // Добавление пустого параграфа для отступа
                        Paragraph emptyParagraph1 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph1.Range.Text = "";
                        emptyParagraph1.Range.Font.Name = fontName;
                        emptyParagraph1.Range.Font.Size = fontSiz;
                        emptyParagraph1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        // Заголовок "СТИПЕНДИЯ"
                        Paragraph titleParagraph = wordDoc.Content.Paragraphs.Add();
                        titleParagraph.Range.Text = "СТИПЕНДИЯ\n";
                        titleParagraph.Range.Font.Name = fontName;
                        titleParagraph.Range.Font.Size = fontSize;
                        titleParagraph.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Добавление пустого параграфа для отступа
                        Paragraph emptyParagraph2 = wordDoc.Content.Paragraphs.Add();
                        emptyParagraph2.Range.Text = ""; // Пустой абзац
                        emptyParagraph2.Range.Font.Name = fontName;
                        emptyParagraph2.Range.Font.Size = fontSize;
                        emptyParagraph2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Добавление параграфов с данными
                        for (int i = 1; i < 9; i++) // Предполагается, что есть 9 колонок данных
                        {
                            // Получаем значение ячейки, если пустое — ставим "N/A"
                            string cellValue = dataGridView4.Rows[selectedRowIndex].Cells[i].Value?.ToString() ?? "N/A";
                            Paragraph dataParagraph = wordDoc.Content.Paragraphs.Add();
                            dataParagraph.Range.Text = $"{dataGridView4.Columns[i].HeaderText}: {cellValue}\n";
                            dataParagraph.Range.Font.Name = fontName;
                            dataParagraph.Range.Font.Size = fontSize;
                        }
                    }
                    else
                    {
                        // Если индекс не в пределах допустимого диапазона
                        MessageBox.Show("Для создания отчета Word выберите строку, нажмите на нее и нажмите кнопку " +
                                        "'Отчет Word'.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}