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
using Npgsql;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace RadMirosh
{
    public partial class Form1 : Form
    {


       
        private string connString;
        public Form1()
        {
            InitializeComponent();
            // Замените эти значения на ваши данные
            string serverName = "localhost";
            string dbName = "mirosh";
            string userName = "postgres";
            string password = "1234";

            // Строка подключения
            connString = String.Format("Server={0};Database={1};User Id={2};Password={3};", serverName, dbName, userName, password);
        }

        private void Form1_Load(object sender, EventArgs e)
        {




            using (NpgsqlConnection conn = new NpgsqlConnection(connString))
            {
                try
                {
                    conn.Open(); // открываем соединение
                    LoadClientsToComboBox();
                    LoadProductsToComboBox();
                    LoadPaymentsToComboBox();
                    LoadClientsToCheckedListBox();
                    MessageBox.Show("Соединение с базой данных установлено");
                    conn.Close(); // закрываем соединение
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string query = "";
            
            switch (comboBox1.SelectedItem.ToString())
            {
                case "Client":
                    query = "SELECT * FROM Client";
                   
                    break;
                case "Payment":
                    
                    query = "SELECT * FROM Payment";
                    
                    
                    break;
                case "Product":
                    
                    query = "SELECT * FROM Product";
                    break;
                case "Contract_Info":
                    query = "SELECT * FROM contractinfo";
                   
                    break;
                default:
                    MessageBox.Show("Неизвестная таблица");
                    return;
            }

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();
                    using (NpgsqlDataAdapter da = new NpgsqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;
                        
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    string name = nameText.Text;
                    string lastName = lastNameText.Text;
                    decimal number = decimal.Parse(numberText.Text);
                    string sql = $"INSERT INTO Client(name, last_name, phone_number) VALUES('{name}', '{lastName}', {number})";
                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Клиент добавлен");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    // Получаем новые данные из полей ввода
                    string name = nameText.Text;
                    string lastName = lastNameText.Text;
                    decimal phoneNumber = decimal.Parse(numberText.Text);

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для обновления записи
                        string sql = $"UPDATE Client SET name = '{name}', last_name = '{lastName}', phone_number = {phoneNumber} WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Данные клиента обновлены");

                    // Обновляем данные в dataGridView и comboBox после обновления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadClientsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите клиента для обновления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для удаления записи
                        string sql = $"DELETE FROM Client WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Клиент удален");

                    // Обновляем данные в dataGridView и comboBox после удаления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadClientsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите клиента для удаления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    // Получаем новые данные из полей ввода и ComboBox-ов
                    int clientId = (int)comboBoxClient.SelectedValue;
                    int productId = (int)comboBoxProduct.SelectedValue;
                    int paymentId = (int)comboBoxPayment.SelectedValue;
                    int quantity = int.Parse(textBoxQuantity.Text);
                    decimal contractPrice = decimal.Parse(textBoxContractPrice.Text);
                    DateTime contractDate = dateTimePickerContractDate.Value;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для обновления записи
                        string sql = $"UPDATE ContractInfo SET client_id = @clientId, product_id = @productId, payment_id = @paymentId, " +
                                     $"quantity = @quantity, contract_price = @contractPrice, contract_date = @contractDate WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@clientId", clientId);
                            cmd.Parameters.AddWithValue("@productId", productId);
                            cmd.Parameters.AddWithValue("@paymentId", paymentId);
                            cmd.Parameters.AddWithValue("@quantity", quantity);
                            cmd.Parameters.AddWithValue("@contractPrice", contractPrice);
                            cmd.Parameters.AddWithValue("@contractDate", contractDate);

                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Договор обновлен");

                    // Обновляем данные в dataGridView после обновления
                    comboBox1_SelectedIndexChanged(null, null);
                }
                else
                {
                    MessageBox.Show("Выберите договор для обновления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    string name = textBoxProductName.Text;
                    decimal price = decimal.Parse(textBoxProductPrice.Text);
                    bool isCargo = checkBoxIsCargo.Checked;
                    bool isDelivered = checkBoxIsDelivered.Checked;

                    string sql = $"INSERT INTO Product(name, price, isCargo, isDelivered) VALUES('{name}', {price}, {isCargo}, {isDelivered})";
                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Товар добавлен");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    string pay_Name = payName.Text;
                    

                    string sql = $"INSERT INTO Payment(type) VALUES('{pay_Name}')";
                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Тип платежа добавлен");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void LoadClientsToComboBox()
        {
            string query = "SELECT id, name, last_name FROM Client";
            LoadToComboBox(comboBoxClient, query, "Id", "last_name");
        }

        private void LoadProductsToComboBox()
        {
            string query = "SELECT id, name FROM Product";
            LoadToComboBox(comboBoxProduct, query, "Id", "Name");
        }

        private void LoadPaymentsToComboBox()
        {
            string query = "SELECT id, type FROM Payment";
            LoadToComboBox(comboBoxPayment, query, "Id", "Type");
        }
        private void LoadToComboBox(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();
                    DataTable dt = new DataTable();
                    using (NpgsqlDataAdapter da = new NpgsqlDataAdapter(query, conn))
                    {
                        da.Fill(dt);
                    }

                    comboBox.DataSource = dt;
                    comboBox.ValueMember = valueMember;
                    comboBox.DisplayMember = displayMember;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для удаления записи
                        string sql = $"DELETE FROM Contractinfo WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Договор удален");

                    // Обновляем данные в dataGridView и comboBox после удаления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadClientsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите договор для удаления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                // Получаем данные из полей ввода
                int clientId = (int)comboBoxClient.SelectedValue;
                int productId = (int)comboBoxProduct.SelectedValue;
                int paymentId = (int)comboBoxPayment.SelectedValue;
                int quantity = int.Parse(textBoxQuantity.Text);
                decimal contractPrice = decimal.Parse(textBoxContractPrice.Text);
                DateTime contractDate = dateTimePickerContractDate.Value;

                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    // Создаем SQL-запрос для вставки новой записи
                    string sql = $"INSERT INTO ContractInfo (client_id, product_id, payment_id, quantity, contract_price, contract_date) " +
                                 $"VALUES (@clientId, @productId, @paymentId, @quantity, @contractPrice, @contractDate)";

                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@clientId", clientId);
                        cmd.Parameters.AddWithValue("@productId", productId);
                        cmd.Parameters.AddWithValue("@paymentId", paymentId);
                        cmd.Parameters.AddWithValue("@quantity", quantity);
                        cmd.Parameters.AddWithValue("@contractPrice", contractPrice);
                        cmd.Parameters.AddWithValue("@contractDate", contractDate);

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Договор добавлен");

                // Обновляем данные в dataGridView после добавления
                comboBox1_SelectedIndexChanged(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    // Получаем новые данные из поля ввода
                    string paymentType = payName.Text;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для обновления записи
                        string sql = $"UPDATE Payment SET type = @paymentType WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@paymentType", paymentType);

                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Тип платежа обновлен");

                    // Обновляем данные в dataGridView и comboBox после обновления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadPaymentsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите тип платежа для обновления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для удаления записи
                        string sql = $"DELETE FROM Payment WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Тип платежа удален");

                    // Обновляем данные в dataGridView и comboBox после удаления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadPaymentsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите тип платежа для удаления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    // Получаем новые данные из полей ввода
                    string name = textBoxProductName.Text;
                    decimal price = decimal.Parse(textBoxProductPrice.Text);
                    bool isCargo = checkBoxIsCargo.Checked;
                    bool isDelivered = checkBoxIsDelivered.Checked;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для обновления записи
                        string sql = $"UPDATE Product SET name = @name, price = @price, isCargo = @isCargo, isDelivered = @isDelivered WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@name", name);
                            cmd.Parameters.AddWithValue("@price", price);
                            cmd.Parameters.AddWithValue("@isCargo", isCargo);
                            cmd.Parameters.AddWithValue("@isDelivered", isDelivered);

                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Товар обновлен");

                    // Обновляем данные в dataGridView и comboBox после обновления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadProductsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите товар для обновления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Получаем id выбранной записи
                    int id = (int)dataGridView1.SelectedRows[0].Cells[0].Value;

                    using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        // Создаем SQL-запрос для удаления записи
                        string sql = $"DELETE FROM Product WHERE id = {id}";

                        using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Товар удален");

                    // Обновляем данные в dataGridView и comboBox после удаления
                    comboBox1_SelectedIndexChanged(null, null);
                    LoadProductsToComboBox();
                }
                else
                {
                    MessageBox.Show("Выберите товар для удаления");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void LoadClientsToCheckedListBox()
        {
            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();
                    string sql = "SELECT id, name, last_name FROM Client";
                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        using (NpgsqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string clientInfo = $"{reader.GetInt32(0)}: {reader.GetString(1)} {reader.GetString(2)}";
                                checkedListBoxClients.Items.Add(clientInfo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            LoadClientsToComboBox();
            LoadProductsToComboBox();
            LoadPaymentsToComboBox();
            LoadClientsToCheckedListBox();
        }

        private DataTable GenerateReportForClient(int clientId, DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();

            try
            {
                using (NpgsqlConnection conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    string sql = "SELECT * FROM material_report(@startDate::date, @endDate::date, @clientId)";

                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@startDate", startDate.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@endDate", endDate.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@clientId", clientId);

                        using (NpgsqlDataReader reader = cmd.ExecuteReader())
                        {
                            dt.Load(reader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Получаем выбранные даты
                DateTime startDate = dateTimePickerStartDate.Value;
                DateTime endDate = dateTimePickerEndDate.Value;

                foreach (var item in checkedListBoxClients.CheckedItems)
                {
                    // Извлекаем id клиента из строки
                    int clientId = int.Parse(item.ToString().Split(':')[0]);

                    // Генерируем отчет для этого клиента
                    DataTable dt = GenerateReportForClient(clientId, startDate, endDate);

                    // Экспортируем результат в Excel
                    ExportDataTableToExcel(dt);
                }

                MessageBox.Show("Отчеты экспортированы в Excel");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }
        }

        private void ExportDataTableToExcel(DataTable dt)
        {
            IWorkbook workbook;
            ISheet sheet;

            // Создаем книгу Excel
            if (Path.GetExtension("output.xlsx") == ".xls")
            {
                workbook = new HSSFWorkbook(); // для .xls
            }
            else
            {
                workbook = new XSSFWorkbook(); // для .xlsx
            }

            // Создаем лист в книге
            sheet = workbook.CreateSheet("Отчет");

            // Добавляем заголовки в таблицу
            var headerRow = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            // Добавляем данные в таблицу
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            // Автоматически устанавливаем ширину столбцов, чтобы все данные были видны
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            // Сохраняем книгу
            using (var fs = new FileStream(@"C:\Users\moonb\Documents\output.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
}
