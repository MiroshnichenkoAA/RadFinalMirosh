
namespace RadMirosh
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.dateTimePickerStartDate = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.nameText = new System.Windows.Forms.TextBox();
            this.lastNameText = new System.Windows.Forms.TextBox();
            this.numberText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxProductName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxProductPrice = new System.Windows.Forms.TextBox();
            this.checkBoxIsCargo = new System.Windows.Forms.CheckBox();
            this.checkBoxIsDelivered = new System.Windows.Forms.CheckBox();
            this.payName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBoxClient = new System.Windows.Forms.ComboBox();
            this.comboBoxProduct = new System.Windows.Forms.ComboBox();
            this.comboBoxPayment = new System.Windows.Forms.ComboBox();
            this.textBoxQuantity = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.dateTimePickerContractDate = new System.Windows.Forms.DateTimePicker();
            this.textBoxContractPrice = new System.Windows.Forms.TextBox();
            this.dateTimePickerEndDate = new System.Windows.Forms.DateTimePicker();
            this.checkedListBoxClients = new System.Windows.Forms.CheckedListBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.button14 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // dateTimePickerStartDate
            // 
            this.dateTimePickerStartDate.Location = new System.Drawing.Point(794, 127);
            this.dateTimePickerStartDate.Name = "dateTimePickerStartDate";
            this.dateTimePickerStartDate.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerStartDate.TabIndex = 4;
            this.dateTimePickerStartDate.ValueChanged += new System.EventHandler(this.dateTimePicker4_ValueChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(794, 192);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(200, 65);
            this.button1.TabIndex = 5;
            this.button1.Text = "EXCEL EXPORT";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Client",
            "Payment",
            "Product",
            "Contract_Info"});
            this.comboBox1.Location = new System.Drawing.Point(72, 27);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(248, 21);
            this.comboBox1.TabIndex = 6;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 315);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1025, 317);
            this.dataGridView1.TabIndex = 7;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(303, 39);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(124, 35);
            this.button2.TabIndex = 8;
            this.button2.Text = "Insert";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(303, 80);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(124, 35);
            this.button3.TabIndex = 9;
            this.button3.Text = "Update";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(303, 121);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(124, 35);
            this.button4.TabIndex = 10;
            this.button4.Text = "Delete";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(68, 69);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(441, 214);
            this.tabControl1.TabIndex = 11;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.numberText);
            this.tabPage1.Controls.Add(this.lastNameText);
            this.tabPage1.Controls.Add(this.nameText);
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(433, 188);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Клиент";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.textBoxContractPrice);
            this.tabPage2.Controls.Add(this.dateTimePickerContractDate);
            this.tabPage2.Controls.Add(this.label13);
            this.tabPage2.Controls.Add(this.label11);
            this.tabPage2.Controls.Add(this.label10);
            this.tabPage2.Controls.Add(this.label9);
            this.tabPage2.Controls.Add(this.label8);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.textBoxQuantity);
            this.tabPage2.Controls.Add(this.comboBoxPayment);
            this.tabPage2.Controls.Add(this.comboBoxProduct);
            this.tabPage2.Controls.Add(this.comboBoxClient);
            this.tabPage2.Controls.Add(this.button5);
            this.tabPage2.Controls.Add(this.button6);
            this.tabPage2.Controls.Add(this.button7);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(433, 188);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Договор";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.payName);
            this.tabPage3.Controls.Add(this.button8);
            this.tabPage3.Controls.Add(this.button9);
            this.tabPage3.Controls.Add(this.button10);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(433, 188);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Платеж";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.checkBoxIsDelivered);
            this.tabPage4.Controls.Add(this.checkBoxIsCargo);
            this.tabPage4.Controls.Add(this.textBoxProductPrice);
            this.tabPage4.Controls.Add(this.label5);
            this.tabPage4.Controls.Add(this.label4);
            this.tabPage4.Controls.Add(this.textBoxProductName);
            this.tabPage4.Controls.Add(this.button11);
            this.tabPage4.Controls.Add(this.button12);
            this.tabPage4.Controls.Add(this.button13);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(433, 188);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Товар";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(303, 78);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(124, 35);
            this.button5.TabIndex = 12;
            this.button5.Text = "Update";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(303, 119);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(124, 35);
            this.button6.TabIndex = 13;
            this.button6.Text = "Delete";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(303, 39);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(124, 35);
            this.button7.TabIndex = 11;
            this.button7.Text = "Insert";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(303, 77);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(124, 35);
            this.button8.TabIndex = 15;
            this.button8.Text = "Update";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(303, 118);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(124, 35);
            this.button9.TabIndex = 16;
            this.button9.Text = "Delete";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(303, 38);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(124, 35);
            this.button10.TabIndex = 14;
            this.button10.Text = "Insert";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(303, 75);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(124, 35);
            this.button11.TabIndex = 18;
            this.button11.Text = "Update";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(303, 116);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(124, 35);
            this.button12.TabIndex = 19;
            this.button12.Text = "Delete";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(303, 36);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(124, 35);
            this.button13.TabIndex = 17;
            this.button13.Text = "Insert";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // nameText
            // 
            this.nameText.Location = new System.Drawing.Point(82, 47);
            this.nameText.Name = "nameText";
            this.nameText.Size = new System.Drawing.Size(155, 20);
            this.nameText.TabIndex = 12;
            this.nameText.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // lastNameText
            // 
            this.lastNameText.Location = new System.Drawing.Point(82, 73);
            this.lastNameText.Name = "lastNameText";
            this.lastNameText.Size = new System.Drawing.Size(155, 20);
            this.lastNameText.TabIndex = 13;
            this.lastNameText.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // numberText
            // 
            this.numberText.Location = new System.Drawing.Point(82, 99);
            this.numberText.Name = "numberText";
            this.numberText.Size = new System.Drawing.Size(155, 20);
            this.numberText.TabIndex = 14;
            this.numberText.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 15;
            this.label1.Text = "Имя";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Фамилия";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(29, 99);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Номер";
            // 
            // textBoxProductName
            // 
            this.textBoxProductName.Location = new System.Drawing.Point(77, 44);
            this.textBoxProductName.Name = "textBoxProductName";
            this.textBoxProductName.Size = new System.Drawing.Size(141, 20);
            this.textBoxProductName.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(42, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Имя";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(42, 75);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(33, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "Цена";
            // 
            // textBoxProductPrice
            // 
            this.textBoxProductPrice.Location = new System.Drawing.Point(77, 72);
            this.textBoxProductPrice.Name = "textBoxProductPrice";
            this.textBoxProductPrice.Size = new System.Drawing.Size(141, 20);
            this.textBoxProductPrice.TabIndex = 22;
            // 
            // checkBoxIsCargo
            // 
            this.checkBoxIsCargo.AutoSize = true;
            this.checkBoxIsCargo.Location = new System.Drawing.Point(45, 98);
            this.checkBoxIsCargo.Name = "checkBoxIsCargo";
            this.checkBoxIsCargo.Size = new System.Drawing.Size(129, 17);
            this.checkBoxIsCargo.TabIndex = 23;
            this.checkBoxIsCargo.Text = "Положен к отгрузке";
            this.checkBoxIsCargo.UseVisualStyleBackColor = true;
            // 
            // checkBoxIsDelivered
            // 
            this.checkBoxIsDelivered.AutoSize = true;
            this.checkBoxIsDelivered.Location = new System.Drawing.Point(45, 121);
            this.checkBoxIsDelivered.Name = "checkBoxIsDelivered";
            this.checkBoxIsDelivered.Size = new System.Drawing.Size(75, 17);
            this.checkBoxIsDelivered.TabIndex = 24;
            this.checkBoxIsDelivered.Text = "Отгружен";
            this.checkBoxIsDelivered.UseVisualStyleBackColor = true;
            this.checkBoxIsDelivered.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // payName
            // 
            this.payName.Location = new System.Drawing.Point(77, 53);
            this.payName.Name = "payName";
            this.payName.Size = new System.Drawing.Size(100, 20);
            this.payName.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(45, 56);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(26, 13);
            this.label6.TabIndex = 18;
            this.label6.Text = "Тип";
            // 
            // comboBoxClient
            // 
            this.comboBoxClient.FormattingEnabled = true;
            this.comboBoxClient.Location = new System.Drawing.Point(83, 22);
            this.comboBoxClient.Name = "comboBoxClient";
            this.comboBoxClient.Size = new System.Drawing.Size(121, 21);
            this.comboBoxClient.TabIndex = 14;
            this.comboBoxClient.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // comboBoxProduct
            // 
            this.comboBoxProduct.FormattingEnabled = true;
            this.comboBoxProduct.Location = new System.Drawing.Point(83, 49);
            this.comboBoxProduct.Name = "comboBoxProduct";
            this.comboBoxProduct.Size = new System.Drawing.Size(121, 21);
            this.comboBoxProduct.TabIndex = 15;
            this.comboBoxProduct.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            // 
            // comboBoxPayment
            // 
            this.comboBoxPayment.FormattingEnabled = true;
            this.comboBoxPayment.Location = new System.Drawing.Point(83, 76);
            this.comboBoxPayment.Name = "comboBoxPayment";
            this.comboBoxPayment.Size = new System.Drawing.Size(121, 21);
            this.comboBoxPayment.TabIndex = 16;
            this.comboBoxPayment.SelectedIndexChanged += new System.EventHandler(this.comboBox4_SelectedIndexChanged);
            // 
            // textBoxQuantity
            // 
            this.textBoxQuantity.Location = new System.Drawing.Point(83, 104);
            this.textBoxQuantity.Name = "textBoxQuantity";
            this.textBoxQuantity.Size = new System.Drawing.Size(121, 20);
            this.textBoxQuantity.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(31, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 18;
            this.label7.Text = "Клиент";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(36, 52);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(38, 13);
            this.label8.TabIndex = 19;
            this.label8.Text = "Товар";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 84);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(72, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "Тип платежа";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(3, 107);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(79, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "Кол-во товара";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(-7, 130);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(91, 13);
            this.label11.TabIndex = 22;
            this.label11.Text = "Сумма договора";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(36, 155);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(33, 13);
            this.label13.TabIndex = 24;
            this.label13.Text = "Дата";
            // 
            // dateTimePickerContractDate
            // 
            this.dateTimePickerContractDate.Location = new System.Drawing.Point(83, 149);
            this.dateTimePickerContractDate.Name = "dateTimePickerContractDate";
            this.dateTimePickerContractDate.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerContractDate.TabIndex = 12;
            // 
            // textBoxContractPrice
            // 
            this.textBoxContractPrice.Location = new System.Drawing.Point(83, 127);
            this.textBoxContractPrice.Name = "textBoxContractPrice";
            this.textBoxContractPrice.Size = new System.Drawing.Size(121, 20);
            this.textBoxContractPrice.TabIndex = 25;
            // 
            // dateTimePickerEndDate
            // 
            this.dateTimePickerEndDate.Location = new System.Drawing.Point(794, 159);
            this.dateTimePickerEndDate.Name = "dateTimePickerEndDate";
            this.dateTimePickerEndDate.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEndDate.TabIndex = 12;
            // 
            // checkedListBoxClients
            // 
            this.checkedListBoxClients.FormattingEnabled = true;
            this.checkedListBoxClients.Location = new System.Drawing.Point(794, 27);
            this.checkedListBoxClients.Name = "checkedListBoxClients";
            this.checkedListBoxClients.Size = new System.Drawing.Size(200, 94);
            this.checkedListBoxClients.TabIndex = 13;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(69, 9);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(165, 13);
            this.label12.TabIndex = 14;
            this.label12.Text = "Выбор таблицы для просмотра";
            this.label12.Click += new System.EventHandler(this.label12_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(75, 53);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(129, 13);
            this.label14.TabIndex = 15;
            this.label14.Text = "Редактирование таблиц";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(825, 9);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(122, 13);
            this.label15.TabIndex = 16;
            this.label15.Text = "Формирование отчета";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(734, 35);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(54, 13);
            this.label16.TabIndex = 17;
            this.label16.Text = "Клиенты:";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(765, 127);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(23, 13);
            this.label17.TabIndex = 18;
            this.label17.Text = "От:";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(765, 159);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(25, 13);
            this.label18.TabIndex = 19;
            this.label18.Text = "До:";
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(337, 14);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(162, 49);
            this.button14.TabIndex = 20;
            this.button14.Text = "REFRESH";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1049, 678);
            this.Controls.Add(this.button14);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.checkedListBoxClients);
            this.Controls.Add(this.dateTimePickerEndDate);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dateTimePickerStartDate);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePickerStartDate;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.TextBox nameText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox numberText;
        private System.Windows.Forms.TextBox lastNameText;
        private System.Windows.Forms.CheckBox checkBoxIsDelivered;
        private System.Windows.Forms.CheckBox checkBoxIsCargo;
        private System.Windows.Forms.TextBox textBoxProductPrice;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxProductName;
        private System.Windows.Forms.ComboBox comboBoxPayment;
        private System.Windows.Forms.ComboBox comboBoxProduct;
        private System.Windows.Forms.ComboBox comboBoxClient;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox payName;
        private System.Windows.Forms.DateTimePicker dateTimePickerContractDate;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBoxQuantity;
        private System.Windows.Forms.TextBox textBoxContractPrice;
        private System.Windows.Forms.DateTimePicker dateTimePickerEndDate;
        private System.Windows.Forms.CheckedListBox checkedListBoxClients;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Button button14;
    }
}

