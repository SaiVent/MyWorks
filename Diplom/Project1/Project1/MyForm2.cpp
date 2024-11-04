#include "MyForm.h" 
#include "MyForm2.h" 
#include "MyForm3.h" 


using namespace System;
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;



System::Void Project1::MyForm2::MyForm2_Load(System::Object^ sender, System::EventArgs^ e) // отображение таблицы платежи
{
    textBox1->Text = LogPol;
    dataGridView1->Rows->Clear();
    //Подключаем к БД
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ selectedClient;
    if (LogPol == "120")
    {
        selectedClient = "Склад Дзержинского";
    }
    if (LogPol == "123")
    {
        selectedClient = "Склад Ленина";
    }
    if (LogPol == "133")
    {
        selectedClient = "Склад Кутузовской";
    }
    if (LogPol == "134")
    {
        selectedClient = "Склад Суворовской";
    }
    //Выборка и вывод таблицы.[столбца] с показаниями клиента
    String^ query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги]) WHERE Client.[Клиент] = '" + selectedClient + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Показания"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = false;

    Column1->HeaderText = "Клиент";
    Column2->HeaderText = "Название услуги";
    Column3->HeaderText = "Дата оплаты";
    Column4->HeaderText = "Показания";
    String^ testQuery = "SELECT [Код услуги] FROM [Services]"; //заполнение ввода услуг
    OleDbCommand^ testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    OleDbDataReader^ testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox1->Items->Add(testReader["Код услуги"]->ToString());
    }
    comboBox1->SelectedIndex = 0;
    testReader->Close();
    dbReader->Close();
    dbConnection->Close();
	return System::Void();
}

System::Void Project1::MyForm2::выходИзПрограммыToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) // выход из программы
{
    Application::Exit();
}

System::Void Project1::MyForm2::выходИзУчетнойЗаписиToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //выход из учетной записи
{
    MyForm2::Close();
    MyForm^ form = gcnew MyForm();
    form->Show(); 
    return System::Void();
}

System::Void Project1::MyForm2::оПрограммеToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    MessageBox::Show("Это программа Учета комунально жилых услуг предприятия.С помощью этой информационной системы пользователи могут легко управлять информацией о клиентах, услугах, платежах, а также генерировать квитанции для оплаты.", "Внимание!");
    return System::Void();
}

System::Void Project1::MyForm2::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) 
{
}

System::Void Project1::MyForm2::button1_Click(System::Object^ sender, System::EventArgs^ e) //вывод платежей
{
    dataGridView1->Rows->Clear();
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    dbConnection->Open();
    //Выборка и вывод таблицы.[столбца] с показаниями клиента
    String^ selectedClient;
    if (LogPol == "120")
    {
        selectedClient = "Воинская Часть Дзержинского";
    }
    if (LogPol == "123")
    {
        selectedClient = "Склад Ленина";
    }
    if (LogPol == "133")
    {
        selectedClient = "Воинская Часть Кутузовской";
    }
    if (LogPol == "134")
    {
        selectedClient = "Склад Суворовской";
    }
    String^ query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги]) WHERE Client.[Клиент] = '" + selectedClient + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Показания"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = false;

    Column1->HeaderText = "Клиент";
    Column2->HeaderText = "Название услуги";
    Column3->HeaderText = "Дата оплаты";
    Column4->HeaderText = "Показания";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm2::button2_Click(System::Object^ sender, System::EventArgs^ e) //вывод услуг
{
    dataGridView1->Rows->Clear();
    //Подключаем к БД
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ query = "SELECT * FROM [Services]";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {
            dataGridView1->Rows->Add(dbReader["Код услуги"], dbReader["Название услуги"], dbReader["Цена за еденицу"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = false;
    Column5->Visible = false;

    Column1->HeaderText = "Код услуги";
    Column2->HeaderText = "Название услуги";
    Column3->HeaderText = "Цена за еденицу";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm2::button3_Click(System::Object^ sender, System::EventArgs^ e) // ввод показаний 
{
    if (String::IsNullOrWhiteSpace(textBox2->Text))
    {
        MessageBox::Show("Поле показания не может быть пустым.", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
    else
    {
        String^ client = textBox1->Text;
        String^ service = comboBox1->Text;
        DateTime dateValue = dateTimePicker1->Value;
        dateValue = dateValue.AddDays(7); // 7 дней к выбранной дате
        String^ date = dateValue.ToString("yyyy-MM-dd");
        String^ reading = textBox2->Text;
        String^ description = "нет";
        // Добавление новой записи в DataGridView
        dataGridView1->Rows->Add(client, service, date, reading);
        // Вставка новой записи в базу данных
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        dbConnection->Open();
        String^ query = "INSERT INTO Payments ([Клиент], [Название услуги], [Дата оплаты], [Показания], [Описание]) VALUES ('" + client + "', '" + service + "', #" + date + "#, '" + reading + "', '" + description + "')";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);


        if (dbComand->ExecuteNonQuery() != 1) {
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        }
        else {
            MessageBox::Show("Данные добавлены!", "Успех!");
        }
        
        textBox1->Text = LogPol;
        dataGridView1->Rows->Clear();
        String^ selectedClient;
        if (LogPol == "120")
        {
            selectedClient = "Склад Дзержинского";
        }
        if (LogPol == "123")
        {
            selectedClient = "Склад Ленина";
        }
        if (LogPol == "133")
        {
            selectedClient = "Склад Кутузовской";
        }
        if (LogPol == "134")
        {
            selectedClient = "Склад Суворовской";
        }
        Column1->HeaderText = "Клиент";
        Column2->HeaderText = "Название услуги";
        Column3->HeaderText = "Дата оплаты";
        Column4->HeaderText = "Показания";
        query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги]) WHERE Client.[Клиент] = '" + selectedClient + "'";
        dbComand = gcnew OleDbCommand(query, dbConnection);
        OleDbDataReader^ dbReader = dbComand->ExecuteReader();
        //проверка данных
        if (dbReader->HasRows == false) {
            MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
        }
        else {
            while (dbReader->Read())
            {

                dataGridView1->Rows->Add(dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Показания"]);
            }
        }
        dbReader->Close();
        dbConnection->Close();
    }
      return System::Void();
}
