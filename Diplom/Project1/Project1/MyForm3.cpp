#include "MyForm.h" 
#include "MyForm2.h" 
#include "MyForm3.h"

using namespace System; 
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;



int tab1, tab2, tab3;

System::Void Project1::MyForm3::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
}

System::Void Project1::MyForm3::оПрограммеToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) // о программе
{
    MessageBox::Show("Это программа Учета комунально жилых услуг предприятия.С помощью этой информационной системы пользователи могут легко управлять информацией о клиентах, услугах, платежах, а также генерировать квитанции для оплаты.","Внимание!");
    return System::Void();
}

System::Void Project1::MyForm3::MyForm3_Load(System::Object^ sender, System::EventArgs^ e) //запуск окна работника и загрузка таблицы платежи
{
    tab1 = 0;
    tab2 = 0;
    tab3 = 1;
    dataGridView1->Rows->Clear();
    //Подключаем к БД
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги],Payments.[Описание], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги])";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["Номер показаний"], dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Описание"], dbReader["Показания"]);
        }
    }
    Column0->Visible = true;
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = true;
    Column0->ReadOnly = true;
    Column1->ReadOnly = true;
    Column2->ReadOnly = true;
    Column3->ReadOnly = true;
    Column4->ReadOnly = true;
    Column5->ReadOnly = true;
    button1->Enabled = false;

    Column0->HeaderText = "Номер показаний";
    Column1->HeaderText = "Клиент";
    Column2->HeaderText = "Название услуги";
    Column3->HeaderText = "Дата оплаты";
    Column4->HeaderText = "Описание";
    Column5->HeaderText = "Показания";

    String^ testQuery = "SELECT [Код услуги] FROM [Services]"; //заполнение ввода услуг
    OleDbCommand^ testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    OleDbDataReader^ testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox1->Items->Add(testReader["Код услуги"]->ToString());
    }
    comboBox1->SelectedIndex = 0;
    testQuery = "SELECT [Лицевой счет] FROM [Client]"; //заполнение ввода услуг
    testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox2->Items->Add(testReader["Лицевой счет"]->ToString());
    }
    comboBox2->SelectedIndex = 0;
   
    testReader->Close();
    dbReader->Close();
    dbConnection->Close();
    
}

System::Void Project1::MyForm3::button1_Click_1(System::Object^ sender, System::EventArgs^ e) //Загрузка данных в таблицу
{
    if (dataGridView1->SelectedRows->Count != 1) //проверка выборки и ввода

    {
        MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
        return;
    }
    if (tab1 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr || 
            dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ Cel4 = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "INSERT INTO [Client] VALUES ("+ Cel1 +", '"+ Cel2 +"','" + Cel3 + "'," + Cel4 + " )";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
   
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else
            MessageBox::Show("Данные добавлены!", "Ошибка!");

        dbConnection->Close();
    }

    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr )
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "INSERT INTO [Services] VALUES (" + Cel1 + ", '" + Cel2 + "'," + Cel3 + ")";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else
            MessageBox::Show("Данные добавлены!", "Ошибка!");

        dbConnection->Close();
    }
    if (tab3 == 1)
    {

    }
    return System::Void();
}

System::Void Project1::MyForm3::button2_Click(System::Object^ sender, System::EventArgs^ e) // удаление данных в таблице
{
    if (dataGridView1->SelectedRows->Count != 1) //проверка выборки и ввода

    {
        MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
        return;
    }
    if (tab1 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr )
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "DELETE FROM [Client] WHERE [Лицевой счет] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else {
            MessageBox::Show("Данные удалены!", "Ошибка!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    }
    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "DELETE FROM [Services] WHERE [Код услуги] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else {
            MessageBox::Show("Данные удалены!", "Ошибка!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    }

    if (tab3 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "DELETE FROM [Payments] WHERE [Номер показаний] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else {
            MessageBox::Show("Данные удалены!", "Ошибка!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    } 
    return System::Void();
}

System::Void Project1::MyForm3::button3_Click(System::Object^ sender, System::EventArgs^ e) // Обновление данных в таблице
{
    if (dataGridView1->SelectedRows->Count != 1) //проверка выборки и ввода

    {
        MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
        return;
    }
    if (tab1 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ Cel4 = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "UPDATE [Client] SET Клиент = '" + Cel2 + "',Адрес = '" + Cel3 + "',[Номер телефона] = " + Cel4 + " WHERE [Лицевой счет] ="  + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else
            MessageBox::Show("Данные измененны!", "Ошибка!");

        dbConnection->Close();
    }
    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr)
        {
            MessageBox::Show("Не все данные введены!", "Внимание!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        //Подключаем к БД
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //Выполняем запрос к БД
        dbConnection->Open();
        String^ query = "UPDATE [Services] SET [Название услуги] = '" + Cel2 + "',[Цена за еденицу] = " + Cel3 + " WHERE [Код услуги] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
        else
            MessageBox::Show("Данные обновленны!", "Ошибка!");

        dbConnection->Close();
    }
    if (tab3 == 1)
    {

    }
    return System::Void();
}

System::Void Project1::MyForm3::button4_Click(System::Object^ sender, System::EventArgs^ e) //Загрузка таблицы клиентов
{
    tab1 = 1;
    tab2 = 0;
    tab3 = 0;
    dataGridView1->Rows->Clear();
    //Подключаем к БД
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ query = "SELECT * FROM [Client]";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["Лицевой счет"], dbReader["Клиент"], dbReader["Адрес"], dbReader["Номер телефона"]);
        }
    }
    Column0->Visible = true;
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = false;
    Column5->Visible = false;
    Column0->ReadOnly = false;
    Column1->ReadOnly = false;
    Column2->ReadOnly = false;
    Column3->ReadOnly = false;
    Column4->ReadOnly = false;
    Column5->ReadOnly = false;
    button1->Enabled = true;
    button8->Enabled = false;

    Column0->HeaderText = "Лицевой счет";
    Column1->HeaderText = "Клиент";
    Column2->HeaderText = "Адрес";
    Column3->HeaderText = "Номер телефона";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button5_Click(System::Object^ sender, System::EventArgs^ e) //Загрузка таблицы услуг
{
    tab1 = 0;
    tab2 = 1;
    tab3 = 0;
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
    Column0->Visible = true;
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = false;
    Column4->Visible = false;
    Column5->Visible = false;
    Column0->ReadOnly = false;
    Column1->ReadOnly = false;
    Column2->ReadOnly = false;
    Column3->ReadOnly = false;
    Column4->ReadOnly = false;
    Column5->ReadOnly = false;
    button1->Enabled = true;
    button8->Enabled = false;

    Column0->HeaderText = "Код услуги";
    Column1->HeaderText = "Название услуги";
    Column2->HeaderText = "Цена за еденицу";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button6_Click(System::Object^ sender, System::EventArgs^ e) //Загрузка таблицы платежей
{
    tab1 = 0;
    tab2 = 0;
    tab3 = 1;
    dataGridView1->Rows->Clear();
    //Подключаем к БД
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги],Payments.[Описание], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги])";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //проверка данных
    if (dbReader->HasRows == false) {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["Номер показаний"], dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Описание"], dbReader["Показания"]);
        }
    }
    Column0->Visible = true;
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = true;
    Column0->ReadOnly = true;
    Column1->ReadOnly = true;
    Column2->ReadOnly = true;
    Column3->ReadOnly = true;
    Column4->ReadOnly = true;
    Column5->ReadOnly = true;
    button1->Enabled = false;
    button8->Enabled = true;
    
    Column0->HeaderText = "Номер показаний";
    Column1->HeaderText = "Клиент";
    Column2->HeaderText = "Название услуги";
    Column3->HeaderText = "Дата оплаты";
    Column4->HeaderText = "Описание";
    Column5->HeaderText = "Показания";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button7_Click(System::Object^ sender, System::EventArgs^ e) //внесение показаний
{
    
    if (String::IsNullOrWhiteSpace(comboBox2->Text))
    {
        MessageBox::Show("Поле показания не может быть пустым.", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
    else
    {
        String^ client = comboBox2->Text;
        String^ service = comboBox1->Text;
        DateTime dateValue = dateTimePicker1->Value;
        dateValue = dateValue.AddDays(7); // 7 дней к выбранной дате
        String^ date = dateValue.ToString("yyyy-MM-dd");
        String^ reading = textBox1->Text;
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

        if (tab3 == 1)
        {
        dataGridView1->Rows->Clear();
        query = "SELECT Payments.[Номер показаний], Client.[Клиент], Services.[Название услуги],Payments.[Описание], Payments.[Дата оплаты], Payments.[Показания] FROM ((Payments INNER JOIN Client ON Payments.[Клиент] = Client.[Лицевой счет]) INNER JOIN Services ON Payments.[Название услуги] = Services.[Код услуги])";
        dbComand = gcnew OleDbCommand(query, dbConnection);
        OleDbDataReader^ dbReader = dbComand->ExecuteReader();
        //проверка данных
        if (dbReader->HasRows == false) {
            MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
        }
        else {
            while (dbReader->Read())
            {

                dataGridView1->Rows->Add(dbReader["Номер показаний"], dbReader["Клиент"], dbReader["Название услуги"], dbReader["Дата оплаты"], dbReader["Описание"], dbReader["Показания"]);
            }
        }
        Column0->HeaderText = "Номер показаний";
        Column1->HeaderText = "Клиент";
        Column2->HeaderText = "Название услуги";
        Column3->HeaderText = "Дата оплаты";
        Column4->HeaderText = "Описание";
        Column5->HeaderText = "Показания";
        dbReader->Close();
        }


        dbConnection->Close();
    }
    return System::Void();
}

System::Void Project1::MyForm3::выходИзУчетнойЗаписиToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //выход из учетной записи
{
    MyForm3::Close();
    MyForm^ form = gcnew MyForm();
    form->Show();
}

System::Void Project1::MyForm3::выходИзПрограммыToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //выход из программы
{
    Application::Exit();
}

System::Void Project1::MyForm3::button8_Click(System::Object^ sender, System::EventArgs^ e) //заполение квитанции
{
    //Загрузка строк в переменные, и расчет и выбор услуги
    if (dataGridView1->SelectedRows->Count != 1)
    {
        MessageBox::Show("Выберите одну строку для создания квитанции!", "Внимание!");
        return;
    }
    int index = dataGridView1->SelectedRows[0]->Index;
    int servicecost;
    int readingint;
    String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
    String^ client = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
    String^ service = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
    String^ date = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
    String^ description = dataGridView1->Rows[index]->Cells[4]->Value->ToString();
    String^ reading = dataGridView1->Rows[index]->Cells[5]->Value->ToString();
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //Выполняем запрос к БД
    dbConnection->Open();
    String^ query = "SELECT [Цена за еденицу] FROM [Services] WHERE  [Название услуги] = ""'" + service + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader(); 
    if (dbReader->Read()) {
        servicecost = Convert::ToInt32(dbReader[0]);
        
    }
    else {
        MessageBox::Show("Ошибка ввода данных!", "Ошибка!");
    }
    dbReader->Close();
    dbConnection->Close();
    int servicepay = servicecost * Convert::ToInt32(reading);
    // Загрузка существующей картинки
    System::Drawing::Bitmap^ bmp = gcnew System::Drawing::Bitmap("kvitansiya.png");
    System::Drawing::Graphics^ g = System::Drawing::Graphics::FromImage(bmp);
    System::Drawing::Font^ font = gcnew System::Drawing::Font("Aptos", 10);
    System::Drawing::SolidBrush^ brush = gcnew System::Drawing::SolidBrush(System::Drawing::Color::Black);
    g->DrawString(servicepay.ToString(), font, brush, 450, 203);
    g->DrawString(servicecost.ToString(), font, brush, 350,203);
    g->DrawString(id, font, brush, 150, 77);
    g->DrawString(client, font, brush, 75, 120);
    g->DrawString(service, font, brush, 30, 203);
    g->DrawString("Оплатить до: " + date, font, brush, 25, 345);
    g->DrawString(reading, font, brush, 270, 203);
    delete g;
    // Сохраняем измененную картинку по выбранному пути
    SaveFileDialog^ saveFileDialog1 = gcnew SaveFileDialog();
    saveFileDialog1->Filter = "PNG files|*.png";
    saveFileDialog1->Title = "Save an Image File";
    saveFileDialog1->InitialDirectory = "C:\\";
    if (saveFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
    {
        String^ filePath = saveFileDialog1->FileName;
        bmp->Save(filePath, System::Drawing::Imaging::ImageFormat::Png);
        MessageBox::Show("Квитанция успешно создана и сохранена в файле " + filePath, "Успех!");
    }
    delete bmp;
    return System::Void();
}
