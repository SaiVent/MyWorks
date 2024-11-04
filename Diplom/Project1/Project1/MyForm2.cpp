#include "MyForm.h" 
#include "MyForm2.h" 
#include "MyForm3.h" 


using namespace System;
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;



System::Void Project1::MyForm2::MyForm2_Load(System::Object^ sender, System::EventArgs^ e) // ����������� ������� �������
{
    textBox1->Text = LogPol;
    dataGridView1->Rows->Clear();
    //���������� � ��
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //��������� ������ � ��
    dbConnection->Open();
    String^ selectedClient;
    if (LogPol == "120")
    {
        selectedClient = "����� ������������";
    }
    if (LogPol == "123")
    {
        selectedClient = "����� ������";
    }
    if (LogPol == "133")
    {
        selectedClient = "����� �����������";
    }
    if (LogPol == "134")
    {
        selectedClient = "����� �����������";
    }
    //������� � ����� �������.[�������] � ����������� �������
    String^ query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������]) WHERE Client.[������] = '" + selectedClient + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["���������"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = false;

    Column1->HeaderText = "������";
    Column2->HeaderText = "�������� ������";
    Column3->HeaderText = "���� ������";
    Column4->HeaderText = "���������";
    String^ testQuery = "SELECT [��� ������] FROM [Services]"; //���������� ����� �����
    OleDbCommand^ testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    OleDbDataReader^ testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox1->Items->Add(testReader["��� ������"]->ToString());
    }
    comboBox1->SelectedIndex = 0;
    testReader->Close();
    dbReader->Close();
    dbConnection->Close();
	return System::Void();
}

System::Void Project1::MyForm2::����������������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) // ����� �� ���������
{
    Application::Exit();
}

System::Void Project1::MyForm2::��������������������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //����� �� ������� ������
{
    MyForm2::Close();
    MyForm^ form = gcnew MyForm();
    form->Show(); 
    return System::Void();
}

System::Void Project1::MyForm2::����������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    MessageBox::Show("��� ��������� ����� ���������� ����� ����� �����������.� ������� ���� �������������� ������� ������������ ����� ����� ��������� ����������� � ��������, �������, ��������, � ����� ������������ ��������� ��� ������.", "��������!");
    return System::Void();
}

System::Void Project1::MyForm2::�����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) 
{
}

System::Void Project1::MyForm2::button1_Click(System::Object^ sender, System::EventArgs^ e) //����� ��������
{
    dataGridView1->Rows->Clear();
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    dbConnection->Open();
    //������� � ����� �������.[�������] � ����������� �������
    String^ selectedClient;
    if (LogPol == "120")
    {
        selectedClient = "�������� ����� ������������";
    }
    if (LogPol == "123")
    {
        selectedClient = "����� ������";
    }
    if (LogPol == "133")
    {
        selectedClient = "�������� ����� �����������";
    }
    if (LogPol == "134")
    {
        selectedClient = "����� �����������";
    }
    String^ query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������]) WHERE Client.[������] = '" + selectedClient + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["���������"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = true;
    Column5->Visible = false;

    Column1->HeaderText = "������";
    Column2->HeaderText = "�������� ������";
    Column3->HeaderText = "���� ������";
    Column4->HeaderText = "���������";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm2::button2_Click(System::Object^ sender, System::EventArgs^ e) //����� �����
{
    dataGridView1->Rows->Clear();
    //���������� � ��
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //��������� ������ � ��
    dbConnection->Open();
    String^ query = "SELECT * FROM [Services]";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {
            dataGridView1->Rows->Add(dbReader["��� ������"], dbReader["�������� ������"], dbReader["���� �� �������"]);
        }
    }
    Column1->Visible = true;
    Column2->Visible = true;
    Column3->Visible = true;
    Column4->Visible = false;
    Column5->Visible = false;

    Column1->HeaderText = "��� ������";
    Column2->HeaderText = "�������� ������";
    Column3->HeaderText = "���� �� �������";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm2::button3_Click(System::Object^ sender, System::EventArgs^ e) // ���� ��������� 
{
    if (String::IsNullOrWhiteSpace(textBox2->Text))
    {
        MessageBox::Show("���� ��������� �� ����� ���� ������.", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
    else
    {
        String^ client = textBox1->Text;
        String^ service = comboBox1->Text;
        DateTime dateValue = dateTimePicker1->Value;
        dateValue = dateValue.AddDays(7); // 7 ���� � ��������� ����
        String^ date = dateValue.ToString("yyyy-MM-dd");
        String^ reading = textBox2->Text;
        String^ description = "���";
        // ���������� ����� ������ � DataGridView
        dataGridView1->Rows->Add(client, service, date, reading);
        // ������� ����� ������ � ���� ������
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        dbConnection->Open();
        String^ query = "INSERT INTO Payments ([������], [�������� ������], [���� ������], [���������], [��������]) VALUES ('" + client + "', '" + service + "', #" + date + "#, '" + reading + "', '" + description + "')";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);


        if (dbComand->ExecuteNonQuery() != 1) {
            MessageBox::Show("������ ���������� �������!", "������!");
        }
        else {
            MessageBox::Show("������ ���������!", "�����!");
        }
        
        textBox1->Text = LogPol;
        dataGridView1->Rows->Clear();
        String^ selectedClient;
        if (LogPol == "120")
        {
            selectedClient = "����� ������������";
        }
        if (LogPol == "123")
        {
            selectedClient = "����� ������";
        }
        if (LogPol == "133")
        {
            selectedClient = "����� �����������";
        }
        if (LogPol == "134")
        {
            selectedClient = "����� �����������";
        }
        Column1->HeaderText = "������";
        Column2->HeaderText = "�������� ������";
        Column3->HeaderText = "���� ������";
        Column4->HeaderText = "���������";
        query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������]) WHERE Client.[������] = '" + selectedClient + "'";
        dbComand = gcnew OleDbCommand(query, dbConnection);
        OleDbDataReader^ dbReader = dbComand->ExecuteReader();
        //�������� ������
        if (dbReader->HasRows == false) {
            MessageBox::Show("������ ����� ������!", "������!");
        }
        else {
            while (dbReader->Read())
            {

                dataGridView1->Rows->Add(dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["���������"]);
            }
        }
        dbReader->Close();
        dbConnection->Close();
    }
      return System::Void();
}
