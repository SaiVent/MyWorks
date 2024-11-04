#include "MyForm.h" 
#include "MyForm2.h" 
#include "MyForm3.h"

using namespace System; 
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;



int tab1, tab2, tab3;

System::Void Project1::MyForm3::�����ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
}

System::Void Project1::MyForm3::����������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) // � ���������
{
    MessageBox::Show("��� ��������� ����� ���������� ����� ����� �����������.� ������� ���� �������������� ������� ������������ ����� ����� ��������� ����������� � ��������, �������, ��������, � ����� ������������ ��������� ��� ������.","��������!");
    return System::Void();
}

System::Void Project1::MyForm3::MyForm3_Load(System::Object^ sender, System::EventArgs^ e) //������ ���� ��������� � �������� ������� �������
{
    tab1 = 0;
    tab2 = 0;
    tab3 = 1;
    dataGridView1->Rows->Clear();
    //���������� � ��
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //��������� ������ � ��
    dbConnection->Open();
    String^ query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������],Payments.[��������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������])";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["����� ���������"], dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["��������"], dbReader["���������"]);
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

    Column0->HeaderText = "����� ���������";
    Column1->HeaderText = "������";
    Column2->HeaderText = "�������� ������";
    Column3->HeaderText = "���� ������";
    Column4->HeaderText = "��������";
    Column5->HeaderText = "���������";

    String^ testQuery = "SELECT [��� ������] FROM [Services]"; //���������� ����� �����
    OleDbCommand^ testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    OleDbDataReader^ testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox1->Items->Add(testReader["��� ������"]->ToString());
    }
    comboBox1->SelectedIndex = 0;
    testQuery = "SELECT [������� ����] FROM [Client]"; //���������� ����� �����
    testCommand = gcnew OleDbCommand(testQuery, dbConnection);
    testReader = testCommand->ExecuteReader();
    while (testReader->Read())
    {
        comboBox2->Items->Add(testReader["������� ����"]->ToString());
    }
    comboBox2->SelectedIndex = 0;
   
    testReader->Close();
    dbReader->Close();
    dbConnection->Close();
    
}

System::Void Project1::MyForm3::button1_Click_1(System::Object^ sender, System::EventArgs^ e) //�������� ������ � �������
{
    if (dataGridView1->SelectedRows->Count != 1) //�������� ������� � �����

    {
        MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
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
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ Cel4 = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "INSERT INTO [Client] VALUES ("+ Cel1 +", '"+ Cel2 +"','" + Cel3 + "'," + Cel4 + " )";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
   
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else
            MessageBox::Show("������ ���������!", "������!");

        dbConnection->Close();
    }

    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr )
        {
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "INSERT INTO [Services] VALUES (" + Cel1 + ", '" + Cel2 + "'," + Cel3 + ")";
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else
            MessageBox::Show("������ ���������!", "������!");

        dbConnection->Close();
    }
    if (tab3 == 1)
    {

    }
    return System::Void();
}

System::Void Project1::MyForm3::button2_Click(System::Object^ sender, System::EventArgs^ e) // �������� ������ � �������
{
    if (dataGridView1->SelectedRows->Count != 1) //�������� ������� � �����

    {
        MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
        return;
    }
    if (tab1 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr )
        {
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "DELETE FROM [Client] WHERE [������� ����] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else {
            MessageBox::Show("������ �������!", "������!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    }
    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
        {
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "DELETE FROM [Services] WHERE [��� ������] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else {
            MessageBox::Show("������ �������!", "������!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    }

    if (tab3 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;
        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
        {
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "DELETE FROM [Payments] WHERE [����� ���������] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else {
            MessageBox::Show("������ �������!", "������!");
            dataGridView1->Rows->RemoveAt(index);
        }
        dbConnection->Close();
    } 
    return System::Void();
}

System::Void Project1::MyForm3::button3_Click(System::Object^ sender, System::EventArgs^ e) // ���������� ������ � �������
{
    if (dataGridView1->SelectedRows->Count != 1) //�������� ������� � �����

    {
        MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
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
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        String^ Cel4 = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "UPDATE [Client] SET ������ = '" + Cel2 + "',����� = '" + Cel3 + "',[����� ��������] = " + Cel4 + " WHERE [������� ����] ="  + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else
            MessageBox::Show("������ ���������!", "������!");

        dbConnection->Close();
    }
    if (tab2 == 1)
    {
        int index = dataGridView1->SelectedRows[0]->Index;

        if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
            dataGridView1->Rows[index]->Cells[2]->Value == nullptr)
        {
            MessageBox::Show("�� ��� ������ �������!", "��������!");
            return;
        }
        String^ Cel1 = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
        String^ Cel2 = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
        String^ Cel3 = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
        //���������� � ��
        String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
        OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
        //��������� ������ � ��
        dbConnection->Open();
        String^ query = "UPDATE [Services] SET [�������� ������] = '" + Cel2 + "',[���� �� �������] = " + Cel3 + " WHERE [��� ������] =" + Cel1;
        OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

        if (dbComand->ExecuteNonQuery() != 1)
            MessageBox::Show("������ ���������� �������!", "������!");
        else
            MessageBox::Show("������ ����������!", "������!");

        dbConnection->Close();
    }
    if (tab3 == 1)
    {

    }
    return System::Void();
}

System::Void Project1::MyForm3::button4_Click(System::Object^ sender, System::EventArgs^ e) //�������� ������� ��������
{
    tab1 = 1;
    tab2 = 0;
    tab3 = 0;
    dataGridView1->Rows->Clear();
    //���������� � ��
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //��������� ������ � ��
    dbConnection->Open();
    String^ query = "SELECT * FROM [Client]";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["������� ����"], dbReader["������"], dbReader["�����"], dbReader["����� ��������"]);
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

    Column0->HeaderText = "������� ����";
    Column1->HeaderText = "������";
    Column2->HeaderText = "�����";
    Column3->HeaderText = "����� ��������";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button5_Click(System::Object^ sender, System::EventArgs^ e) //�������� ������� �����
{
    tab1 = 0;
    tab2 = 1;
    tab3 = 0;
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

    Column0->HeaderText = "��� ������";
    Column1->HeaderText = "�������� ������";
    Column2->HeaderText = "���� �� �������";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button6_Click(System::Object^ sender, System::EventArgs^ e) //�������� ������� ��������
{
    tab1 = 0;
    tab2 = 0;
    tab3 = 1;
    dataGridView1->Rows->Clear();
    //���������� � ��
    String^ connectionsString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.accdb";
    OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionsString);
    //��������� ������ � ��
    dbConnection->Open();
    String^ query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������],Payments.[��������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������])";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader();
    //�������� ������
    if (dbReader->HasRows == false) {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    else {
        while (dbReader->Read())
        {

            dataGridView1->Rows->Add(dbReader["����� ���������"], dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["��������"], dbReader["���������"]);
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
    
    Column0->HeaderText = "����� ���������";
    Column1->HeaderText = "������";
    Column2->HeaderText = "�������� ������";
    Column3->HeaderText = "���� ������";
    Column4->HeaderText = "��������";
    Column5->HeaderText = "���������";
    dbReader->Close();
    dbConnection->Close();
    return System::Void();
}

System::Void Project1::MyForm3::button7_Click(System::Object^ sender, System::EventArgs^ e) //�������� ���������
{
    
    if (String::IsNullOrWhiteSpace(comboBox2->Text))
    {
        MessageBox::Show("���� ��������� �� ����� ���� ������.", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }
    else
    {
        String^ client = comboBox2->Text;
        String^ service = comboBox1->Text;
        DateTime dateValue = dateTimePicker1->Value;
        dateValue = dateValue.AddDays(7); // 7 ���� � ��������� ����
        String^ date = dateValue.ToString("yyyy-MM-dd");
        String^ reading = textBox1->Text;
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

        if (tab3 == 1)
        {
        dataGridView1->Rows->Clear();
        query = "SELECT Payments.[����� ���������], Client.[������], Services.[�������� ������],Payments.[��������], Payments.[���� ������], Payments.[���������] FROM ((Payments INNER JOIN Client ON Payments.[������] = Client.[������� ����]) INNER JOIN Services ON Payments.[�������� ������] = Services.[��� ������])";
        dbComand = gcnew OleDbCommand(query, dbConnection);
        OleDbDataReader^ dbReader = dbComand->ExecuteReader();
        //�������� ������
        if (dbReader->HasRows == false) {
            MessageBox::Show("������ ����� ������!", "������!");
        }
        else {
            while (dbReader->Read())
            {

                dataGridView1->Rows->Add(dbReader["����� ���������"], dbReader["������"], dbReader["�������� ������"], dbReader["���� ������"], dbReader["��������"], dbReader["���������"]);
            }
        }
        Column0->HeaderText = "����� ���������";
        Column1->HeaderText = "������";
        Column2->HeaderText = "�������� ������";
        Column3->HeaderText = "���� ������";
        Column4->HeaderText = "��������";
        Column5->HeaderText = "���������";
        dbReader->Close();
        }


        dbConnection->Close();
    }
    return System::Void();
}

System::Void Project1::MyForm3::��������������������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //����� �� ������� ������
{
    MyForm3::Close();
    MyForm^ form = gcnew MyForm();
    form->Show();
}

System::Void Project1::MyForm3::����������������ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) //����� �� ���������
{
    Application::Exit();
}

System::Void Project1::MyForm3::button8_Click(System::Object^ sender, System::EventArgs^ e) //��������� ���������
{
    //�������� ����� � ����������, � ������ � ����� ������
    if (dataGridView1->SelectedRows->Count != 1)
    {
        MessageBox::Show("�������� ���� ������ ��� �������� ���������!", "��������!");
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
    //��������� ������ � ��
    dbConnection->Open();
    String^ query = "SELECT [���� �� �������] FROM [Services] WHERE  [�������� ������] = ""'" + service + "'";
    OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
    OleDbDataReader^ dbReader = dbComand->ExecuteReader(); 
    if (dbReader->Read()) {
        servicecost = Convert::ToInt32(dbReader[0]);
        
    }
    else {
        MessageBox::Show("������ ����� ������!", "������!");
    }
    dbReader->Close();
    dbConnection->Close();
    int servicepay = servicecost * Convert::ToInt32(reading);
    // �������� ������������ ��������
    System::Drawing::Bitmap^ bmp = gcnew System::Drawing::Bitmap("kvitansiya.png");
    System::Drawing::Graphics^ g = System::Drawing::Graphics::FromImage(bmp);
    System::Drawing::Font^ font = gcnew System::Drawing::Font("Aptos", 10);
    System::Drawing::SolidBrush^ brush = gcnew System::Drawing::SolidBrush(System::Drawing::Color::Black);
    g->DrawString(servicepay.ToString(), font, brush, 450, 203);
    g->DrawString(servicecost.ToString(), font, brush, 350,203);
    g->DrawString(id, font, brush, 150, 77);
    g->DrawString(client, font, brush, 75, 120);
    g->DrawString(service, font, brush, 30, 203);
    g->DrawString("�������� ��: " + date, font, brush, 25, 345);
    g->DrawString(reading, font, brush, 270, 203);
    delete g;
    // ��������� ���������� �������� �� ���������� ����
    SaveFileDialog^ saveFileDialog1 = gcnew SaveFileDialog();
    saveFileDialog1->Filter = "PNG files|*.png";
    saveFileDialog1->Title = "Save an Image File";
    saveFileDialog1->InitialDirectory = "C:\\";
    if (saveFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
    {
        String^ filePath = saveFileDialog1->FileName;
        bmp->Save(filePath, System::Drawing::Imaging::ImageFormat::Png);
        MessageBox::Show("��������� ������� ������� � ��������� � ����� " + filePath, "�����!");
    }
    delete bmp;
    return System::Void();
}
