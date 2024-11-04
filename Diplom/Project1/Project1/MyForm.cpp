#include "MyForm.h" 
#include "MyForm2.h" 
#include "MyForm3.h"

using namespace System; //Подключение простраства имен.
using namespace System::Windows::Forms;

[STAThread] //Включение программы
int main(array<String^>^ arg) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);

	Project1::MyForm Form;
	Application::Run(% Form);
}