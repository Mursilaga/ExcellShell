
#include "ExcelShell.h"

int main()
{
	HRESULT hr;

	//start excel and open file
	std::wstring path = L"C:\\work\\table.xlsx";
	xlsh::xls_t *xls = new xlsh::xls_t;
	hr = proc_beg(path, xls);

	//write in any cell
	char * char_str = "char";
	wchar_t *wchar_str = L"wchar_t";
	std::string str = "string";
	std::wstring wstr = L"wstring";
	double dbl_value = 3.14159;

	hr = write(xls, 1, xlsh::B_COLUMN, char_str);
	hr = write(xls, 1, xlsh::C_COLUMN, wchar_str);
	hr = write(xls, 1, xlsh::D_COLUMN, str);
	hr = write(xls, 1, xlsh::E_COLUMN, wstr);
	hr = write(xls, 1, xlsh::F_COLUMN, dbl_value);

	//set bold in B1 cell. For italic is similar 
	hr = xlsh::set_bold(xls->ws, 1, xlsh::B_COLUMN, true);
	//set bold in range of cells
	hr = set_bold_range(xls, 5, xlsh::A_COLUMN, 10, xlsh::B_COLUMN, true);
	//get bold state
	bool boldState;
	hr = xlsh::get_bold(xls->ws, 1, xlsh::B_COLUMN, &boldState);
	std::cout << "\nbold in B1 is " << boldState << std::endl;
	//or if you don't worry about hresult ret code
	boldState = xlsh::get_bold(xls->ws, 5, xlsh::B_COLUMN);
	std::cout << "bold in B5 is " << boldState << std::endl;
	

	//set interior color. For inter color is similar
	hr = xlsh::set_inter_color(xls->ws, 1, xlsh::B_COLUMN, xlsh::GRAY);
	//set interior color in range of cells
	hr = set_inter_color_range(xls, 5, xlsh::A_COLUMN, 10, xlsh::B_COLUMN, xlsh::LIGHT_BROWN);
	//get interior color
	int interiorColor;
	hr = xlsh::get_inter_color(xls->ws, 1, xlsh::B_COLUMN, &interiorColor);
	std::cout << "\ninterior color in B1 is " << interiorColor << std::endl;
	//or if you don't worry about hresult ret code
	interiorColor = xlsh::get_inter_color(xls->ws, 5, xlsh::A_COLUMN);
	std::cout << "interior color in A5 is " << interiorColor << std::endl;

	int a;
	std::cin >> a; //pause for check results, before excel process will finished
	//end proc, save and close document on default
	xlsh::proc_end(hr, xls);

    return 0;
}

