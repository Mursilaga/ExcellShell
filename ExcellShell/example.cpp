
#include "ExcelShell.h"

int main()
{
	HRESULT hr;

	//start excel and open file
	std::wstring path = L"C:\\work\\table.xlsx";
	xls_t *xls = new xls_t;
	hr = proc_beg(path, xls);

	//write std::wstring in any cell
	hr = write(xls, 1, B_COLUMN, L"test");

	//set bold in B1 cell. For italic is the same 
	hr = set_bold(xls->ws, 1, B_COLUMN, true);
	//set bold in range of cells
	hr = set_bold_range(xls, 5, A_COLUMN, 10, B_COLUMN, true);
	//get bold state
	bool boldState;
	hr = get_bold(xls->ws, 1, B_COLUMN, &boldState);
	std::cout << "\nbold in B1 is " << boldState << std::endl;
	//or if you don't worry about hresult ret code
	boldState = get_bold(xls->ws, 5, B_COLUMN);
	std::cout << "bold in B5 is " << boldState << std::endl;
	

	//set interior color
	hr = set_inter_color(xls->ws, 1, B_COLUMN, GRAY);
	//set interior color in range of cells
	hr = set_inter_color_range(xls, 5, A_COLUMN, 10, B_COLUMN, LIGHT_BROWN);
	//get interior color
	int interiorColor;
	hr = get_inter_color(xls->ws, 1, B_COLUMN, &interiorColor);
	std::cout << "\ninterior color in B1 is " << interiorColor << std::endl;
	//or if you don't worry about hresult ret code
	interiorColor = get_inter_color(xls->ws, 5, A_COLUMN);
	std::cout << "interior color in A5 is " << interiorColor << std::endl;

	int a;
	std::cin >> a; //pause for check results, before excel process will finished
	//end proc, save and close document on default
	proc_end(hr, xls);

    return 0;
}

