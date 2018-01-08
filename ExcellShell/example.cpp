
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

	//set italic in B1 cell
	hr = set_italic(xls->ws, 1, B_COLUMN, true);

	//set italic in range of cells
	hr = set_italic_range(xls, 5, A_COLUMN, 10, B_COLUMN, true);

	//get italic state
	bool italicState;
	hr = get_italic(xls->ws, 1, B_COLUMN, &italicState);
	std::cout << "italic in B1 is " << italicState << std::endl;

	//or if you don't worry about hresult ret code
	italicState = get_italic(xls->ws, 5, B_COLUMN);
	std::cout << "italic in B1 is " << italicState << std::endl;
	
	int a;
	std::cin >> a;
	//end proc, save and close document on default
	proc_end(hr, xls);

    return 0;
}

