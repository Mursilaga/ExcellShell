
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


	//end proc, save and close document on default
	proc_end(hr, xls);

    return 0;
}

