#include "ExcelShell.h"

namespace xlsh
{

	HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
	{
		if (!pDisp) return E_FAIL;

		va_list marker;
		va_start(marker, cArgs);

		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		DISPID dispidNamed = DISPID_PROPERTYPUT;
		DISPID dispID;
		char szName[200];

		// Convert down to ANSI
		WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

		// Get DISPID for name passed...
		HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
		if (FAILED(hr)) {
			return hr;
		}
		// Allocate memory for arguments...
		VARIANT *pArgs = new VARIANT[cArgs + 1];
		// Extract arguments...
		for (int i = 0; i < cArgs; i++) {
			pArgs[i] = va_arg(marker, VARIANT);
		}

		// Build DISPPARAMS
		dp.cArgs = cArgs;
		dp.rgvarg = pArgs;

		// Handle special-case for property-puts!
		if (autoType & DISPATCH_PROPERTYPUT) {
			dp.cNamedArgs = 1;
			dp.rgdispidNamedArgs = &dispidNamed;
		}

		// Make the call!
		hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
		if (FAILED(hr)) {
			return hr;
		}

		// End variable-argument section...
		va_end(marker);

		delete[] pArgs;

		return hr;
	}

	HRESULT proc_beg(const std::wstring &path, xls_t * const xls, bool visible)
	{
		HRESULT hr;

		VARIANT x;
		VARIANT _path;

		VariantInit(&xls->app);
		VariantInit(&xls->wbs);
		VariantInit(&xls->wb);
		VariantInit(&_path);
		VariantInit(&xls->wss);
		VariantInit(&xls->ws);

		// Initialize COM for this thread
		if (S_OK != CoInitialize(NULL))
			std::cout << "initialization failed" << std::endl;
		else std::cout << "initialization succeed" << std::endl;

		// Get CLSID for our server
		CLSID clsid;
		hr = CLSIDFromProgID(L"Excel.Application", &clsid);
		if (FAILED(hr))
			std::cout << "Excel application opening failed" << std::endl;
		else std::cout << "Excel application succesfully opened" << std::endl;

		// Start server and get IDispatch
		xls->app.vt = VT_DISPATCH;
		xls->app.pdispVal = 0;
		hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&xls->app.pdispVal);
		if (FAILED(hr))
			std::cout << "create instance failed" << std::endl;
		else std::cout << "create instance succeed" << std::endl;

		// Make it visible/invisible 
		x.vt = VT_I4;
		x.boolVal = visible;
		AutoWrap(DISPATCH_PROPERTYPUT, NULL, xls->app.pdispVal, L"Visible", 1, x);
		if (visible)
			std::cout << "excel was opened visible" << std::endl;
		else std::cout << "excel was opened invisible" << std::endl;

		// Get Workbooks collection
		hr = AutoWrap(DISPATCH_PROPERTYGET, &xls->wbs, xls->app.pdispVal, L"Workbooks", 0);
		if (FAILED(hr))
			std::cout << "getting of workbooks collection failed" << std::endl;
		else std::cout << "workbooks collection get" << std::endl;

		if (path.size())
		{
			//	std::cout << "try to open document" << std::endl;
			_path.vt = VT_BSTR;
			//_path.bstrVal = SysAllocString(str_to_wstr(path).c_str());
			_path.bstrVal = SysAllocString(path.c_str());
			hr = AutoWrap(DISPATCH_METHOD, &xls->wb, xls->wbs.pdispVal, L"Open", 1, _path);
			if (FAILED(hr))
				std::cout << "opening document failed" << std::endl;
			else std::cout << "document successfully opened" << std::endl;
			VariantClear(&_path);
		}

		else
		{
			//	std::cout << "try to add document" << std::endl;
			hr = AutoWrap(DISPATCH_METHOD, &xls->wb, xls->wbs.pdispVal, L"Add", 0);
			if (FAILED(hr))
				std::cout << "adding document failed" << std::endl;
			else std::cout << "document successfully added" << std::endl;
		}

		hr = AutoWrap(DISPATCH_PROPERTYGET, &xls->wss, xls->wb.pdispVal, L"Worksheets", 0);
		if (FAILED(hr))
			std::cout << "getting of worksheets collection failed" << std::endl;
		else std::cout << "worksheets collection get" << std::endl;

		x.lVal = 1;
		hr = AutoWrap(DISPATCH_PROPERTYGET, &xls->ws, xls->wss.pdispVal, L"Item", 1, x);
		if (FAILED(hr))
			std::cout << "getting of items collection failed" << std::endl;
		else std::cout << "item collection get" << std::endl;

		return hr;
	}

	HRESULT proc_end(HRESULT hr, xls_t * const xls, bool save, bool close)
	{
		if (!FAILED(hr))
		{
			if (save)
			{
				hr = AutoWrap(DISPATCH_METHOD, NULL, xls->wb.pdispVal, L"Save", 0);
				if (FAILED(hr))
					std::cout << "saving document failed" << std::endl;
				else std::cout << "document saved" << std::endl;
			}

			if (close)
			{
				hr = AutoWrap(DISPATCH_METHOD, NULL, xls->wb.pdispVal, L"Close", 0);
				if (FAILED(hr))
					std::cout << "closing document failed" << std::endl;
				else std::cout << "document closed" << std::endl;

				hr = AutoWrap(DISPATCH_METHOD, NULL, xls->app.pdispVal, L"Quit", 0);
				if (FAILED(hr))
					std::cout << "quit failed" << std::endl;
				else std::cout << "quit from excel done\n" << std::endl;
			}


		}

		VariantClear(&xls->app);
		VariantClear(&xls->wbs);
		VariantClear(&xls->wb);
		VariantClear(&xls->wss);
		VariantClear(&xls->ws);

		return hr;
	}

	HRESULT activate_sheet(xls_t * const xls, int sheet_num)
	{
		HRESULT hr = NULL;
		VARIANT result;
		VariantInit(&result);

		VARIANT parm;
		parm.vt = VT_I4;
		parm.lVal = sheet_num;
		while (true)
		{
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &xls->ws, xls->wss.pdispVal, L"Item", 1, parm))
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_METHOD, &result, xls->ws.pdispVal, L"Activate", 0))
			}

			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}
			break;
		}
		return hr;
	}

	std::wstring str_to_wstr(const std::string &s, const unsigned cp)
	{
		std::wstring res;
		unsigned length =
			MultiByteToWideChar
			(
				cp, //CodePage
				0, //dwFlags
				s.c_str(), //lpMultiByteStr
				-1, //cchMultiByte
				0, //lpWideCharStr
				0 //cchWideChar
			);
		wchar_t *buffer = new wchar_t[length];
		if
			(
				MultiByteToWideChar
				(
					cp, //CodePage
					0, //dwFlags
					s.c_str(), //lpMultiByteStr
					-1, //cchMultiByte
					buffer, //lpWideCharStr
					length //cchWideChar
				)
				)
			res = buffer;
		delete[] buffer;
		return res;
	}

	std::string wstr_to_str(const std::wstring &s, const unsigned cp)
	{
		std::string res;
		unsigned length =
			WideCharToMultiByte
			(
				cp, //CodePage
				0, //dwFlags
				s.c_str(), //lpWideCharStr
				-1, //cchWideChar
				0, //lpMultiByteStr
				0, //cchMultiByte
				0, //lpDefaultChar
				0 //lpUsedDefaultChar
			);
		char *buffer = new char[length];
		if
			(
				WideCharToMultiByte
				(
					cp, //CodePage
					0, //dwFlags
					s.c_str(), //lpWideCharStr
					-1, //cchWideChar
					buffer, //lpMultiByteStr
					length, //cchMultiByte
					0, //lpDefaultChar
					0 //lpUsedDefaultChar
				)
				)
			res = buffer;
		delete[] buffer;
		return res;
	}

	HRESULT read(VARIANT ws, int _r, int _c, std::wstring * const _x)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT x;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&x);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))

				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &x, cell.pdispVal, L"Value", 0))
				if (x.vt == VT_BSTR)
				{
					*_x = x.bstrVal;
				}
				else
				{
					VARIANT tmp;
					VariantInit(&tmp);

					if (SUCCEEDED(VariantChangeType(&tmp, &x, 0, VT_BSTR)))
					{
						*_x = tmp.bstrVal;
						VariantClear(&tmp);
					}
					else
					{
						_x->clear();
					}
				}

			break;
		}

		VariantClear(&cell);
		VariantClear(&x);

		return hr;
	}

	int read_int(VARIANT ws, int row, int col)
	{
		std::wstring cell_wstr;
		int cell_int;
		read(ws, row, col, &cell_wstr);
		try
		{
			cell_int = std::stoi(cell_wstr);
		}
		catch (...)
		{
			cell_int = 0;
		}
		return cell_int;
	}

	HRESULT write(xls_t * const xls, int _r, int _c, int value)
	{
		VARIANT tmp;
		tmp.vt = VT_INT;
		tmp.intVal = value;
		return write_in_table(xls, _r, _c, &tmp);
	}

	HRESULT write(xls_t * const xls, int _r, int _c, float value)
	{
		VARIANT tmp;
		tmp.vt = VT_R4;
		tmp.fltVal = value;
		return write_in_table(xls, _r, _c, &tmp);
	}

	HRESULT write(xls_t * const xls, int _r, int _c, double value)
	{
		VARIANT tmp;
		tmp.vt = VT_R8;
		tmp.dblVal = value;
		return write_in_table(xls, _r, _c, &tmp);
	}

	HRESULT write(xls_t * const xls, int _r, int _c, char* char_str)
	{
		const size_t size = strlen(char_str) + 1;
		wchar_t* wchar_str = new wchar_t[size];
		mbstowcs(wchar_str, char_str, size);
		return write(xls, _r, _c, wchar_str);
	}

	HRESULT write(xls_t * const xls, int _r, int _c, std::string str)
	{
		return write(xls, _r, _c, std::wstring(str.begin(), str.end()));
	}

	HRESULT write(xls_t * const xls, int _r, int _c, std::wstring wstr)
	{
		wchar_t *wchar_str = (wchar_t *)wstr.c_str();
		return write(xls, _r, _c, wchar_str);
	}

	HRESULT write(xls_t * const xls, int _r, int _c, wchar_t* wchar_str)
	{
		VARIANT tmp;
		tmp.vt = VT_BSTR;
		tmp.bstrVal = ::SysAllocString(wchar_str);

		return write_in_table(xls, _r, _c, &tmp);
	}

	static HRESULT write_in_table(xls_t * const xls, int _r, int _c, VARIANT *value)
	{
		HRESULT hr = NULL;
		VariantInit(&xls->app);

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				std::wstring cell = get_cell(_r, _c);
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(cell.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, value))
			break;
		}
		return hr;
	}

	std::wstring get_cell(int r, int c)
	{
		std::wstring symb_for_excel[MAX_COLUMN + 1] = { L"", L"a", L"b", L"c", L"d", L"e", L"f", L"g", L"h", L"i", L"j", L"k", L"l", L"m", L"n", L"o", L"p", L"q", L"r", L"s", L"t", L"u", L"v", L"w", L"x", L"y", L"z", L"aa", L"ab", L"ac", L"ad", L"ae", L"af", L"ag", L"ah", L"ai", L"aj", L"ak", L"al", L"am", L"an" };

		std::wstring res = symb_for_excel[c];
		res += std::to_wstring(r);
		return res;
	}

	int get_font_color(VARIANT ws, int _r, int _c)
	{
		int color;
		HRESULT hr;
		hr = get_font_color(ws, _r, _c, &color);
		return color;
	}

	HRESULT get_font_color(VARIANT ws, int _r, int _c, int *color_value)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT color;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		VariantInit(&color);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &color, in.pdispVal, L"Color", 0))
				*color_value = color.dblVal;
			break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_font_color(VARIANT ws, int _r, int _c, const int color_value)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT color;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		color.vt = VT_I4;
		color.lVal = color_value;

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Color", 1, color))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_font_color_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, const int color_value)
	{
		std::wstring range = get_cell(r_since, c_since);
		range += L":";
		range += get_cell(r_before, c_before);

		HRESULT hr;

		VARIANT cell;
		VARIANT in;
		VARIANT color;
		VariantInit(&in);
		color.vt = VT_I4;
		color.lVal = color_value;

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(range.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, pXlRange, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Color", 1, color))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	int get_inter_color(VARIANT ws, int _r, int _c)
	{
		int color;
		HRESULT hr;
		hr = get_inter_color(ws, _r, _c, &color);
		return color;
	}

	HRESULT get_inter_color(VARIANT ws, int _r, int _c, int *color_value)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT color;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		VariantInit(&color);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Interior", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &color, in.pdispVal, L"Color", 0))
				*color_value = color.dblVal;
			break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_inter_color(VARIANT ws, int _r, int _c, const int color_value)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT color;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		color.vt = VT_I4;
		color.lVal = color_value;

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Interior", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Color", 1, color))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_inter_color_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, const int color_value)
	{
		std::wstring range = get_cell(r_since, c_since);
		range += L":";
		range += get_cell(r_before, c_before);

		HRESULT hr;

		VARIANT cell;
		VARIANT in;
		VARIANT color;
		VariantInit(&in);
		color.vt = VT_I4;
		color.lVal = color_value;

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(range.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, pXlRange, L"Interior", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Color", 1, color))
				break;
		}
		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	bool get_italic(VARIANT ws, int _r, int _c)
	{
		bool state;
		HRESULT hr;
		hr = get_italic(ws, _r, _c, &state);
		return state;
	}

	HRESULT get_italic(VARIANT ws, int _r, int _c, bool* state)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT italic;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		VariantInit(&italic);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &italic, in.pdispVal, L"Italic", 0))
				*state = italic.boolVal;
			break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_italic(VARIANT ws, int _r, int _c, bool state)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT italic;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		italic.vt = VT_I4;
		italic.lVal = state;

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Italic", 1, italic))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_italic_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, bool state)
	{
		std::wstring range = get_cell(r_since, c_since);
		range += L":";
		range += get_cell(r_before, c_before);

		HRESULT hr;

		VARIANT cell;
		VARIANT in;
		VARIANT italic_state;
		VariantInit(&in);
		italic_state.vt = VT_BOOL;
		italic_state.boolVal = state;

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(range.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, pXlRange, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Italic", 1, italic_state))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	bool get_bold(VARIANT ws, int _r, int _c)
	{
		bool state;
		HRESULT hr;
		hr = get_bold(ws, _r, _c, &state);
		return state;
	}

	HRESULT get_bold(VARIANT ws, int _r, int _c, bool* state)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT bold;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		VariantInit(&bold);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &bold, in.pdispVal, L"Bold", 0))
				*state = bold.boolVal;
			break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_bold(VARIANT ws, int _r, int _c, bool state)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT in;
		VARIANT bold;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&in);
		bold.vt = VT_I4;
		bold.lVal = state;

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, cell.pdispVal, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Bold", 1, bold))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	HRESULT set_bold_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, bool state)
	{
		std::wstring range = get_cell(r_since, c_since);
		range += L":";
		range += get_cell(r_before, c_before);

		HRESULT hr;

		VARIANT cell;
		VARIANT in;
		VARIANT bold_state;
		VariantInit(&in);
		bold_state.vt = VT_BOOL;
		bold_state.boolVal = state;

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(range.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &in, pXlRange, L"Font", 0))
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYPUT, 0, in.pdispVal, L"Bold", 1, bold_state))
				break;
		}

		VariantClear(&cell);
		VariantClear(&in);

		return hr;
	}

	void erase_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before)
	{
		std::wstring range = get_cell(r_since, c_since);
		range += L":";
		range += get_cell(r_before, c_before);

		HRESULT hr = NULL;

		VariantInit(&xls->app);
		VariantInit(&xls->wbs);
		VariantInit(&xls->wb);
		VariantInit(&xls->wss);
		VariantInit(&xls->ws);

		VARIANT arr;
		arr.vt = VT_ARRAY | VT_VARIANT;
		{
			SAFEARRAYBOUND sab[1];
			sab[0].lLbound = 1; sab[0].cElements = 1;
			arr.parray = SafeArrayCreate(VT_VARIANT, 1, sab);
		}

		while (true)
		{
			IDispatch *pXlSheet;
			{
				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, xls->app.pdispVal, L"ActiveSheet", 0))
					pXlSheet = result.pdispVal;
			}

			IDispatch *pXlRange;
			{
				VARIANT parm;
				parm.vt = VT_BSTR;
				parm.bstrVal = ::SysAllocString(range.c_str());

				VARIANT result;
				VariantInit(&result);
				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm))
					VariantClear(&parm);

				pXlRange = result.pdispVal;
			}

			BREAK_ON_FAIL(AutoWrap(DISPATCH_METHOD, NULL, pXlRange, L"ClearContents", 0))
				break;
		}
	}

	HRESULT read_formula(VARIANT ws, int _r, int _c, std::wstring * const _x)
	{
		HRESULT hr;

		VARIANT cell;
		VARIANT r;
		VARIANT c;
		VARIANT x;

		VariantInit(&cell);
		r.vt = VT_I4;
		c.vt = VT_I4;
		VariantInit(&x);

		while (true)
		{
			r.lVal = _r;
			c.lVal = _c;
			BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &cell, ws.pdispVal, L"Cells", 2, c, r))

				BREAK_ON_FAIL(AutoWrap(DISPATCH_PROPERTYGET, &x, cell.pdispVal, L"Formula", 0))
				if (x.vt == VT_BSTR)
				{
					*_x = x.bstrVal;
				}
				else
				{
					VARIANT tmp;
					VariantInit(&tmp);

					if (SUCCEEDED(VariantChangeType(&tmp, &x, 0, VT_BSTR)))
					{
						*_x = tmp.bstrVal;
						VariantClear(&tmp);
					}
					else
					{
						_x->clear();
					}
				}
			break;
		}

		VariantClear(&cell);
		VariantClear(&x);

		return hr;
	}

}