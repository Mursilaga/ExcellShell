#pragma once

#include <string>
#include <locale>
#include <windows.h>
#include <algorithm> // all_of, copy, fill, find, for_each, none_of, remove, reverse, transform
#include <iostream> // istream, ostream
#include <sstream> // stringstream
#include <fstream>
#include <vector>

#define BREAK_ON_FAIL(x) if (FAILED(hr = x)) break;
#define MAX_COLUMN 40
#define VISIBLE 1 //app.visible = 1, invisible = 0;

namespace xlsh
{

	enum columns
	{
		A_COLUMN = 1, B_COLUMN, C_COLUMN, D_COLUMN, E_COLUMN,
		F_COLUMN, G_COLUMN, H_COLUMN, I_COLUMN, J_COLUMN,
		K_COLUMN, L_COLUMN, M_COLUMN, N_COLUMN, O_COLUMN,
		P_COLUMN, Q_COLUMN, R_COLUMN, S_COLUMN, T_COLUMN,
		U_COLUMN, V_COLUMN, W_COLUMN, X_COLUMN, Y_COLUMN,
		Z_COLUMN,
		AA_COLUMN, AB_COLUMN, AC_COLUMN, AD_COLUMN, AE_COLUMN,
		AF_COLUMN, AG_COLUMN, AH_COLUMN, AI_COLUMN, AJ_COLUMN,
		AK_COLUMN, AL_COLUMN, AM_COLUMN, AN_COLUMN, AO_COLUMN,
		AP_COLUMN, AQ_COLUMN, AR_COLUMN, AS_COLUMN, AT_COLUMN,
		AU_COLUMN, AV_COLUMN, AW_COLUMN, AX_COLUMN, AY_COLUMN,
		AZ_COLUMN
	};

	enum colors
	{
		BLACK = 0x0,
		WHITE = 0xFFFFFF,
		DARK_GREEN = 0x008000,
		BROWN = 0xC3C83,
		BLUE = 0xFF0000,
		GRAY = 14277081,
		RED = 0x0000FF,
		LIGHT_BROWN = 8696052,
		PINK = 13408767,
		CREAM = 0xADCBF8
	};

	struct xls_t
	{
		VARIANT app;
		VARIANT wbs;
		VARIANT wb;
		VARIANT wss;
		VARIANT ws;
	};

	HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...);

	std::wstring str_to_wstr(const std::string &s, const unsigned cp = CP_ACP);
	std::string wstr_to_str(const std::wstring &s, const unsigned cp = CP_ACP);

	HRESULT proc_beg(const std::wstring &path, xls_t * const xls, bool visible = true); //open if path.size() != 0 or add if path.size() == 0
	HRESULT proc_end(HRESULT hr, xls_t * const xls, bool save = true, bool close = true);

	HRESULT activate_sheet(xls_t * const xls, int sheetnum);

	HRESULT read(VARIANT ws, int row, int column, std::wstring * const _x);
	HRESULT read_formula(VARIANT ws, int row, int column, std::wstring * const _x);
	int read_int(VARIANT ws, int row, int column);

	HRESULT write(xls_t * const xls, int row, int column, int value);
	HRESULT write(xls_t * const xls, int row, int column, float value);
	HRESULT write(xls_t * const xls, int row, int column, double value);
	HRESULT write(xls_t * const xls, int row, int column, std::string str);
	HRESULT write(xls_t * const xls, int row, int column, std::wstring wstr);
	HRESULT write(xls_t * const xls, int row, int column, char* char_str);
	HRESULT write(xls_t * const xls, int row, int column, wchar_t* wchar_str);
	static HRESULT write_in_table(xls_t * const xls, int row, int column, VARIANT *value);

	int get_font_color(VARIANT ws, int row, int column);
	HRESULT get_font_color(VARIANT ws, int row, int column, int *color_value);
	HRESULT set_font_color(VARIANT ws, int row, int column, const int color_value);
	HRESULT set_font_color_range(xls_t * const xls, int row_since, int column_since, int row_before, int column_before, const int color_value);

	int get_inter_color(VARIANT ws, int row, int column);
	HRESULT get_inter_color(VARIANT ws, int row, int column, int *color_value);
	HRESULT set_inter_color(VARIANT ws, int row, int column, const int color_value);
	HRESULT set_inter_color_range(xls_t * const xls, int r_since, int c_since, int row_before, int column_before, const int color_value);

	bool get_italic(VARIANT ws, int row, int column);
	HRESULT get_italic(VARIANT ws, int row, int column, bool* state);
	HRESULT set_italic(VARIANT ws, int row, int column, bool state);
	HRESULT set_italic_range(xls_t * const xls, int row_since, int column_since, int row_before, int column_before, bool state);

	bool get_bold(VARIANT ws, int row, int column);
	HRESULT get_bold(VARIANT ws, int row, int column, bool* state);
	HRESULT set_bold(VARIANT ws, int row, int column, bool state);
	HRESULT set_bold_range(xls_t * const xls, int row_since, int column_since, int row_before, int column_before, bool state);

	std::wstring get_cell(int row, int column);

	void erase_range(xls_t * const xls, int row_since, int column_since, int row_before, int column_before);

}