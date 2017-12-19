#pragma once
// Включение SDKDDKVer.h обеспечивает определение самой последней доступной платформы Windows.

// Если требуется выполнить сборку приложения для предыдущей версии Windows, включите WinSDKVer.h и
// задайте для макроса _WIN32_WINNT значение поддерживаемой платформы перед включением SDKDDKVer.h.
#include <SDKDDKVer.h>

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

#define A_COLUMN 1
#define B_COLUMN 2
#define C_COLUMN 3
#define D_COLUMN 4
#define E_COLUMN 5
#define F_COLUMN 6
#define G_COLUMN 7
#define H_COLUMN 8
#define I_COLUMN 9
#define J_COLUMN 10
#define L_COLUMN 12
#define N_COLUMN 14
#define P_COLUMN 16
#define Q_COLUMN 17
#define R_COLUMN 18
#define S_COLUMN 19
#define T_COLUMN 20
#define U_COLUMN 21
#define V_COLUMN 22
#define W_COLUMN 23
#define X_COLUMN 24
#define Y_COLUMN 25
#define Z_COLUMN 26
#define AA_COLUMN 27
#define AB_COLUMN 28
#define AC_COLUMN 29
#define AD_COLUMN 30
#define AF_COLUMN 32
#define AG_COLUMN 33
#define AH_COLUMN 34
#define AI_COLUMN 35
#define AM_COLUMN 39

#define BLACK 0x0
#define WHITE 0xFFFFFF
#define DARK_GREEN 0x008000
#define BROWN 0xC3C83
#define BLUE 0xFF0000
#define GRAY 14277081
#define RED 0x0000FF
#define LIGHT_BROWN 8696052
#define PINK 13408767
#define CREAM 0xADCBF8

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

HRESULT proc_beg(const std::wstring &path, xls_t * const xls, bool visible = false); //open if path.size() != 0 or add if path.size() == 0
HRESULT proc_end(HRESULT hr, xls_t * const xls, bool save = true, bool close = true);

HRESULT activate_sheet(xls_t * const xls, int sheetnum);

HRESULT read(VARIANT ws, int _r, int _c, std::wstring * const _x);
int read_int(VARIANT ws, int _r, int _c);
HRESULT write(xls_t * const xls, int _r, int _c, std::wstring wstr);

HRESULT set_color(VARIANT ws, int _r, int _c, const int x);
HRESULT set_font_color_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, const int x);
HRESULT set_inter_color_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before, const int x);
HRESULT get_inter_color(VARIANT ws, int _r, int _c, int *x);
int get_inter_color(VARIANT ws, int _r, int _c);

HRESULT set_bold_range(xls_t * const xls, bool state, int r_since, int c_since, int r_before, int c_before);

std::wstring get_cell(int r, int c);

HRESULT set_font_color(VARIANT ws, int _r, int _c, const int x);
HRESULT get_font_color(VARIANT ws, int _r, int _c, int *x);
int get_font_color(VARIANT ws, int _r, int _c);

bool get_italic(VARIANT ws, int _r, int _c);
bool get_bold(VARIANT ws, int _r, int _c);

void erase_range(xls_t * const xls, int r_since, int c_since, int r_before, int c_before);

HRESULT read_formula(VARIANT ws, int _r, int _c, std::wstring * const _x);