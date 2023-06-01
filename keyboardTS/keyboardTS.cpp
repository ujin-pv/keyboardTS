#include "framework.h"
#include "keyboardTS.h"
#include <vector>
#include <sstream>
#include <string>
#include "libxl.h"

#define MAX_LOADSTRING 100

// Глобальные переменные:
HINSTANCE hInst;                                
WCHAR szTitle[MAX_LOADSTRING];                  
WCHAR szWindowClass[MAX_LOADSTRING];            
HWND hInput;
HWND hBtn, hBtnCl;
HWND hLabelVK, hLabelScan;
HWND hTextFieldVK, hTextFieldScan;
std::string textOut;
std::string textX, textVK, textScan;

ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);


    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_KEYBOARDTS, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    if (!InitInstance(hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_KEYBOARDTS));

    MSG msg;


    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}


ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_KEYBOARDTS));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = nullptr;
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}


BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance;

    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_SYSMENU,
        CW_USEDEFAULT, 0, 280, 130, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd)
    {
        return FALSE;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}





std::string DecimalToHexadecimal(int decimalValue)
{
    std::stringstream ss;
    ss << std::hex << std::uppercase << decimalValue;
    return ss.str();
}

void xlsADD()
{
    std::stringstream inpText(textX);
    std::stringstream VKText(textVK);
    std::stringstream ScanText(textScan);
    std::vector<std::string> elementsINP;
    std::vector<std::string> elementsVK;
    std::vector<std::string> elementsScan;
    std::string elementINP, elementVK, elementScan;

    while (std::getline(inpText, elementINP, ' ') && std::getline(VKText, elementVK, ' ') && std::getline(ScanText, elementScan, ' '))
    {
        elementsINP.push_back(elementINP);
        elementsVK.push_back(elementVK);
        elementsScan.push_back(elementScan);
    }

    libxl::Book* book = xlCreateBook();
    if (book)
    {
        libxl::Sheet* sheet = book->addSheet("Sheet1");

        sheet->writeStr(1, 0, "Input");
        sheet->writeStr(2, 0, "VK cd");
        sheet->writeStr(3, 0, "Scan cd");

        for (int i = 0; i < elementsINP.size(); i++)
        {
            sheet->writeStr(1, i + 1, elementsINP[i].c_str());
            sheet->writeStr(2, i + 1, elementsVK[i].c_str());
            sheet->writeStr(3, i + 1, elementsScan[i].c_str());

        }

        book->save("TABLE.xls");
        book->release();
    }
}





LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{

    switch (message)
    {
    case WM_CREATE:
    {

        hInput = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER, 10, 10, 240, 20, hWnd, (HMENU)9, hInst, NULL);

        hLabelVK = CreateWindowW(L"STATIC", L" VK code", WS_VISIBLE | WS_CHILD | WS_BORDER, 10, 35, 75, 20, hWnd, (HMENU)10, hInst, NULL);
        hLabelScan = CreateWindowW(L"STATIC", L" Scan code", WS_VISIBLE | WS_CHILD | WS_BORDER, 10, 60, 75, 20, hWnd, (HMENU)11, hInst, NULL);
        hTextFieldVK = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER, 90, 35, 50, 20, hWnd, (HMENU)12, hInst, NULL);
        hTextFieldScan = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER, 90, 60, 50, 20, hWnd, (HMENU)13, hInst, NULL);

        hBtn = CreateWindowW(L"BUTTON", L"xls", WS_VISIBLE | WS_CHILD | WS_BORDER, 145, 35, 50, 45, hWnd, (HMENU)14, hInst, NULL);
        hBtnCl = CreateWindowW(L"BUTTON", L"clear", WS_VISIBLE | WS_CHILD | WS_BORDER, 200, 35, 50, 45, hWnd, (HMENU)15, hInst, NULL);
    }
    break;

    case WM_COMMAND:
    {
        RECT rect;
        GetClientRect(hInput, &rect);
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case 14:
            xlsADD();
            SetFocus(hWnd);
            break;
        case 15:
            SetWindowTextA(hInput, "");
            SetWindowTextA(hTextFieldVK, "");
            SetWindowTextA(hTextFieldScan, "");
            textOut.clear();
            textX.clear();
            textScan.clear();
            textVK.clear();
            SetFocus(hWnd);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }

    }
    break;

    case WM_KEYDOWN:
    {
        char name[12];

        GetKeyNameTextA(lParam, name, 10);

        textOut += name;
        if (textOut.length() < 20)
            SetWindowTextA(hInput, textOut.c_str());
        else
            SetWindowTextA(hInput, textOut.substr(textOut.length() - 20).c_str());

        textX += name;
        textX += " ";
        std::string VKstr = DecimalToHexadecimal(wParam);
        std::string SkanStr = DecimalToHexadecimal(MapVirtualKey(wParam, MAPVK_VK_TO_VSC));
        SetWindowTextA(hTextFieldVK, VKstr.c_str());
        SetWindowTextA(hTextFieldScan, SkanStr.c_str());
        textVK += VKstr + " ";
        textScan += SkanStr + " ";

    }
    break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}
