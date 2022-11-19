#include <Windows.h>
#include <iostream>
#include <comdef.h>
#include <WbemCli.h>
#include <tchar.h>
#include <vector>
#include <string>

#include "resource.h"

#pragma comment(lib, "wbemuuid.lib")

// Глобальні змінні:
HINSTANCE hInst;
LPCTSTR szWindowClass = L"LB-2";
LPCTSTR szTitle = L"LB-2";

IWbemLocator* pLocator = NULL;
IWbemServices* pService = NULL;
IEnumWbemClassObject* pEnumerator = NULL;
VARIANT vRet;
ULONG uReturn = 0;
IWbemClassObject* clsObj = NULL;

std::wstring bufwstring; // global string for Messagebox 1 task
std::wstring bufwstring1; // task 2
std::wstring bufwstring2; // task 3
std::wstring bufwstring3; // task 4

// Попередній опис функцій
ATOM MyRegisterClass(HINSTANCE hInstance);
BOOL InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void ConvertFuncSTR(VARIANT& vRet);
void ConvertFuncSTR1(VARIANT& vRet);
void ConvertFuncSTR2(VARIANT& vRet);
void ConvertFuncSTR3(VARIANT& vRet);
//std::vector<pair<wstring, int>> processes = getTop10ProcessWithMaxActiveThreads();

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	MSG msg;

	// Реєстрація класу вікна 
	MyRegisterClass(hInstance);

	// Створення вікна програми
	if (!InitInstance(hInstance, nCmdShow))
	{
		return FALSE;
	}
	// Цикл обробки повідомлень
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;
	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.style = CS_HREDRAW | CS_VREDRAW; 		//стиль вікна
	wcex.lpfnWndProc = (WNDPROC)WndProc; 		//віконна процедура
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hInstance = hInstance; 			//дескриптор програми
	wcex.hIcon = LoadIcon(NULL, IDI_INFORMATION); 		//визначення іконки
	wcex.hCursor = LoadCursor(NULL, IDC_ARROW); 	//визначення курсору
	wcex.hbrBackground = (HBRUSH)CreateSolidBrush(RGB(128, 166, 255)); //установка фону
	wcex.lpszMenuName = MAKEINTRESOURCE(IDR_MENU1); 				//визначення меню
	wcex.lpszClassName = szWindowClass; 		//ім’я класу
	wcex.hIconSm = NULL;

	return RegisterClassEx(&wcex); 			//реєстрація класу вікна
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;
	hInst = hInstance; //зберігає дескриптор додатка в змінній hInst
	hWnd = CreateWindow(szWindowClass, 	// ім’я класу вікна
		szTitle, 				// назва програми
		WS_OVERLAPPEDWINDOW,			// стиль вікна
		520, 			// положення по Х	
		180,			// положення по Y	
		500, 			// розмір по Х
		500, 			// розмір по Y
		NULL, 					// дескриптор батьківського вікна	
		NULL, 					// дескриптор меню вікна
		hInstance, 				// дескриптор програми
		NULL); 				// параметри створення.

	if (!hWnd) 	//Якщо вікно не творилось, функція повертає FALSE
	{
		return FALSE;
	}
	ShowWindow(hWnd, nCmdShow); 		//Показати вікно
	UpdateWindow(hWnd); 				//Оновити вікно
	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	HDC hdc;
	RECT rt;
	HRESULT hRes;
	std::wstring Manufacturer = L"Manufacturer: ";
	std::wstring Name = L"Processor Name: ";
	std::wstring PowerManagementSupported = L"PowerManagementSupported: ";
	std::wstring LoadPercent = L"CPU Load: ";
	std::wstring Description = L"Description: ";
	std::wstring Status = L"Status: ";
	std::wstring Characteristics = L"Characteristics: ";
	std::wstring Family = L"Family: ";
	std::wstring CreationDate = L"CreationDate: ";
	std::wstring Priority = L"Priority: ";
	std::wstring ProcessId = L"ProcessId: ";
	std::wstring ThreadCount = L"ThreadCount: "; 
	std::wstring ExecutablePath = L"ExecutablePath: ";
	std::wstring ProcessHandle = L"ProcessHandle: ";
	std::wstring ThreadPriority = L"Thread Priority: ";
	std::wstring PriorityBase = L"PriorityBase: ";
	std::wstring ElapsedTime = L"ElapsedTime: "; 
	std::wstring ThreadState = L"ThreadState: ";

	switch (message)
	{
	case WM_CREATE: 				//Повідомлення приходить при створенні вікна
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case ID_TASK1_PROCESSORINFO:
			//TASK 1
			bufwstring.clear();

			//Manufacturer
			clsObj->Get(L"Manufacturer", 0, &vRet, NULL, NULL);
			bufwstring += Manufacturer.c_str();
			ConvertFuncSTR(vRet);
			//Name
			clsObj->Get(L"Name", 0, &vRet, NULL, NULL);
			bufwstring += Name.c_str();
			ConvertFuncSTR(vRet);
			//PowerManagementSupported
			clsObj->Get(L"PowerManagementSupported", 0, &vRet, 0, 0);
			bufwstring += PowerManagementSupported.c_str();
			bufwstring += std::to_wstring(vRet.boolVal);
			bufwstring += '\n';

			MessageBox(hWnd, bufwstring.c_str(), L"Task1", LB_OKAY);
			VariantClear(&vRet);
			break;
		case ID_TASK2_CONNECTEDDEVICESINFO:
			//TASK 2
			bufwstring1.clear();

			//Description
			clsObj->Get(L"Description", 0, &vRet, NULL, NULL);
			bufwstring1 += Description.c_str();
			ConvertFuncSTR1(vRet);
			//Status
			clsObj->Get(L"Status", 0, &vRet, NULL, NULL);
			bufwstring1 += Status.c_str();
			ConvertFuncSTR1(vRet);
			//LoadPercentage
			
			clsObj->Get(L"LoadPercentage", 0, &vRet, NULL, NULL);
			bufwstring1 += LoadPercent.c_str();
			bufwstring1 += std::to_wstring(vRet.uintVal);
			bufwstring1 += '\n';
			//Characteristics
			clsObj->Get(L"Characteristics", 0, &vRet, NULL, NULL);
			bufwstring1 += Characteristics.c_str();
			bufwstring1 += std::to_wstring(vRet.uintVal);
			bufwstring1 += '\n';
			//Family
			clsObj->Get(L"Family", 0, &vRet, NULL, NULL);
			bufwstring1 += Family.c_str();
			bufwstring1 += std::to_wstring(vRet.uintVal);
			bufwstring1 += '\n';

			MessageBox(hWnd, bufwstring1.c_str(), L"Task2", LB_OKAY);
			VariantClear(&vRet);
			break;
		case ID_TASK3_PROCESSINFO:
			bufwstring2.clear();
			//TASK 3 
			//WIN32_PROCESS
			hRes = pService->ExecQuery(bstr_t("WQL"), bstr_t("SELECT * FROM Win32_Process WHERE Name = 'WINWORD.exe'"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator);
			VariantInit(&vRet);
			pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, &uReturn);

			//ExecutablePath
			clsObj->Get(L"ExecutablePath", 0, &vRet, NULL, NULL);
			if (vRet.bstrVal == nullptr) {
				MessageBox(hWnd, L"Process hasn't been found!", L"ERROR", LB_OKAY);
			}
			else {
				bufwstring2 += ExecutablePath.c_str();
				ConvertFuncSTR2(vRet);
				//CreationDate
				clsObj->Get(L"CreationDate", 0, &vRet, NULL, NULL);
				bufwstring2 += CreationDate.c_str();
				ConvertFuncSTR2(vRet);
				//Priority
				clsObj->Get(L"Priority", 0, &vRet, NULL, NULL);
				bufwstring2 += Priority.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';
				//ProcessId
				clsObj->Get(L"ProcessId", 0, &vRet, NULL, NULL);
				bufwstring2 += ProcessId.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';
				//ThreadCount
				clsObj->Get(L"ThreadCount", 0, &vRet, NULL, NULL);
				bufwstring2 += ThreadCount.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';

				//WIN32_THREAD
				hRes = pService->ExecQuery(bstr_t("WQL"), bstr_t("SELECT * FROM Win32_Thread WHERE ProcessHandle = SELECT ProcessId FROM Win32_Process WHERE Name = 'WINWORD.exe'"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator); // SELECT * FROM Win32_Thread WHERE ProcessHandle = SELECT ProcessId FROM Win32_Process WHERE Name = 'WINWORD.exe'
				VariantInit(&vRet);
				pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, &uReturn);

				//ProcessHandle
				clsObj->Get(L"ProcessHandle", 0, &vRet, NULL, NULL);
				bufwstring2 += ProcessHandle.c_str();
				ConvertFuncSTR2(vRet);
				//Priority
				clsObj->Get(L"Priority", 0, &vRet, NULL, NULL);
				bufwstring2 += ThreadPriority.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';
				//PriorityBase
				clsObj->Get(L"PriorityBase", 0, &vRet, NULL, NULL);
				bufwstring2 += PriorityBase.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';
				//ElapsedTime
				clsObj->Get(L"ElapsedTime", 0, &vRet, NULL, NULL);
				bufwstring2 += ElapsedTime.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';
				//ThreadState
				clsObj->Get(L"ThreadState", 0, &vRet, NULL, NULL);
				bufwstring2 += ThreadState.c_str();
				bufwstring2 += std::to_wstring(vRet.uintVal);
				bufwstring2 += '\n';

				MessageBox(hWnd, bufwstring2.c_str(), L"Task3", LB_OKAY);
				VariantClear(&vRet);
			}
			break;
		case ID_TASK4_PROCESSWITHHIGHESTTHREADQUANTITY:
			bufwstring3.clear();
			hRes = pService->ExecQuery(bstr_t("WQL"), bstr_t("SELECT Name, ThreadCount FROM Win32_Process WHERE (ThreadCount > 2000)"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator); // WHERE ProcessHandle = (SELECT ProcessId FROM Win32_Process WHERE Name = 'WINWORD.exe')
			VariantInit(&vRet);
			pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, &uReturn);

			MessageBox(hWnd, bufwstring3.c_str(), L"Task4", LB_OKAY);
			VariantClear(&vRet);
			break;
		}
		break;
	case WM_PAINT:
	{
		// DON'T DELETE !!!!
		hdc = BeginPaint(hWnd, &ps);
		GetClientRect(hWnd, &rt);
		DrawText(hdc, L" LB #2 ", -1, &rt, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
		EndPaint(hWnd, &ps);

		//CHECKING
		hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
		if (FAILED(hRes))
		{
			MessageBox(hWnd, L"COM", L"ERROR", LB_OKAY);
			return 1;
		}

		if ((FAILED(hRes = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_CONNECT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, 0))))
		{
			MessageBox(hWnd, L"Security issue", L"ERROR", LB_OKAY);
			return 1;
		}

		if (FAILED(hRes = CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pLocator))))
		{
			MessageBox(hWnd, L"CLSID_WbemLocator", L"ERROR", LB_OKAY);
			return 1;
		}

		if (FAILED(hRes = pLocator->ConnectServer(_bstr_t("root\\CIMV2"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &pService)))
		{
			pLocator->Release();
			MessageBox(hWnd, L"ConnectServer", L"ERROR", LB_OKAY);
			return 1;
		}

		if (FAILED(hRes = CoSetProxyBlanket(pService, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE)))
		{
			MessageBox(hWnd, L"ProxyBlanket", L"ERROR", LB_OKAY);
			pService->Release();
			pLocator->Release();
			CoUninitialize();
			return 1;
		}
		hRes = pService->ExecQuery(bstr_t("WQL"), bstr_t("SELECT * FROM Win32_Processor"), WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator);
		VariantInit(&vRet);
		pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, &uReturn);
		if (FAILED(hRes))
		{
			pLocator->Release();
			pService->Release();
			MessageBox(hWnd, L"ExecQuery", L"ERROR", LB_OKAY);
			return 1;
		}
	}
	break;
	case WM_DESTROY: 				//Завершення роботи
		clsObj->Release();
		pEnumerator->Release();
		pService->Release();
		pLocator->Release();
		PostQuitMessage(0);
		break;
	default:
		//Обробка повідомлень, які не оброблені користувачем
		VariantClear(&vRet);
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

void ConvertFuncSTR(VARIANT& vRet) {
	assert(vRet.bstrVal != nullptr);
	std::wstring tempstr(vRet.bstrVal, SysStringLen(vRet.bstrVal)); // VARIANT to wstring

	std::vector<char> infoVector(tempstr.begin(), tempstr.end());
	infoVector.push_back('\n');
	for (int i = 0; i < infoVector.size(); i++) {
		bufwstring += infoVector[i];
	}
}

void ConvertFuncSTR1(VARIANT& vRet) {
	assert(vRet.bstrVal != nullptr);
	std::wstring tempstr(vRet.bstrVal, SysStringLen(vRet.bstrVal)); // VARIANT to wstring

	std::vector<char> infoVector(tempstr.begin(), tempstr.end());
	infoVector.push_back('\n');
	for (int i = 0; i < infoVector.size(); i++) {
		bufwstring1 += infoVector[i];
	}
}

void ConvertFuncSTR2(VARIANT& vRet) {
	assert(vRet.bstrVal != nullptr);
	std::wstring tempstr(vRet.bstrVal, SysStringLen(vRet.bstrVal)); // VARIANT to wstring

	std::vector<char> infoVector(tempstr.begin(), tempstr.end());
	infoVector.push_back('\n');
	for (int i = 0; i < infoVector.size(); i++) {
		bufwstring2 += infoVector[i];
	}
}

void ConvertFuncSTR3(VARIANT& vRet) {
	assert(vRet.bstrVal != nullptr);
	std::wstring tempstr(vRet.bstrVal, SysStringLen(vRet.bstrVal)); // VARIANT to wstring

	std::vector<char> infoVector(tempstr.begin(), tempstr.end());
	infoVector.push_back('\n');
	for (int i = 0; i < infoVector.size(); i++) {
		bufwstring3 += infoVector[i];
	}
}