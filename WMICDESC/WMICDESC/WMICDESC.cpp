#include <Windows.h>
#include <wbemidl.h>
#include <comutil.h>
#include <iostream>
#include <thread>
#include <chrono>
#include <iomanip> 
#pragma comment(lib, "wbemuuid.lib")

void PrintProcessorInfo(IWbemClassObject* pclsObj)
{
    VARIANT vtProp;
    VariantInit(&vtProp);

    HRESULT hr = pclsObj->Get(L"MaxClockSpeed", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        std::wcout << L"Max Turbo Frequency: " << vtProp.ulVal << L" Mhz" << std::endl;
        VariantClear(&vtProp);
    }


    hr = pclsObj->Get(L"LoadPercentage", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        std::wcout << L"Load Percentage: " << vtProp.ulVal << L" %" << std::endl;
        VariantClear(&vtProp);
    }

    hr = pclsObj->Get(L"CurrentVoltage", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        double voltageMillivolts = vtProp.dblVal;  // значення в мілівольтах ?????????????
        double voltageVolts = voltageMillivolts / 1000.0;  // Переведення у вольти !??

        if (voltageVolts != 0.0) {
            std::wcout << L"Current Voltage: " << voltageVolts << L" V" << std::endl;
        }
        else {
            std::wcout << L"Current Voltage: N/A" << std::endl;
        }

        VariantClear(&vtProp);
    }

    hr = pclsObj->Get(L"CurrentClockSpeed", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        std::wcout << L"Current Clock Speed: " << vtProp.ulVal << L" Mhz" << std::endl;
        VariantClear(&vtProp);
    }


    hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        std::wcout << L"Description: " << vtProp.bstrVal << std::endl;
        VariantClear(&vtProp);
    }
}
void PrintMemoryInfo(IWbemServices* pSvc)
{
    IEnumWbemClassObject* pEnumerator = NULL;
    HRESULT hres = pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT * FROM Win32_PhysicalMemory"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        std::cerr << "Query for memory failed" << std::endl;
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;


    HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
        &pclsObj, &uReturn);



    VARIANT vtProp;
    VariantInit(&vtProp);

    hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
    if (SUCCEEDED(hr))
    {
        std::wcout << L"RAM Capacity: " << vtProp.ullVal / (1024 * 1024) << L" MB" << std::endl;
        VariantClear(&vtProp);
    }

    pclsObj->Release();


    pEnumerator->Release();
}

void PrintVideoCardInfo(IWbemServices* pSvc)
{
    IEnumWbemClassObject* pEnumerator = NULL;
    HRESULT hres = pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT * FROM Win32_VideoController"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        std::cerr << "Query for video controller failed" << std::endl;
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (hr != WBEM_S_NO_ERROR || uReturn == 0)
        {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr))
        {
            std::wcout << L"GPU Description: " << vtProp.bstrVal << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
    }

    pEnumerator->Release();
}
void UpdateProcessorData(IWbemServices* pSvc)
{
    IEnumWbemClassObject* pEnumerator = NULL;
    HRESULT hres = pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        std::cerr << "Query for processors failed" << std::endl;
        return;
    }

    while (true)
    {
        IWbemClassObject* pclsObj = NULL;
        ULONG uReturn = 0;

        while (pEnumerator)
        {
            HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            PrintProcessorInfo(pclsObj);


            pclsObj->Release();
        }

        pEnumerator->Release();

        // Оновлення кожні 4 секунд
        std::this_thread::sleep_for(std::chrono::seconds(4));

        // Знову отримати перелік процесорів
        hres = pSvc->ExecQuery(
            _bstr_t(L"WQL"),
            _bstr_t(L"SELECT * FROM Win32_Processor"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator);

        if (FAILED(hres))
        {
            std::cerr << "Query for processors failed" << std::endl;
            return;
        }
    }
}

int main(int argc, char* argv[])
{
    HRESULT hres;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres))
    {
        std::cerr << "Failed to initialize COM library" << std::endl;
        return 1;
    }

    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL
    );

    IWbemLocator* pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres))
    {
        std::cerr << "Failed to create IWbemLocator object" << std::endl;
        CoUninitialize();
        return 1;
    }

    IWbemServices* pSvc = NULL;

    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        &pSvc);

    if (FAILED(hres))
    {
        std::cerr << "Could not connect to WMI" << std::endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );

    if (FAILED(hres))
    {
        std::cerr << "Could not set proxy blanket" << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Створення окремого потоку для оновлення даних
    std::thread updateThread(UpdateProcessorData, pSvc);
    PrintMemoryInfo(pSvc);
    PrintVideoCardInfo(pSvc);
    // Заспати головний потік на деякий час, щоб обидва потоки мали можливість працювати
    std::this_thread::sleep_for(std::chrono::seconds(60));

    // Завершити потік оновлення
    updateThread.join();



    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    system("pause");
    return 0;
}
