#define _WIN32_DCOM
#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>
#include <fstream>
using namespace std; 

#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char** argv)
{
    wofstream MyFile("data.txt");

    HRESULT hres;

    // Step 1: --------------------------------------------------
    // Initialize COM. ------------------------------------------

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres))
    {
        cout << "Failed to initialize COM library. Error code = 0x"
            << hex << hres << endl;
        return 1; // Program has failed.
    }

    // Step 2: --------------------------------------------------
    // Set general COM security levels --------------------------

    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities
        NULL                         // Reserved
    );

    if (FAILED(hres))
    {
        cout << "Failed to initialize security. Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1; // Program has failed.
    }

    // Step 3: ---------------------------------------------------
    // Obtain the initial locator to WMI -------------------------

    IWbemLocator* pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres))
    {
        cout << "Failed to create IWbemLocator object."
            << " Err code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1; // Program has failed.
    }

    // Step 4: -----------------------------------------------------
    // Connect to WMI through the IWbemLocator::ConnectServer method

    IWbemServices* pSvc = NULL;

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
        NULL,                    // User name. NULL = current user
        NULL,                    // User password. NULL = current
        0,                       // Locale. NULL indicates current
        NULL,                    // Security flags.
        0,                       // Authority (for example, Kerberos)
        0,                       // Context object
        &pSvc                    // pointer to IWbemServices proxy
    );

    if (FAILED(hres))
    {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }

    cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl
        << endl;

    // Step 5: --------------------------------------------------
    // Set security levels on the proxy -------------------------

    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities
    );

    if (FAILED(hres))
    {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }

    // Step 6: --------------------------------------------------
    // Use the IWbemServices pointer to make requests of WMI ----

    // For example, get the name of the operating system
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_OperatingSystem"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
            << " Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }

    // Step 7: -------------------------------------------------
    // Get the data from the query in step 6 -------------------

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    MyFile << "\t-------------------------------------------------------------------------------------------" << endl
        << endl;
    MyFile << "\t***    OS details like Name, Version, OSLanguage, BuildNumber:     ***" << endl
        << endl;
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wstring oSName = vtProp.bstrVal;
        MyFile << " 1. OS Name : " << oSName << endl
            << endl;
        wcout << "OS Name written Successfully!!!" << endl;
        VariantClear(&vtProp);

        // Get the value of the Version property
        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        wstring version = vtProp.bstrVal;
        MyFile << " 2. Version : " << version << endl
            << endl;
        wcout << "Version written Successfully!!!" << endl;
        VariantClear(&vtProp);

        // Get the value of the OSLanguage property
        hr = pclsObj->Get(L"OSLanguage", 0, &vtProp, 0, 0);
        MyFile << " 3. OSLanguage : " << PtrToInt(vtProp.bstrVal) << endl
            << endl;
        wcout << "OSLanguage written Successfully!!!" << endl;
        VariantClear(&vtProp);

        // Get the value of the BuildNumber property
        hr = pclsObj->Get(L"BuildNumber", 0, &vtProp, 0, 0);
        wstring buildNumber = vtProp.bstrVal;
        MyFile << " 4. BuildNumber : " << buildNumber << endl
            << endl;
        wcout << "BuildNumber written Successfully!!!" << endl;
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_PhysicalMemory"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;
        MyFile << "\t-------------------------------------------------------------------------------------------" << endl
            << endl;
        MyFile << "\t***    Machine details like RAM:     ***" << endl
            << endl;
        while (pEnumerator)
        {
            hres = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"DeviceLocator", 0, &vtProp, 0, 0);
            MyFile << " 1. DeviceLocator : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            MyFile << " 2. Name : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            // pclsObj = NULL;
        }
        wcout << "RAM written successfully!!!" << endl;
    }

    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;
        MyFile << "\t-------------------------------------------------------------------------------------------" << endl
            << endl;
        MyFile << "\t***    Machine details like Processor:     ***" << endl
            << endl;
        while (pEnumerator)
        {
            hres = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
            MyFile << " 1. Caption : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            MyFile << " 2. Name : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"DeviceID", 0, &vtProp, 0, 0);
            MyFile << " 3. DeviceID : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"Status", 0, &vtProp, 0, 0);
            MyFile << " 4. Status: " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"SystemName", 0, &vtProp, 0, 0);
            MyFile << " 5. SystemName : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            // pclsObj = NULL;
        }
        wcout << "Processor written successfully!!!" << endl;
    }

    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_DiskDrive"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;
        MyFile << "\t-------------------------------------------------------------------------------------------" << endl
            << endl;
        MyFile << "\t***    Machine details like Disk:     ***" << endl
            << endl;
        while (pEnumerator)
        {
            hres = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
            MyFile << " Caption : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
            MyFile << " Description : " << vtProp.bstrVal << endl
                << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"DeviceID", 0, &vtProp, 0, 0);
            MyFile << " DeviceID : " << vtProp.bstrVal << endl
                << endl
                << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            // pclsObj = NULL;
        }
        wcout << "Disk written successfully!!!" << endl;
    }

    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_Process"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1; // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;
        MyFile << "\t-------------------------------------------------------------------------------------------" << endl
            << endl;
        MyFile << "\t***    Processes details like Name, CommandLine, ThreadCount:     ***" << endl
            << endl;
        while (pEnumerator)
        {
            hres = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            MyFile << " Process Name : " << vtProp.bstrVal << endl;
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"CommandLine", 0, &vtProp, 0, 0);
            if (!(vtProp.vt == VT_NULL))
            {
                //wstring commandLine = vtProp.bstrVal;
                MyFile << " CommandLine : " << PtrToLong(vtProp.bstrVal) << endl;
            }
            VariantClear(&vtProp);

            // Get the value of the ThreadCount property
            hres = pclsObj->Get(L"ThreadCount", 0, &vtProp, 0, 0);
            MyFile << " ThreadCount : " << PtrToInt(vtProp.bstrVal) << endl
                << endl;
            VariantClear(&vtProp);

            pclsObj->Release();
            // pclsObj = NULL;
        }
        wcout << "Processes written successfully!!!" << endl;
    }

    // Cleanup
    // ========

    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();
    cout << endl
        << "File written Successfully!!!" << endl
        << endl;
    MyFile.close();

    return 0; // Program successfully completed.
}