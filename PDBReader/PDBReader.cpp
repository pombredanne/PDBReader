// PDBReader.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//
#include <iostream>
#include "PDBReader.h"
#include <filesystem>

int main() {
    std::cout << "Hello World!\n" << std::endl;

    auto ret = PDBReader::DownloadPDBForFile(L"C:\\Windows\\System32\\win32kfull.sys", L"Symbols");
    if (!ret) {
        std::cout << "Download pdb failed\n" << std::endl;
    }
    else {
        std::cout << "Download pdb succeed\n" << std::endl;
    }

    system("pause");
}

PDBReader::PDBReader() {
    ;
}

HRESULT PDBReader::COINIT(DWORD init_flag) {
    return CoInitializeEx(0, init_flag);
}

bool PDBReader::DownloadPDBForFile(std::wstring executable_name, std::wstring symbol_folder) {
    CComPtr<IDiaDataSource> pSource;
    HRESULT hr;

    hr = CreateDiaDataSourceWithoutComRegistration(&pSource);
    if (FAILED(hr)) {
        throw std::exception("Could not CoCreate CLSID_DiaSource. Register msdia80.dll.");
    }

    std::filesystem::path sym_cache_folder = std::filesystem::absolute(symbol_folder);
    std::wstring MS_SYMBOL_SERVER = L"https://msdl.microsoft.com/download/symbols";
    std::wstring seachPath = std::wstring(L"srv*") + sym_cache_folder.c_str() + L"*" + MS_SYMBOL_SERVER;

    hr = pSource->loadDataForExe(executable_name.c_str(), seachPath.c_str(), 0);
    if (FAILED(hr)) {
        return false;
    }

    return true;
}

// We dont want to regsvr32 msdia140.dll on client's machine...
HRESULT PDBReader::CreateDiaDataSourceWithoutComRegistration(IDiaDataSource** data_source) {
    HMODULE hmodule = LoadLibraryW(PDBReader::dia_dll_name.c_str());

    // try to load dia dll in the same folder as main executable's
    if (!hmodule) {
        wchar_t temp[MAX_PATH] = { 0 };
        GetModuleFileNameW(0, temp, MAX_PATH);
        std::filesystem::path dllPath(temp);
        dllPath = dllPath.parent_path() / L"msdia140.dll";

        hmodule = LoadLibraryW(dllPath.c_str());
        if (!hmodule)
            return HRESULT_FROM_WIN32(GetLastError()); // library not found
    }

    BOOL(WINAPI * DllGetClassObject)(REFCLSID, REFIID, LPVOID*) =
        (BOOL(WINAPI*)(REFCLSID, REFIID, LPVOID*))GetProcAddress(hmodule, "DllGetClassObject");
    if (!DllGetClassObject)
        return E_FAIL;

    IClassFactory* pClassFactory;
    HRESULT hr = DllGetClassObject(CLSID_DiaSource, IID_IClassFactory, (LPVOID *)&pClassFactory);
    if (FAILED(hr))
        return hr;

    hr = pClassFactory->CreateInstance(NULL, IID_IDiaDataSource, (void **)data_source);
    if (FAILED(hr)) {
        pClassFactory->Release();
        return hr;
    }
        
    pClassFactory->Release();
    return S_OK;
}

void PDBReader::SetDiaDllName(std::wstring name) {
    PDBReader::dia_dll_name = name;
}
