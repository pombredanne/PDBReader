#pragma once
#include <string>
#include <atlbase.h>
#include <dia2.h>

class PDBReader {
public:
    PDBReader();

    // ~PDBReader();

    static HRESULT COINIT(DWORD init_flag);

    static bool DownloadPDBForFile(std::wstring executable_name, std::wstring symbol_folder);

    static HRESULT CreateDiaDataSourceWithoutComRegistration(IDiaDataSource** data_source);

    static void SetDiaDllName(std::wstring name);

private:
    static inline std::wstring dia_dll_name = L"msdia140.dll";
};
