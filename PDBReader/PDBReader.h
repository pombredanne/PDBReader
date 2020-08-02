#pragma once
#include <string>
#include <atlbase.h>
#include <dia2.h>

class PDBReader {
public:
    PDBReader(std::wstring pdb_name);

    PDBReader(std::wstring executable_name, std::wstring search_path);

    size_t FindSymbol(std::wstring sym, DWORD &type);

    size_t FindConst(std::wstring const_name);

    size_t FindFunction(std::wstring func);

    // ~PDBReader();

    static HRESULT COINIT(DWORD init_flag);

    static bool DownloadPDBForFile(std::wstring executable_name, std::wstring symbol_folder);

    static HRESULT CreateDiaDataSourceWithoutComRegistration(IDiaDataSource** data_source);

    static void SetDiaDllName(std::wstring name);

private:
    static inline std::wstring dia_dll_name = L"msdia140.dll";

    CComPtr<IDiaSession> pSession;

    CComPtr<IDiaSymbol> pGlobal;
};
