#define _WIN32_DCOM

#include <iostream>
#include <windows.h>
#include <comdef.h>
#include <string>
#include <iomanip> // if it's not included in your code already
#include <codecvt>  // Needed for conversion between std::wstring and std::string
#include <locale>   // Same as above
#import "msado15.dll" no_namespace rename("EOF", "EndOfFile")

class ConnectionWrapper {
public:
    ConnectionWrapper(_ConnectionPtr conn) : connection(conn) {}
    ~ConnectionWrapper() {
        if (connection && connection->State == adStateOpen) {
            connection->Close();
        }
    }
private:
    _ConnectionPtr connection;
};

class RecordsetWrapper {
public:
    RecordsetWrapper(_RecordsetPtr rec) : recordset(rec) {}
    ~RecordsetWrapper() {
        if (recordset && recordset->State == adStateOpen) {
            recordset->Close();
        }
    }
private:
    _RecordsetPtr recordset;
};

std::wstring trim(const std::wstring& str) {
    size_t first = str.find_first_not_of(L' ');
    if (std::wstring::npos == first) {
        return str;
    }
    size_t last = str.find_last_not_of(L' ');
    return str.substr(first, (last - first + 1));
}

std::string ConvertWideToUTF8(const std::wstring& wstr)
{
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    return converter.to_bytes(wstr);
}

std::string trim_and_convert_bstr(_bstr_t bstr)
{
    std::wstring wstr(bstr.GetBSTR());  // Convert to wstring
    wstr = trim(wstr);  // Trim the wstring
    std::string str = ConvertWideToUTF8(wstr);  // Convert wstring to string
    return str;
}

int main(int argc, char* argv[]) {
    // Initialize COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM library. Error code = 0x"
            << std::hex << hr << std::endl;
        return -1;
    }

    // Command line arguments check
    if (argc != 3) {
        std::cerr << "Usage: QueryApp <container-path> <query>" << std::endl;
        CoUninitialize();
        return -1;
    }

    std::string container_path = argv[1];
    std::string query = argv[2];

    std::string conn_str = "Provider=vfpoledb;Data Source=" + container_path + ";Collating Sequence=machine;Exclusive=No;";

    _ConnectionPtr pConnection = nullptr;
    _RecordsetPtr pRecordset = nullptr;

    try {
        hr = pConnection.CreateInstance(__uuidof(Connection));
        if (FAILED(hr))
            throw _com_error(hr);

        pConnection->Open(conn_str.c_str(), "", "", adConnectUnspecified);

        ConnectionWrapper cw(pConnection);

        hr = pRecordset.CreateInstance(__uuidof(Recordset));
        if (FAILED(hr))
            throw _com_error(hr);

        pRecordset->Open(query.c_str(), _variant_t((IDispatch*)pConnection, true), adOpenStatic, adLockReadOnly, adCmdText);

        RecordsetWrapper rw(pRecordset);

        // print column names
        std::cout << "";
        for (long i = 0; i < pRecordset->Fields->Count; i++) {
            FieldPtr field = pRecordset->Fields->GetItem(i);
            _bstr_t fieldName(field->Name);
            std::string strFieldName = trim_and_convert_bstr(fieldName);
            std::cout << strFieldName << "|";
        }
        std::cout << std::endl;

        if (!pRecordset->EndOfFile) {
            pRecordset->MoveFirst();

            while (!pRecordset->EndOfFile) {

                for (long i = 0; i < pRecordset->Fields->Count; i++) {
                    FieldPtr field = pRecordset->Fields->GetItem(i);
                    _variant_t value = field->Value;

                    // Determine the data type of the field and print accordingly
                    switch (value.vt) {
                    case VT_I1:
                        std::cout << "int: " << static_cast<int>(value.cVal) << "|"; // Signed char
                        break;
                    case VT_I2:
                        std::cout << "int: " << value.iVal << "|"; // 2-byte signed int
                        break;
                    case VT_I4:
                    case VT_INT:
                        std::cout << "int: " << value.intVal << "|"; // 4-byte signed int
                        break;
                    case VT_I8:
                        std::cout << "int: " << value.llVal << "|"; // 8-byte signed int
                        break;
                    case VT_UI1:
                        std::cout << "int: " << static_cast<int>(value.bVal) << "|"; // Unsigned char
                        break;
                    case VT_UI2:
                        std::cout << "int: " << value.uiVal << "|"; // 2-byte unsigned int
                        break;
                    case VT_UI4:
                    case VT_UINT:
                        std::cout << "int: " << value.uintVal << "|"; // 4-byte unsigned int
                        break;
                    case VT_UI8:
                        std::cout << "int: " << value.ullVal << "|"; // 8-byte unsigned int
                        break;
                    case VT_R4:
                        std::cout << "float: " << value.fltVal << "|"; // 8-byte unsigned int
                        break;
                    case VT_R8:
                        std::cout << "float: " << value.dblVal << "|"; // 8-byte unsigned int
                        break;
                    case VT_DECIMAL:
                        double dblVal;
                        VarR8FromDec(&value.decVal, &dblVal);
                        std::cout << "float: " << dblVal << "|";
                        break;
                    case VT_BOOL:
                        std::cout << "bool: " << (value.boolVal ? 1 : 0) << "|";
                        break;
                    case VT_DATE: {
                        SYSTEMTIME st;
                        VariantTimeToSystemTime(value.date, &st);
                        char date[30];
                        sprintf_s(date, "%04d-%02d-%02d %02d:%02d:%02d",
                            st.wYear, st.wMonth, st.wDay,
                            st.wHour, st.wMinute, st.wSecond);
                        std::cout << "date: " << date << "|";
                        break;
                    }
                    case VT_BSTR:
                        std::cout << "string: " << trim_and_convert_bstr(value.bstrVal) << "|";
                        break;
                    default: // Everything else we'll just print as string
                        std::cout << "Value: ";
                        if (value.vt != VT_NULL) {
                            std::cout << trim_and_convert_bstr(value.bstrVal);
                        }
                        else {
                            std::cout << "NULL";
                        }
                        std::cout << "|";
                        break;
                    }
                }
                std::cout << std::endl;
                pRecordset->MoveNext();
            }
        }

        pRecordset->Close();
        pConnection->Close();
    }
    catch (_com_error& e) {
        IErrorInfo* errorInfo = e.ErrorInfo();
        if (errorInfo) {
            BSTR description;
            errorInfo->GetDescription(&description);
            _bstr_t bstrDescription(description);
            std::cerr << "COM Error:\n";
            std::cerr << e.ErrorMessage() << std::endl;
            std::cerr << "Error description: " << (const char*)bstrDescription << std::endl;
            errorInfo->Release();
        }
        else {
            std::cerr << "COM Error (no description available):\n";
            std::cerr << e.ErrorMessage() << std::endl;
        }
        CoUninitialize();
        return -1;
    }
    catch (const std::exception& e) {
        std::cerr << "Standard exception: " << e.what() << std::endl;
        CoUninitialize();
        return -1;
    }
    catch (...) {
        std::cerr << "Unknown error occurred.\n";
        CoUninitialize();
        return -1;
    }

    // Uninitialize COM
    CoUninitialize();
    return 0;
}
