#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <oleauto.h>

#include "point_excel_import.h"

#include <algorithm>
#include <cwctype>
#include <fstream>
#include <new>
#include <sstream>
#include <stdexcept>
#include <utility>

namespace point {
namespace {

std::string utf8_from_wide(const std::wstring& value) {
    if (value.empty()) return {};
    const int required = WideCharToMultiByte(
        CP_UTF8, WC_ERR_INVALID_CHARS,
        value.data(), static_cast<int>(value.size()),
        nullptr, 0, nullptr, nullptr);
    if (required <= 0)
        return "Excel returned an unreadable Unicode error";
    std::string output(static_cast<std::size_t>(required), '\0');
    const int written = WideCharToMultiByte(
        CP_UTF8, WC_ERR_INVALID_CHARS,
        value.data(), static_cast<int>(value.size()),
        output.data(), required, nullptr, nullptr);
    if (written != required)
        return "Excel returned an unreadable Unicode error";
    return output;
}

class Variant {
public:
    Variant() { VariantInit(&value_); }
    explicit Variant(const std::wstring& value) : Variant() {
        value_.vt = VT_BSTR;
        value_.bstrVal =
            SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
        if (!value_.bstrVal) throw std::bad_alloc();
    }
    explicit Variant(long value) : Variant() {
        value_.vt = VT_I4;
        value_.lVal = value;
    }
    explicit Variant(bool value) : Variant() {
        value_.vt = VT_BOOL;
        value_.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    }
    Variant(Variant&& other) noexcept : Variant() {
        value_ = other.value_;
        VariantInit(&other.value_);
    }
    Variant& operator=(Variant&& other) noexcept {
        if (this != &other) {
            VariantClear(&value_);
            value_ = other.value_;
            VariantInit(&other.value_);
        }
        return *this;
    }
    Variant(const Variant&) = delete;
    Variant& operator=(const Variant&) = delete;
    ~Variant() { VariantClear(&value_); }

    VARIANT* get() { return &value_; }
    const VARIANT& value() const { return value_; }

    static Variant missing() {
        Variant result;
        result.value_.vt = VT_ERROR;
        result.value_.scode = DISP_E_PARAMNOTFOUND;
        return result;
    }

private:
    VARIANT value_{};
};

class Dispatch {
public:
    Dispatch() = default;
    explicit Dispatch(IDispatch* value) : value_(value) {}
    Dispatch(Dispatch&& other) noexcept
        : value_(std::exchange(other.value_, nullptr)) {}
    Dispatch& operator=(Dispatch&& other) noexcept {
        if (this != &other) {
            if (value_) value_->Release();
            value_ = std::exchange(other.value_, nullptr);
        }
        return *this;
    }
    Dispatch(const Dispatch&) = delete;
    Dispatch& operator=(const Dispatch&) = delete;
    ~Dispatch() { if (value_) value_->Release(); }
    IDispatch* get() const { return value_; }
    explicit operator bool() const { return value_ != nullptr; }
private:
    IDispatch* value_ = nullptr;
};

DISPID member_id(IDispatch* object, const wchar_t* name) {
    if (!object) throw std::runtime_error("Excel object is unavailable");
    LPOLESTR names[] = {const_cast<LPOLESTR>(name)};
    DISPID id = 0;
    const HRESULT result = object->GetIDsOfNames(
        IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(result))
        throw std::runtime_error("Excel member is unavailable");
    return id;
}

Variant invoke(IDispatch* object, const wchar_t* name, WORD flags,
               std::vector<Variant> arguments = {}) {
    const DISPID id = member_id(object, name);
    std::vector<VARIANTARG> reversed;
    reversed.reserve(arguments.size());
    for (auto it = arguments.rbegin(); it != arguments.rend(); ++it) {
        VARIANTARG copied{};
        VariantInit(&copied);
        const HRESULT copied_result =
            VariantCopy(&copied, const_cast<VARIANT*>(&it->value()));
        if (FAILED(copied_result))
            throw std::runtime_error("Unable to prepare Excel argument");
        reversed.push_back(copied);
    }
    DISPPARAMS parameters{};
    parameters.rgvarg = reversed.empty() ? nullptr : reversed.data();
    parameters.cArgs = static_cast<UINT>(reversed.size());
    DISPID property_put = DISPID_PROPERTYPUT;
    if ((flags & DISPATCH_PROPERTYPUT) != 0) {
        parameters.rgdispidNamedArgs = &property_put;
        parameters.cNamedArgs = 1;
    }
    Variant result;
    EXCEPINFO exception{};
    UINT argument_error = 0;
    const HRESULT status = object->Invoke(
        id, IID_NULL, LOCALE_USER_DEFAULT, flags,
        &parameters, result.get(), &exception, &argument_error);
    for (auto& value : reversed) VariantClear(&value);
    if (FAILED(status)) {
        if (exception.bstrDescription) {
            const std::wstring description(exception.bstrDescription);
            SysFreeString(exception.bstrDescription);
            throw std::runtime_error(utf8_from_wide(description));
        }
        throw std::runtime_error("Excel rejected an automation request");
    }
    return result;
}

Dispatch dispatch_property(IDispatch* object, const wchar_t* name) {
    Variant result = invoke(object, name, DISPATCH_PROPERTYGET);
    if (result.value().vt != VT_DISPATCH ||
        result.value().pdispVal == nullptr)
        throw std::runtime_error("Excel returned an invalid object");
    result.value().pdispVal->AddRef();
    return Dispatch(result.value().pdispVal);
}

long long_property(IDispatch* object, const wchar_t* name) {
    Variant result = invoke(object, name, DISPATCH_PROPERTYGET);
    VARIANT converted{};
    VariantInit(&converted);
    const HRESULT status = VariantChangeType(
        &converted, const_cast<VARIANT*>(&result.value()), 0, VT_I4);
    if (FAILED(status))
        throw std::runtime_error("Excel returned an invalid number");
    const long value = converted.lVal;
    VariantClear(&converted);
    return value;
}

std::wstring string_property(IDispatch* object, const wchar_t* name) {
    Variant result = invoke(object, name, DISPATCH_PROPERTYGET);
    VARIANT converted{};
    VariantInit(&converted);
    const HRESULT status = VariantChangeType(
        &converted, const_cast<VARIANT*>(&result.value()), 0, VT_BSTR);
    if (FAILED(status))
        throw std::runtime_error("Excel returned invalid text");
    std::wstring value =
        converted.bstrVal ? converted.bstrVal : L"";
    VariantClear(&converted);
    return value;
}

void put(IDispatch* object, const wchar_t* name, Variant value) {
    std::vector<Variant> arguments;
    arguments.push_back(std::move(value));
    (void)invoke(object, name, DISPATCH_PROPERTYPUT,
                 std::move(arguments));
}

Dispatch dispatch_result(Variant result) {
    if (result.value().vt != VT_DISPATCH ||
        result.value().pdispVal == nullptr)
        throw std::runtime_error("Excel returned an invalid object");
    result.value().pdispVal->AddRef();
    return Dispatch(result.value().pdispVal);
}

bool excel_extension(const std::filesystem::path& path) {
    std::wstring extension = path.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
                   [](wchar_t value) {
                       return static_cast<wchar_t>(std::towlower(value));
                   });
    return extension == L".xlsx" || extension == L".xls" ||
           extension == L".xlsm";
}

bool csv_extension(const std::filesystem::path& path) {
    std::wstring extension = path.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(),
                   [](wchar_t value) {
                       return static_cast<wchar_t>(std::towlower(value));
                   });
    return extension == L".csv";
}

std::wstring safe_component(std::wstring value) {
    for (wchar_t& ch : value) {
        if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' ||
            ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' ||
            ch == L'?' || ch == L'*')
            ch = L'_';
    }
    while (!value.empty() &&
           (value.back() == L' ' || value.back() == L'.'))
        value.pop_back();
    if (value.empty()) value = L"Sheet";
    if (value.size() > 80) value.resize(80);
    return value;
}

Dispatch open_workbook(IDispatch* workbooks,
                       const std::filesystem::path& path) {
    std::vector<Variant> arguments;
    arguments.emplace_back(path.wstring());  // Filename
    arguments.emplace_back(0L);              // UpdateLinks
    arguments.emplace_back(true);            // ReadOnly
    arguments.push_back(Variant::missing()); // Format
    arguments.push_back(Variant::missing()); // Password
    arguments.push_back(Variant::missing()); // WriteResPassword
    arguments.emplace_back(true);            // IgnoreReadOnlyRecommended
    arguments.push_back(Variant::missing()); // Origin
    arguments.push_back(Variant::missing()); // Delimiter
    arguments.emplace_back(false);           // Editable
    arguments.emplace_back(false);           // Notify
    arguments.push_back(Variant::missing()); // Converter
    arguments.emplace_back(false);           // AddToMru
    arguments.emplace_back(true);            // Local
    arguments.emplace_back(1L);              // CorruptLoad: normal load
    return dispatch_result(invoke(
        workbooks, L"Open", DISPATCH_METHOD, std::move(arguments)));
}

void close_workbook(IDispatch* workbook) {
    std::vector<Variant> arguments;
    arguments.emplace_back(false);
    (void)invoke(workbook, L"Close", DISPATCH_METHOD,
                 std::move(arguments));
}

std::wstring variant_text(const VARIANT& input) {
    if (input.vt == VT_EMPTY || input.vt == VT_NULL) return {};
    if (input.vt == VT_BSTR)
        return input.bstrVal ? input.bstrVal : L"";
    VARIANT converted{};
    VariantInit(&converted);
    const HRESULT status = VariantChangeType(
        &converted, const_cast<VARIANT*>(&input), 0, VT_BSTR);
    if (FAILED(status)) return {};
    std::wstring output =
        converted.bstrVal ? converted.bstrVal : L"";
    VariantClear(&converted);
    return output;
}

std::string csv_field(const std::wstring& input) {
    const std::string value = utf8_from_wide(input);
    if (value.find_first_of(",\"\r\n") == std::string::npos)
        return value;
    std::string output = "\"";
    for (char ch : value) {
        if (ch == '"') output += "\"\"";
        else output.push_back(ch);
    }
    output.push_back('"');
    return output;
}

void write_worksheet_csv(
        IDispatch* worksheet,
        const std::filesystem::path& output) {
    Dispatch used_range =
        dispatch_property(worksheet, L"UsedRange");
    Variant values = invoke(
        used_range.get(), L"Value2", DISPATCH_PROPERTYGET);
    std::ofstream stream(output, std::ios::binary | std::ios::trunc);
    if (!stream)
        throw std::runtime_error(
            "Point could not create the worksheet cache");
    stream << "\xEF\xBB\xBF";

    if ((values.value().vt & VT_ARRAY) == 0) {
        stream << csv_field(variant_text(values.value())) << "\r\n";
        return;
    }

    SAFEARRAY* array = values.value().parray;
    if (!array || SafeArrayGetDim(array) != 2)
        throw std::runtime_error(
            "Excel returned an unsupported worksheet array");
    VARTYPE element_type = VT_EMPTY;
    if (FAILED(SafeArrayGetVartype(array, &element_type)) ||
        element_type != VT_VARIANT)
        throw std::runtime_error(
            "Excel returned unsupported cell value types");

    LONG first_row = 0, last_row = -1;
    LONG first_column = 0, last_column = -1;
    if (FAILED(SafeArrayGetLBound(array, 1, &first_row)) ||
        FAILED(SafeArrayGetUBound(array, 1, &last_row)) ||
        FAILED(SafeArrayGetLBound(array, 2, &first_column)) ||
        FAILED(SafeArrayGetUBound(array, 2, &last_column)))
        throw std::runtime_error(
            "Excel returned invalid worksheet boundaries");
    const unsigned long long row_count =
        static_cast<unsigned long long>(
            static_cast<long long>(last_row) - first_row + 1);
    const unsigned long long column_count =
        static_cast<unsigned long long>(
            static_cast<long long>(last_column) - first_column + 1);
    if (row_count > 2'000'001ull || column_count > 2'000ull)
        throw std::runtime_error(
            "Worksheet exceeds Point's row or column safety limit");

    for (LONG row = first_row; row <= last_row; ++row) {
        for (LONG column = first_column;
             column <= last_column; ++column) {
            if (column != first_column) stream.put(',');
            LONG indices[2] = {row, column};
            VARIANT cell{};
            VariantInit(&cell);
            const HRESULT status =
                SafeArrayGetElement(array, indices, &cell);
            if (FAILED(status)) {
                VariantClear(&cell);
                throw std::runtime_error(
                    "Excel cell extraction failed");
            }
            stream << csv_field(variant_text(cell));
            VariantClear(&cell);
        }
        stream << "\r\n";
        if (!stream)
            throw std::runtime_error(
                "Worksheet cache write failed");
    }
}

}  // namespace

ExcelImportResult prepare_import_sources(
        const std::filesystem::path& inbox,
        const std::filesystem::path& cache,
        const std::function<void(
            std::size_t, std::size_t,
            const std::filesystem::path&, bool)>& progress,
        const std::function<bool()>& cancelled) {
    ExcelImportResult result;
    std::filesystem::create_directories(cache);

    std::vector<std::filesystem::path> workbooks_to_import;
    for (const auto& entry :
         std::filesystem::directory_iterator(inbox)) {
        std::error_code error;
        if (!entry.is_regular_file(error) || error ||
            entry.is_symlink(error) || error)
            continue;
        if (csv_extension(entry.path()))
            result.csv_sources.push_back(entry.path());
        else if (excel_extension(entry.path()))
            workbooks_to_import.push_back(entry.path());
    }
    if (workbooks_to_import.empty()) return result;

    std::sort(workbooks_to_import.begin(), workbooks_to_import.end());
    std::ostringstream signature;
    for (const auto& workbook : workbooks_to_import) {
        std::error_code error;
        const auto size = std::filesystem::file_size(workbook, error);
        if (error) continue;
        const auto modified =
            std::filesystem::last_write_time(workbook, error);
        if (error) continue;
        signature << workbook.filename().string() << '\t' << size << '\t'
                  << modified.time_since_epoch().count() << '\n';
    }
    const auto manifest = cache / "point-excel-cache.manifest";
    std::ifstream existing_manifest(manifest, std::ios::binary);
    const std::string existing_signature{
        std::istreambuf_iterator<char>(existing_manifest),
        std::istreambuf_iterator<char>()};
    if (!existing_signature.empty() &&
        existing_signature == signature.str()) {
        for (const auto& entry :
             std::filesystem::directory_iterator(cache)) {
            std::error_code error;
            if (entry.is_regular_file(error) && !error &&
                csv_extension(entry.path()))
                result.csv_sources.push_back(entry.path());
        }
        std::sort(result.csv_sources.begin(), result.csv_sources.end());
        if (!result.csv_sources.empty()) {
            result.workbook_count = workbooks_to_import.size();
            result.worksheet_count = result.csv_sources.size();
            result.cache_reused = true;
            if (progress) {
                for (std::size_t index = 0;
                     index < workbooks_to_import.size(); ++index)
                    progress(index + 1, workbooks_to_import.size(),
                             workbooks_to_import[index], true);
            }
            return result;
        }
    }

    for (const auto& entry :
         std::filesystem::directory_iterator(cache)) {
        std::error_code error;
        if (entry.is_regular_file(error) &&
            csv_extension(entry.path()))
            std::filesystem::remove(entry.path(), error);
    }

    const HRESULT initialized =
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool must_uninitialize = SUCCEEDED(initialized);
    if (FAILED(initialized) && initialized != RPC_E_CHANGED_MODE) {
        result.issues.push_back(
            "Windows COM could not be initialized");
        return result;
    }

    Dispatch application;
    try {
        CLSID excel_class{};
        if (FAILED(CLSIDFromProgID(L"Excel.Application", &excel_class)))
            throw std::runtime_error(
                "Microsoft Excel is not installed");
        IDispatch* raw_application = nullptr;
        const HRESULT created = CoCreateInstance(
            excel_class, nullptr, CLSCTX_LOCAL_SERVER,
            IID_IDispatch,
            reinterpret_cast<void**>(&raw_application));
        if (FAILED(created) || !raw_application)
            throw std::runtime_error(
                "Microsoft Excel could not be started");
        application = Dispatch(raw_application);

        put(application.get(), L"Visible", Variant(false));
        put(application.get(), L"DisplayAlerts", Variant(false));
        put(application.get(), L"EnableEvents", Variant(false));
        put(application.get(), L"AskToUpdateLinks", Variant(false));
        put(application.get(), L"AutomationSecurity", Variant(3L));
        Dispatch workbooks =
            dispatch_property(application.get(), L"Workbooks");

        std::size_t workbook_ordinal = 0;
        for (const auto& workbook_path : workbooks_to_import) {
            if (cancelled && cancelled())
                throw std::runtime_error("Refresh cancelled");
            ++workbook_ordinal;
            if (progress)
                progress(workbook_ordinal, workbooks_to_import.size(),
                         workbook_path, false);
            try {
                Dispatch workbook =
                    open_workbook(workbooks.get(), workbook_path);
                ++result.workbook_count;
                Dispatch worksheets =
                    dispatch_property(workbook.get(), L"Worksheets");
                const long count =
                    long_property(worksheets.get(), L"Count");
                for (long index = 1; index <= count; ++index) {
                    if (cancelled && cancelled())
                        throw std::runtime_error("Refresh cancelled");
                    std::vector<Variant> item_arguments;
                    item_arguments.emplace_back(index);
                    Dispatch worksheet = dispatch_result(invoke(
                        worksheets.get(), L"Item",
                        DISPATCH_PROPERTYGET,
                        std::move(item_arguments)));
                    const auto sheet_name =
                        string_property(worksheet.get(), L"Name");
                    const auto output =
                        cache /
                        (std::to_wstring(workbook_ordinal) + L"__" +
                         safe_component(workbook_path.stem().wstring()) +
                         L"__" + safe_component(sheet_name) + L".csv");
                    write_worksheet_csv(worksheet.get(), output);
                    result.csv_sources.push_back(output);
                    ++result.worksheet_count;
                }
                close_workbook(workbook.get());
            } catch (const std::exception& ex) {
                result.issues.push_back(
                    workbook_path.filename().string() + ": " + ex.what());
            }
        }
        (void)invoke(application.get(), L"Quit", DISPATCH_METHOD);
    } catch (const std::exception& ex) {
        result.issues.push_back(ex.what());
        if (application) {
            try {
                (void)invoke(application.get(), L"Quit", DISPATCH_METHOD);
            } catch (...) {}
        }
    }
    if (result.issues.empty()) {
        std::ofstream output(manifest, std::ios::binary | std::ios::trunc);
        output << signature.str();
    }
    if (must_uninitialize) CoUninitialize();
    return result;
}

}  // namespace point
