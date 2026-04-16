/**
 * 自动化计算工具 - C++ Win32版
 * 版本: 3.0
 * 开发者: 德龙轧钢自动化团队
 * 
 * 功能: 打包数据查询、状态字查询、内存映象网计算、模拟量计算、速度转换
 */

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <string>
#include <vector>
#include <map>
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <cmath>
#include <cstdlib>
#include <memory>
#include <functional>

#include "resource.h"
#include "json.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

//=============================================================================
// 全局常量
//=============================================================================

static const wchar_t* WINDOW_TITLE = L"自动化计算工具-王国强-202603";
static const int WINDOW_WIDTH = 533;
static const int WINDOW_HEIGHT = 630;

// 颜色定义
static const COLORREF COLOR_TITLE = RGB(0, 0, 200);
static const COLORREF COLOR_SEPARATOR = RGB(128, 128, 128);
static const COLORREF COLOR_FAULT = RGB(139, 0, 0);
static const COLORREF COLOR_RESULT = RGB(0, 100, 0);
static const COLORREF COLOR_ERROR = RGB(231, 76, 60);
static const COLORREF COLOR_INFO = RGB(52, 152, 219);
static const COLORREF COLOR_SUCCESS = RGB(39, 174, 96);
static const COLORREF COLOR_BG = RGB(240, 240, 240);
static const COLORREF COLOR_TITLE_BG = RGB(230, 243, 255);

//=============================================================================
// 数据结构
//=============================================================================

struct BitInfo {
    std::wstring description;
    std::string display_on;
};

struct GroupData {
    std::wstring first_name = L"第一个字";
    std::wstring second_name = L"第二个字";
    std::map<int, BitInfo> first_word;
    std::map<int, BitInfo> second_word;
    std::string word_displayon = "0";
};

//=============================================================================
// 全局变量
//=============================================================================

static HINSTANCE g_hInst = nullptr;
static HWND g_hMainWnd = nullptr;
static HWND g_hTab = nullptr;
static HWND g_hTabPages[5] = {nullptr, nullptr, nullptr, nullptr, nullptr};
static int g_currentTab = 0;
static HFONT g_hFont = nullptr;
static HFONT g_hFontBold = nullptr;

// 打包数据相关
static std::map<std::wstring, GroupData> g_groups;
static std::wstring g_currentGroup;
static std::wstring g_firstName = L"第一个字";
static std::wstring g_secondName = L"第二个字";
static std::map<int, BitInfo> g_firstWord;
static std::map<int, BitInfo> g_secondWord;
static std::string g_wordDisplayon = "0";

// 状态字相关
static std::map<std::wstring, std::wstring> g_statusWordMap;
static std::map<std::wstring, std::map<std::wstring, std::wstring>> g_statusGroups;
static std::wstring g_currentStatusGroup;
static std::wstring g_currentStatusWord;
static std::wstring g_currentStatusName = L"状态字";

// 内存计算相关
static int g_memCalcType = 1;
static std::map<int, HWND> g_memEntries;
static HWND g_hMemInputFrame = nullptr;

// 置顶状态
static bool g_isPinned = false;

//=============================================================================
// 前向声明
//=============================================================================

static void LoadGroup(const std::wstring& groupName);
static void LoadStatusGroup(const std::wstring& groupName);
static void LoadStatusWord(const std::wstring& statusKey);

//=============================================================================
// 工具函数
//=============================================================================

static std::wstring Utf8ToWstring(const std::string& str) {
    if (str.empty()) return L"";
    if (str.size() > static_cast<size_t>(INT_MAX)) return L"";
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), nullptr, 0);
    if (size_needed <= 0) return L"";
    std::wstring result(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), static_cast<int>(str.size()), &result[0], size_needed);
    return result;
}

static std::string WstringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    if (wstr.size() > static_cast<size_t>(INT_MAX)) return "";
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), static_cast<int>(wstr.size()), nullptr, 0, nullptr, nullptr);
    if (size_needed <= 0) return "";
    std::string result(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), static_cast<int>(wstr.size()), &result[0], size_needed, nullptr, nullptr);
    return result;
}

static std::wstring GbkToWstring(const std::string& str) {
    if (str.empty()) return L"";
    if (str.size() > static_cast<size_t>(INT_MAX)) return L"";
    int size_needed = MultiByteToWideChar(CP_ACP, 0, str.c_str(), static_cast<int>(str.size()), nullptr, 0);
    if (size_needed <= 0) return L"";
    std::wstring result(size_needed, 0);
    MultiByteToWideChar(CP_ACP, 0, str.c_str(), static_cast<int>(str.size()), &result[0], size_needed);
    return result;
}

static std::wstring SafeJsonString(const json::Value& val, const std::wstring& defaultVal = L"") {
    if (val.is_string()) {
        return Utf8ToWstring(val.as_string());
    } else if (val.is_number()) {
        std::wostringstream ss;
        ss << val.as_double();
        return ss.str();
    } else if (val.is_bool()) {
        return val.as_bool() ? L"true" : L"false";
    }
    return defaultVal;
}

static std::string SafeJsonStringUtf8(const json::Value& val, const std::string& defaultVal = "") {
    if (val.is_string()) {
        return val.as_string();
    } else if (val.is_number()) {
        std::ostringstream ss;
        ss << val.as_double();
        return ss.str();
    } else if (val.is_bool()) {
        return val.as_bool() ? "true" : "false";
    }
    return defaultVal;
}

static int g_dpiScale = 100;
static HBRUSH g_hBrushBg = nullptr;
static HBRUSH g_hBrushWhite = nullptr;

static int Scale(int value) {
    return MulDiv(value, g_dpiScale, 100);
}

static std::wstring IntToWstr(int val) {
    return std::to_wstring(val);
}

static std::wstring IntToWstr(size_t val) {
    return std::to_wstring(val);
}

static std::wstring DoubleToWstr(double val, int precision = 2) {
    std::wostringstream ss;
    ss << std::fixed << std::setprecision(precision) << val;
    return ss.str();
}

enum class ParseError { None, Empty, InvalidFormat, OutOfRange };

static ParseError TryParseIntEx(const std::wstring& s, int& outVal) {
    if (s.empty()) return ParseError::Empty;
    try {
        size_t pos = 0;
        long long val = std::stoll(s, &pos, 10);
        if (pos != s.length()) return ParseError::InvalidFormat;
        if (val < INT_MIN || val > INT_MAX) return ParseError::OutOfRange;
        outVal = static_cast<int>(val);
        return ParseError::None;
    } catch (const std::invalid_argument&) {
        return ParseError::InvalidFormat;
    } catch (const std::out_of_range&) {
        return ParseError::OutOfRange;
    } catch (...) {
        return ParseError::InvalidFormat;
    }
}

static bool TryParseInt(const std::wstring& s, int& outVal) {
    return TryParseIntEx(s, outVal) == ParseError::None;
}

static ParseError TryParseDoubleEx(const std::wstring& s, double& outVal) {
    if (s.empty()) return ParseError::Empty;
    try {
        size_t pos = 0;
        outVal = std::stod(s, &pos);
        if (pos != s.length()) return ParseError::InvalidFormat;
        return ParseError::None;
    } catch (const std::invalid_argument&) {
        return ParseError::InvalidFormat;
    } catch (const std::out_of_range&) {
        return ParseError::OutOfRange;
    } catch (...) {
        return ParseError::InvalidFormat;
    }
}

static bool TryParseDouble(const std::wstring& s, double& outVal) {
    return TryParseDoubleEx(s, outVal) == ParseError::None;
}

static bool TryParseHex(const std::wstring& s, int& outVal) {
    if (s.empty()) return false;
    try {
        std::wstring upper = s;
        for (auto& c : upper) c = towupper(c);
        size_t pos = 0;
        outVal = std::stoi(upper, &pos, 16);
        return pos == upper.length();
    } catch (...) {
        return false;
    }
}

static std::wstring IntToHex(int val) {
    wchar_t buf[16];
    swprintf_s(buf, L"%04X", val);
    return buf;
}

static std::wstring DecimalToBinary(int num, int maxBits) {
    std::wstring result;
    result.reserve(maxBits);
    
    // 直接按位处理，支持0和负数
    unsigned int unum = static_cast<unsigned int>(num);
    for (int i = maxBits - 1; i >= 0; --i) {
        result += (wchar_t)('0' + ((unum >> i) & 1));
    }
    
    return result;
}

static std::wstring FormatBinary(const std::wstring& bits) {
    std::wstring result;
    for (size_t i = 0; i < bits.size(); ++i) {
        if (i > 0 && i % 4 == 0) result += L' ';
        result += bits[i];
    }
    return result;
}

static std::wstring Base64Decode(const std::string& encoded) {
    static const std::string base64_chars = 
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    
    std::vector<int> T(256, -1);
    for (int i = 0; i < 64; i++) T[base64_chars[i]] = i;
    
    std::string decoded;
    size_t len = encoded.size();
    
    // 跳过前导空白
    size_t start = 0;
    while (start < len && (encoded[start] == ' ' || encoded[start] == '\r' || encoded[start] == '\n' || encoded[start] == '\t')) {
        start++;
    }
    
    // 按4字节块处理
    for (size_t i = start; i < len; ) {
        int block[4] = {-1, -1, -1, -1};
        int validCount = 0;
        
        // 读取4个有效字符
        for (int j = 0; j < 4 && i < len; ) {
            unsigned char c = encoded[i++];
            if (c == '=') {
                block[j++] = -2; // 填充标记
                validCount++;
            } else if (T[c] >= 0) {
                block[j++] = T[c];
                validCount++;
            }
            // 跳过空白字符
        }
        
        if (validCount == 0) continue;
        if (validCount < 2) break; // 无效的Base64
        
        // 解码
        int val = (block[0] << 18) | (block[1] << 12);
        decoded.push_back(char((val >> 16) & 0xFF));
        
        if (block[2] >= 0) {
            val |= (block[2] << 6);
            decoded.push_back(char((val >> 8) & 0xFF));
            
            if (block[3] >= 0) {
                val |= block[3];
                decoded.push_back(char(val & 0xFF));
            }
        }
        
        // 遇到填充字符则停止
        if (block[2] == -2 || block[3] == -2) break;
    }
    
    return Utf8ToWstring(decoded);
}

//=============================================================================
// 控件创建辅助函数
//=============================================================================

static HWND CreateStatic(HWND parent, const wchar_t* text, int x, int y, int w, int h, DWORD style = WS_VISIBLE | WS_CHILD) {
    HWND hCtrl = CreateWindowExW(0, L"STATIC", text, style, x, y, w, h, parent, nullptr, g_hInst, nullptr);
    if (hCtrl && g_hFont) SendMessageW(hCtrl, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hCtrl;
}

static HWND CreateEdit(HWND parent, int id, int x, int y, int w, int h, const wchar_t* text = L"", bool multiline = false, bool readonly = false) {
    DWORD style = WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP;
    if (multiline) style |= ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL;
    if (readonly) style |= ES_READONLY;
    
    HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", text, style, x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
    if (hEdit && g_hFont) SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hEdit;
}

static HWND CreateButton(HWND parent, int id, const wchar_t* text, int x, int y, int w, int h) {
    HWND hBtn = CreateWindowExW(0, L"BUTTON", text, WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | WS_TABSTOP, x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
    if (hBtn && g_hFont) SendMessageW(hBtn, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hBtn;
}

static HWND CreateGroupBox(HWND parent, int id, const wchar_t* text, int x, int y, int w, int h) {
    HWND hGroup = CreateWindowExW(0, L"BUTTON", text, WS_VISIBLE | WS_CHILD | BS_GROUPBOX, x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
    if (hGroup && g_hFont) SendMessageW(hGroup, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hGroup;
}

static HWND CreateRadioButton(HWND parent, int id, const wchar_t* text, int x, int y, int w, int h, bool checked = false, bool firstInGroup = false) {
    DWORD style = WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON;
    if (firstInGroup) style |= WS_GROUP;
    HWND hRadio = CreateWindowExW(0, L"BUTTON", text, style, x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
    if (hRadio && g_hFont) SendMessageW(hRadio, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    if (checked) SendMessageW(hRadio, BM_SETCHECK, BST_CHECKED, 0);
    return hRadio;
}

static HWND CreateComboBox(HWND parent, int id, int x, int y, int w, int h) {
    HWND hCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_VSCROLL, x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
    if (hCombo && g_hFont) SendMessageW(hCombo, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hCombo;
}

static void SetText(HWND hCtrl, const std::wstring& text) {
    SetWindowTextW(hCtrl, text.c_str());
}

static std::wstring GetText(HWND hCtrl) {
    int len = GetWindowTextLengthW(hCtrl);
    if (len == 0) return L"";
    std::wstring text(len + 1, 0);
    GetWindowTextW(hCtrl, &text[0], len + 1);
    text.resize(len);
    return text;
}

static void AppendText(HWND hEdit, const std::wstring& text) {
    int len = GetWindowTextLengthW(hEdit);
    SendMessageW(hEdit, EM_SETSEL, len, len);
    SendMessageW(hEdit, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

static void ClearText(HWND hEdit) {
    SetWindowTextW(hEdit, L"");
}

static void ShowInfo(HWND parent, const wchar_t* msg) {
    MessageBoxW(parent, msg, L"提示", MB_OK | MB_ICONINFORMATION);
}

static void ShowError(HWND parent, const wchar_t* msg) {
    MessageBoxW(parent, msg, L"错误", MB_OK | MB_ICONERROR);
}

//=============================================================================
// JSON数据加载
//=============================================================================

static void ProcessGroupedData(const json::Value& groups, bool isStatusFormat, const std::wstring& decodeMethod);
static void ProcessUngroupedData(const json::Value& data, const std::wstring& decodeMethod);

static void LoadJsonData(const std::wstring& filePath) {
    HANDLE hFile = CreateFileW(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        ShowError(g_hMainWnd, L"无法打开文件");
        return;
    }
    
    DWORD fileSize = GetFileSize(hFile, nullptr);
    if (fileSize == INVALID_FILE_SIZE || fileSize > 10 * 1024 * 1024) {
        CloseHandle(hFile);
        ShowError(g_hMainWnd, L"文件大小无效或过大（最大10MB）");
        return;
    }
    
    std::string content(fileSize, 0);
    DWORD bytesRead;
    BOOL readOk = ReadFile(hFile, &content[0], fileSize, &bytesRead, nullptr);
    CloseHandle(hFile);
    
    if (!readOk || bytesRead != fileSize) {
        ShowError(g_hMainWnd, L"读取文件失败");
        return;
    }
    
    // 检测BOM头确定编码
    bool isUtf8Bom = false;
    bool isGbk = false;
    size_t startPos = 0;
    
    if (bytesRead >= 3 && (unsigned char)content[0] == 0xEF && 
        (unsigned char)content[1] == 0xBB && (unsigned char)content[2] == 0xBF) {
        isUtf8Bom = true;
        startPos = 3;
    }
    
    // 移除尾部空白
    content.erase(std::remove(content.begin(), content.end(), '\r'), content.end());
    while (!content.empty() && (content.back() == '\n' || content.back() == ' ' || content.back() == '\t')) {
        content.pop_back();
    }
    
    json::Value data;
    bool isBase64 = false;
    std::wstring decodeMethod;
    
    // 尝试UTF-8解析
    try {
        json::Parser parser;
        data = parser.parse(content.substr(startPos));
        decodeMethod = isUtf8Bom ? L"UTF-8 BOM格式" : L"UTF-8格式";
    } catch (...) {
        // 尝试Base64解码
        try {
            std::wstring decoded = Base64Decode(content);
            if (decoded.empty()) {
                // 尝试GBK编码
                try {
                    json::Parser parser;
                    std::wstring wcontent = GbkToWstring(content.substr(startPos));
                    data = parser.parse(WstringToUtf8(wcontent));
                    decodeMethod = L"GBK格式";
                } catch (...) {
                    ShowError(g_hMainWnd, L"无法解析文件内容\n支持的格式：UTF-8、UTF-8 BOM、GBK、Base64编码JSON");
                    return;
                }
            } else {
                json::Parser parser;
                data = parser.parse(WstringToUtf8(decoded));
                isBase64 = true;
                decodeMethod = L"Base64编码格式";
            }
        } catch (...) {
            // 最后尝试GBK
            try {
                json::Parser parser;
                std::wstring wcontent = GbkToWstring(content.substr(startPos));
                data = parser.parse(WstringToUtf8(wcontent));
                decodeMethod = L"GBK格式";
            } catch (...) {
                ShowError(g_hMainWnd, L"无法解析文件内容\n支持的格式：UTF-8、UTF-8 BOM、GBK、Base64编码JSON");
                return;
            }
        }
    }
    
    if (data.has("groups")) {
        auto& groups = data["groups"];
        if (!groups.is_object()) {
            ShowError(g_hMainWnd, L"JSON格式错误：groups字段必须是对象");
            return;
        }
        
        bool isStatusFormat = false;
        for (const auto& name : groups.keys()) {
            auto& g = groups[name];
            if (g.has("status_map") || g.has("status_name")) {
                isStatusFormat = true;
                break;
            }
        }
        
        ProcessGroupedData(groups, isStatusFormat, decodeMethod);
    } else {
        ProcessUngroupedData(data, decodeMethod);
    }
}

static void ProcessGroupedData(const json::Value& groups, bool isStatusFormat, const std::wstring& decodeMethod) {
    if (isStatusFormat) {
        g_statusGroups.clear();
        
        for (const auto& name : groups.keys()) {
            auto& g = groups[name];
            if (!g.is_object()) continue;
            
            std::wstring gname = Utf8ToWstring(name);
            std::map<std::wstring, std::wstring> statusMap;
            
            if (g.has("status_map") && g["status_map"].is_object()) {
                auto& sm = g["status_map"];
                for (const auto& key : sm.keys()) {
                    statusMap[Utf8ToWstring(key)] = SafeJsonString(sm[key]);
                }
            } else if (g.has("first_word") && g["first_word"].is_object()) {
                auto& fw = g["first_word"];
                for (const auto& key : fw.keys()) {
                    statusMap[Utf8ToWstring(key)] = SafeJsonString(fw[key]);
                }
            }
            
            if (!statusMap.empty()) {
                g_statusGroups[gname] = statusMap;
            }
        }
        
        if (!g_statusGroups.empty()) {
            LoadStatusGroup(g_statusGroups.begin()->first);
            std::wstring msg = L"成功导入 " + IntToWstr(g_statusGroups.size()) + L" 个状态字分组！(" + decodeMethod + L")";
            ShowInfo(g_hMainWnd, msg.c_str());
        }
    } else {
        g_groups.clear();
        
        for (const auto& name : groups.keys()) {
            auto& g = groups[name];
            if (!g.is_object()) continue;
            
            GroupData gd;
            if (g.has("first_name")) gd.first_name = SafeJsonString(g["first_name"], L"第一个字");
            if (g.has("second_name")) gd.second_name = SafeJsonString(g["second_name"], L"第二个字");
            if (g.has("displayon")) gd.word_displayon = SafeJsonStringUtf8(g["displayon"], "0");
            
            if (g.has("first_word") && g["first_word"].is_object()) {
                auto& fw = g["first_word"];
                for (const auto& key : fw.keys()) {
                    int bit = 0;
                    if (!TryParseInt(Utf8ToWstring(key), bit)) continue;
                    
                    auto& info = fw[key];
                    BitInfo bi;
                    if (info.is_string()) {
                        bi.description = SafeJsonString(info);
                    } else if (info.is_object()) {
                        if (info.has("description")) bi.description = SafeJsonString(info["description"]);
                        if (info.has("display_on")) bi.display_on = SafeJsonStringUtf8(info["display_on"], "1");
                    }
                    gd.first_word[bit] = bi;
                }
            }
            
            if (g.has("second_word") && g["second_word"].is_object()) {
                auto& sw = g["second_word"];
                for (const auto& key : sw.keys()) {
                    int bit = 0;
                    if (!TryParseInt(Utf8ToWstring(key), bit)) continue;
                    
                    auto& info = sw[key];
                    BitInfo bi;
                    if (info.is_string()) {
                        bi.description = SafeJsonString(info);
                    } else if (info.is_object()) {
                        if (info.has("description")) bi.description = SafeJsonString(info["description"]);
                        if (info.has("display_on")) bi.display_on = SafeJsonStringUtf8(info["display_on"], "1");
                    }
                    gd.second_word[bit] = bi;
                }
            }
            
            g_groups[Utf8ToWstring(name)] = gd;
        }
        
        if (!g_groups.empty()) {
            LoadGroup(g_groups.begin()->first);
            std::wstring msg = L"成功导入 " + IntToWstr(g_groups.size()) + L" 个分组！(" + decodeMethod + L")";
            ShowInfo(g_hMainWnd, msg.c_str());
        }
    }
}

static void ProcessUngroupedData(const json::Value& data, const std::wstring& decodeMethod) {
    bool isStatusFormat = data.is_object();
    for (const auto& key : data.keys()) {
        if (!data[key].is_string()) {
            isStatusFormat = false;
            break;
        }
    }
    
    if (isStatusFormat && !data.has("first_word") && !data.has("second_word")) {
        g_statusWordMap.clear();
        for (const auto& key : data.keys()) {
            g_statusWordMap[Utf8ToWstring(key)] = SafeJsonString(data[key]);
        }
        if (!g_statusWordMap.empty()) {
            LoadStatusWord(g_statusWordMap.begin()->first);
        }
    } else {
        GroupData gd;
        if (data.has("first_name")) gd.first_name = SafeJsonString(data["first_name"], L"第一个字");
        if (data.has("second_name")) gd.second_name = SafeJsonString(data["second_name"], L"第二个字");
        if (data.has("displayon")) gd.word_displayon = SafeJsonStringUtf8(data["displayon"], "0");
        
        if (data.has("first_word") && data["first_word"].is_object()) {
            auto& fw = data["first_word"];
            for (const auto& key : fw.keys()) {
                int bit = 0;
                if (!TryParseInt(Utf8ToWstring(key), bit)) continue;
                
                auto& info = fw[key];
                BitInfo bi;
                if (info.is_string()) {
                    bi.description = SafeJsonString(info);
                } else if (info.is_object()) {
                    if (info.has("description")) bi.description = SafeJsonString(info["description"]);
                    if (info.has("display_on")) bi.display_on = SafeJsonStringUtf8(info["display_on"], "1");
                }
                gd.first_word[bit] = bi;
            }
        }
        
        if (data.has("second_word") && data["second_word"].is_object()) {
            auto& sw = data["second_word"];
            for (const auto& key : sw.keys()) {
                int bit = 0;
                if (!TryParseInt(Utf8ToWstring(key), bit)) continue;
                
                auto& info = sw[key];
                BitInfo bi;
                if (info.is_string()) {
                    bi.description = SafeJsonString(info);
                } else if (info.is_object()) {
                    if (info.has("description")) bi.description = SafeJsonString(info["description"]);
                    if (info.has("display_on")) bi.display_on = SafeJsonStringUtf8(info["display_on"], "1");
                }
                gd.second_word[bit] = bi;
            }
        }
        
        g_groups[L"默认"] = gd;
        LoadGroup(L"默认");
        ShowInfo(g_hMainWnd, (L"打包数据对应表导入完成！(" + decodeMethod + L")").c_str());
    }
}

//=============================================================================
// 数据加载函数
//=============================================================================

static void LoadGroup(const std::wstring& groupName) {
    auto it = g_groups.find(groupName);
    if (it == g_groups.end()) {
        ShowError(g_hMainWnd, (L"分组 '" + groupName + L"' 不存在").c_str());
        return;
    }
    
    g_currentGroup = groupName;
    g_firstName = it->second.first_name;
    g_secondName = it->second.second_name;
    g_firstWord = it->second.first_word;
    g_secondWord = it->second.second_word;
    g_wordDisplayon = it->second.word_displayon;
    
    // 更新UI
    HWND hPage = g_hTabPages[0];
    SetText(GetDlgItem(hPage, IDC_GROUP_LABEL), L"当前选择的打包字：" + groupName);
    SetText(GetDlgItem(hPage, IDC_FIRST_NAME), g_firstName);
    SetText(GetDlgItem(hPage, IDC_SECOND_NAME), g_secondName);
    
    HWND hResult = GetDlgItem(hPage, IDC_RESULT_TEXT);
    ClearText(hResult);
    AppendText(hResult, L"已加载打包字：" + groupName + L"\r\n");
    AppendText(hResult, L"第一个字名称：" + g_firstName + L"\r\n");
    AppendText(hResult, L"第二个字名称：" + g_secondName + L"\r\n");
    AppendText(hResult, L"════════════════════════\r\n");
    AppendText(hResult, L"请在输入框中输入十进制数值进行查询\r\n");
    
    // 强制刷新UI
    InvalidateRect(hPage, nullptr, TRUE);
    UpdateWindow(hPage);
}

static void LoadStatusGroup(const std::wstring& groupName) {
    auto it = g_statusGroups.find(groupName);
    if (it == g_statusGroups.end()) {
        ShowError(g_hMainWnd, (L"状态字分组 '" + groupName + L"' 不存在").c_str());
        return;
    }
    
    g_currentStatusGroup = groupName;
    g_statusWordMap = it->second;
    
    HWND hPage = g_hTabPages[1];
    SetText(GetDlgItem(hPage, IDC_STATUS_LABEL), L"当前选择的状态字分组：" + groupName);
    
    HWND hResult = GetDlgItem(hPage, IDC_STATUS_RESULT);
    ClearText(hResult);
    AppendText(hResult, L"已加载状态字分组：" + groupName + L"\r\n");
    AppendText(hResult, L"════════════════════════\r\n");
    AppendText(hResult, L"状态字映射列表：\r\n\r\n");
    
    std::vector<std::wstring> keys;
    for (const auto& p : g_statusWordMap) keys.push_back(p.first);
    std::sort(keys.begin(), keys.end(), [](const std::wstring& a, const std::wstring& b) {
        int ia = 0, ib = 0;
        if (TryParseInt(a, ia) && TryParseInt(b, ib)) return ia < ib;
        return a < b;
    });
    
    for (const auto& key : keys) {
        AppendText(hResult, L"        \"" + key + L"\": \"" + g_statusWordMap[key] + L"\",\r\n");
    }
}

static void LoadStatusWord(const std::wstring& statusKey) {
    auto it = g_statusWordMap.find(statusKey);
    if (it == g_statusWordMap.end()) {
        ShowError(g_hMainWnd, (L"状态字 '" + statusKey + L"' 不存在").c_str());
        return;
    }
    
    g_currentStatusWord = statusKey;
    
    HWND hPage = g_hTabPages[1];
    if (!g_currentStatusGroup.empty()) {
        SetText(GetDlgItem(hPage, IDC_STATUS_LABEL), L"当前选择的状态字分组：" + g_currentStatusGroup);
    } else {
        SetText(GetDlgItem(hPage, IDC_STATUS_LABEL), L"当前选择的状态字：" + statusKey);
    }
    
    HWND hResult = GetDlgItem(hPage, IDC_STATUS_RESULT);
    ClearText(hResult);
    AppendText(hResult, L"已加载状态字：" + statusKey + L"\r\n");
    AppendText(hResult, L"状态值：" + statusKey + L"\r\n");
    AppendText(hResult, L"描述：" + it->second + L"\r\n");
    AppendText(hResult, L"════════════════════════\r\n");
}

//=============================================================================
// 功能函数
//=============================================================================

static void ImportJson() {
    wchar_t filePath[MAX_PATH] = {0};
    
    // 获取exe所在目录作为初始目录
    wchar_t exePath[MAX_PATH] = {0};
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);
    std::wstring exeDir = exePath;
    size_t lastSlash = exeDir.find_last_of(L"\\/");
    if (lastSlash != std::wstring::npos) {
        exeDir = exeDir.substr(0, lastSlash);
    }
    
    OPENFILENAMEW ofn = {sizeof(ofn)};
    ofn.hwndOwner = g_hMainWnd;
    ofn.lpstrFilter = L"JSON文件\0*.json\0所有文件\0*.*\0";
    ofn.lpstrFile = filePath;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrInitialDir = exeDir.c_str();
    ofn.lpstrTitle = L"选择配置JSON文件";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    
    if (GetOpenFileNameW(&ofn)) {
        LoadJsonData(filePath);
    }
}

static void CheckFault(int wordType) {
    HWND hPage = g_hTabPages[0];
    HWND hResult = GetDlgItem(hPage, IDC_RESULT_TEXT);
    ClearText(hResult);
    
    HWND hEntry = GetDlgItem(hPage, wordType == 1 ? IDC_FIRST_ENTRY : IDC_SECOND_ENTRY);
    std::wstring inputStr = GetText(hEntry);
    
    int num = 0;
    if (!TryParseInt(inputStr, num)) {
        AppendText(hResult, L"错误：请输入有效的十进制整数\r\n");
        return;
    }
    
    auto& faultMap = wordType == 1 ? g_firstWord : g_secondWord;
    std::wstring title = wordType == 1 ? L"第一个字检查结果" : L"第二个字检查结果";
    
    int maxKey = 0;
    for (const auto& p : faultMap) {
        if (p.first > maxKey) maxKey = p.first;
    }
    
    std::wstring binary = DecimalToBinary(num, maxKey + 1);
    std::wstring reversed = binary;
    std::reverse(reversed.begin(), reversed.end());
    std::wstring formatted = FormatBinary(reversed);
    
    AppendText(hResult, title + L"\r\n");
    AppendText(hResult, L"十进制值：" + IntToWstr(num) + L"\r\n");
    AppendText(hResult, L"二进制表示（低位在前）：" + formatted + L"\r\n");
    AppendText(hResult, L"════════════════════════\r\n");
    AppendText(hResult, L"你查询的结果如下：\r\n");
    
    bool hasMatch = false;
    for (size_t bitPos = 0; bitPos < reversed.size(); ++bitPos) {
        auto it = faultMap.find((int)bitPos);
        if (it == faultMap.end()) continue;
        
        const BitInfo& info = it->second;
        if (info.description.empty()) continue;
        
        wchar_t currentBit = reversed[bitPos];
        std::string displayOn = info.display_on.empty() ? "1" : info.display_on;
        
        bool isMatch = (currentBit == L'1' && displayOn == "1") || 
                       (currentBit == L'0' && displayOn == "0");
        
        if (g_wordDisplayon == "1") {
            if (isMatch) {
                AppendText(hResult, L"  • 第" + IntToWstr(bitPos) + L"位: " + info.description + L"\r\n");
                hasMatch = true;
            }
        } else {
            if (currentBit == (wchar_t)(displayOn[0])) {
                AppendText(hResult, L"  • 第" + IntToWstr(bitPos) + L"位: " + info.description + L"\r\n");
                hasMatch = true;
            }
        }
    }
    
    if (!hasMatch) {
        AppendText(hResult, L"无对应数据位或打包数据表未导入\r\n");
    }
}

static void ShowPopupMenu(const std::vector<std::wstring>& items, bool isStatus) {
    HWND hBtn = GetDlgItem(g_hMainWnd, IDC_BTN_LIST);
    RECT rc;
    GetWindowRect(hBtn, &rc);
    
    HMENU hMenu = CreatePopupMenu();
    for (size_t i = 0; i < items.size() && i < 100; ++i) {
        AppendMenuW(hMenu, MF_STRING, IDM_POPUP_BASE + (int)i, items[i].c_str());
    }
    
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_LEFTALIGN, rc.left, rc.bottom, 0, g_hMainWnd, nullptr);
    DestroyMenu(hMenu);
    
    int idx = cmd - IDM_POPUP_BASE;
    if (idx >= 0 && idx < (int)items.size()) {
        if (isStatus) {
            if (!g_statusGroups.empty()) {
                LoadStatusGroup(items[idx]);
            } else {
                LoadStatusWord(items[idx]);
            }
        } else {
            LoadGroup(items[idx]);
        }
    }
}

static void ShowGroupMenu() {
    int tabIdx = TabCtrl_GetCurSel(g_hTab);
    
    if (tabIdx == 0) {
        if (g_groups.empty()) {
            ShowInfo(g_hMainWnd, L"请先导入打包数据JSON配置文件");
            return;
        }
        std::vector<std::wstring> items;
        for (const auto& p : g_groups) items.push_back(p.first);
        ShowPopupMenu(items, false);
    } else if (tabIdx == 1) {
        if (g_statusGroups.empty() && g_statusWordMap.empty()) {
            ShowInfo(g_hMainWnd, L"请先导入状态字JSON配置文件");
            return;
        }
        std::vector<std::wstring> items;
        if (!g_statusGroups.empty()) {
            for (const auto& p : g_statusGroups) items.push_back(p.first);
        } else {
            for (const auto& p : g_statusWordMap) items.push_back(p.first);
        }
        ShowPopupMenu(items, true);
    } else {
        ShowInfo(g_hMainWnd, L"请在'打包数据'或'状态字'标签页中使用列表功能");
    }
}

static void OpenCalculator() {
    ShellExecuteW(nullptr, L"open", L"calc.exe", nullptr, nullptr, SW_SHOWNORMAL);
}

static void ShowAbout() {
    MessageBoxW(g_hMainWnd,
        L"自动化计算工具-王国强-202603\n\n"
        L"版本：3.0 (C++ Win32版)\n"
        L"开发者：德龙轧钢自动化团队\n\n"
        L"本软件用于查看和分析打包数据及状态字，\n"
        L"支持导入JSON格式的配置文件，\n"
        L"帮助工程师快速定位和诊断设备状态。\n\n"
        L"功能说明：\n"
        L"• 导入：导入JSON格式的打包数据或状态字配置文件\n"
        L"• 列表：选择不同的打包字分组或状态字分组\n"
        L"• 计算器：打开系统计算器\n"
        L"• 关于：显示本帮助信息\n\n"
        L"支持的文件格式：\n"
        L"• 明文JSON格式\n"
        L"• Base64编码JSON格式\n\n"
        L"© 2026 德龙轧钢自动化团队",
        L"关于", MB_OK | MB_ICONINFORMATION);
}

static void TogglePinTop() {
    g_isPinned = !g_isPinned;
    SetWindowPos(g_hMainWnd, g_isPinned ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    HWND hBtn = GetDlgItem(g_hMainWnd, IDC_BTN_PIN);
    SetText(hBtn, g_isPinned ? L"  取消置顶  " : L"  置顶  ");
}

//=============================================================================
// 内存映象网计算
//=============================================================================

static void CreateMemoryInputFields() {
    if (!g_hMemInputFrame) return;
    
    // 销毁旧的输入框
    for (auto& p : g_memEntries) {
        if (p.second) DestroyWindow(p.second);
    }
    g_memEntries.clear();
    
    // 销毁旧的标签
    HWND hChild = GetWindow(g_hMemInputFrame, GW_CHILD);
    while (hChild) {
        HWND hNext = GetWindow(hChild, GW_HWNDNEXT);
        DestroyWindow(hChild);
        hChild = hNext;
    }
    
    // 强制刷新避免重影
    InvalidateRect(g_hMemInputFrame, nullptr, TRUE);
    UpdateWindow(g_hMemInputFrame);
    
    int y = 25;
    
    struct FieldInfo {
        std::wstring label;
        int id;
    };
    
    std::vector<FieldInfo> fields;
    
    switch (g_memCalcType) {
        case 1:
            fields = {{L"初始内存映象网地址", 100}, {L"初始R地址", 101}, {L"终止R地址", 102}};
            break;
        case 2:
            fields = {{L"初始内存映象网地址", 100}, {L"终止内存映象网地址", 101}, {L"初始R地址", 102}};
            break;
        case 3:
            fields = {{L"初始内存映象网地址", 100}, {L"初始m地址", 101}, {L"终止m地址", 102}};
            break;
        case 4:
            fields = {{L"初始内存映象网地址", 100}, {L"终止内存映象网地址", 101}, {L"初始m地址", 102}, {L"位编号", 103}};
            break;
    }
    
    for (const auto& f : fields) {
        CreateStatic(g_hMemInputFrame, f.label.c_str(), 15, y, 130, 20);
        
        HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", 
            WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 
            150, y, 150, 22, g_hMemInputFrame, nullptr, g_hInst, nullptr);
        if (hEdit && g_hFont) SendMessageW(hEdit, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        g_memEntries[f.id] = hEdit;
        
        y += 35;
    }
}

static void CalculateMemory() {
    HWND hPage = g_hTabPages[2];
    HWND hResult = GetDlgItem(hPage, IDC_MEM_RESULT);
    ClearText(hResult);
    
    auto getVal = [](int id) -> std::wstring {
        auto it = g_memEntries.find(id);
        if (it == g_memEntries.end() || !it->second) return L"";
        return GetText(it->second);
    };
    
    std::wstring result;
    
    switch (g_memCalcType) {
        case 1: {
            std::wstring initMem = getVal(100);
            std::wstring initR = getVal(101);
            std::wstring endR = getVal(102);
            
            if (initMem.empty() || initR.empty() || endR.empty()) {
                AppendText(hResult, L"错误：所有输入框必须填写");
                return;
            }
            
            int initMemVal = 0, initRVal = 0, endRVal = 0;
            if (!TryParseHex(initMem, initMemVal)) {
                AppendText(hResult, L"错误：初始内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseInt(initR, initRVal)) {
                AppendText(hResult, L"错误：初始R地址不是有效的整数");
                return;
            }
            if (!TryParseInt(endR, endRVal)) {
                AppendText(hResult, L"错误：终止R地址不是有效的整数");
                return;
            }
            
            int offset = (endRVal - initRVal) * 2;
            int endMemVal = initMemVal + offset;
            result = L"终止内存映象网地址: " + IntToHex(endMemVal);
            break;
        }
        case 2: {
            std::wstring initMem = getVal(100);
            std::wstring endMem = getVal(101);
            std::wstring initR = getVal(102);
            
            if (initMem.empty() || endMem.empty() || initR.empty()) {
                AppendText(hResult, L"错误：所有输入框必须填写");
                return;
            }
            
            int initMemVal = 0, endMemVal = 0, initRVal = 0;
            if (!TryParseHex(initMem, initMemVal)) {
                AppendText(hResult, L"错误：初始内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseHex(endMem, endMemVal)) {
                AppendText(hResult, L"错误：终止内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseInt(initR, initRVal)) {
                AppendText(hResult, L"错误：初始R地址不是有效的整数");
                return;
            }
            
            int offset = (endMemVal - initMemVal) / 2;
            int endRVal = initRVal + offset;
            result = L"终止R地址: " + IntToWstr(endRVal);
            break;
        }
        case 3: {
            std::wstring initMem = getVal(100);
            std::wstring initM = getVal(101);
            std::wstring endM = getVal(102);
            
            if (initMem.empty() || initM.empty() || endM.empty()) {
                AppendText(hResult, L"错误：所有输入框必须填写");
                return;
            }
            
            int initMemVal = 0, initMVal = 0, endMVal = 0;
            if (!TryParseHex(initMem, initMemVal)) {
                AppendText(hResult, L"错误：初始内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseInt(initM, initMVal)) {
                AppendText(hResult, L"错误：初始m地址不是有效的整数");
                return;
            }
            if (!TryParseInt(endM, endMVal)) {
                AppendText(hResult, L"错误：终止m地址不是有效的整数");
                return;
            }
            
            int offset = (endMVal - initMVal) / 8;
            int bitPos = (endMVal - initMVal) % 8;
            int endMemVal = initMemVal + offset;
            result = L"终止内存映象网地址: " + IntToHex(endMemVal) + L", 位编号: " + IntToWstr(bitPos);
            break;
        }
        case 4: {
            std::wstring initMem = getVal(100);
            std::wstring endMem = getVal(101);
            std::wstring initM = getVal(102);
            std::wstring bitPos = getVal(103);
            
            if (initMem.empty() || endMem.empty() || initM.empty() || bitPos.empty()) {
                AppendText(hResult, L"错误：所有输入框必须填写");
                return;
            }
            
            int initMemVal = 0, endMemVal = 0, initMVal = 0, bitPosVal = 0;
            if (!TryParseHex(initMem, initMemVal)) {
                AppendText(hResult, L"错误：初始内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseHex(endMem, endMemVal)) {
                AppendText(hResult, L"错误：终止内存映象网地址不是有效的16进制");
                return;
            }
            if (!TryParseInt(initM, initMVal)) {
                AppendText(hResult, L"错误：初始m地址不是有效的整数");
                return;
            }
            if (!TryParseInt(bitPos, bitPosVal)) {
                AppendText(hResult, L"错误：位编号不是有效的整数");
                return;
            }
            
            int offset = (endMemVal - initMemVal) * 8;
            int endMVal = initMVal + offset + bitPosVal;
            result = L"终止m地址: " + IntToWstr(endMVal);
            break;
        }
    }
    
    AppendText(hResult, result);
}

//=============================================================================
// 模拟量计算
//=============================================================================

static void CalculateAnalog() {
    HWND hPage = g_hTabPages[3];
    HWND hResult = GetDlgItem(hPage, IDC_ANALOG_RESULT);
    
    auto getVal = [hPage](int id) -> double {
        std::wstring s = GetText(GetDlgItem(hPage, id));
        double val = 0;
        return TryParseDouble(s, val) ? val : 0;
    };
    
    double rawLow = getVal(IDC_ANALOG_RAW_LOW);
    double rawHigh = getVal(IDC_ANALOG_RAW_HIGH);
    double engLow = getVal(IDC_ANALOG_ENG_LOW);
    double engHigh = getVal(IDC_ANALOG_ENG_HIGH);
    
    double rawRange = rawHigh - rawLow;
    double engRange = engHigh - engLow;
    
    // 精确的浮点数比较，避免除零
    const double EPSILON = 1e-9;
    if (std::abs(rawRange) < EPSILON) {
        SetText(hResult, L"错误：原始值范围无效（上限等于下限）");
        return;
    }
    if (std::abs(engRange) < EPSILON) {
        SetText(hResult, L"错误：工程量范围无效（上限等于下限）");
        return;
    }
    if (rawRange < 0 || engRange < 0) {
        SetText(hResult, L"错误：上限必须大于下限");
        return;
    }
    
    double raw = getVal(IDC_ANALOG_RAW);
    double eng = getVal(IDC_ANALOG_ENG);
    double current = getVal(IDC_ANALOG_CURRENT);
    
    std::wstring result;
    
    if (raw != 0) {
        if (raw < rawLow || raw > rawHigh) {
            result = L"原始值超出范围！";
        } else {
            double engVal = (raw - rawLow) * engRange / rawRange + engLow;
            double curVal = 4 + (raw - rawLow) * 16 / rawRange;
            result = L"电流：" + DoubleToWstr(curVal) + L" mA\r\n";
            result += L"原始值：" + IntToWstr((int)raw) + L"\r\n";
            result += L"工程量：" + DoubleToWstr(engVal);
        }
    } else if (eng != 0) {
        if (eng < engLow || eng > engHigh) {
            result = L"工程量超出范围！";
        } else {
            double rawVal = (eng - engLow) * rawRange / engRange + rawLow;
            double curVal = 4 + (eng - engLow) * 16 / engRange;
            result = L"电流：" + DoubleToWstr(curVal) + L" mA\r\n";
            result += L"原始值：" + IntToWstr((int)rawVal) + L"\r\n";
            result += L"工程量：" + DoubleToWstr(eng);
        }
    } else if (current != 0) {
        if (current < 4 || current > 20) {
            result = L"电流超出4-20mA范围！";
        } else {
            double rawVal = rawLow + (current - 4) * rawRange / 16;
            double engVal = engLow + (current - 4) * engRange / 16;
            result = L"电流：" + DoubleToWstr(current) + L" mA\r\n";
            result += L"原始值：" + IntToWstr((int)rawVal) + L"\r\n";
            result += L"工程量：" + DoubleToWstr(engVal);
        }
    } else {
        result = L"请输入原始值、工程量或电流进行计算";
    }
    
    SetText(hResult, result);
}

static void ResetAnalog() {
    HWND hPage = g_hTabPages[3];
    SetText(GetDlgItem(hPage, IDC_ANALOG_RAW), L"");
    SetText(GetDlgItem(hPage, IDC_ANALOG_ENG), L"");
    SetText(GetDlgItem(hPage, IDC_ANALOG_CURRENT), L"");
    SetText(GetDlgItem(hPage, IDC_ANALOG_RESULT), L"请输入原始值、工程量或电流进行计算");
}

//=============================================================================
// 速度转换
//=============================================================================

static void CalculateSpeed() {
    HWND hPage = g_hTabPages[4];
    HWND hResult = GetDlgItem(hPage, IDC_SPEED_RESULT);
    
    std::wstring dStr = GetText(GetDlgItem(hPage, IDC_SPEED_DIAMETER));
    std::wstring iStr = GetText(GetDlgItem(hPage, IDC_SPEED_RATIO));
    std::wstring inputStr = GetText(GetDlgItem(hPage, IDC_SPEED_INPUT_VALUE));
    
    if (dStr.empty() || iStr.empty() || inputStr.empty()) {
        SetText(hResult, L"错误：请填写所有参数");
        return;
    }
    
    double d = 0, i = 1.0, inputVal = 0;
    
    if (!TryParseDouble(dStr, d) || d <= 0) {
        SetText(hResult, L"错误：辊直径必须是正数");
        return;
    }
    
    // 解析减速比
    size_t pos = iStr.find(L'(');
    if (pos != std::wstring::npos) {
        if (!TryParseDouble(iStr.substr(0, pos), i)) {
            SetText(hResult, L"错误：减速比格式无效");
            return;
        }
    } else {
        if (!TryParseDouble(iStr, i)) {
            SetText(hResult, L"错误：减速比必须是数字");
            return;
        }
    }
    
    if (!TryParseDouble(inputStr, inputVal) || inputVal <= 0) {
        SetText(hResult, L"错误：输入值必须是正数");
        return;
    }
    
    if (i < 1) {
        SetText(hResult, L"错误：减速比必须大于等于1");
        return;
    }
    
    HWND hUnit = GetDlgItem(hPage, IDC_SPEED_UNIT);
    int unitIdx = (int)SendMessageW(hUnit, CB_GETCURSEL, 0, 0);
    
    HWND hType = GetDlgItem(hPage, IDC_SPEED_INPUT_TYPE);
    BOOL isMotorRpm = (SendMessageW(hType, BM_GETCHECK, 0, 0) == BST_CHECKED);
    
    const double PI = 3.14159265358979323846;
    double nRoller, v;
    
    if (!isMotorRpm) {
        if (unitIdx == 1) {
            v = inputVal * 60;
        } else {
            v = inputVal;
        }
        nRoller = (v * 1000) / (PI * d);
    } else {
        double nMotor;
        if (unitIdx == 1) {
            nMotor = inputVal * 60;
        } else {
            nMotor = inputVal;
        }
        nRoller = nMotor / i;
        v = (PI * d * nRoller) / 1000;
    }
    
    double t = nRoller > 0 ? 60 / nRoller : 0;
    
    std::wstring result = L"辊转速：" + DoubleToWstr(nRoller) + L" rpm\r\n";
    result += L"线速度：" + DoubleToWstr(v) + L" m/min\r\n";
    result += L"每转时间：" + DoubleToWstr(t) + L" 秒";
    
    SetText(hResult, result);
}

static void ResetSpeed() {
    HWND hPage = g_hTabPages[4];
    SetText(GetDlgItem(hPage, IDC_SPEED_DIAMETER), L"");
    SetText(GetDlgItem(hPage, IDC_SPEED_INPUT_VALUE), L"");
    SendMessageW(GetDlgItem(hPage, IDC_SPEED_RATIO), CB_SETCURSEL, 0, 0);
    SendMessageW(GetDlgItem(hPage, IDC_SPEED_UNIT), CB_SETCURSEL, 0, 0);
    SendMessageW(GetDlgItem(hPage, IDC_SPEED_INPUT_TYPE), BM_SETCHECK, BST_CHECKED, 0);
    SetText(GetDlgItem(hPage, IDC_SPEED_RESULT), L"请输入参数后点击计算");
}

//=============================================================================
// Tab页面创建
//=============================================================================

static void CreateDataTab(HWND hParent) {
    HWND hLabel = CreateStatic(hParent, L"当前选择的打包字：未选择", 20, 10, 450, 20);
    SetWindowLongPtr(hLabel, GWLP_ID, IDC_GROUP_LABEL);
    
    CreateStatic(hParent, L"第一个字(十进制):", 20, 45, 120, 20);
    CreateEdit(hParent, IDC_FIRST_ENTRY, 140, 43, 150, 22);
    CreateButton(hParent, IDC_FIRST_QUERY, L"转换并查询", 300, 42, 90, 24);
    HWND hFirstName = CreateStatic(hParent, L"第一个字", 140, 65, 150, 16);
    SetWindowLongPtr(hFirstName, GWLP_ID, IDC_FIRST_NAME);
    
    CreateStatic(hParent, L"第二个字(十进制):", 20, 95, 120, 20);
    CreateEdit(hParent, IDC_SECOND_ENTRY, 140, 93, 150, 22);
    CreateButton(hParent, IDC_SECOND_QUERY, L"转换并查询", 300, 92, 90, 24);
    HWND hSecondName = CreateStatic(hParent, L"第二个字", 140, 115, 150, 16);
    SetWindowLongPtr(hSecondName, GWLP_ID, IDC_SECOND_NAME);
    
    CreateEdit(hParent, IDC_RESULT_TEXT, 20, 140, 470, 380, L"", true, true);
}

static void CreateStatusTab(HWND hParent) {
    HWND hLabel = CreateStatic(hParent, L"当前选择的状态字：未选择", 20, 10, 450, 20);
    SetWindowLongPtr(hLabel, GWLP_ID, IDC_STATUS_LABEL);
    
    CreateEdit(hParent, IDC_STATUS_RESULT, 20, 40, 470, 480, L"", true, true);
}

static void CreateMemoryTab(HWND hParent) {
    HWND hTypeGroup = CreateGroupBox(hParent, 0, L"计算类型", 10, 10, 150, 200);
    
    int y = 30;
    const wchar_t* types[] = {
        L"模拟量计算",
        L"  16进制地址",
        L"模拟量计算",
        L"  R地址",
        L"数字量计算",
        L"  16进制地址和位数",
        L"数字量计算",
        L"  M地址"
    };
    
    for (int i = 0; i < 4; ++i) {
        HWND hRadio = CreateRadioButton(hTypeGroup, IDC_MEM_CALC_TYPE + i, types[i*2], 10, y, 130, 18, i == 0, i == 0);
        y += 20;
        CreateStatic(hTypeGroup, types[i*2+1], 10, y, 130, 18);
        y += 30;
    }
    
    g_hMemInputFrame = CreateGroupBox(hParent, IDC_MEM_INPUT_FRAME, L"输入参数", 170, 10, 330, 160);
    CreateMemoryInputFields();
    
    CreateButton(hParent, IDC_MEM_CALC_BTN, L"开始计算", 280, 180, 100, 28);
    
    CreateGroupBox(hParent, 0, L"计算结果", 170, 220, 330, 80);
    CreateEdit(hParent, IDC_MEM_RESULT, 180, 245, 310, 50, L"", true, true);
}

static void CreateAnalogTab(HWND hParent) {
    CreateGroupBox(hParent, 0, L"量程设置", 10, 10, 490, 80);
    
    CreateStatic(hParent, L"原始值范围：", 20, 25, 80, 18);
    CreateStatic(hParent, L"原始值下限", 105, 25, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_RAW_LOW, 170, 23, 60, 22, L"5530");
    CreateStatic(hParent, L"~", 235, 25, 15, 18);
    CreateStatic(hParent, L"原始值上限", 255, 25, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_RAW_HIGH, 320, 23, 60, 22, L"27648");
    
    CreateStatic(hParent, L"工程量范围：", 20, 55, 80, 18);
    CreateStatic(hParent, L"工程量下限", 105, 55, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_ENG_LOW, 170, 53, 60, 22, L"0.0");
    CreateStatic(hParent, L"~", 235, 55, 15, 18);
    CreateStatic(hParent, L"工程量上限", 255, 55, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_ENG_HIGH, 320, 53, 60, 22, L"100.0");
    
    CreateGroupBox(hParent, 0, L"输入数值（任填一项自动计算）", 10, 100, 490, 90);
    
    CreateStatic(hParent, L"原始值：", 20, 120, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_RAW, 85, 118, 100, 22);
    
    CreateStatic(hParent, L"工程量：", 200, 120, 60, 18);
    CreateEdit(hParent, IDC_ANALOG_ENG, 265, 118, 100, 22);
    
    CreateStatic(hParent, L"电流 (mA)：", 20, 150, 70, 18);
    CreateEdit(hParent, IDC_ANALOG_CURRENT, 95, 148, 90, 22);
    
    CreateButton(hParent, IDC_ANALOG_RESET, L"重置", 200, 200, 80, 28);
    
    CreateGroupBox(hParent, 0, L"计算结果", 10, 240, 490, 100);
    CreateEdit(hParent, IDC_ANALOG_RESULT, 20, 260, 470, 70, L"请输入原始值、工程量或电流进行计算", true, true);
}

static void CreateSpeedTab(HWND hParent) {
    CreateGroupBox(hParent, 0, L"参数设置", 10, 10, 490, 150);
    
    CreateStatic(hParent, L"辊直径：", 20, 28, 60, 18);
    CreateEdit(hParent, IDC_SPEED_DIAMETER, 85, 26, 100, 22);
    CreateStatic(hParent, L"mm", 190, 28, 30, 18);
    
    CreateStatic(hParent, L"减速比：", 20, 58, 60, 18);
    HWND hRatio = CreateComboBox(hParent, IDC_SPEED_RATIO, 85, 56, 150, 200);
    SendMessageW(hRatio, CB_ADDSTRING, 0, (LPARAM)L"1");
    SendMessageW(hRatio, CB_ADDSTRING, 0, (LPARAM)L"3.25 (F1电机)");
    SendMessageW(hRatio, CB_ADDSTRING, 0, (LPARAM)L"2.24 (F2电机)");
    SendMessageW(hRatio, CB_ADDSTRING, 0, (LPARAM)L"1.44 (F3电机)");
    SendMessageW(hRatio, CB_SETCURSEL, 0, 0);
    
    CreateStatic(hParent, L"输入类型：", 20, 88, 60, 18);
    CreateRadioButton(hParent, IDC_SPEED_INPUT_TYPE, L"电机转速", 85, 86, 80, 20, true, true);
    CreateRadioButton(hParent, 0, L"线速度", 175, 86, 70, 20, false, false);
    
    CreateStatic(hParent, L"输入值：", 20, 118, 50, 18);
    CreateEdit(hParent, IDC_SPEED_INPUT_VALUE, 75, 116, 80, 22);
    
    CreateStatic(hParent, L"单位：", 165, 118, 40, 18);
    HWND hUnit = CreateComboBox(hParent, IDC_SPEED_UNIT, 210, 116, 80, 200);
    SendMessageW(hUnit, CB_ADDSTRING, 0, (LPARAM)L"rpm");
    SendMessageW(hUnit, CB_ADDSTRING, 0, (LPARAM)L"rps");
    SendMessageW(hUnit, CB_SETCURSEL, 0, 0);
    
    CreateButton(hParent, IDC_SPEED_CALC, L"计算", 150, 170, 80, 28);
    CreateButton(hParent, IDC_SPEED_RESET, L"重置", 250, 170, 80, 28);
    
    CreateGroupBox(hParent, 0, L"计算结果", 10, 210, 490, 120);
    CreateEdit(hParent, IDC_SPEED_RESULT, 20, 230, 470, 90, L"请输入参数后点击计算", true, true);
}

//=============================================================================
// Tab页面管理
//=============================================================================

static void ResizeTabPages() {
    if (!g_hTab) return;
    
    RECT rc;
    GetClientRect(g_hTab, &rc);
    TabCtrl_AdjustRect(g_hTab, FALSE, &rc);
    
    for (int i = 0; i < 5; ++i) {
        if (g_hTabPages[i]) {
            SetWindowPos(g_hTabPages[i], nullptr, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);
        }
    }
}

static void SwitchTabPage(int newIndex) {
    if (newIndex < 0 || newIndex >= 5) return;
    if (newIndex == g_currentTab) return;
    
    ShowWindow(g_hTabPages[g_currentTab], SW_HIDE);
    ShowWindow(g_hTabPages[newIndex], SW_SHOW);
    g_currentTab = newIndex;
}

//=============================================================================
// 窗口过程
//=============================================================================

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_CREATE: {
            // 初始化通用控件
            INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_TAB_CLASSES | ICC_STANDARD_CLASSES};
            InitCommonControlsEx(&icc);
            
            // 创建画刷
            g_hBrushBg = CreateSolidBrush(COLOR_BG);
            g_hBrushWhite = CreateSolidBrush(RGB(255, 255, 255));
            
            // 创建字体
            NONCLIENTMETRICSW ncm = {sizeof(ncm)};
            SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
            g_hFont = CreateFontIndirectW(&ncm.lfMessageFont);
            if (g_hFont) {
                LOGFONTW lfBold = ncm.lfMessageFont;
                lfBold.lfWeight = FW_BOLD;
                g_hFontBold = CreateFontIndirectW(&lfBold);
            }
            
            // 创建顶部按钮
            CreateButton(hWnd, IDC_BTN_IMPORT, L"  导入  ", 10, 10, 70, 28);
            CreateButton(hWnd, IDC_BTN_LIST, L"  列表  ", 90, 10, 70, 28);
            CreateButton(hWnd, IDC_BTN_CALCULATOR, L"  计算器  ", 170, 10, 80, 28);
            CreateButton(hWnd, IDC_BTN_ABOUT, L"  关于  ", 260, 10, 70, 28);
            CreateButton(hWnd, IDC_BTN_PIN, L"  置顶  ", 340, 10, 80, 28);
            
            // 创建Tab控件
            g_hTab = CreateWindowExW(0, WC_TABCONTROLW, L"", 
                WS_VISIBLE | WS_CHILD | WS_CLIPSIBLINGS, 
                5, 45, WINDOW_WIDTH - 15, WINDOW_HEIGHT - 55, 
                hWnd, (HMENU)IDC_TAB_CONTROL, g_hInst, nullptr);
            if (g_hTab && g_hFont) SendMessageW(g_hTab, WM_SETFONT, (WPARAM)g_hFont, TRUE);
            
            // 添加Tab页
            const wchar_t* tabNames[] = {
                L"  打包数据  ", 
                L"  状态字  ", 
                L"  内存映象网计算  ", 
                L"  模拟量计算  ", 
                L"  速度转换  "
            };
            for (int i = 0; i < 5; ++i) {
                TCITEMW tci = {TCIF_TEXT};
                tci.pszText = (LPWSTR)tabNames[i];
                TabCtrl_InsertItem(g_hTab, i, &tci);
            }
            
            // 创建Tab页面容器（使用WS_EX_CONTROLPARENT以支持Tab键导航）
            for (int i = 0; i < 5; ++i) {
                g_hTabPages[i] = CreateWindowExW(
                    WS_EX_CONTROLPARENT, 
                    L"STATIC", L"", 
                    WS_VISIBLE | WS_CHILD | WS_CLIPSIBLINGS, 
                    0, 0, 0, 0, 
                    hWnd, nullptr, g_hInst, nullptr
                );
            }
            
            // 创建各Tab页内容
            CreateDataTab(g_hTabPages[0]);
            CreateStatusTab(g_hTabPages[1]);
            CreateMemoryTab(g_hTabPages[2]);
            CreateAnalogTab(g_hTabPages[3]);
            CreateSpeedTab(g_hTabPages[4]);
            
            // 调整页面大小并显示第一个
            ResizeTabPages();
            ShowWindow(g_hTabPages[0], SW_SHOW);
            for (int i = 1; i < 5; ++i) ShowWindow(g_hTabPages[i], SW_HIDE);
            
            break;
        }
        
        case WM_NOTIFY: {
            LPNMHDR pnmh = (LPNMHDR)lParam;
            if (pnmh->code == TCN_SELCHANGE && pnmh->idFrom == IDC_TAB_CONTROL) {
                int sel = TabCtrl_GetCurSel(g_hTab);
                SwitchTabPage(sel);
            }
            break;
        }
        
        case WM_COMMAND: {
            int id = LOWORD(wParam);
            int code = HIWORD(wParam);
            
            // Edit控件获得焦点时全选
            if (code == EN_SETFOCUS) {
                SendMessageW((HWND)lParam, EM_SETSEL, 0, -1);
                break;
            }
            
            switch (id) {
                case IDC_BTN_IMPORT: ImportJson(); break;
                case IDC_BTN_LIST: ShowGroupMenu(); break;
                case IDC_BTN_CALCULATOR: OpenCalculator(); break;
                case IDC_BTN_ABOUT: ShowAbout(); break;
                case IDC_BTN_PIN: TogglePinTop(); break;
                
                case IDC_FIRST_QUERY: CheckFault(1); break;
                case IDC_SECOND_QUERY: CheckFault(2); break;
                
                case IDC_MEM_CALC_BTN: CalculateMemory(); break;
                
                case IDC_ANALOG_RESET: ResetAnalog(); break;
                
                case IDC_SPEED_CALC: CalculateSpeed(); break;
                case IDC_SPEED_RESET: ResetSpeed(); break;
                
                default:
                    // 处理内存计算类型选择
                    if (id >= IDC_MEM_CALC_TYPE && id < IDC_MEM_CALC_TYPE + 4) {
                        g_memCalcType = id - IDC_MEM_CALC_TYPE + 1;
                        CreateMemoryInputFields();
                    }
                    break;
            }
            break;
        }
        
        case WM_SIZE:
            ResizeTabPages();
            break;
            
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT: {
            HDC hdcStatic = (HDC)wParam;
            HWND hCtrl = (HWND)lParam;
            
            // 检查是否为只读编辑框
            LONG style = GetWindowLongW(hCtrl, GWL_STYLE);
            if (style & ES_READONLY) {
                SetBkColor(hdcStatic, RGB(255, 255, 255));
                SetTextColor(hdcStatic, RGB(0, 0, 0));
                return (INT_PTR)g_hBrushWhite;
            }
            
            // 默认背景
            SetBkColor(hdcStatic, COLOR_BG);
            return (INT_PTR)g_hBrushBg;
        }
            
        case WM_DPICHANGED: {
            // 更新DPI缩放比例
            int newDpi = HIWORD(wParam);
            g_dpiScale = MulDiv(newDpi, 100, 96);
            
            // 重新创建字体
            if (g_hFont) DeleteObject(g_hFont);
            if (g_hFontBold) DeleteObject(g_hFontBold);
            
            NONCLIENTMETRICSW ncm = {sizeof(ncm)};
            SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
            ncm.lfMessageFont.lfHeight = MulDiv(ncm.lfMessageFont.lfHeight, newDpi, 96);
            g_hFont = CreateFontIndirectW(&ncm.lfMessageFont);
            if (g_hFont) {
                LOGFONTW lfBold = ncm.lfMessageFont;
                lfBold.lfWeight = FW_BOLD;
                g_hFontBold = CreateFontIndirectW(&lfBold);
            }
            
            // 更新所有控件字体
            for (int i = 0; i < 5; ++i) {
                if (g_hTabPages[i]) {
                    EnumChildWindows(g_hTabPages[i], [](HWND hChild, LPARAM lParam) -> BOOL {
                        HFONT hFont = (HFONT)lParam;
                        SendMessageW(hChild, WM_SETFONT, (WPARAM)hFont, TRUE);
                        return TRUE;
                    }, (LPARAM)g_hFont);
                }
            }
            
            // 调整窗口大小
            RECT* pRect = (RECT*)lParam;
            SetWindowPos(hWnd, nullptr, pRect->left, pRect->top, 
                pRect->right - pRect->left, pRect->bottom - pRect->top, 
                SWP_NOZORDER | SWP_NOACTIVATE);
            break;
        }
            
        case WM_DESTROY:
            // 清理资源
            if (g_hFont) DeleteObject(g_hFont);
            if (g_hFontBold) DeleteObject(g_hFontBold);
            if (g_hBrushBg) DeleteObject(g_hBrushBg);
            if (g_hBrushWhite) DeleteObject(g_hBrushWhite);
            
            // 清理全局数据
            g_groups.clear();
            g_statusGroups.clear();
            g_statusWordMap.clear();
            g_firstWord.clear();
            g_secondWord.clear();
            g_memEntries.clear();
            
            PostQuitMessage(0);
            break;
            
        default:
            return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

//=============================================================================
// 程序入口
//=============================================================================

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    g_hInst = hInstance;
    
    // 设置DPI感知（优先使用Per-Monitor v2）
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32) {
        typedef DPI_AWARENESS_CONTEXT (WINAPI *SetThreadDpiAwarenessContextFunc)(DPI_AWARENESS_CONTEXT);
        SetThreadDpiAwarenessContextFunc setDpiContext = 
            (SetThreadDpiAwarenessContextFunc)GetProcAddress(hUser32, "SetThreadDpiAwarenessContext");
        if (setDpiContext) {
            setDpiContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }
    }
    
    // 备用方案：使用shcore.dll
    HMODULE hShCore = LoadLibraryW(L"shcore.dll");
    if (hShCore) {
        typedef HRESULT (WINAPI *SetProcessDpiAwarenessFunc)(int);
        SetProcessDpiAwarenessFunc setDpiAwareness = 
            (SetProcessDpiAwarenessFunc)GetProcAddress(hShCore, "SetProcessDpiAwareness");
        if (setDpiAwareness) {
            setDpiAwareness(2); // PROCESS_PER_MONITOR_DPI_AWARE
        }
        FreeLibrary(hShCore);
    }
    
    // 获取初始DPI缩放比例
    HDC hdc = GetDC(nullptr);
    if (hdc) {
        g_dpiScale = MulDiv(GetDeviceCaps(hdc, LOGPIXELSX), 100, 96);
        ReleaseDC(nullptr, hdc);
    }
    
    // 注册窗口类
    WNDCLASSEXW wcex = {sizeof(wcex)};
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wcex.lpszClassName = L"AutoCalcTool";
    wcex.hIconSm = LoadIcon(nullptr, IDI_APPLICATION);
    RegisterClassExW(&wcex);
    
    // 创建主窗口（使用DPI缩放后的尺寸）
    int scaledWidth = Scale(WINDOW_WIDTH);
    int scaledHeight = Scale(WINDOW_HEIGHT);
    
    g_hMainWnd = CreateWindowExW(0, L"AutoCalcTool", WINDOW_TITLE,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, scaledWidth, scaledHeight,
        nullptr, nullptr, hInstance, nullptr);
    
    if (!g_hMainWnd) return FALSE;
    
    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);
    
    // 消息循环（支持Tab键导航）
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(g_hMainWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    
    return (int)msg.wParam;
}
