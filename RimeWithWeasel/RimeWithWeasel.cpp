/******************************************************************************
 * RimeWithWeasel.cpp - 智能输入法核心处理器实现
 *
 * 本文件实现了RimeWithWeaselHandler类，是智能输入法(Weasel)的核心组件。
 * 负责处理与Rime引擎的交互、UI更新、候选词管理、AI助手功能等。
 *
 * 主要功能模块：
 * 1. Rime引擎初始化与管理
 * 2. 按键事件处理与响应
 * 3. 候选词列表管理
 * 4. 焦点管理（焦点进入/离开）
 * 5. UI更新与渲染
 * 6. AI助手面板功能（可选）
 *
 * C++ 语法要点：
 * - std::function: 函数包装器，用于回调
 * - std::unique_ptr/std::shared_ptr: 智能指针管理资源
 * - lambda表达式: 匿名函数用于回调
 * - std::map/std::vector: STL容器
 * - Windows API: WinHTTP, COM, Window消息处理
 ******************************************************************************/

#include "stdafx.h"
#include <logging.h>
#include <RimeWithWeasel.h>
#include <StringAlgorithm.hpp>
#include <WeaselConstants.h>
#include <WeaselUtility.h>
#include <winhttp.h>
#include <shellapi.h>

#include <cctype>
#include <cstdlib>
#include <cwctype>
#include <array>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <map>
#include <random>
#include <regex>
#include <thread>
#include <vector>
#include <rime_api.h>

#if __has_include("../third_party/webview2/pkg/build/native/include/WebView2.h")
#include "../third_party/webview2/pkg/build/native/include/WebView2.h"
#define WEASEL_HAS_WEBVIEW2 1
#else
#define WEASEL_HAS_WEBVIEW2 0
#endif

#if WEASEL_HAS_WEBVIEW2
#include <wrl.h>
#include <wrl/client.h>
#endif

#ifndef RAPIDJSON_NOMINMAX
#define RAPIDJSON_NOMINMAX
#endif
#ifndef _SILENCE_CXX17_ITERATOR_BASE_CLASS_DEPRECATION_WARNING
#define _SILENCE_CXX17_ITERATOR_BASE_CLASS_DEPRECATION_WARNING
#endif
#ifdef Bool
#pragma push_macro("Bool")
#undef Bool
#define WEASEL_RESTORE_RIME_BOOL_MACRO
#endif
#include "../librime/deps/opencc/deps/rapidjson-1.1.0/rapidjson/document.h"
#ifdef WEASEL_RESTORE_RIME_BOOL_MACRO
#pragma pop_macro("Bool")
#undef WEASEL_RESTORE_RIME_BOOL_MACRO
#endif

#pragma comment(lib, "winhttp.lib")

#define TRANSPARENT_COLOR 0x00000000
#define ARGB2ABGR(value)                                 \
  ((value & 0xff000000) | ((value & 0x000000ff) << 16) | \
   (value & 0x0000ff00) | ((value & 0x00ff0000) >> 16))
#define RGBA2ABGR(value)                                   \
  (((value & 0xff) << 24) | ((value & 0xff000000) >> 24) | \
   ((value & 0x00ff0000) >> 8) | ((value & 0x0000ff00) << 8))
typedef enum { COLOR_ABGR = 0, COLOR_ARGB, COLOR_RGBA } ColorFormat;

using namespace weasel;

static RimeApi* rime_api;
WeaselSessionId _GenerateNewWeaselSessionId(SessionStatusMap sm, DWORD pid) {
  if (sm.empty())
    return (WeaselSessionId)(pid + 1);
  return (WeaselSessionId)(sm.rbegin()->first + 1);
}

int expand_ibus_modifier(int m) {
  return (m & 0xff) | ((m & 0xff00) << 16);
}

namespace {

#if WEASEL_HAS_WEBVIEW2
using Microsoft::WRL::Callback;
using Microsoft::WRL::ComPtr;
using CreateCoreWebView2EnvironmentWithOptionsFunc = HRESULT(STDAPICALLTYPE*)(
    PCWSTR,
    PCWSTR,
    ICoreWebView2EnvironmentOptions*,
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
#endif

constexpr wchar_t kAIPanelWindowClass[] = L"WeaselAIAssistantPanelWindow";
constexpr int kAIPanelWidth = 540;
constexpr int kAIPanelHeight = 360;
constexpr int kAIPanelMinWidth = 180;
constexpr int kAIPanelMaxWidth = 860;
constexpr int kAIPanelMinHeight = 220;
constexpr int kAIPanelMaxHeight = 560;
constexpr int kAIPanelScreenMargin = 8;
constexpr int kAIPanelPadding = 0;
constexpr int kAIPanelCornerRadius = 12;
constexpr int kAIPanelButtonWidth = 96;
constexpr int kAIPanelButtonHeight = 30;
constexpr UINT kAIPanelControlStatus = 4101;
constexpr UINT kAIPanelControlOutput = 4102;
constexpr UINT kAIPanelControlRequest = 4103;
constexpr UINT kAIPanelControlConfirm = 4104;
constexpr UINT kAIPanelControlCancel = 4105;
constexpr UINT WM_AI_STREAM_APPEND = WM_APP + 401;
constexpr UINT WM_AI_STREAM_DONE = WM_APP + 402;
constexpr UINT WM_AI_STREAM_ERROR = WM_APP + 403;
constexpr UINT WM_AI_WEBVIEW_INIT = WM_APP + 404;
constexpr UINT WM_AI_PANEL_DESTROY = WM_APP + 405;
constexpr UINT WM_AI_WEBVIEW_SYNC = WM_APP + 406;
constexpr UINT WM_AI_PANEL_DRAG = WM_APP + 407;
constexpr char kDefaultAIAssistantPrompt[] =
    "Continue the user's existing text. Reply with continuation only, with no "
    "explanation, no markdown, and no quotation marks.";
constexpr wchar_t kSystemCommandCommitPrefix[] = L"__weasel_syscmd__:";
constexpr const char* kAllowedSystemCommandIds[] = {
    "jsq",       "calc",     "notepad",   "mspaint", "explorer",
    "txt",       "md",       "gh",        "bd",      "wb",
    "g",         "yt",       "rili",      "calendar", "cal", "kb"};
struct AIPanelTextMessage {
  uint64_t request_id = 0;
  std::wstring text;
};

struct AIPanelDoneMessage {
  uint64_t request_id = 0;
  bool has_error = false;
  std::wstring text;
};

struct AIPanelUiCommand {
  std::string type;
  std::wstring text;
  std::wstring institution_id;
  int panel_width = 0;
  int panel_height = 0;
  std::string resize_reason;
};

bool IsAllowedSystemCommandId(const std::string& command_id) {
  for (const char* allowed_id : kAllowedSystemCommandIds) {
    if (command_id == allowed_id) {
      return true;
    }
  }
  return false;
}

bool TryParseSystemCommandMarker(const std::wstring& commit_text,
                                 std::string* command_id) {
  if (!command_id) {
    return false;
  }
  const std::wstring prefix(kSystemCommandCommitPrefix);
  if (!starts_with(commit_text, prefix) || commit_text.size() <= prefix.size()) {
    return false;
  }
  const std::string candidate_id = wtou8(commit_text.substr(prefix.size()));
  if (!IsAllowedSystemCommandId(candidate_id)) {
    return false;
  }
  *command_id = candidate_id;
  return true;
}

std::wstring ExpandEnvironmentVariables(const std::wstring& input) {
  if (input.empty()) {
    return input;
  }
  const DWORD required = ExpandEnvironmentStringsW(input.c_str(), nullptr, 0);
  if (required == 0) {
    return input;
  }
  std::wstring expanded(required, L'\0');
  const DWORD written = ExpandEnvironmentStringsW(
      input.c_str(), expanded.data(), static_cast<DWORD>(expanded.size()));
  if (written == 0) {
    return input;
  }
  expanded.resize(written > 0 ? written - 1 : 0);
  return expanded;
}

void TryDisableAIPanelWindowBorder(HWND hwnd) {
  if (!hwnd) {
    return;
  }
  HMODULE dwmapi = LoadLibraryW(L"dwmapi.dll");
  if (!dwmapi) {
    return;
  }
  using DwmSetWindowAttributeFn =
      HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
  auto set_window_attribute = reinterpret_cast<DwmSetWindowAttributeFn>(
      GetProcAddress(dwmapi, "DwmSetWindowAttribute"));
  if (set_window_attribute) {
    // 仅隐藏标题栏按钮区域，保留 resize 边框
    constexpr DWORD kDWMWAAttribute = 2;  // DWMWA_EXTENDED_FRAME_BOUNDS
    set_window_attribute(hwnd, kDWMWAAttribute, nullptr, 0);
  }
  FreeLibrary(dwmapi);
}

bool TryApplyAIPanelDwmRoundedCorner(HWND hwnd) {
  if (!hwnd) {
    return false;
  }
  HMODULE dwmapi = LoadLibraryW(L"dwmapi.dll");
  if (!dwmapi) {
    return false;
  }
  using DwmSetWindowAttributeFn =
      HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
  auto set_window_attribute = reinterpret_cast<DwmSetWindowAttributeFn>(
      GetProcAddress(dwmapi, "DwmSetWindowAttribute"));
  if (!set_window_attribute) {
    FreeLibrary(dwmapi);
    return false;
  }
  constexpr DWORD kDWMWAWindowCornerPreference = 33;
  constexpr DWORD kDWMWCPRound = 2;
  const DWORD corner_preference = kDWMWCPRound;
  const HRESULT hr =
      set_window_attribute(hwnd, kDWMWAWindowCornerPreference,
                           &corner_preference, sizeof(corner_preference));
  FreeLibrary(dwmapi);
  return SUCCEEDED(hr);
}

void ApplyAIPanelRoundedRegion(HWND hwnd) {
  if (!hwnd || !IsWindow(hwnd)) {
    return;
  }
  if (TryApplyAIPanelDwmRoundedCorner(hwnd)) {
    // Prefer DWM compositor rounded corners to avoid aliasing artifacts.
    SetWindowRgn(hwnd, nullptr, TRUE);
    return;
  }
  RECT rect = {0};
  if (!GetClientRect(hwnd, &rect)) {
    return;
  }
  const int width = rect.right - rect.left;
  const int height = rect.bottom - rect.top;
  if (width <= 0 || height <= 0) {
    return;
  }
  HRGN region = CreateRoundRectRgn(0, 0, width + 1, height + 1,
                                   kAIPanelCornerRadius * 2,
                                   kAIPanelCornerRadius * 2);
  if (!region) {
    return;
  }
  if (!SetWindowRgn(hwnd, region, TRUE)) {
    DeleteObject(region);
  }
}

std::wstring ResolveAIAssistantDumpPath(const AIAssistantConfig& config) {
  std::wstring path = u8tow(config.debug_dump_path);
  if (path.empty()) {
    path = L"ai_context_dump.txt";
  }
  std::filesystem::path fs_path(path);
  if (!fs_path.is_absolute()) {
    fs_path = WeaselUserDataPath() / fs_path;
  }
  return fs_path.wstring();
}

void AppendAIAssistantContextDump(const AIAssistantConfig& config,
                                  const std::string& source,
                                  HWND target_hwnd,
                                  const std::wstring& context) {
  if (!config.debug_dump_context) {
    return;
  }

  const std::wstring dump_path = ResolveAIAssistantDumpPath(config);
  std::ofstream output(std::filesystem::path(dump_path),
                       std::ios::binary | std::ios::app);
  if (!output.is_open()) {
    return;
  }

  SYSTEMTIME now = {0};
  GetLocalTime(&now);
  char header[256] = {0};
  _snprintf_s(header, sizeof(header), _TRUNCATE,
              "[%04u-%02u-%02u %02u:%02u:%02u.%03u] source=%s hwnd=0x%p "
              "chars=%zu\n",
              static_cast<unsigned>(now.wYear),
              static_cast<unsigned>(now.wMonth),
              static_cast<unsigned>(now.wDay),
              static_cast<unsigned>(now.wHour),
              static_cast<unsigned>(now.wMinute),
              static_cast<unsigned>(now.wSecond),
              static_cast<unsigned>(now.wMilliseconds), source.c_str(),
              target_hwnd, context.size());
  output << header;
  output << wtou8(context) << "\n\n";
}

void LayoutAIPanelControls(HWND hwnd) {
  RECT client_rect = {0};
  if (!GetClientRect(hwnd, &client_rect)) {
    return;
  }
  HWND status_hwnd = GetDlgItem(hwnd, kAIPanelControlStatus);
  HWND output_hwnd = GetDlgItem(hwnd, kAIPanelControlOutput);
  HWND request_hwnd = GetDlgItem(hwnd, kAIPanelControlRequest);
  HWND confirm_hwnd = GetDlgItem(hwnd, kAIPanelControlConfirm);
  HWND cancel_hwnd = GetDlgItem(hwnd, kAIPanelControlCancel);
  if (!output_hwnd) {
    return;
  }
  const int client_width = client_rect.right - client_rect.left;
  const int client_height = client_rect.bottom - client_rect.top;
  const bool hide_native_header_footer =
      status_hwnd && request_hwnd && confirm_hwnd && cancel_hwnd &&
      !IsWindowVisible(status_hwnd) && !IsWindowVisible(request_hwnd) &&
      !IsWindowVisible(confirm_hwnd) && !IsWindowVisible(cancel_hwnd);
  if (hide_native_header_footer) {
    MoveWindow(output_hwnd, kAIPanelPadding, kAIPanelPadding,
               max(1, client_width - kAIPanelPadding * 2),
               max(1, client_height - kAIPanelPadding * 2), TRUE);
    return;
  }
  const int button_gap = 8;
  const int status_height = 22;
  const int output_top = kAIPanelPadding + status_height + 8;
  const int button_top = client_height - kAIPanelPadding - kAIPanelButtonHeight;
  const int output_height = max(80, button_top - output_top - 10);
  const int button_y = client_height - kAIPanelPadding - kAIPanelButtonHeight;
  const int cancel_x =
      client_width - kAIPanelPadding - kAIPanelButtonWidth;
  const int confirm_x = cancel_x - button_gap - kAIPanelButtonWidth;
  const int request_x = confirm_x - button_gap - kAIPanelButtonWidth;

  if (status_hwnd) {
    MoveWindow(status_hwnd, kAIPanelPadding, kAIPanelPadding,
               client_width - kAIPanelPadding * 2, status_height, TRUE);
  }
  MoveWindow(output_hwnd, kAIPanelPadding, output_top,
             client_width - kAIPanelPadding * 2, output_height, TRUE);
  if (request_hwnd) {
    MoveWindow(request_hwnd, request_x, button_y, kAIPanelButtonWidth,
               kAIPanelButtonHeight, TRUE);
  }
  if (confirm_hwnd) {
    MoveWindow(confirm_hwnd, confirm_x, button_y, kAIPanelButtonWidth,
               kAIPanelButtonHeight, TRUE);
  }
  if (cancel_hwnd) {
    MoveWindow(cancel_hwnd, cancel_x, button_y, kAIPanelButtonWidth,
               kAIPanelButtonHeight, TRUE);
  }
}

#if WEASEL_HAS_WEBVIEW2
CreateCoreWebView2EnvironmentWithOptionsFunc ResolveWebView2Factory() {
  static HMODULE loader = nullptr;
  static CreateCoreWebView2EnvironmentWithOptionsFunc factory = nullptr;
  static bool initialized = false;
  if (initialized) {
    return factory;
  }
  initialized = true;

  loader = LoadLibraryW(L"WebView2Loader.dll");
  if (!loader) {
    wchar_t module_path[MAX_PATH] = {0};
    if (GetModuleFileNameW(nullptr, module_path, MAX_PATH) > 0) {
      const fs::path exe_dir = fs::path(module_path).parent_path();
      const fs::path local_loader =
          exe_dir /
          L"..\\third_party\\webview2\\pkg\\build\\native\\x64\\WebView2Loader.dll";
      loader = LoadLibraryW(local_loader.wstring().c_str());
    }
  }
  if (!loader) {
    return nullptr;
  }

  factory = reinterpret_cast<CreateCoreWebView2EnvironmentWithOptionsFunc>(
      GetProcAddress(loader, "CreateCoreWebView2EnvironmentWithOptions"));
  return factory;
}
#endif

struct ScopedWinHttpHandle {
  ScopedWinHttpHandle() : handle(nullptr) {}
  explicit ScopedWinHttpHandle(HINTERNET value) : handle(value) {}
  ~ScopedWinHttpHandle() {
    if (handle) {
      WinHttpCloseHandle(handle);
    }
  }

  operator HINTERNET() const { return handle; }
  HINTERNET get() const { return handle; }
  bool valid() const { return handle != nullptr; }

  HINTERNET handle;
};

bool ReadConfigString(RimeConfig* config, const char* path, std::string* value) {
  if (!config || !value) {
    return false;
  }
  const int kBufferSize = 4096;
  char buffer[kBufferSize] = {0};
  if (!rime_api->config_get_string(config, path, buffer, kBufferSize - 1)) {
    return false;
  }
  *value = buffer;
  return true;
}

bool ReadConfigInt(RimeConfig* config, const char* path, int* value) {
  if (!config || !value) {
    return false;
  }
  return !!rime_api->config_get_int(config, path, value);
}

std::string EscapeJsonString(const std::string& input) {
  std::string escaped;
  escaped.reserve(input.size() + input.size() / 4);
  static const char kHex[] = "0123456789abcdef";
  for (unsigned char ch : input) {
    switch (ch) {
      case '\\':
        escaped += "\\\\";
        break;
      case '"':
        escaped += "\\\"";
        break;
      case '\b':
        escaped += "\\b";
        break;
      case '\f':
        escaped += "\\f";
        break;
      case '\n':
        escaped += "\\n";
        break;
      case '\r':
        escaped += "\\r";
        break;
      case '\t':
        escaped += "\\t";
        break;
      default:
        if (ch < 0x20) {
          escaped += "\\u00";
          escaped.push_back(kHex[(ch >> 4) & 0x0f]);
          escaped.push_back(kHex[ch & 0x0f]);
        } else {
          escaped.push_back(static_cast<char>(ch));
        }
        break;
    }
  }
  return escaped;
}

std::string BuildAIAssistantUserInput(const std::wstring& context_text) {
  std::string prompt =
      "Context:\n" + wtou8(context_text) +
      "\n\nContinue writing from the end of that context.";
  return prompt;
}

std::string BuildAIAssistantRequestBody(const AIAssistantConfig& config,
                                        const std::wstring& context_text,
                                        bool stream) {
  const std::string system_prompt =
      config.system_prompt.empty() ? kDefaultAIAssistantPrompt
                                   : config.system_prompt;
  std::string body;
  body.reserve(1024 + context_text.size() * 3);
  body += "{\"model\":\"";
  body += EscapeJsonString(config.model);
  body += "\"";
  if (!config.reasoning_effort.empty()) {
    body += ",\"reasoning\":{\"effort\":\"";
    body += EscapeJsonString(config.reasoning_effort);
    body += "\"}";
  }
  body += ",\"input\":[";
  body += "{\"role\":\"system\",\"content\":[{\"type\":\"input_text\",\"text\":\"";
  body += EscapeJsonString(system_prompt);
  body += "\"}]},";
  body += "{\"role\":\"user\",\"content\":[{\"type\":\"input_text\",\"text\":\"";
  body += EscapeJsonString(BuildAIAssistantUserInput(context_text));
  body += "\"}]}";
  body += stream ? "],\"stream\":true}" : "],\"stream\":false}";
  return body;
}

void AppendResponseText(const rapidjson::Value& value, std::string* text) {
  if (!text) {
    return;
  }
  if (value.IsString()) {
    text->append(value.GetString(), value.GetStringLength());
    return;
  }
  if (!value.IsObject() && !value.IsArray()) {
    return;
  }
  if (value.IsObject()) {
    auto member = value.FindMember("text");
    if (member != value.MemberEnd() && member->value.IsString()) {
      text->append(member->value.GetString(), member->value.GetStringLength());
    }
    member = value.FindMember("output_text");
    if (member != value.MemberEnd()) {
      AppendResponseText(member->value, text);
    }
    member = value.FindMember("content");
    if (member != value.MemberEnd()) {
      AppendResponseText(member->value, text);
    }
    member = value.FindMember("answer");
    if (member != value.MemberEnd()) {
      AppendResponseText(member->value, text);
    }
    member = value.FindMember("message");
    if (member != value.MemberEnd()) {
      AppendResponseText(member->value, text);
    }
    return;
  }
  for (auto i = value.Begin(); i != value.End(); ++i) {
    AppendResponseText(*i, text);
  }
}

bool ParseAIAssistantResponse(const std::string& response_body,
                              std::wstring* output_text,
                              std::string* error_message) {
  rapidjson::Document document;
  document.Parse(response_body.c_str(), response_body.size());
  if (document.HasParseError()) {
    if (error_message) {
      *error_message = "AI response is not valid JSON.";
    }
    return false;
  }
  if (document.IsObject()) {
    auto error = document.FindMember("error");
    if (error != document.MemberEnd() && error->value.IsObject()) {
      auto message = error->value.FindMember("message");
      if (message != error->value.MemberEnd() && message->value.IsString()) {
        if (error_message) {
          *error_message = message->value.GetString();
        }
        return false;
      }
    }
  }

  std::string output_utf8;
  if (document.IsObject()) {
    auto output_text_member = document.FindMember("output_text");
    if (output_text_member != document.MemberEnd()) {
      AppendResponseText(output_text_member->value, &output_utf8);
    }
    if (output_utf8.empty()) {
      auto output_member = document.FindMember("output");
      if (output_member != document.MemberEnd()) {
        AppendResponseText(output_member->value, &output_utf8);
      }
    }
    if (output_utf8.empty()) {
      auto choices_member = document.FindMember("choices");
      if (choices_member != document.MemberEnd()) {
        AppendResponseText(choices_member->value, &output_utf8);
      }
    }
    if (output_utf8.empty()) {
      auto answer_member = document.FindMember("answer");
      if (answer_member != document.MemberEnd()) {
        AppendResponseText(answer_member->value, &output_utf8);
      }
    }
    if (output_utf8.empty()) {
      auto data_member = document.FindMember("data");
      if (data_member != document.MemberEnd()) {
        AppendResponseText(data_member->value, &output_utf8);
      }
    }
  }

  if (output_utf8.empty()) {
    if (error_message) {
      *error_message = "AI response did not contain any text output.";
    }
    return false;
  }

  if (output_text) {
    *output_text = u8tow(output_utf8);
  }
  return true;
}

bool InvokeAIAssistant(const AIAssistantConfig& config,
                       const std::wstring& context_text,
                       std::wstring* output_text,
                       std::string* error_message) {
  if (config.endpoint.empty()) {
    if (error_message) {
      *error_message = "AI endpoint is empty.";
    }
    return false;
  }

  std::wstring endpoint = u8tow(config.endpoint);
  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
    if (error_message) {
      *error_message = "Unable to parse AI endpoint URL.";
    }
    return false;
  }

  const std::wstring host(parts.lpszHostName, parts.dwHostNameLength);
  std::wstring object_name =
      parts.dwUrlPathLength > 0
          ? std::wstring(parts.lpszUrlPath, parts.dwUrlPathLength)
          : std::wstring(L"/");
  if (parts.dwExtraInfoLength > 0) {
    object_name.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);
  }

  ScopedWinHttpHandle session(
      WinHttpOpen(L"WeaselAIAssistant/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                  WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
  if (!session.valid()) {
    if (error_message) {
      *error_message = "WinHTTP session creation failed.";
    }
    return false;
  }
  WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                     config.timeout_ms, config.timeout_ms);

  ScopedWinHttpHandle connection(
      WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
  if (!connection.valid()) {
    if (error_message) {
      *error_message = "WinHTTP connection failed.";
    }
    return false;
  }

  const DWORD request_flags =
      parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
  ScopedWinHttpHandle request(
      WinHttpOpenRequest(connection.get(), L"POST", object_name.c_str(), NULL,
                         WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                         request_flags));
  if (!request.valid()) {
    if (error_message) {
      *error_message = "WinHTTP request creation failed.";
    }
    return false;
  }

  std::wstring headers =
      L"Content-Type: application/json\r\nAccept: application/json\r\n";
  if (!config.api_key.empty()) {
    headers += L"Authorization: Bearer ";
    headers += u8tow(config.api_key);
    headers += L"\r\n";
  }
  WinHttpAddRequestHeaders(request.get(), headers.c_str(),
                           static_cast<DWORD>(-1),
                           WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

  std::string request_body =
      BuildAIAssistantRequestBody(config, context_text, false);
  if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                          request_body.empty()
                              ? WINHTTP_NO_REQUEST_DATA
                              : (LPVOID)request_body.data(),
                          static_cast<DWORD>(request_body.size()),
                          static_cast<DWORD>(request_body.size()), 0) ||
      !WinHttpReceiveResponse(request.get(), nullptr)) {
    if (error_message) {
      *error_message = "Sending the AI request failed.";
    }
    return false;
  }

  DWORD status_code = 0;
  DWORD status_code_size = sizeof(status_code);
  WinHttpQueryHeaders(request.get(),
                      WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                      WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                      &status_code_size, WINHTTP_NO_HEADER_INDEX);

  std::string response_body;
  for (;;) {
    DWORD available = 0;
    if (!WinHttpQueryDataAvailable(request.get(), &available)) {
      if (error_message) {
        *error_message = "Failed to read the AI response.";
      }
      return false;
    }
    if (available == 0) {
      break;
    }
    std::string chunk(available, '\0');
    DWORD downloaded = 0;
    if (!WinHttpReadData(request.get(), chunk.data(), available, &downloaded)) {
      if (error_message) {
        *error_message = "Failed while downloading the AI response.";
      }
      return false;
    }
    chunk.resize(downloaded);
    response_body += chunk;
  }

  std::string parsed_error;
  if (status_code < 200 || status_code >= 300) {
    if (!ParseAIAssistantResponse(response_body, nullptr, &parsed_error) ||
        parsed_error.empty()) {
      parsed_error = "AI request returned HTTP " + std::to_string(status_code) +
                     ".";
    }
    if (error_message) {
      *error_message = parsed_error;
    }
    return false;
  }

  if (!ParseAIAssistantResponse(response_body, output_text, &parsed_error)) {
    if (error_message) {
      *error_message = parsed_error;
    }
    return false;
  }

  return true;
}

bool ParseAIAssistantStreamEvent(const std::string& event_json,
                                 std::wstring* delta_text,
                                 bool* is_complete) {
  if (delta_text) {
    delta_text->clear();
  }
  if (is_complete) {
    *is_complete = false;
  }
  if (event_json.empty()) {
    return true;
  }

  rapidjson::Document document;
  document.Parse(event_json.c_str(), event_json.size());
  if (document.HasParseError() || !document.IsObject()) {
    return false;
  }

  const auto type_member = document.FindMember("type");
  if (type_member != document.MemberEnd() && type_member->value.IsString()) {
    const std::string event_type = type_member->value.GetString();
    if (event_type == "response.completed" || event_type == "done") {
      if (is_complete) {
        *is_complete = true;
      }
      return true;
    }
  }
  const auto event_member = document.FindMember("event");
  if (event_member != document.MemberEnd() && event_member->value.IsString()) {
    const std::string event_type = event_member->value.GetString();
    if (event_type == "message_end" || event_type == "done") {
      if (is_complete) {
        *is_complete = true;
      }
      return true;
    }
  }

  std::string delta_utf8;
  const auto delta_member = document.FindMember("delta");
  if (delta_member != document.MemberEnd() && delta_member->value.IsString()) {
    delta_utf8.append(delta_member->value.GetString(),
                      delta_member->value.GetStringLength());
  }
  if (delta_utf8.empty()) {
    const auto text_member = document.FindMember("text");
    if (text_member != document.MemberEnd() && text_member->value.IsString()) {
      delta_utf8.append(text_member->value.GetString(),
                        text_member->value.GetStringLength());
    }
  }
  if (delta_utf8.empty()) {
    const auto answer_member = document.FindMember("answer");
    if (answer_member != document.MemberEnd() && answer_member->value.IsString()) {
      delta_utf8.append(answer_member->value.GetString(),
                        answer_member->value.GetStringLength());
    }
  }
  if (delta_utf8.empty()) {
    const auto output_text_member = document.FindMember("output_text");
    if (output_text_member != document.MemberEnd()) {
      AppendResponseText(output_text_member->value, &delta_utf8);
    }
  }

  if (!delta_utf8.empty() && delta_text) {
    *delta_text = u8tow(delta_utf8);
  }
  return true;
}

bool InvokeAIAssistantStream(
    const AIAssistantConfig& config,
    const std::wstring& context_text,
    const std::function<void(const std::wstring&)>& on_delta,
    std::string* error_message) {
  if (config.endpoint.empty()) {
    if (error_message) {
      *error_message = "AI endpoint is empty.";
    }
    return false;
  }

  std::wstring endpoint = u8tow(config.endpoint);
  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
    if (error_message) {
      *error_message = "Unable to parse AI endpoint URL.";
    }
    return false;
  }

  const std::wstring host(parts.lpszHostName, parts.dwHostNameLength);
  std::wstring object_name =
      parts.dwUrlPathLength > 0
          ? std::wstring(parts.lpszUrlPath, parts.dwUrlPathLength)
          : std::wstring(L"/");
  if (parts.dwExtraInfoLength > 0) {
    object_name.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);
  }

  ScopedWinHttpHandle session(
      WinHttpOpen(L"WeaselAIAssistant/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                  WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
  if (!session.valid()) {
    if (error_message) {
      *error_message = "WinHTTP session creation failed.";
    }
    return false;
  }
  WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                     config.timeout_ms, config.timeout_ms);

  ScopedWinHttpHandle connection(
      WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
  if (!connection.valid()) {
    if (error_message) {
      *error_message = "WinHTTP connection failed.";
    }
    return false;
  }

  const DWORD request_flags =
      parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
  ScopedWinHttpHandle request(
      WinHttpOpenRequest(connection.get(), L"POST", object_name.c_str(), NULL,
                         WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                         request_flags));
  if (!request.valid()) {
    if (error_message) {
      *error_message = "WinHTTP request creation failed.";
    }
    return false;
  }

  std::wstring headers = config.stream
                             ? L"Content-Type: application/json\r\nAccept: "
                               L"text/event-stream\r\n"
                             : L"Content-Type: application/json\r\nAccept: "
                               L"application/json\r\n";
  if (!config.api_key.empty()) {
    headers += L"Authorization: Bearer ";
    headers += u8tow(config.api_key);
    headers += L"\r\n";
  }
  WinHttpAddRequestHeaders(request.get(), headers.c_str(),
                           static_cast<DWORD>(-1),
                           WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

  std::string request_body =
      BuildAIAssistantRequestBody(config, context_text, config.stream);
  if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                          request_body.empty()
                              ? WINHTTP_NO_REQUEST_DATA
                              : (LPVOID)request_body.data(),
                          static_cast<DWORD>(request_body.size()),
                          static_cast<DWORD>(request_body.size()), 0) ||
      !WinHttpReceiveResponse(request.get(), nullptr)) {
    if (error_message) {
      *error_message = "Sending the AI request failed.";
    }
    return false;
  }

  DWORD status_code = 0;
  DWORD status_code_size = sizeof(status_code);
  WinHttpQueryHeaders(request.get(),
                      WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                      WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                      &status_code_size, WINHTTP_NO_HEADER_INDEX);

  std::string body;
  body.reserve(2048);
  std::string line_buffer;
  std::string event_payload;
  line_buffer.reserve(1024);
  bool saw_stream_frame = false;
  const bool should_parse_stream =
      config.stream && status_code >= 200 && status_code < 300;

  const auto flush_event_payload = [&](const std::string& payload) -> bool {
    if (payload.empty()) {
      return true;
    }
    saw_stream_frame = true;
    if (payload == "[DONE]") {
      return true;
    }
    std::wstring delta;
    if (!ParseAIAssistantStreamEvent(payload, &delta, nullptr)) {
      return false;
    }
    if (!delta.empty() && on_delta) {
      on_delta(delta);
    }
    return true;
  };

  const auto append_data_line = [&](const std::string& line) {
    std::string data_part = line.substr(5);
    size_t trim_index = 0;
    while (trim_index < data_part.size() &&
           std::isspace(static_cast<unsigned char>(data_part[trim_index]))) {
      ++trim_index;
    }
    if (trim_index > 0) {
      data_part.erase(0, trim_index);
    }
    if (!event_payload.empty()) {
      event_payload.push_back('\n');
    }
    event_payload += data_part;
  };

  for (;;) {
    DWORD available = 0;
    if (!WinHttpQueryDataAvailable(request.get(), &available)) {
      if (error_message) {
        *error_message = "Failed to read the AI response.";
      }
      return false;
    }
    if (available == 0) {
      break;
    }
    std::string chunk(available, '\0');
    DWORD downloaded = 0;
    if (!WinHttpReadData(request.get(), chunk.data(), available, &downloaded)) {
      if (error_message) {
        *error_message = "Failed while downloading the AI response.";
      }
      return false;
    }
    chunk.resize(downloaded);
    body += chunk;

    if (!should_parse_stream) {
      continue;
    }

    for (char ch : chunk) {
      if (ch == '\r') {
        continue;
      }
      if (ch != '\n') {
        line_buffer.push_back(ch);
        continue;
      }

      if (line_buffer.empty()) {
        if (!event_payload.empty()) {
          if (!flush_event_payload(event_payload)) {
            if (error_message) {
              *error_message = "AI stream event payload is invalid.";
            }
            return false;
          }
          event_payload.clear();
        }
      } else if (line_buffer.rfind("data:", 0) == 0) {
        append_data_line(line_buffer);
      }
      line_buffer.clear();
    }
  }

  if (should_parse_stream) {
    if (!line_buffer.empty() && line_buffer.rfind("data:", 0) == 0) {
      append_data_line(line_buffer);
    }
    if (!event_payload.empty() && !flush_event_payload(event_payload)) {
      if (error_message) {
        *error_message = "AI stream tail payload is invalid.";
      }
      return false;
    }
  }

  std::string parsed_error;
  if (status_code < 200 || status_code >= 300) {
    if (!ParseAIAssistantResponse(body, nullptr, &parsed_error) ||
        parsed_error.empty()) {
      parsed_error = "AI request returned HTTP " + std::to_string(status_code) +
                     ".";
    }
    if (error_message) {
      *error_message = parsed_error;
    }
    return false;
  }

  if (should_parse_stream && saw_stream_frame) {
    return true;
  }

  std::wstring output_text;
  if (!ParseAIAssistantResponse(body, &output_text, &parsed_error)) {
    if (error_message) {
      *error_message = parsed_error;
    }
    return false;
  }
  if (!output_text.empty() && on_delta) {
    on_delta(output_text);
  }
  return true;
}

std::wstring CaptureCurrentPreeditText(RimeSessionId session_id) {
  RIME_STRUCT(RimeContext, ctx);
  std::wstring preedit_text;
  if (rime_api->get_context(session_id, &ctx)) {
    if (ctx.commit_text_preview) {
      preedit_text = u8tow(ctx.commit_text_preview);
    } else if (ctx.composition.preedit) {
      preedit_text = u8tow(ctx.composition.preedit);
    }
    rime_api->free_context(&ctx);
  }
  return preedit_text;
}

std::wstring NormalizePreeditForAsciiToggle(const std::wstring& preedit_text) {
  if (preedit_text.empty()) {
    return preedit_text;
  }
  std::wstring normalized;
  normalized.reserve(preedit_text.size());
  for (wchar_t ch : preedit_text) {
    // Preedit can contain segmentation separators; strip all whitespace before
    // committing raw ascii text on Ctrl+Space toggle.
    if (std::iswspace(static_cast<wint_t>(ch)) || ch == 0x00A0 ||
        ch == 0x2007 || ch == 0x202F || ch == 0x3000) {
      continue;
    }
    normalized.push_back(ch);
  }
  return normalized;
}

std::wstring CaptureRawCompositionPreeditText(RimeSessionId session_id) {
  RIME_STRUCT(RimeContext, ctx);
  std::wstring preedit_text;
  if (rime_api->get_context(session_id, &ctx)) {
    if (ctx.composition.preedit && ctx.composition.length > 0) {
      preedit_text = u8tow(ctx.composition.preedit);
    }
    rime_api->free_context(&ctx);
  }
  return NormalizePreeditForAsciiToggle(preedit_text);
}

bool IsAIAssistantTriggerKey(const AIAssistantConfig& config,
                             const KeyEvent& key_event) {
  if (!config.enabled) {
    return false;
  }
  const auto normalize_keycode = [](UINT keycode) {
    if (keycode >= 'A' && keycode <= 'Z') {
      return keycode - 'A' + 'a';
    }
    if (keycode == ibus::grave) {
      return static_cast<UINT>('`');
    }
    return keycode;
  };
  if (normalize_keycode(key_event.keycode) !=
      normalize_keycode(config.trigger_keycode)) {
    return false;
  }
  constexpr UINT kHotkeyModifiers = ibus::CONTROL_MASK | ibus::SHIFT_MASK |
                                    ibus::ALT_MASK | ibus::SUPER_MASK |
                                    ibus::HYPER_MASK | ibus::META_MASK;
  const UINT modifiers = key_event.mask & kHotkeyModifiers;
  if ((modifiers & config.trigger_modifiers) != config.trigger_modifiers) {
    return false;
  }
  return (modifiers & (~config.trigger_modifiers)) == 0;
}

bool IsCtrlSpaceToggleKey(const KeyEvent& key_event) {
  if (key_event.keycode != ibus::space) {
    return false;
  }
  const auto modifiers = key_event.mask & ibus::MODIFIER_MASK;
  const bool has_ctrl = (modifiers & ibus::CONTROL_MASK) != 0;
  const bool has_other_modifiers =
      (modifiers & (ibus::SHIFT_MASK | ibus::ALT_MASK | ibus::SUPER_MASK |
                    ibus::HYPER_MASK | ibus::META_MASK)) != 0;
  return has_ctrl && !has_other_modifiers;
}

std::string BuildInputContentPreviewForLog(const std::wstring& text,
                                           size_t max_chars = 80) {
  if (text.empty()) {
    return std::string();
  }
  const size_t take = min(max_chars, text.size());
  const std::string utf8 = wtou8(text.substr(0, take));
  std::string escaped;
  escaped.reserve(utf8.size() + utf8.size() / 4);
  for (char ch : utf8) {
    switch (ch) {
      case '\\':
        escaped += "\\\\";
        break;
      case '\'':
        escaped += "\\'";
        break;
      case '\r':
        escaped += "\\r";
        break;
      case '\n':
        escaped += "\\n";
        break;
      case '\t':
        escaped += "\\t";
        break;
      default:
        escaped.push_back(ch);
        break;
    }
  }
  if (text.size() > take) {
    escaped += "...";
  }
  return escaped;
}

std::mutex g_input_content_log_mutex;

bool AppendLineToFile(const fs::path& file, const std::string& line) {
  HANDLE file_handle = CreateFileW(
      file.c_str(), GENERIC_WRITE,
      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
      OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file_handle == INVALID_HANDLE_VALUE) {
    return false;
  }
  SetFilePointer(file_handle, 0, nullptr, FILE_END);
  DWORD written = 0;
  const BOOL ok = WriteFile(file_handle, line.data(),
                            static_cast<DWORD>(line.size()), &written, nullptr);
  CloseHandle(file_handle);
  return ok != FALSE && written == line.size();
}

bool IsRimeInfoLogFileName(const std::wstring& file_name) {
  if (file_name.empty()) {
    return false;
  }
  const std::wstring prefix = L"rime.weasel.";
  const std::wstring marker = L".log.INFO.";
  const std::wstring suffix = L".log";
  if (file_name.size() <= prefix.size() + marker.size() + suffix.size()) {
    return false;
  }
  if (file_name.compare(0, prefix.size(), prefix) != 0) {
    return false;
  }
  if (file_name.find(marker) == std::wstring::npos) {
    return false;
  }
  if (file_name.compare(file_name.size() - suffix.size(), suffix.size(),
                        suffix) != 0) {
    return false;
  }
  return true;
}

bool IsCurrentProcessInfoLog(const std::wstring& file_name, DWORD pid) {
  const std::wstring pid_token = L"." + std::to_wstring(pid) + L".log";
  return file_name.find(pid_token) != std::wstring::npos;
}

fs::path FindLatestRimeInfoLogFileForCurrentProcess(const fs::path& log_dir) {
  std::error_code ec;
  fs::path latest_file;
  fs::file_time_type latest_time;
  bool found = false;
  const DWORD pid = GetCurrentProcessId();
  for (fs::directory_iterator it(log_dir, ec), end; !ec && it != end;
       it.increment(ec)) {
    const fs::directory_entry& entry = *it;
    if (!entry.is_regular_file(ec) || ec) {
      ec.clear();
      continue;
    }
    const std::wstring file_name = entry.path().filename().wstring();
    if (!IsRimeInfoLogFileName(file_name) ||
        !IsCurrentProcessInfoLog(file_name, pid)) {
      continue;
    }
    const fs::file_time_type write_time = entry.last_write_time(ec);
    if (ec) {
      ec.clear();
      continue;
    }
    if (!found || write_time > latest_time) {
      latest_time = write_time;
      latest_file = entry.path();
      found = true;
    }
  }
  return latest_file;
}

void AppendInputContentInfoLogLine(const std::string& message) {
  std::unique_lock<std::mutex> lock(g_input_content_log_mutex, std::try_to_lock);
  if (!lock.owns_lock()) {
    return;
  }
  const fs::path log_dir = WeaselLogPath();
  std::error_code ec;
  fs::create_directories(log_dir, ec);
  const fs::path dedicated_log_file =
      log_dir / L"rime.weasel.input_content.INFO.log";
  SYSTEMTIME st = {0};
  GetLocalTime(&st);
  char time_buffer[64] = {0};
  _snprintf_s(time_buffer, _countof(time_buffer), _TRUNCATE,
              "%04d-%02d-%02d %02d:%02d:%02d.%03d", st.wYear, st.wMonth,
              st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
  const std::string line =
      std::string(time_buffer) + " [INFO] " + message + "\r\n";
  if (AppendLineToFile(dedicated_log_file, line)) {
    return;
  }
  WCHAR temp_path[MAX_PATH] = {0};
  DWORD temp_len = GetTempPathW(_countof(temp_path), temp_path);
  if (temp_len == 0 || temp_len >= _countof(temp_path)) {
    return;
  }
  const fs::path fallback_path =
      fs::path(temp_path) / L"weasel_input_content_fallback.log";
  AppendLineToFile(fallback_path, line);
}

std::wstring ResolveAIAssistantLoginStatePath(const AIAssistantConfig& config) {
  std::wstring path = u8tow(config.login_state_path);
  if (path.empty()) {
    path = L"ai_login_state.json";
  }
  std::filesystem::path fs_path(path);
  if (!fs_path.is_absolute()) {
    fs_path = WeaselUserDataPath() / fs_path;
  }
  return fs_path.wstring();
}

std::string TrimAsciiWhitespace(const std::string& input) {
  size_t first = 0;
  while (first < input.size() && std::isspace(static_cast<unsigned char>(input[first]))) {
    ++first;
  }
  size_t last = input.size();
  while (last > first &&
         std::isspace(static_cast<unsigned char>(input[last - 1]))) {
    --last;
  }
  return input.substr(first, last - first);
}

std::string ToLowerAscii(std::string input) {
  std::transform(input.begin(), input.end(), input.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
  return input;
}

bool ParseAIAssistantTriggerKeyToken(const std::string& key_token,
                                     UINT* keycode) {
  if (!keycode || key_token.empty()) {
    return false;
  }
  if (key_token.size() == 1) {
    const unsigned char ch = static_cast<unsigned char>(key_token[0]);
    if (std::isalnum(ch) || ch == '/' || ch == '\\' || ch == ',' || ch == '.' ||
        ch == ';' || ch == '\'' || ch == '-' || ch == '=' || ch == '`') {
      *keycode = (ch == '`') ? static_cast<UINT>(ibus::grave)
                             : static_cast<UINT>(ch);
      return true;
    }
  }
  if (key_token == "space" || key_token == "spacebar") {
    *keycode = ibus::space;
    return true;
  }
  if (key_token == "grave" || key_token == "backquote" ||
      key_token == "backtick") {
    *keycode = ibus::grave;
    return true;
  }
  if (key_token == "slash") {
    *keycode = '/';
    return true;
  }
  if (key_token == "backslash") {
    *keycode = '\\';
    return true;
  }
  if (key_token == "comma") {
    *keycode = ',';
    return true;
  }
  if (key_token == "period" || key_token == "dot") {
    *keycode = '.';
    return true;
  }
  if (key_token == "semicolon") {
    *keycode = ';';
    return true;
  }
  if (key_token == "quote" || key_token == "apostrophe") {
    *keycode = '\'';
    return true;
  }
  if (key_token == "minus" || key_token == "hyphen") {
    *keycode = '-';
    return true;
  }
  if (key_token == "equal" || key_token == "equals") {
    *keycode = '=';
    return true;
  }
  if (key_token == "tab") {
    *keycode = ibus::Tab;
    return true;
  }
  if (key_token == "enter" || key_token == "return") {
    *keycode = ibus::Return;
    return true;
  }
  if (key_token == "esc" || key_token == "escape") {
    *keycode = ibus::Escape;
    return true;
  }
  if (key_token.size() > 1 && key_token[0] == 'f') {
    bool all_digits = true;
    for (size_t i = 1; i < key_token.size(); ++i) {
      if (!std::isdigit(static_cast<unsigned char>(key_token[i]))) {
        all_digits = false;
        break;
      }
    }
    if (all_digits) {
      const int fn = std::stoi(key_token.substr(1));
      if (fn >= 1 && fn <= 35) {
        *keycode = static_cast<UINT>(ibus::F1 + fn - 1);
        return true;
      }
    }
  }
  return false;
}

bool TryParseAIAssistantTriggerHotkey(const std::string& raw_hotkey,
                                      UINT* keycode,
                                      UINT* modifiers) {
  if (!keycode || !modifiers) {
    return false;
  }
  const std::string hotkey = TrimAsciiWhitespace(raw_hotkey);
  if (hotkey.empty()) {
    return false;
  }
  bool has_key = false;
  UINT parsed_keycode = 0;
  UINT parsed_modifiers = 0;
  size_t begin = 0;
  while (begin <= hotkey.size()) {
    const size_t end = hotkey.find('+', begin);
    const std::string token_raw =
        end == std::string::npos ? hotkey.substr(begin)
                                 : hotkey.substr(begin, end - begin);
    const std::string token = ToLowerAscii(TrimAsciiWhitespace(token_raw));
    if (token.empty()) {
      return false;
    }
    UINT modifier = 0;
    if (token == "ctrl" || token == "control") {
      modifier = ibus::CONTROL_MASK;
    } else if (token == "shift") {
      modifier = ibus::SHIFT_MASK;
    } else if (token == "alt") {
      modifier = ibus::ALT_MASK;
    } else if (token == "super" || token == "win" || token == "windows") {
      modifier = ibus::SUPER_MASK;
    } else if (token == "meta") {
      modifier = ibus::META_MASK;
    } else if (token == "hyper") {
      modifier = ibus::HYPER_MASK;
    }
    if (modifier != 0) {
      parsed_modifiers |= modifier;
    } else {
      if (has_key || !ParseAIAssistantTriggerKeyToken(token, &parsed_keycode)) {
        return false;
      }
      has_key = true;
    }
    if (end == std::string::npos) {
      break;
    }
    begin = end + 1;
  }
  if (!has_key) {
    return false;
  }
  *keycode = parsed_keycode;
  *modifiers = parsed_modifiers;
  return true;
}

std::string GenerateLoginClientId() {
  static std::mutex mutex;
  static std::mt19937_64 rng(
      static_cast<uint64_t>(std::chrono::high_resolution_clock::now()
                                .time_since_epoch()
                                .count()) ^
      static_cast<uint64_t>(GetCurrentProcessId()) << 32);
  uint64_t a = 0;
  uint64_t b = 0;
  {
    std::lock_guard<std::mutex> lock(mutex);
    a = rng();
    b = rng();
  }
  char buffer[64] = {0};
  _snprintf_s(buffer, sizeof(buffer), _TRUNCATE,
              "%08llx-%04llx-%04llx-%04llx-%012llx",
              static_cast<unsigned long long>((a >> 32) & 0xffffffffULL),
              static_cast<unsigned long long>((a >> 16) & 0xffffULL),
              static_cast<unsigned long long>(a & 0xffffULL),
              static_cast<unsigned long long>((b >> 48) & 0xffffULL),
              static_cast<unsigned long long>(b & 0xffffffffffffULL));
  return buffer;
}

std::wstring ReplaceUuidPlaceholder(const std::wstring& text,
                                    const std::wstring& uuid) {
  const std::wstring placeholder = L"{uuid}";
  if (text.find(placeholder) == std::wstring::npos) {
    return text;
  }
  std::wstring result = text;
  size_t pos = 0;
  while ((pos = result.find(placeholder, pos)) != std::wstring::npos) {
    result.replace(pos, placeholder.size(), uuid);
    pos += uuid.size();
  }
  return result;
}

std::wstring BuildLoginUrlForBrowser(const AIAssistantConfig& config,
                                     const std::string& client_id) {
  std::wstring base_url = u8tow(config.login_url);
  if (base_url.empty()) {
    return std::wstring();
  }
  const std::wstring uuid = u8tow(client_id);
  std::wstring url = ReplaceUuidPlaceholder(base_url, uuid);
  if (url.find(L"{uuid}") == std::wstring::npos &&
      base_url.find(L"{uuid}") == std::wstring::npos) {
    const wchar_t separator = url.find(L'?') == std::wstring::npos ? L'?' : L'&';
    url.push_back(separator);
    url += L"uuid=";
    url += uuid;
    if (url.find(L"operationType=") == std::wstring::npos) {
      url += L"&operationType=login";
    }
    if (url.find(L"fromType=") == std::wstring::npos) {
      url += L"&fromType=plugIn";
    }
  }
  return url;
}

std::string BuildMqttTopicForClient(const AIAssistantConfig& config,
                                    const std::string& client_id) {
  std::string topic = config.mqtt_topic_template;
  if (topic.empty()) {
    return std::string();
  }
  const std::string placeholder = "{uuid}";
  const size_t pos = topic.find(placeholder);
  if (pos != std::string::npos) {
    topic.replace(pos, placeholder.size(), client_id);
    return topic;
  }
  if (!topic.empty() && topic.back() != '/') {
    topic.push_back('/');
  }
  topic += client_id;
  return topic;
}

bool LoadAIAssistantLoginIdentity(const AIAssistantConfig& config,
                                  std::string* token,
                                  std::string* tenant_id,
                                  std::string* refresh_token = nullptr);

bool LoadAIAssistantLoginToken(const AIAssistantConfig& config,
                               std::string* token) {
  if (!token) {
    return false;
  }
  std::string tenant_id;
  token->clear();
  tenant_id.clear();
  return LoadAIAssistantLoginIdentity(config, token, &tenant_id, nullptr);
}

bool ReadStringLikeJsonMember(const rapidjson::Value& value,
                              const char* key,
                              std::string* out) {
  if (!key || !out || !value.IsObject()) {
    return false;
  }
  const auto it = value.FindMember(key);
  if (it == value.MemberEnd()) {
    return false;
  }
  if (it->value.IsString()) {
    *out = it->value.GetString();
    return !out->empty();
  }
  if (it->value.IsInt64()) {
    *out = std::to_string(it->value.GetInt64());
    return true;
  }
  if (it->value.IsUint64()) {
    *out = std::to_string(it->value.GetUint64());
    return true;
  }
  if (it->value.IsInt()) {
    *out = std::to_string(it->value.GetInt());
    return true;
  }
  if (it->value.IsUint()) {
    *out = std::to_string(it->value.GetUint());
    return true;
  }
  return false;
}

std::string ReadFirstStringLikeJsonMember(
    const rapidjson::Value& value,
    const std::initializer_list<const char*>& keys) {
  std::string result;
  for (const char* key : keys) {
    if (ReadStringLikeJsonMember(value, key, &result) && !result.empty()) {
      return result;
    }
  }
  return std::string();
}

void ExtractLoginIdentityFromJson(const rapidjson::Value& value,
                                  const std::string& preferred_token_key,
                                  std::string* token,
                                  std::string* tenant_id,
                                  std::string* refresh_token) {
  if (!value.IsObject()) {
    return;
  }
  if (token && token->empty()) {
    if (!preferred_token_key.empty()) {
      ReadStringLikeJsonMember(value, preferred_token_key.c_str(), token);
    }
    if (token->empty()) {
      *token = ReadFirstStringLikeJsonMember(
          value, {"token", "accessToken", "saToken"});
    }
  }
  if (tenant_id && tenant_id->empty()) {
    *tenant_id = ReadFirstStringLikeJsonMember(
        value, {"tenantId", "tenantID", "tenant_id", "tenant", "insId",
                "insCode"});
  }
  if (refresh_token && refresh_token->empty()) {
    *refresh_token = ReadFirstStringLikeJsonMember(
        value, {"refreshToken", "refresh_token", "rtoken"});
  }

  const char* nested_keys[] = {"data", "result", "body", "content"};
  for (const char* nested_key : nested_keys) {
    const auto nested_it = value.FindMember(nested_key);
    if (nested_it == value.MemberEnd()) {
      continue;
    }
    const rapidjson::Value& nested_value = nested_it->value;
    if (nested_value.IsObject()) {
      ExtractLoginIdentityFromJson(nested_value, preferred_token_key, token,
                                   tenant_id, refresh_token);
    } else if (nested_value.IsArray()) {
      for (auto it = nested_value.Begin(); it != nested_value.End(); ++it) {
        if (it->IsObject()) {
          ExtractLoginIdentityFromJson(*it, preferred_token_key, token,
                                       tenant_id, refresh_token);
        }
      }
    }
  }
}

void ExtractLoginIdentityFromPayload(const std::string& payload,
                                     const std::string& preferred_token_key,
                                     std::string* token,
                                     std::string* tenant_id,
                                     std::string* refresh_token) {
  const std::string trimmed = TrimAsciiWhitespace(payload);
  if (trimmed.empty()) {
    return;
  }
  rapidjson::Document payload_doc;
  if (payload_doc.Parse(trimmed.c_str()).HasParseError() ||
      !payload_doc.IsObject()) {
    if (token && token->empty()) {
      *token = trimmed;
    }
    return;
  }
  ExtractLoginIdentityFromJson(payload_doc, preferred_token_key, token,
                               tenant_id, refresh_token);
}

bool LoadAIAssistantLoginIdentity(const AIAssistantConfig& config,
                                  std::string* token,
                                  std::string* tenant_id,
                                  std::string* refresh_token) {
  if (!token) {
    return false;
  }
  token->clear();
  if (tenant_id) {
    tenant_id->clear();
  }
  if (refresh_token) {
    refresh_token->clear();
  }
  const fs::path state_path = ResolveAIAssistantLoginStatePath(config);
  std::ifstream input(state_path, std::ios::binary);
  if (!input) {
    return false;
  }
  std::string json((std::istreambuf_iterator<char>(input)),
                   std::istreambuf_iterator<char>());
  if (json.empty()) {
    return false;
  }
  rapidjson::Document doc;
  if (doc.Parse(json.c_str()).HasParseError() || !doc.IsObject()) {
    return false;
  }
  ExtractLoginIdentityFromJson(doc, config.login_token_key, token, tenant_id,
                               refresh_token);
  const auto payload_it = doc.FindMember("payload");
  if (payload_it != doc.MemberEnd() && payload_it->value.IsString()) {
    ExtractLoginIdentityFromPayload(payload_it->value.GetString(),
                                    config.login_token_key, token, tenant_id,
                                    refresh_token);
  }
  if (refresh_token && refresh_token->empty()) {
    ReadStringLikeJsonMember(doc, "refreshToken", refresh_token);
  }
  return !token->empty() || (refresh_token && !refresh_token->empty());
}

bool SaveAIAssistantLoginState(const AIAssistantConfig& config,
                               const std::string& token,
                               const std::string& tenant_id,
                               const std::string& refresh_token,
                               const std::string& client_id,
                               const std::string& topic,
                               const std::string& payload) {
  if (token.empty() && refresh_token.empty()) {
    return false;
  }
  const fs::path state_path = ResolveAIAssistantLoginStatePath(config);
  std::error_code ec;
  fs::create_directories(state_path.parent_path(), ec);

  SYSTEMTIME st = {0};
  GetLocalTime(&st);
  char time_buffer[64] = {0};
  _snprintf_s(time_buffer, _countof(time_buffer), _TRUNCATE,
              "%04d-%02d-%02d %02d:%02d:%02d", st.wYear, st.wMonth, st.wDay,
              st.wHour, st.wMinute, st.wSecond);

  std::ofstream output(state_path, std::ios::binary | std::ios::trunc);
  if (!output) {
    return false;
  }

  const std::string key =
      config.login_token_key.empty() ? "token" : config.login_token_key;
  output << "{";
  output << "\"token\":\"" << EscapeJsonString(token) << "\",";
  if (!key.empty() && key != "token") {
    output << "\"" << EscapeJsonString(key) << "\":\""
           << EscapeJsonString(token) << "\",";
  }
  output << "\"tenantId\":\"" << EscapeJsonString(tenant_id) << "\",";
  output << "\"refreshToken\":\"" << EscapeJsonString(refresh_token) << "\",";
  output << "\"client_id\":\"" << EscapeJsonString(client_id) << "\",";
  output << "\"topic\":\"" << EscapeJsonString(topic) << "\",";
  output << "\"updated_at\":\"" << EscapeJsonString(time_buffer) << "\",";
  output << "\"payload\":\"" << EscapeJsonString(payload) << "\"";
  output << "}";
  return output.good();
}

std::string BuildAIAssistantInstitutionListRequestBody(
    const AIAssistantConfig& config) {
  (void)config;
  std::string body;
  body.reserve(128);
  body += "{\"queryContent\":\"\",\"insShowType\":\"";
  body += "2";
  body += "\"}";
  return body;
}

const rapidjson::Value* FindInstitutionArrayInResponse(
    const rapidjson::Value& value) {
  if (value.IsArray()) {
    return &value;
  }
  if (!value.IsObject()) {
    return nullptr;
  }

  const char* direct_array_keys[] = {"data", "rows", "list", "items"};
  for (const char* key : direct_array_keys) {
    const auto it = value.FindMember(key);
    if (it != value.MemberEnd() && it->value.IsArray()) {
      return &it->value;
    }
  }

  const char* nested_object_keys[] = {"data", "result"};
  for (const char* key : nested_object_keys) {
    const auto it = value.FindMember(key);
    if (it == value.MemberEnd() || !it->value.IsObject()) {
      continue;
    }
    const rapidjson::Value& nested = it->value;
    const char* nested_array_keys[] = {"list", "rows", "items", "data"};
    for (const char* nested_key : nested_array_keys) {
      const auto nested_it = nested.FindMember(nested_key);
      if (nested_it != nested.MemberEnd() && nested_it->value.IsArray()) {
        return &nested_it->value;
      }
    }
  }
  return nullptr;
}

bool ParseAIAssistantInstitutionList(
    const std::string& response_body,
    std::vector<AIPanelInstitutionOption>* options,
    std::string* error_message) {
  if (!options) {
    if (error_message) {
      *error_message = "Institution options output is null.";
    }
    return false;
  }
  options->clear();
  rapidjson::Document document;
  document.Parse(response_body.c_str(), response_body.size());
  if (document.HasParseError()) {
    if (error_message) {
      *error_message = "Institution list response is not valid JSON.";
    }
    return false;
  }

  const rapidjson::Value* list = FindInstitutionArrayInResponse(document);
  if (!list || !list->IsArray()) {
    if (error_message) {
      *error_message = "Institution list response did not contain array data.";
    }
    return false;
  }

  for (auto it = list->Begin(); it != list->End(); ++it) {
    if (!it->IsObject()) {
      continue;
    }
    std::string id = ReadFirstStringLikeJsonMember(
        *it, {"insId", "id", "insCode", "code", "tenantId", "orgId"});
    std::string name = ReadFirstStringLikeJsonMember(
        *it, {"insName", "name", "tenantName", "title", "label"});
    std::string app_key =
        ReadFirstStringLikeJsonMember(*it, {"appKey", "app_key", "appkey"});
    if (name.empty()) {
      continue;
    }
    if (app_key.empty()) {
      continue;
    }
    if (id.empty()) {
      id = name;
    }
    options->push_back(
        AIPanelInstitutionOption(u8tow(id), u8tow(name), u8tow(app_key)));
  }

  if (options->empty()) {
    if (error_message) {
      *error_message = "Institution list is empty.";
    }
    return false;
  }
  return true;
}

bool FetchAIAssistantInstitutionOptions(
    const AIAssistantConfig& config,
    const std::string& token,
    const std::string& tenant_id,
    std::vector<AIPanelInstitutionOption>* options,
    std::string* error_message,
    int* http_status_code = nullptr) {
  (void)config;
  (void)token;
  (void)tenant_id;
  if (!options) {
    if (error_message) {
      *error_message = "Institution options output is null.";
    }
    return false;
  }
  options->clear();
  if (http_status_code) {
    *http_status_code = 0;
  }
  if (error_message) {
    *error_message =
        "Institution list fetch is disabled; panel loads agent list itself.";
  }
  return false;
}

std::wstring UrlEncodeQueryComponent(const std::string& input) {
  static const wchar_t kHex[] = L"0123456789ABCDEF";
  std::wstring output;
  output.reserve(input.size() * 3);
  for (unsigned char ch : input) {
    const bool is_alnum =
        (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') ||
        (ch >= '0' && ch <= '9');
    if (is_alnum || ch == '-' || ch == '_' || ch == '.' || ch == '~') {
      output.push_back(static_cast<wchar_t>(ch));
      continue;
    }
    output.push_back(L'%');
    output.push_back(kHex[(ch >> 4) & 0x0F]);
    output.push_back(kHex[ch & 0x0F]);
  }
  return output;
}

std::wstring BuildAIPanelAuthQueryString(const std::string& token,
                                         const std::string& tenant_id,
                                         const std::string& refresh_token) {
  std::wstring query;
  auto append_pair = [&query](const wchar_t* key, const std::string& value) {
    if (!key || !*key) {
      return;
    }
    if (!query.empty()) {
      query += L"&";
    }
    query += key;
    query += L"=";
    query += UrlEncodeQueryComponent(value);
  };

  append_pair(L"tenantid", tenant_id);
  append_pair(L"token", token);
  append_pair(L"tenantId", tenant_id);
  append_pair(L"refreshToken", refresh_token);
  return query;
}

std::wstring AppendAIPanelAuthToUrl(const std::wstring& panel_url,
                                    const std::string& token,
                                    const std::string& tenant_id,
                                    const std::string& refresh_token) {
  if (panel_url.empty()) {
    return panel_url;
  }
  const std::wstring query =
      BuildAIPanelAuthQueryString(token, tenant_id, refresh_token);

  const size_t hash_pos = panel_url.find(L'#');
  if (hash_pos == std::wstring::npos) {
    std::wstring url = panel_url;
    if (url.find(L'?') == std::wstring::npos) {
      url += L"?";
    } else if (!url.empty() && url.back() != L'?' && url.back() != L'&') {
      url += L"&";
    }
    url += query;
    return url;
  }

  std::wstring prefix = panel_url.substr(0, hash_pos);
  std::wstring fragment = panel_url.substr(hash_pos + 1);
  const size_t query_pos = fragment.find(L'?');
  if (query_pos != std::wstring::npos) {
    fragment = fragment.substr(0, query_pos);
  }
  if (fragment.empty()) {
    fragment = L"/rime-with-weasel";
  }

  std::wstring result = prefix;
  result += L"#";
  result += fragment;
  result += L"?";
  result += query;
  return result;
}

bool ParseHttpUrlOrigin(const std::wstring& url,
                        std::wstring* scheme,
                        std::wstring* host,
                        INTERNET_PORT* port) {
  if (!scheme || !host || !port || url.empty()) {
    return false;
  }
  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(url.c_str(), 0, 0, &parts)) {
    return false;
  }
  if (!parts.lpszScheme || parts.dwSchemeLength == 0 || !parts.lpszHostName ||
      parts.dwHostNameLength == 0) {
    return false;
  }
  scheme->assign(parts.lpszScheme, parts.dwSchemeLength);
  host->assign(parts.lpszHostName, parts.dwHostNameLength);
  *port = parts.nPort;
  if (*port == 0) {
    if (_wcsicmp(scheme->c_str(), L"https") == 0) {
      *port = INTERNET_DEFAULT_HTTPS_PORT;
    } else {
      *port = INTERNET_DEFAULT_HTTP_PORT;
    }
  }
  return true;
}

std::wstring BuildHttpOrigin(const std::wstring& scheme,
                             const std::wstring& host,
                             INTERNET_PORT port) {
  if (scheme.empty() || host.empty()) {
    return std::wstring();
  }
  std::wstring origin = scheme;
  origin += L"://";
  origin += host;
  const bool is_https = _wcsicmp(scheme.c_str(), L"https") == 0;
  const bool is_http = _wcsicmp(scheme.c_str(), L"http") == 0;
  const bool use_default_port =
      (is_https && port == INTERNET_DEFAULT_HTTPS_PORT) ||
      (is_http && port == INTERNET_DEFAULT_HTTP_PORT);
  if (!use_default_port) {
    origin += L":";
    origin += std::to_wstring(static_cast<unsigned>(port));
  }
  return origin;
}

std::wstring ResolveAIPanelAllowedOrigin(const AIAssistantConfig& config) {
  if (!config.panel_allowed_origin.empty()) {
    return u8tow(config.panel_allowed_origin);
  }
  if (config.panel_url.empty()) {
    return std::wstring();
  }
  std::wstring scheme;
  std::wstring host;
  INTERNET_PORT port = 0;
  if (!ParseHttpUrlOrigin(u8tow(config.panel_url), &scheme, &host, &port)) {
    return std::wstring();
  }
  return BuildHttpOrigin(scheme, host, port);
}

bool IsAIPanelMessageOriginAllowed(const std::wstring& expected_origin,
                                   const std::wstring& source_url) {
  if (expected_origin.empty()) {
    return true;
  }
  std::wstring scheme;
  std::wstring host;
  INTERNET_PORT port = 0;
  if (!ParseHttpUrlOrigin(source_url, &scheme, &host, &port)) {
    return false;
  }
  const std::wstring origin = BuildHttpOrigin(scheme, host, port);
  if (origin.empty()) {
    return false;
  }
  return _wcsicmp(origin.c_str(), expected_origin.c_str()) == 0;
}

#if WEASEL_HAS_WEBVIEW2
bool ExtractWebViewMessageText(
    ICoreWebView2WebMessageReceivedEventArgs* args,
    std::wstring* message) {
  if (!args || !message) {
    return false;
  }
  message->clear();
  LPWSTR message_raw = nullptr;
  if (SUCCEEDED(args->TryGetWebMessageAsString(&message_raw)) && message_raw) {
    *message = message_raw;
    CoTaskMemFree(message_raw);
    return true;
  }
  LPWSTR message_json_raw = nullptr;
  if (SUCCEEDED(args->get_WebMessageAsJson(&message_json_raw)) &&
      message_json_raw) {
    std::wstring json_text(message_json_raw);
    CoTaskMemFree(message_json_raw);
    const std::string json_utf8 = wtou8(json_text);
    rapidjson::Document json_doc;
    json_doc.Parse(json_utf8.c_str(), json_utf8.size());
    if (!json_doc.HasParseError() && json_doc.IsString()) {
      *message = u8tow(std::string(json_doc.GetString(),
                                   json_doc.GetStringLength()));
    } else {
      *message = json_text;
    }
    return true;
  }
  return false;
}
#endif

bool ParseAIPanelUiCommand(const std::wstring& message,
                           AIPanelUiCommand* command) {
  if (!command) {
    return false;
  }
  command->type.clear();
  command->text.clear();
  command->institution_id.clear();
  command->panel_width = 0;
  command->panel_height = 0;
  command->resize_reason.clear();
  if (message.empty()) {
    return false;
  }

  const std::string message_utf8 = wtou8(message);
  rapidjson::Document document;
  document.Parse(message_utf8.c_str(), message_utf8.size());
  if (!document.HasParseError() && document.IsObject()) {
    std::string type;
    ReadStringLikeJsonMember(document, "type", &type);
    if (type.empty()) {
      return false;
    }
    command->type = type;
    const rapidjson::Value* payload = nullptr;
    const auto payload_it = document.FindMember("payload");
    if (payload_it != document.MemberEnd() && payload_it->value.IsObject()) {
      payload = &payload_it->value;
    }
    const rapidjson::Value& payload_ref = payload ? *payload : document;
    const auto read_int_like_member = [](const rapidjson::Value& object,
                                         const char* key,
                                         int* out) -> bool {
      if (!out || !object.IsObject()) {
        return false;
      }
      const auto it = object.FindMember(key);
      if (it == object.MemberEnd()) {
        return false;
      }
      const rapidjson::Value& value = it->value;
      if (value.IsInt()) {
        *out = value.GetInt();
        return true;
      }
      if (value.IsUint()) {
        *out = static_cast<int>(value.GetUint());
        return true;
      }
      if (value.IsInt64()) {
        *out = static_cast<int>(value.GetInt64());
        return true;
      }
      if (value.IsUint64()) {
        *out = static_cast<int>(value.GetUint64());
        return true;
      }
      if (value.IsDouble()) {
        *out = static_cast<int>(value.GetDouble());
        return true;
      }
      if (value.IsString()) {
        const char* text = value.GetString();
        if (!text || !*text) {
          return false;
        }
        *out = std::atoi(text);
        return true;
      }
      return false;
    };

    if (type == "ui.context.changed") {
      std::string text;
      if (ReadStringLikeJsonMember(payload_ref, "text", &text)) {
        command->text = u8tow(text);
      }
    } else if (type == "ui.select.institution") {
      std::string id;
      if (!ReadStringLikeJsonMember(payload_ref, "id", &id)) {
        ReadStringLikeJsonMember(payload_ref, "institutionId", &id);
      }
      command->institution_id = u8tow(id);
    } else if (type == "ui.writeback.confirm") {
      std::string text;
      if (ReadStringLikeJsonMember(payload_ref, "text", &text)) {
        command->text = u8tow(text);
      }
    } else if (type == "ui.panel.resize") {
      int width = 0;
      int height = 0;
      read_int_like_member(payload_ref, "width", &width);
      read_int_like_member(payload_ref, "height", &height);
      command->panel_width = width;
      command->panel_height = height;
      ReadStringLikeJsonMember(payload_ref, "reason", &command->resize_reason);
    } else if (type == "ui.system_command") {
      std::string text;
      if (ReadStringLikeJsonMember(payload_ref, "text", &text)) {
        command->text = u8tow(text);
      }
    }
    return true;
  }

  if (message == L"ai_request") {
    command->type = "ui.request";
    return true;
  }
  if (message.rfind(L"ai_request_text:", 0) == 0) {
    command->type = "ui.request";
    command->text = message.substr(16);
    return true;
  }
  if (message == L"ai_confirm") {
    command->type = "ui.confirm";
    return true;
  }
  if (message == L"ai_cancel") {
    command->type = "ui.cancel";
    return true;
  }
  if (message == L"ai_drag_start") {
    command->type = "ui.drag.start";
    return true;
  }
  if (message.rfind(L"ai_context_changed:", 0) == 0) {
    command->type = "ui.context.changed";
    command->text = message.substr(19);
    return true;
  }
  if (message.rfind(L"ai_select_ins:", 0) == 0) {
    command->type = "ui.select.institution";
    command->institution_id = message.substr(14);
    return true;
  }
  return false;
}

std::wstring NormalizeReferencedContextText(const std::wstring& context_text) {
  if (context_text.empty()) {
    return context_text;
  }

  const std::string utf8 = wtou8(context_text);
  rapidjson::Document doc;
  if (doc.Parse(utf8.c_str()).HasParseError() || !doc.IsArray()) {
    return context_text;
  }

  if (doc.Empty()) {
    return std::wstring();
  }

  const rapidjson::Value& last = doc[doc.Size() - 1];
  if (last.IsString()) {
    return u8tow(last.GetString());
  }
  if (last.IsObject()) {
    const char* keys[] = {"text", "content", "value"};
    for (const char* key : keys) {
      const auto member = last.FindMember(key);
      if (member != last.MemberEnd() && member->value.IsString()) {
        return u8tow(member->value.GetString());
      }
    }
  }
  return std::wstring();
}

std::string BuildAIPanelHostSyncMessage(
    const std::wstring& context_text,
    const std::wstring& status_text,
    const std::wstring& output_text,
    bool requesting,
    bool institutions_loading,
    const std::vector<AIPanelInstitutionOption>& institution_options,
    const std::wstring& selected_institution_id,
    const std::string& token,
    const std::string& tenant_id,
    const std::string& refresh_token) {
  const std::wstring normalized_context =
      NormalizeReferencedContextText(context_text);
  std::string json;
  json.reserve(1024 + normalized_context.size() * 3 + output_text.size() * 2 +
               institution_options.size() * 96);
  json += "{\"v\":\"1.0\",\"type\":\"host.sync\",\"payload\":{";
  json += "\"context\":\"";
  json += EscapeJsonString(wtou8(normalized_context));
  json += "\",\"status\":\"";
  json += EscapeJsonString(wtou8(status_text));
  json += "\",\"output\":\"";
  json += EscapeJsonString(wtou8(output_text));
  json += "\",\"requesting\":";
  json += requesting ? "true" : "false";
  json += ",\"institutionsLoading\":";
  json += institutions_loading ? "true" : "false";
  json += ",\"panelLimits\":{";
  json += "\"minWidth\":";
  json += std::to_string(kAIPanelMinWidth);
  json += ",\"maxWidth\":";
  json += std::to_string(kAIPanelMaxWidth);
  json += ",\"minHeight\":";
  json += std::to_string(kAIPanelMinHeight);
  json += ",\"maxHeight\":";
  json += std::to_string(kAIPanelMaxHeight);
  json += ",\"margin\":";
  json += std::to_string(kAIPanelScreenMargin);
  json += "}";
  json += ",\"selectedInstitutionId\":\"";
  json += EscapeJsonString(wtou8(selected_institution_id));
  json += "\",\"auth\":{\"token\":\"";
  json += EscapeJsonString(token);
  json += "\",\"tenantId\":\"";
  json += EscapeJsonString(tenant_id);
  json += "\",\"refreshToken\":\"";
  json += EscapeJsonString(refresh_token);
  json += "\"},\"institutions\":[";
  bool first = true;
  for (const auto& option : institution_options) {
    if (!first) {
      json += ",";
    }
    first = false;
    json += "{\"id\":\"";
    json += EscapeJsonString(wtou8(option.id));
    json += "\",\"name\":\"";
    json += EscapeJsonString(wtou8(option.name));
    json += "\",\"appKey\":\"";
    json += EscapeJsonString(wtou8(option.app_key));
    json += "\"}";
  }
  json += "]}}";
  return json;
}

void AddUniqueEndpoint(std::vector<std::wstring>* endpoints,
                       const std::wstring& endpoint) {
  if (!endpoints || endpoint.empty()) {
    return;
  }
  for (const auto& existing : *endpoints) {
    if (_wcsicmp(existing.c_str(), endpoint.c_str()) == 0) {
      return;
    }
  }
  endpoints->push_back(endpoint);
}

void AddRefreshEndpointCandidatesFromUrl(const std::wstring& url,
                                         std::vector<std::wstring>* endpoints) {
  std::wstring scheme;
  std::wstring host;
  INTERNET_PORT port = 0;
  if (!ParseHttpUrlOrigin(url, &scheme, &host, &port)) {
    return;
  }
  const std::wstring origin = BuildHttpOrigin(scheme, host, port);
  if (origin.empty()) {
    return;
  }
  AddUniqueEndpoint(endpoints, origin + L"/api/oauth/anyTenant/refresh");
  AddUniqueEndpoint(endpoints, origin + L"/lamp-api/oauth/anyTenant/refresh");
  if (port == 85) {
    const std::wstring alt_origin =
        BuildHttpOrigin(scheme, host, static_cast<INTERNET_PORT>(84));
    AddUniqueEndpoint(endpoints,
                      alt_origin + L"/api/oauth/anyTenant/refresh");
  }
}

std::vector<std::wstring> BuildRefreshTokenEndpointCandidates(
    const AIAssistantConfig& config) {
  std::vector<std::wstring> endpoints;
  if (!config.refresh_token_endpoint.empty()) {
    AddUniqueEndpoint(&endpoints, u8tow(config.refresh_token_endpoint));
  }
  if (!config.login_url.empty()) {
    AddRefreshEndpointCandidatesFromUrl(u8tow(config.login_url), &endpoints);
  }
  return endpoints;
}

bool ParseRefreshTokenResponse(const std::string& response_body,
                               std::string* access_token,
                               std::string* refresh_token,
                               std::string* error_message) {
  if (!access_token || !refresh_token) {
    if (error_message) {
      *error_message = "Refresh token response target is null.";
    }
    return false;
  }
  access_token->clear();
  refresh_token->clear();

  rapidjson::Document doc;
  doc.Parse(response_body.c_str(), response_body.size());
  if (doc.HasParseError() || !doc.IsObject()) {
    if (error_message) {
      *error_message = "Refresh token response is not valid JSON.";
    }
    return false;
  }

  auto read_code = [&doc]() -> int {
    const auto code_it = doc.FindMember("code");
    if (code_it == doc.MemberEnd()) {
      return 200;
    }
    if (code_it->value.IsInt()) {
      return code_it->value.GetInt();
    }
    if (code_it->value.IsString()) {
      return std::atoi(code_it->value.GetString());
    }
    return 200;
  };

  const int code = read_code();
  if (code != 0 && code != 200) {
    if (error_message) {
      std::string message = "Refresh token response code " +
                            std::to_string(code) + ".";
      const auto msg_it = doc.FindMember("msg");
      if (msg_it != doc.MemberEnd() && msg_it->value.IsString()) {
        message = msg_it->value.GetString();
      }
      *error_message = message;
    }
    return false;
  }

  const rapidjson::Value* payload = &doc;
  const auto data_it = doc.FindMember("data");
  if (data_it != doc.MemberEnd() && data_it->value.IsObject()) {
    payload = &data_it->value;
  }

  *access_token = ReadFirstStringLikeJsonMember(
      *payload, {"token", "accessToken", "saToken"});
  *refresh_token = ReadFirstStringLikeJsonMember(
      *payload, {"refreshToken", "refresh_token", "rtoken"});
  if (access_token->empty()) {
    if (error_message) {
      *error_message = "Refresh token response did not include token.";
    }
    return false;
  }
  return true;
}

bool RequestAIAssistantAccessTokenWithRefresh(
    const AIAssistantConfig& config,
    const std::wstring& refresh_endpoint,
    const std::string& refresh_token,
    const std::string& tenant_id,
    std::string* access_token,
    std::string* refreshed_refresh_token,
    std::string* error_message) {
  if (!access_token || !refreshed_refresh_token) {
    if (error_message) {
      *error_message = "Refresh token output is null.";
    }
    return false;
  }
  access_token->clear();
  refreshed_refresh_token->clear();
  if (refresh_endpoint.empty() || refresh_token.empty()) {
    if (error_message) {
      *error_message = "Refresh endpoint or refresh token is empty.";
    }
    return false;
  }

  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(refresh_endpoint.c_str(), 0, 0, &parts)) {
    if (error_message) {
      *error_message = "Unable to parse refresh endpoint URL.";
    }
    return false;
  }

  const std::wstring host(parts.lpszHostName, parts.dwHostNameLength);
  std::wstring object_name =
      parts.dwUrlPathLength > 0
          ? std::wstring(parts.lpszUrlPath, parts.dwUrlPathLength)
          : std::wstring(L"/");
  if (parts.dwExtraInfoLength > 0) {
    object_name.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);
  }
  object_name += (object_name.find(L'?') == std::wstring::npos) ? L"?" : L"&";
  object_name += L"refreshToken=";
  object_name += UrlEncodeQueryComponent(refresh_token);

  ScopedWinHttpHandle session(
      WinHttpOpen(L"WeaselAIAssistantRefresh/1.0",
                  WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
                  WINHTTP_NO_PROXY_BYPASS, 0));
  if (!session.valid()) {
    if (error_message) {
      *error_message = "WinHTTP session creation failed for refresh.";
    }
    return false;
  }
  WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                     config.timeout_ms, config.timeout_ms);

  ScopedWinHttpHandle connection(
      WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
  if (!connection.valid()) {
    if (error_message) {
      *error_message = "WinHTTP connection failed for refresh endpoint.";
    }
    return false;
  }

  const DWORD request_flags =
      parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
  ScopedWinHttpHandle request(
      WinHttpOpenRequest(connection.get(), L"POST", object_name.c_str(), nullptr,
                         WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                         request_flags));
  if (!request.valid()) {
    if (error_message) {
      *error_message = "WinHTTP request creation failed for refresh.";
    }
    return false;
  }

  std::wstring headers = L"Accept: application/json\r\n";
  if (!tenant_id.empty()) {
    headers += L"TenantId: ";
    headers += u8tow(tenant_id);
    headers += L"\r\n";
  }
  WinHttpAddRequestHeaders(request.get(), headers.c_str(),
                           static_cast<DWORD>(-1),
                           WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

  if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                          WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
      !WinHttpReceiveResponse(request.get(), nullptr)) {
    if (error_message) {
      *error_message = "Sending refresh token request failed.";
    }
    return false;
  }

  DWORD status_code = 0;
  DWORD status_code_size = sizeof(status_code);
  WinHttpQueryHeaders(request.get(),
                      WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                      WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                      &status_code_size, WINHTTP_NO_HEADER_INDEX);

  std::string response_body;
  for (;;) {
    DWORD available = 0;
    if (!WinHttpQueryDataAvailable(request.get(), &available)) {
      if (error_message) {
        *error_message = "Failed to read refresh token response.";
      }
      return false;
    }
    if (available == 0) {
      break;
    }
    std::string chunk(available, '\0');
    DWORD downloaded = 0;
    if (!WinHttpReadData(request.get(), chunk.data(), available, &downloaded)) {
      if (error_message) {
        *error_message = "Failed while downloading refresh token response.";
      }
      return false;
    }
    chunk.resize(downloaded);
    response_body += chunk;
  }

  if (status_code < 200 || status_code >= 300) {
    if (error_message) {
      *error_message = "Refresh token request returned HTTP " +
                       std::to_string(status_code) + ".";
    }
    return false;
  }

  const bool parsed = ParseRefreshTokenResponse(response_body, access_token,
                                                refreshed_refresh_token,
                                                error_message);
  if (parsed && refreshed_refresh_token->empty()) {
    *refreshed_refresh_token = refresh_token;
  }
  return parsed;
}

bool RefreshAIAssistantAccessToken(
    const AIAssistantConfig& config,
    const std::string& refresh_token,
    const std::string& tenant_id,
    std::string* access_token,
    std::string* refreshed_refresh_token,
    std::string* selected_endpoint,
    std::string* error_message) {
  if (!access_token || !refreshed_refresh_token) {
    if (error_message) {
      *error_message = "Refresh token output is null.";
    }
    return false;
  }
  access_token->clear();
  refreshed_refresh_token->clear();
  if (selected_endpoint) {
    selected_endpoint->clear();
  }
  if (refresh_token.empty()) {
    if (error_message) {
      *error_message = "Refresh token is empty.";
    }
    return false;
  }

  const std::vector<std::wstring> endpoints =
      BuildRefreshTokenEndpointCandidates(config);
  if (endpoints.empty()) {
    if (error_message) {
      *error_message = "No refresh token endpoint configured.";
    }
    return false;
  }

  std::string last_error;
  for (const auto& endpoint : endpoints) {
    std::string endpoint_error;
    if (RequestAIAssistantAccessTokenWithRefresh(
            config, endpoint, refresh_token, tenant_id, access_token,
            refreshed_refresh_token, &endpoint_error)) {
      if (selected_endpoint) {
        *selected_endpoint = wtou8(endpoint);
      }
      return true;
    }
    last_error = endpoint_error;
  }

  if (error_message) {
    *error_message = last_error.empty() ? "Refresh token request failed."
                                        : last_error;
  }
  return false;
}

struct ParsedWebSocketUrl {
  bool valid = false;
  bool secure = false;
  std::wstring host;
  INTERNET_PORT port = INTERNET_DEFAULT_HTTP_PORT;
  std::wstring path = L"/";
};

ParsedWebSocketUrl ParseWebSocketUrl(const std::wstring& url) {
  ParsedWebSocketUrl result;
  if (url.empty()) {
    return result;
  }
  std::wstring scheme;
  const size_t scheme_pos = url.find(L"://");
  if (scheme_pos != std::wstring::npos) {
    scheme = url.substr(0, scheme_pos);
    std::transform(scheme.begin(), scheme.end(), scheme.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });
  }
  std::wstring crack_url = url;
  if (scheme == L"ws") {
    crack_url = L"http://" + url.substr(scheme_pos + 3);
  } else if (scheme == L"wss") {
    crack_url = L"https://" + url.substr(scheme_pos + 3);
  }

  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(crack_url.c_str(), 0, 0, &parts)) {
    return result;
  }
  if (scheme.empty() && parts.lpszScheme && parts.dwSchemeLength > 0) {
    scheme.assign(parts.lpszScheme, parts.dwSchemeLength);
    std::transform(scheme.begin(), scheme.end(), scheme.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });
  }
  if (scheme != L"ws" && scheme != L"wss" && scheme != L"http" &&
      scheme != L"https") {
    return result;
  }
  result.secure = (scheme == L"wss" || scheme == L"https");
  if (parts.lpszHostName && parts.dwHostNameLength > 0) {
    result.host.assign(parts.lpszHostName, parts.dwHostNameLength);
  }
  result.port = parts.nPort != 0
                    ? parts.nPort
                    : (result.secure ? INTERNET_DEFAULT_HTTPS_PORT
                                     : INTERNET_DEFAULT_HTTP_PORT);
  std::wstring path;
  if (parts.lpszUrlPath && parts.dwUrlPathLength > 0) {
    path.assign(parts.lpszUrlPath, parts.dwUrlPathLength);
  }
  if (parts.lpszExtraInfo && parts.dwExtraInfoLength > 0) {
    path.append(parts.lpszExtraInfo, parts.dwExtraInfoLength);
  }
  if (path.empty()) {
    path = L"/";
  }
  result.path = path;
  result.valid = !result.host.empty();
  return result;
}

void AppendMqttUtf8String(std::vector<uint8_t>* output,
                          const std::string& text) {
  if (!output) {
    return;
  }
  const size_t take = std::min<size_t>(0xffff, text.size());
  output->push_back(static_cast<uint8_t>((take >> 8) & 0xff));
  output->push_back(static_cast<uint8_t>(take & 0xff));
  output->insert(output->end(), text.begin(), text.begin() + take);
}

void AppendMqttRemainingLength(std::vector<uint8_t>* output, size_t value) {
  if (!output) {
    return;
  }
  do {
    uint8_t encoded = static_cast<uint8_t>(value % 128);
    value /= 128;
    if (value > 0) {
      encoded = static_cast<uint8_t>(encoded | 0x80);
    }
    output->push_back(encoded);
  } while (value > 0);
}

bool ParseMqttRemainingLength(const std::vector<uint8_t>& packet,
                              size_t offset,
                              size_t* value,
                              size_t* used_bytes) {
  if (!value || !used_bytes || offset >= packet.size()) {
    return false;
  }
  *value = 0;
  *used_bytes = 0;
  size_t multiplier = 1;
  for (size_t i = 0; i < 4 && offset + i < packet.size(); ++i) {
    const uint8_t encoded = packet[offset + i];
    *value += static_cast<size_t>(encoded & 0x7f) * multiplier;
    multiplier *= 128;
    *used_bytes = i + 1;
    if ((encoded & 0x80) == 0) {
      return true;
    }
  }
  return false;
}

std::vector<uint8_t> BuildMqttConnectPacket(const std::string& client_id,
                                            const std::string& username,
                                            const std::string& password,
                                            uint16_t keepalive_sec) {
  std::vector<uint8_t> variable_header;
  AppendMqttUtf8String(&variable_header, "MQTT");
  variable_header.push_back(0x04);
  uint8_t connect_flags = 0x02;
  if (!username.empty()) {
    connect_flags = static_cast<uint8_t>(connect_flags | 0x80);
  }
  if (!password.empty()) {
    connect_flags = static_cast<uint8_t>(connect_flags | 0x40);
  }
  variable_header.push_back(connect_flags);
  variable_header.push_back(static_cast<uint8_t>((keepalive_sec >> 8) & 0xff));
  variable_header.push_back(static_cast<uint8_t>(keepalive_sec & 0xff));

  std::vector<uint8_t> payload;
  AppendMqttUtf8String(&payload, client_id);
  if (!username.empty()) {
    AppendMqttUtf8String(&payload, username);
  }
  if (!password.empty()) {
    AppendMqttUtf8String(&payload, password);
  }

  std::vector<uint8_t> packet;
  packet.push_back(0x10);
  AppendMqttRemainingLength(&packet, variable_header.size() + payload.size());
  packet.insert(packet.end(), variable_header.begin(), variable_header.end());
  packet.insert(packet.end(), payload.begin(), payload.end());
  return packet;
}

std::vector<uint8_t> BuildMqttSubscribePacket(const std::string& topic,
                                              uint16_t packet_id) {
  std::vector<uint8_t> body;
  body.push_back(static_cast<uint8_t>((packet_id >> 8) & 0xff));
  body.push_back(static_cast<uint8_t>(packet_id & 0xff));
  AppendMqttUtf8String(&body, topic);
  body.push_back(0x00);

  std::vector<uint8_t> packet;
  packet.push_back(0x82);
  AppendMqttRemainingLength(&packet, body.size());
  packet.insert(packet.end(), body.begin(), body.end());
  return packet;
}

enum class WebSocketReceiveResult {
  kOk,
  kTimeout,
  kClosed,
  kError,
};

bool SendWebSocketBinaryMessage(HINTERNET websocket,
                                const std::vector<uint8_t>& message) {
  if (!websocket || message.empty()) {
    return false;
  }
  const DWORD result = WinHttpWebSocketSend(
      websocket, WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE,
      const_cast<uint8_t*>(message.data()), static_cast<DWORD>(message.size()));
  return result == NO_ERROR;
}

WebSocketReceiveResult ReceiveWebSocketBinaryMessage(
    HINTERNET websocket,
    std::vector<uint8_t>* message) {
  if (!websocket || !message) {
    return WebSocketReceiveResult::kError;
  }
  message->clear();
  std::array<uint8_t, 4096> buffer = {0};
  while (true) {
    DWORD bytes_read = 0;
    WINHTTP_WEB_SOCKET_BUFFER_TYPE buffer_type =
        WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE;
    const DWORD result = WinHttpWebSocketReceive(
        websocket, buffer.data(), static_cast<DWORD>(buffer.size()), &bytes_read,
        &buffer_type);
    if (result == ERROR_WINHTTP_TIMEOUT) {
      return WebSocketReceiveResult::kTimeout;
    }
    if (result != NO_ERROR) {
      return WebSocketReceiveResult::kError;
    }
    if (buffer_type == WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE) {
      return WebSocketReceiveResult::kClosed;
    }
    if (buffer_type != WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE &&
        buffer_type != WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE) {
      if (buffer_type == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE ||
          buffer_type == WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE) {
        continue;
      }
      return WebSocketReceiveResult::kError;
    }
    if (bytes_read > 0) {
      message->insert(message->end(), buffer.begin(),
                      buffer.begin() + static_cast<size_t>(bytes_read));
    }
    if (buffer_type == WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE) {
      return WebSocketReceiveResult::kOk;
    }
  }
}

int MqttPacketType(const std::vector<uint8_t>& packet) {
  if (packet.empty()) {
    return -1;
  }
  return static_cast<int>((packet[0] >> 4) & 0x0f);
}

bool IsMqttConnAckOk(const std::vector<uint8_t>& packet) {
  if (MqttPacketType(packet) != 2) {
    return false;
  }
  size_t remaining = 0;
  size_t used = 0;
  if (!ParseMqttRemainingLength(packet, 1, &remaining, &used)) {
    return false;
  }
  const size_t pos = 1 + used;
  if (remaining < 2 || packet.size() < pos + 2) {
    return false;
  }
  return packet[pos + 1] == 0;
}

bool ParseMqttPublishPacket(const std::vector<uint8_t>& packet,
                            std::string* topic,
                            std::string* payload) {
  if (!topic || !payload || MqttPacketType(packet) != 3 || packet.empty()) {
    return false;
  }
  size_t remaining = 0;
  size_t used = 0;
  if (!ParseMqttRemainingLength(packet, 1, &remaining, &used)) {
    return false;
  }
  size_t pos = 1 + used;
  if (packet.size() < pos + remaining || remaining < 2) {
    return false;
  }
  if (packet.size() < pos + 2) {
    return false;
  }
  const size_t topic_len =
      (static_cast<size_t>(packet[pos]) << 8) | packet[pos + 1];
  pos += 2;
  if (packet.size() < pos + topic_len) {
    return false;
  }
  topic->assign(reinterpret_cast<const char*>(&packet[pos]), topic_len);
  pos += topic_len;

  const int qos = static_cast<int>((packet[0] >> 1) & 0x03);
  if (qos > 0) {
    if (packet.size() < pos + 2) {
      return false;
    }
    pos += 2;
  }

  const size_t packet_end = 1 + used + remaining;
  if (packet_end < pos || packet.size() < packet_end) {
    return false;
  }
  payload->assign(reinterpret_cast<const char*>(&packet[pos]), packet_end - pos);
  return true;
}

}  // namespace

RimeWithWeaselHandler::RimeWithWeaselHandler(UI* ui)
    : m_ui(ui),
      m_active_session(0),
      m_disabled(true),
      m_current_dark_mode(false),
      m_global_ascii_mode(false),
      m_show_notifications_time(1200),
      _UpdateUICallback(NULL),
      m_system_command_callback(),
      m_input_active_context_key(),
      m_ai_login_pending(false),
      m_ai_login_stop(false),
      m_last_input_rect{0, 0, 0, 0},
      m_has_last_input_rect(false),
      m_ai_request_seq(0) {
  m_ui->InServer() = true;
  rime_api = rime_get_api();
  assert(rime_api);
  m_pid = GetCurrentProcessId();
  uint16_t msbit = 0;
  for (auto i = 31; i >= 0; i--) {
    if (m_pid & (1 << i)) {
      msbit = i;
      break;
    }
  }
  m_pid = (m_pid << (31 - msbit));
  _Setup();
}

RimeWithWeaselHandler::~RimeWithWeaselHandler() {
  _StopAIAssistantLoginFlow();
  _DestroyAIPanel();
  m_input_content_store.Clear();
  m_input_active_context_key.clear();
  m_show_notifications.clear();
  m_session_status_map.clear();
  m_app_options.clear();
}

bool add_session = false;
void _UpdateUIStyle(RimeConfig* config, UI* ui, bool initialize);
bool _UpdateUIStyleColor(RimeConfig* config,
                         UIStyle& style,
                         const std::string& color = std::string());
void _LoadAppOptions(RimeConfig* config, AppOptionsByAppName& app_options);

void _RefreshTrayIcon(const RimeSessionId session_id,
                      const std::function<void()> _UpdateUICallback) {
  // Dangerous, don't touch
  static char app_name[256] = {0};
  auto ret = rime_api->get_property(session_id, "client_app", app_name,
                                    sizeof(app_name) - 1);
  if (!ret || u8tow(app_name) == std::wstring(L"explorer.exe"))
    boost::thread th([=]() {
      ::Sleep(100);
      if (_UpdateUICallback)
        _UpdateUICallback();
    });
  else if (_UpdateUICallback)
    _UpdateUICallback();
}

void RimeWithWeaselHandler::_Setup() {
  AppendInputContentInfoLogLine("InputContent setup_ready");
  RIME_STRUCT(RimeTraits, weasel_traits);
  std::string shared_dir = wtou8(WeaselSharedDataPath().wstring());
  std::string user_dir = wtou8(WeaselUserDataPath().wstring());
  weasel_traits.shared_data_dir = shared_dir.c_str();
  weasel_traits.user_data_dir = user_dir.c_str();
  weasel_traits.prebuilt_data_dir = weasel_traits.shared_data_dir;
  std::string distribution_name = wtou8(get_weasel_ime_name());
  weasel_traits.distribution_name = distribution_name.c_str();
  weasel_traits.distribution_code_name = WEASEL_CODE_NAME;
  weasel_traits.distribution_version = WEASEL_VERSION;
  weasel_traits.app_name = "rime.weasel";
  std::string log_dir = WeaselLogPath().u8string();
  weasel_traits.log_dir = log_dir.c_str();
  rime_api->setup(&weasel_traits);
  rime_api->set_notification_handler(&RimeWithWeaselHandler::OnNotify, this);
}

void RimeWithWeaselHandler::Initialize() {
  m_disabled = _IsDeployerRunning();
  if (m_disabled) {
    return;
  }

  m_input_active_context_key.clear();
  AppendInputContentInfoLogLine("InputContent logger_ready");
  LOG(INFO) << "Initializing la rime.";
  rime_api->initialize(NULL);
  if (rime_api->start_maintenance(/*full_check = */ False)) {
    m_disabled = true;
    rime_api->join_maintenance_thread();
  }

  RimeConfig config = {NULL};
  if (rime_api->config_open("weasel", &config)) {
    _LoadAIAssistantConfig(&config);
    if (m_ui) {
      _UpdateUIStyle(&config, m_ui, true);
      _UpdateShowNotifications(&config, true);
      m_current_dark_mode = IsUserDarkMode();
      if (m_current_dark_mode) {
        const int BUF_SIZE = 255;
        char buffer[BUF_SIZE + 1] = {0};
        if (rime_api->config_get_string(&config, "style/color_scheme_dark",
                                        buffer, BUF_SIZE)) {
          std::string color_name(buffer);
          _UpdateUIStyleColor(&config, m_ui->style(), color_name);
        }
      }
      m_base_style = m_ui->style();
    }
    Bool global_ascii = false;
    if (rime_api->config_get_bool(&config, "global_ascii", &global_ascii))
      m_global_ascii_mode = !!global_ascii;
    if (!rime_api->config_get_int(&config, "show_notifications_time",
                                  &m_show_notifications_time))
      m_show_notifications_time = 1200;
    _LoadAppOptions(&config, m_app_options);
    rime_api->config_close(&config);
  }
  m_last_schema_id.clear();
}

void RimeWithWeaselHandler::Finalize() {
  _StopAIAssistantLoginFlow();
  _DestroyAIPanel();
  m_active_session = 0;
  m_disabled = true;
  m_input_content_store.Clear();
  m_input_active_context_key.clear();
  m_session_status_map.clear();
  LOG(INFO) << "Finalizing la rime.";
  rime_api->finalize();
}

DWORD RimeWithWeaselHandler::FindSession(WeaselSessionId ipc_id) {
  if (m_disabled)
    return 0;
  Bool found = rime_api->find_session(to_session_id(ipc_id));
  DLOG(INFO) << "Find session: session_id = " << to_session_id(ipc_id)
             << ", found = " << found;
  return found ? (ipc_id) : 0;
}

DWORD RimeWithWeaselHandler::AddSession(LPWSTR buffer, EatLine eat) {
  if (m_disabled) {
    DLOG(INFO) << "Trying to resume service.";
    EndMaintenance();
    if (m_disabled)
      return 0;
  }
  RimeSessionId session_id = (RimeSessionId)rime_api->create_session();
  if (m_global_ascii_mode) {
    for (const auto& pair : m_session_status_map) {
      if (pair.first) {
        rime_api->set_option(session_id, "ascii_mode",
                             !!pair.second.status.is_ascii_mode);
        break;
      }
    }
  }

  WeaselSessionId ipc_id =
      _GenerateNewWeaselSessionId(m_session_status_map, m_pid);
  DLOG(INFO) << "Add session: created session_id = " << session_id
             << ", ipc_id = " << ipc_id;
  SessionStatus& session_status = new_session_status(ipc_id);
  session_status.style = m_base_style;
  session_status.session_id = session_id;
  _ReadClientInfo(ipc_id, buffer);

  RIME_STRUCT(RimeStatus, status);
  if (rime_api->get_status(session_id, &status)) {
    std::string schema_id = status.schema_id;
    m_last_schema_id = schema_id;
    _LoadSchemaSpecificSettings(ipc_id, schema_id);
    _LoadAppInlinePreeditSet(ipc_id, true);
    _UpdateInlinePreeditStatus(ipc_id);
    _RefreshTrayIcon(session_id, _UpdateUICallback);
    session_status.status = status;
    session_status.__synced = false;
    rime_api->free_status(&status);
  }
  m_ui->style() = session_status.style;
  // show session's welcome message :-) if any
  if (eat) {
    _Respond(ipc_id, eat);
  }
  add_session = true;
  _UpdateUI(ipc_id);
  add_session = false;
  m_active_session = ipc_id;
  return ipc_id;
}

DWORD RimeWithWeaselHandler::RemoveSession(WeaselSessionId ipc_id) {
  if (m_ui)
    m_ui->Hide();
  if (m_disabled)
    return 0;
  DLOG(INFO) << "Remove session: session_id = " << to_session_id(ipc_id);
  // TODO: force committing? otherwise current composition would be lost
  rime_api->destroy_session(to_session_id(ipc_id));
  m_session_status_map.erase(ipc_id);
  m_active_session = 0;
  return 0;
}

void RimeWithWeaselHandler::UpdateColorTheme(BOOL darkMode) {
  RimeConfig config = {NULL};
  if (rime_api->config_open("weasel", &config)) {
    if (m_ui) {
      _UpdateUIStyle(&config, m_ui, true);
      m_current_dark_mode = darkMode;
      if (darkMode) {
        const int BUF_SIZE = 255;
        char buffer[BUF_SIZE + 1] = {0};
        if (rime_api->config_get_string(&config, "style/color_scheme_dark",
                                        buffer, BUF_SIZE)) {
          std::string color_name(buffer);
          _UpdateUIStyleColor(&config, m_ui->style(), color_name);
        }
      }
      m_base_style = m_ui->style();
    }
    rime_api->config_close(&config);
  }

  for (auto& pair : m_session_status_map) {
    RIME_STRUCT(RimeStatus, status);
    if (rime_api->get_status(to_session_id(pair.first), &status)) {
      _LoadSchemaSpecificSettings(pair.first, std::string(status.schema_id));
      _LoadAppInlinePreeditSet(pair.first, true);
      _UpdateInlinePreeditStatus(pair.first);
      pair.second.status = status;
      pair.second.__synced = false;
      rime_api->free_status(&status);
    }
  }
  m_ui->style() = get_session_status(m_active_session).style;
}

BOOL RimeWithWeaselHandler::ProcessKeyEvent(KeyEvent keyEvent,
                                            WeaselSessionId ipc_id,
                                            EatLine eat) {
  DLOG(INFO) << "Process key event: keycode = " << keyEvent.keycode
             << ", mask = " << keyEvent.mask << ", ipc_id = " << ipc_id;
  if (m_disabled)
    return FALSE;
  if (!(keyEvent.mask & ibus::RELEASE_MASK) &&
      keyEvent.keycode == ibus::Keycode::Escape) {
    bool panel_open = false;
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      panel_open = m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
                   IsWindowVisible(m_ai_panel.panel_hwnd);
    }
    if (panel_open) {
      _CancelAIPanelOutput();
      return TRUE;
    }
  }
  if (_TryProcessAIAssistantTrigger(keyEvent, ipc_id, eat)) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
  RimeSessionId session_id = to_session_id(ipc_id);
  if (!(keyEvent.mask & ibus::RELEASE_MASK) && IsCtrlSpaceToggleKey(keyEvent) &&
      !rime_api->get_option(session_id, "ascii_mode")) {
    const std::wstring raw_preedit = CaptureRawCompositionPreeditText(session_id);
    if (!raw_preedit.empty()) {
      DLOG(INFO) << "Ctrl+Space commit raw preedit before ascii switch, chars="
                 << raw_preedit.size() << ", ipc_id=" << ipc_id;
      rime_api->clear_composition(session_id);
      rime_api->set_option(session_id, "ascii_mode", True);
      _Respond(ipc_id, eat, &raw_preedit);
      _UpdateUI(ipc_id);
      m_active_session = ipc_id;
      return TRUE;
    }
  }
  Bool handled = rime_api->process_key(session_id, keyEvent.keycode,
                                       expand_ibus_modifier(keyEvent.mask));
  // vim_mode when keydown only
  if (!handled && !(keyEvent.mask & ibus::Modifier::RELEASE_MASK)) {
    bool isVimBackInCommandMode =
        (keyEvent.keycode == ibus::Keycode::Escape) ||
        ((keyEvent.mask & (1 << 2)) &&
         (keyEvent.keycode == ibus::Keycode::XK_c ||
          keyEvent.keycode == ibus::Keycode::XK_C ||
          keyEvent.keycode == ibus::Keycode::XK_bracketleft));
    if (isVimBackInCommandMode &&
        rime_api->get_option(session_id, "vim_mode") &&
        !rime_api->get_option(session_id, "ascii_mode")) {
      rime_api->set_option(session_id, "ascii_mode", True);
    }
  }
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
  return (BOOL)handled;
}

void RimeWithWeaselHandler::CommitComposition(WeaselSessionId ipc_id) {
  DLOG(INFO) << "Commit composition: ipc_id = " << ipc_id;
  if (m_disabled)
    return;
  rime_api->commit_composition(to_session_id(ipc_id));
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
}

void RimeWithWeaselHandler::ClearComposition(WeaselSessionId ipc_id) {
  DLOG(INFO) << "Clear composition: ipc_id = " << ipc_id;
  if (m_disabled)
    return;
  rime_api->clear_composition(to_session_id(ipc_id));
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
}

void RimeWithWeaselHandler::SelectCandidateOnCurrentPage(
    size_t index,
    WeaselSessionId ipc_id) {
  DLOG(INFO) << "select candidate on current page, ipc_id = " << ipc_id
             << ", index = " << index;
  if (m_disabled)
    return;
  rime_api->select_candidate_on_current_page(to_session_id(ipc_id), index);
}

bool RimeWithWeaselHandler::HighlightCandidateOnCurrentPage(
    size_t index,
    WeaselSessionId ipc_id,
    EatLine eat) {
  DLOG(INFO) << "highlight candidate on current page, ipc_id = " << ipc_id
             << ", index = " << index;
  bool res = rime_api->highlight_candidate_on_current_page(
      to_session_id(ipc_id), index);
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  return res;
}

bool RimeWithWeaselHandler::ChangePage(bool backward,
                                       WeaselSessionId ipc_id,
                                       EatLine eat) {
  DLOG(INFO) << "change page, ipc_id = " << ipc_id
             << (backward ? "backward" : "foreward");
  bool res = rime_api->change_page(to_session_id(ipc_id), backward);
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  return res;
}

void RimeWithWeaselHandler::FocusIn(DWORD client_caps, WeaselSessionId ipc_id) {
  DLOG(INFO) << "Focus in: ipc_id = " << ipc_id
             << ", client_caps = " << client_caps;
  if (m_disabled)
    return;
  if (ipc_id != 0) {
    const std::string context_key = _GetInputContentContextKey(ipc_id);
    m_input_content_store.OnContextSwitch(context_key);
    AppendInputContentInfoLogLine("InputContent context_switch: context_key=" +
                                  context_key + ", ipc_id=" +
                                  std::to_string(ipc_id));
  }
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
}

void RimeWithWeaselHandler::FocusOut(DWORD param, WeaselSessionId ipc_id) {
  DLOG(INFO) << "Focus out: ipc_id = " << ipc_id;
  // Clear composition when focus is lost to prevent stale input state
  // This prevents prediction panel from appearing on next focus with certain key combinations
  rime_api->clear_composition(to_session_id(ipc_id));
  if (m_ui)
    m_ui->Hide();
  m_active_session = 0;
}

void RimeWithWeaselHandler::UpdateInputPosition(RECT const& rc,
                                                WeaselSessionId ipc_id) {
  DLOG(INFO) << "Update input position: (" << rc.left << ", " << rc.top
             << "), ipc_id = " << ipc_id
             << ", m_active_session = " << m_active_session;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_last_input_rect = rc;
    m_has_last_input_rect = true;
  }
  if (m_ui)
    m_ui->UpdateInputPosition(rc);
  if (m_disabled)
    return;
  if (m_active_session != ipc_id) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
  }
}

std::string RimeWithWeaselHandler::m_message_type;
std::string RimeWithWeaselHandler::m_message_value;
std::string RimeWithWeaselHandler::m_message_label;
std::string RimeWithWeaselHandler::m_option_name;
std::mutex RimeWithWeaselHandler::m_notifier_mutex;

void RimeWithWeaselHandler::OnNotify(void* context_object,
                                     uintptr_t session_id,
                                     const char* message_type,
                                     const char* message_value) {
  // may be running in a thread when deploying rime
  RimeWithWeaselHandler* self =
      reinterpret_cast<RimeWithWeaselHandler*>(context_object);
  if (!self || !message_type || !message_value)
    return;
  std::lock_guard<std::mutex> lock(m_notifier_mutex);
  m_message_type = message_type;
  m_message_value = message_value;
  if (RIME_API_AVAILABLE(rime_api, get_state_label) &&
      !strcmp(message_type, "option")) {
    Bool state = message_value[0] != '!';
    const char* option_name = message_value + !state;
    m_option_name = option_name;
    const char* state_label =
        rime_api->get_state_label(session_id, option_name, state);
    if (state_label) {
      m_message_label = std::string(state_label);
    }
  }
}

void RimeWithWeaselHandler::_ReadClientInfo(WeaselSessionId ipc_id,
                                            LPWSTR buffer) {
  std::string app_name;
  // parse request text
  wbufferstream bs(buffer, WEASEL_IPC_BUFFER_LENGTH);
  std::wstring line;
  while (bs.good()) {
    std::getline(bs, line);
    if (!bs.good())
      break;
    // file ends
    if (line == L".")
      break;
    const std::wstring kClientAppKey = L"session.client_app=";
    if (starts_with(line, kClientAppKey)) {
      std::wstring lwr = line;
      to_lower(lwr);
      app_name = wtou8(lwr.substr(kClientAppKey.length()));
    }
  }
  SessionStatus& session_status = get_session_status(ipc_id);
  session_status.client_app = app_name;
  RimeSessionId session_id = session_status.session_id;
  // set app specific options
  if (!app_name.empty()) {
    rime_api->set_property(session_id, "client_app", app_name.c_str());

    auto it = m_app_options.find(app_name);
    if (it != m_app_options.end()) {
      AppOptions& options(m_app_options[it->first]);
      for (const auto& pair : options) {
        DLOG(INFO) << "set app option: " << pair.first << " = " << pair.second;
        rime_api->set_option(session_id, pair.first.c_str(), Bool(pair.second));
      }
    }
  }
  // inline preedit
  bool inline_preedit = session_status.style.inline_preedit;
  rime_api->set_option(session_id, "inline_preedit", Bool(inline_preedit));
  // show soft cursor on weasel panel but not inline
  rime_api->set_option(session_id, "soft_cursor", Bool(!inline_preedit));
}

void RimeWithWeaselHandler::_GetCandidateInfo(CandidateInfo& cinfo,
                                              RimeContext& ctx) {
  cinfo.candies.resize(ctx.menu.num_candidates);
  cinfo.comments.resize(ctx.menu.num_candidates);
  cinfo.labels.resize(ctx.menu.num_candidates);
  for (int i = 0; i < ctx.menu.num_candidates; ++i) {
    cinfo.candies[i].str = escape_string(u8tow(ctx.menu.candidates[i].text));
    if (ctx.menu.candidates[i].comment) {
      cinfo.comments[i].str =
          escape_string(u8tow(ctx.menu.candidates[i].comment));
    }
    if (RIME_STRUCT_HAS_MEMBER(ctx, ctx.select_labels) && ctx.select_labels) {
      cinfo.labels[i].str = escape_string(u8tow(ctx.select_labels[i]));
    } else if (ctx.menu.select_keys) {
      cinfo.labels[i].str =
          escape_string(std::wstring(1, ctx.menu.select_keys[i]));
    } else {
      cinfo.labels[i].str = std::to_wstring((i + 1) % 10);
    }
  }
  cinfo.highlighted = ctx.menu.highlighted_candidate_index;
  cinfo.currentPage = ctx.menu.page_no;
  cinfo.is_last_page = ctx.menu.is_last_page;
}

void RimeWithWeaselHandler::StartMaintenance() {
  m_session_status_map.clear();
  Finalize();
  _UpdateUI(0);
}

void RimeWithWeaselHandler::EndMaintenance() {
  if (m_disabled) {
    Initialize();
    _UpdateUI(0);
  }
  m_session_status_map.clear();
}

void RimeWithWeaselHandler::SetOption(WeaselSessionId ipc_id,
                                      const std::string& opt,
                                      bool val) {
  // from no-session client, not actual typing session
  if (!ipc_id) {
    if (m_global_ascii_mode && opt == "ascii_mode") {
      for (auto& pair : m_session_status_map)
        rime_api->set_option(to_session_id(pair.first), "ascii_mode", val);
    } else {
      rime_api->set_option(to_session_id(m_active_session), opt.c_str(), val);
    }
  } else {
    rime_api->set_option(to_session_id(ipc_id), opt.c_str(), val);
  }
}

void RimeWithWeaselHandler::OnUpdateUI(std::function<void()> const& cb) {
  _UpdateUICallback = cb;
}

void RimeWithWeaselHandler::OnSystemCommand(
    std::function<void(const SystemCommandLaunchRequest&)> const& cb) {
  m_system_command_callback = cb;
}

void RimeWithWeaselHandler::_LoadAIAssistantConfig(RimeConfig* config) {
  m_ai_config = AIAssistantConfig();

  Bool enabled = False;
  if (rime_api->config_get_bool(config, "ai_assistant/enabled", &enabled)) {
    m_ai_config.enabled = !!enabled;
  }
  Bool stream = True;
  if (rime_api->config_get_bool(config, "ai_assistant/stream", &stream)) {
    m_ai_config.stream = !!stream;
  }
  Bool login_required = False;
  if (rime_api->config_get_bool(config, "ai_assistant/login_required",
                                &login_required)) {
    m_ai_config.login_required = !!login_required;
  }
  Bool debug_dump_context = False;
  if (rime_api->config_get_bool(config, "ai_assistant/debug_dump_context",
                                &debug_dump_context)) {
    m_ai_config.debug_dump_context = !!debug_dump_context;
  }
  ReadConfigString(config, "ai_assistant/trigger_hotkey",
                   &m_ai_config.trigger_hotkey);
  ReadConfigString(config, "ai_assistant/endpoint", &m_ai_config.endpoint);
  ReadConfigString(config, "ai_assistant/api_key", &m_ai_config.api_key);
  ReadConfigString(config, "ai_assistant/model", &m_ai_config.model);
  ReadConfigString(config, "ai_assistant/debug_dump_path",
                   &m_ai_config.debug_dump_path);
  ReadConfigString(config, "ai_assistant/panel_url", &m_ai_config.panel_url);
  ReadConfigString(config, "ai_assistant/panel_allowed_origin",
                   &m_ai_config.panel_allowed_origin);
  ReadConfigString(config, "ai_assistant/system_prompt",
                   &m_ai_config.system_prompt);
  ReadConfigString(config, "ai_assistant/reasoning_effort",
                   &m_ai_config.reasoning_effort);
  ReadConfigString(config, "ai_assistant/login_url", &m_ai_config.login_url);
  ReadConfigString(config, "ai_assistant/login_state_path",
                   &m_ai_config.login_state_path);
  ReadConfigString(config, "ai_assistant/login_token_key",
                   &m_ai_config.login_token_key);
  ReadConfigString(config, "ai_assistant/refresh_token_endpoint",
                   &m_ai_config.refresh_token_endpoint);
  ReadConfigString(config, "ai_assistant/mqtt_url", &m_ai_config.mqtt_url);
  ReadConfigString(config, "ai_assistant/mqtt_topic_template",
                   &m_ai_config.mqtt_topic_template);
  ReadConfigString(config, "ai_assistant/mqtt_username",
                   &m_ai_config.mqtt_username);
  ReadConfigString(config, "ai_assistant/mqtt_password",
                   &m_ai_config.mqtt_password);
  ReadConfigInt(config, "ai_assistant/max_history_chars",
                &m_ai_config.max_history_chars);
  ReadConfigInt(config, "ai_assistant/timeout_ms", &m_ai_config.timeout_ms);
  ReadConfigInt(config, "ai_assistant/mqtt_timeout_ms",
                &m_ai_config.mqtt_timeout_ms);

  if (m_ai_config.model.empty()) {
    m_ai_config.model = "gpt-5";
  }
  if (m_ai_config.reasoning_effort.empty()) {
    m_ai_config.reasoning_effort = "low";
  }
  if (m_ai_config.trigger_hotkey.empty()) {
    m_ai_config.trigger_hotkey = "Control+3";
  }
  if (!TryParseAIAssistantTriggerHotkey(m_ai_config.trigger_hotkey,
                                        &m_ai_config.trigger_keycode,
                                        &m_ai_config.trigger_modifiers)) {
    LOG(WARNING) << "Invalid ai_assistant/trigger_hotkey: "
                 << m_ai_config.trigger_hotkey
                 << "; fallback to Control+3.";
    m_ai_config.trigger_hotkey = "Control+3";
    m_ai_config.trigger_keycode = '3';
    m_ai_config.trigger_modifiers = ibus::CONTROL_MASK;
  }
  if (m_ai_config.max_history_chars <= 0) {
    m_ai_config.max_history_chars = 2048;
  }
  if (m_ai_config.timeout_ms <= 0) {
    m_ai_config.timeout_ms = 30000;
  }
  if (m_ai_config.mqtt_timeout_ms <= 0) {
    m_ai_config.mqtt_timeout_ms = 120000;
  }

  m_input_content_store.SetLimits(
      static_cast<size_t>(max(1024, m_ai_config.max_history_chars * 2)),
      64, 16);
}

void RimeWithWeaselHandler::_StopAIAssistantLoginFlow() {
  m_ai_login_stop.store(true);
  if (m_ai_login_thread.joinable()) {
    m_ai_login_thread.join();
  }
  m_ai_login_pending.store(false);
}

bool RimeWithWeaselHandler::_IsAIAssistantLoggedIn() {
  if (!m_ai_config.login_required) {
    return true;
  }
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    if (!m_ai_login_token.empty()) {
      return true;
    }
  }
  std::string persisted_token;
  std::string persisted_tenant_id;
  std::string persisted_refresh_token;
  if (!LoadAIAssistantLoginIdentity(m_ai_config, &persisted_token,
                                    &persisted_tenant_id,
                                    &persisted_refresh_token)) {
    return false;
  }
  if (persisted_token.empty() && !persisted_refresh_token.empty() &&
      !persisted_tenant_id.empty()) {
    std::string refreshed_token;
    std::string refreshed_refresh_token;
    std::string refresh_endpoint;
    std::string refresh_error;
    const bool refreshed = RefreshAIAssistantAccessToken(
        m_ai_config, persisted_refresh_token, persisted_tenant_id,
        &refreshed_token, &refreshed_refresh_token, &refresh_endpoint,
        &refresh_error);
    if (refreshed && !refreshed_token.empty()) {
      persisted_token = refreshed_token;
      if (!refreshed_refresh_token.empty()) {
        persisted_refresh_token = refreshed_refresh_token;
      }
      SaveAIAssistantLoginState(m_ai_config, persisted_token, persisted_tenant_id,
                                persisted_refresh_token, std::string(),
                                std::string(), std::string());
    }
  }
  if (persisted_token.empty()) {
    return false;
  }
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    m_ai_login_token = persisted_token;
    m_ai_login_tenant_id = persisted_tenant_id;
    m_ai_login_refresh_token = persisted_refresh_token;
  }
  return true;
}

bool RimeWithWeaselHandler::_StartAIAssistantLoginFlow() {
  if (!m_ai_config.login_required) {
    return true;
  }
  if (m_ai_login_pending.load()) {
    return true;
  }
  if (m_ai_login_thread.joinable()) {
    m_ai_login_thread.join();
  }

  const std::string client_id = GenerateLoginClientId();
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    m_ai_login_client_id = client_id;
  }

  const std::wstring login_url = BuildLoginUrlForBrowser(m_ai_config, client_id);
  if (login_url.empty()) {
    LOG(WARNING) << "AI login required but login_url is empty.";
    return false;
  }

  m_ai_login_stop.store(false);
  if (!m_ai_config.mqtt_url.empty() && !m_ai_config.mqtt_topic_template.empty()) {
    m_ai_login_pending.store(true);
    m_ai_login_thread = std::thread([this, client_id]() {
      _RunAIAssistantLoginListener(client_id);
    });
  } else {
    m_ai_login_pending.store(false);
    LOG(WARNING)
        << "AI login mqtt_url or mqtt_topic_template not configured; "
        << "login callback listener disabled.";
  }

  const auto open_result = reinterpret_cast<intptr_t>(
      ShellExecuteW(nullptr, L"open", login_url.c_str(), nullptr, nullptr,
                    SW_SHOWNORMAL));
  if (open_result <= 32) {
    LOG(WARNING) << "Failed to open login page, code=" << open_result;
    return false;
  }
  LOG(INFO) << "AI login started, client_id=" << client_id;
  return true;
}

bool RimeWithWeaselHandler::_EnsureAIAssistantLogin() {
  if (!m_ai_config.login_required) {
    return true;
  }
  if (_IsAIAssistantLoggedIn()) {
    return true;
  }
  _StartAIAssistantLoginFlow();
  return false;
}

bool RimeWithWeaselHandler::_ForceAIAssistantRelogin() {
  if (!m_ai_config.login_required) {
    return true;
  }

  _StopAIAssistantLoginFlow();

  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    m_ai_login_token.clear();
    m_ai_login_tenant_id.clear();
    m_ai_login_refresh_token.clear();
  }

  const std::filesystem::path state_path(
      ResolveAIAssistantLoginStatePath(m_ai_config));
  std::error_code remove_error;
  std::filesystem::remove(state_path, remove_error);
  if (remove_error) {
    LOG(WARNING) << "AI relogin: failed to remove login state file: "
                 << state_path.u8string()
                 << ", error=" << remove_error.message();
  }

  const bool started = _StartAIAssistantLoginFlow();
  if (!started) {
    LOG(WARNING) << "AI relogin: unable to start login flow.";
  } else {
    LOG(INFO) << "AI relogin: login flow started by ui.auth.refresh_request.";
  }
  return started;
}

void RimeWithWeaselHandler::_RunAIAssistantLoginListener(
    const std::string& client_id) {
  struct PendingGuard {
    std::atomic<bool>* pending = nullptr;
    explicit PendingGuard(std::atomic<bool>* value) : pending(value) {}
    ~PendingGuard() {
      if (pending) {
        pending->store(false);
      }
    }
  } pending_guard(&m_ai_login_pending);

  const AIAssistantConfig config = m_ai_config;
  const std::string expected_topic = BuildMqttTopicForClient(config, client_id);
  if (expected_topic.empty()) {
    LOG(WARNING) << "AI login mqtt topic is empty.";
    return;
  }

  const ParsedWebSocketUrl ws_url = ParseWebSocketUrl(u8tow(config.mqtt_url));
  if (!ws_url.valid) {
    LOG(WARNING) << "AI login mqtt_url invalid: " << config.mqtt_url;
    return;
  }

  HINTERNET session = nullptr;
  HINTERNET connection = nullptr;
  HINTERNET request = nullptr;
  HINTERNET websocket = nullptr;
  auto close_handle = [](HINTERNET* handle) {
    if (handle && *handle) {
      WinHttpCloseHandle(*handle);
      *handle = nullptr;
    }
  };

  session =
      WinHttpOpen(L"WeaselAIAssistantLogin/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                  WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
  if (!session) {
    return;
  }
  WinHttpSetTimeouts(session, 5000, 5000, 5000, 1000);

  connection = WinHttpConnect(session, ws_url.host.c_str(), ws_url.port, 0);
  if (!connection) {
    close_handle(&session);
    return;
  }

  const DWORD open_flags = ws_url.secure ? WINHTTP_FLAG_SECURE : 0;
  request = WinHttpOpenRequest(connection, L"GET", ws_url.path.c_str(), nullptr,
                               WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                               open_flags);
  if (!request) {
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  if (!WinHttpSetOption(request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, nullptr,
                        0)) {
    close_handle(&request);
    close_handle(&connection);
    close_handle(&session);
    return;
  }
  WinHttpAddRequestHeaders(request, L"Sec-WebSocket-Protocol: mqtt", -1,
                           WINHTTP_ADDREQ_FLAG_ADD);

  if (!WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                          WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
      !WinHttpReceiveResponse(request, nullptr)) {
    close_handle(&request);
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  websocket = WinHttpWebSocketCompleteUpgrade(request, 0);
  close_handle(&request);
  if (!websocket) {
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  const auto connect_packet = BuildMqttConnectPacket(
      client_id, config.mqtt_username, config.mqtt_password, 30);
  if (!SendWebSocketBinaryMessage(websocket, connect_packet)) {
    close_handle(&websocket);
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  const auto deadline = std::chrono::steady_clock::now() +
                        std::chrono::milliseconds(
                            max(1000, config.mqtt_timeout_ms));
  std::vector<uint8_t> packet;
  bool connack_ok = false;
  while (!m_ai_login_stop.load() && std::chrono::steady_clock::now() < deadline) {
    const auto receive_result = ReceiveWebSocketBinaryMessage(websocket, &packet);
    if (receive_result == WebSocketReceiveResult::kTimeout) {
      continue;
    }
    if (receive_result != WebSocketReceiveResult::kOk) {
      break;
    }
    if (IsMqttConnAckOk(packet)) {
      connack_ok = true;
      break;
    }
  }
  if (!connack_ok) {
    close_handle(&websocket);
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  const auto subscribe_packet = BuildMqttSubscribePacket(expected_topic, 1);
  if (!SendWebSocketBinaryMessage(websocket, subscribe_packet)) {
    close_handle(&websocket);
    close_handle(&connection);
    close_handle(&session);
    return;
  }

  while (!m_ai_login_stop.load() && std::chrono::steady_clock::now() < deadline) {
    const auto receive_result = ReceiveWebSocketBinaryMessage(websocket, &packet);
    if (receive_result == WebSocketReceiveResult::kTimeout) {
      continue;
    }
    if (receive_result != WebSocketReceiveResult::kOk) {
      break;
    }
    if (MqttPacketType(packet) != 3) {
      continue;
    }
    std::string topic;
    std::string payload;
    if (!ParseMqttPublishPacket(packet, &topic, &payload)) {
      continue;
    }
    if (topic != expected_topic && topic.find(client_id) == std::string::npos) {
      continue;
    }
    std::string token;
    std::string tenant_id;
    std::string refresh_token;
    ExtractLoginIdentityFromPayload(payload, config.login_token_key, &token,
                                    &tenant_id, &refresh_token);
    if (!refresh_token.empty()) {
      std::string refreshed_token;
      std::string refreshed_refresh_token;
      std::string refresh_endpoint;
      std::string refresh_error;
      const bool refreshed = RefreshAIAssistantAccessToken(
          config, refresh_token, tenant_id, &refreshed_token,
          &refreshed_refresh_token, &refresh_endpoint, &refresh_error);
      if (refreshed && !refreshed_token.empty()) {
        token = refreshed_token;
        if (!refreshed_refresh_token.empty()) {
          refresh_token = refreshed_refresh_token;
        }
        LOG(INFO) << "AI login token refreshed via endpoint "
                  << refresh_endpoint;
      } else {
        LOG(WARNING) << "AI login refresh token exchange failed: "
                     << refresh_error;
      }
    }
    if (token.empty()) {
      continue;
    }
    SaveAIAssistantLoginState(config, token, tenant_id, refresh_token,
                              client_id, topic, payload);
    {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      m_ai_login_token = token;
      m_ai_login_tenant_id = tenant_id;
      m_ai_login_refresh_token = refresh_token;
    }
    LOG(INFO) << "AI login success via mqtt, client_id=" << client_id;
    break;
  }

  if (websocket) {
    WinHttpWebSocketClose(websocket, WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS,
                          nullptr, 0);
  }
  close_handle(&websocket);
  close_handle(&connection);
  close_handle(&session);
}

bool RimeWithWeaselHandler::_TryProcessAIAssistantTrigger(KeyEvent keyEvent,
                                                          WeaselSessionId ipc_id,
                                                          EatLine eat) {
  if (!IsAIAssistantTriggerKey(m_ai_config, keyEvent)) {
    return false;
  }
  if (m_ai_config.panel_url.empty()) {
    LOG(WARNING) << "AI panel url is empty; skip AI trigger.";
    return false;
  }
  if (keyEvent.mask & ibus::RELEASE_MASK) {
    return true;
  }
  if (!_EnsureAIAssistantLogin()) {
    return true;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);

  // 【修复】在提交组字之前获取焦点窗口
  // 优先使用 GetFocus() 获取当前焦点窗口（更准确）
  // 如果 GetFocus() 返回空，则使用 GetForegroundWindow()
  HWND target_hwnd = GetFocus();
  if (!target_hwnd || !IsWindow(target_hwnd)) {
    target_hwnd = GetForegroundWindow();
  }
  DLOG(INFO) << "AI assistant: target_hwnd=0x" << std::hex << reinterpret_cast<uintptr_t>(target_hwnd);

  const std::wstring preedit_snapshot = CaptureCurrentPreeditText(session_id);
  std::wstring prefix_commit;

  RIME_STRUCT(RimeStatus, status);
  const bool got_status = rime_api->get_status(session_id, &status);
  const bool had_composition = got_status && !!status.is_composing;
  if (got_status) {
    rime_api->free_status(&status);
  }

  if (had_composition) {
    rime_api->commit_composition(session_id);
    prefix_commit = _TakePendingCommitText(session_id);
    if (!prefix_commit.empty()) {
      _Respond(ipc_id, eat, &prefix_commit);
    }
  }

  const size_t max_context_chars =
      static_cast<size_t>(max(1, m_ai_config.max_history_chars));
  const std::wstring prompt_tail =
      prefix_commit.empty() ? preedit_snapshot : std::wstring();

  const std::string context_source = "ime_history_only";
  std::wstring prompt_context = _CollectAIAssistantContext(ipc_id, prompt_tail);
  DLOG(INFO) << "AI context source: ime_history_only, chars="
             << prompt_context.size();
  AppendAIAssistantContextDump(m_ai_config, context_source, target_hwnd,
                               prompt_context);

  if (_OpenAIPanel(ipc_id, target_hwnd, 0)) {
    const std::wstring normalized_prompt_context =
        NormalizeReferencedContextText(prompt_context);
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      m_ai_panel.context_text = normalized_prompt_context;
    }
    _ResetAIPanelOutput();
    if (normalized_prompt_context.empty()) {
      _SetAIPanelStatus(L"未检测到输入内容，请在前端面板中继续操作。");
    } else {
      _SetAIPanelStatus(L"上下文已就绪，请在前端面板中发起请求。");
    }
    return true;
  }

  LOG(WARNING)
      << "AI panel window unavailable; skip sync fallback to avoid blocking.";
  return true;
}

bool RimeWithWeaselHandler::_EnsureAIPanelWindow() {
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd)) {
      return true;
    }
  }

  HANDLE ready_event = CreateEventW(nullptr, TRUE, FALSE, nullptr);
  if (!ready_event) {
    return false;
  }

  std::thread([this, ready_event]() {
    const HRESULT co_initialize_result =
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool should_uninitialize_com = SUCCEEDED(co_initialize_result);

    HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW wnd_class = {0};
    if (!GetClassInfoExW(instance, kAIPanelWindowClass, &wnd_class)) {
      wnd_class.cbSize = sizeof(wnd_class);
      wnd_class.style = CS_HREDRAW | CS_VREDRAW;
      wnd_class.lpfnWndProc = &RimeWithWeaselHandler::AIAssistantPanelWndProc;
      wnd_class.hInstance = instance;
      wnd_class.hCursor = LoadCursor(nullptr, IDC_ARROW);
      wnd_class.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
      wnd_class.lpszClassName = kAIPanelWindowClass;
      if (!RegisterClassExW(&wnd_class) &&
          GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        SetEvent(ready_event);
        if (should_uninitialize_com) {
          CoUninitialize();
        }
        return;
      }
    }

    const DWORD panel_style = WS_POPUP | WS_CLIPCHILDREN | WS_SIZEBOX;
    HWND panel_hwnd = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST, kAIPanelWindowClass,
        L"Weasel AI Assistant", panel_style, CW_USEDEFAULT, CW_USEDEFAULT,
        kAIPanelWidth, kAIPanelHeight, nullptr, nullptr, instance, this);
    if (!panel_hwnd) {
      SetEvent(ready_event);
      if (should_uninitialize_com) {
        CoUninitialize();
      }
      return;
    }
    TryDisableAIPanelWindowBorder(panel_hwnd);
    ApplyAIPanelRoundedRegion(panel_hwnd);

    HWND status_hwnd = CreateWindowExW(
        0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0,
        panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlStatus)),
        instance, nullptr);
    HWND output_hwnd = CreateWindowExW(
        0, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE |
            ES_AUTOVSCROLL | ES_READONLY,
        0, 0, 0, 0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlOutput)),
        instance, nullptr);
    HWND request_hwnd = CreateWindowExW(
        0, L"BUTTON", L"请求", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0,
        0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlRequest)),
        instance, nullptr);
    HWND confirm_hwnd = CreateWindowExW(
        0, L"BUTTON", L"确认回写", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0,
        0, 0, 0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlConfirm)),
        instance, nullptr);
    HWND cancel_hwnd = CreateWindowExW(
        0, L"BUTTON", L"取消", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlCancel)),
        instance, nullptr);
    if (!status_hwnd || !output_hwnd || !request_hwnd || !confirm_hwnd ||
        !cancel_hwnd) {
      DestroyWindow(panel_hwnd);
      SetEvent(ready_event);
      if (should_uninitialize_com) {
        CoUninitialize();
      }
      return;
    }

    const HFONT default_font =
        static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    SendMessageW(status_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    SendMessageW(output_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    SendMessageW(request_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    SendMessageW(confirm_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    SendMessageW(cancel_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    EnableWindow(request_hwnd, FALSE);
    EnableWindow(confirm_hwnd, FALSE);
    ShowWindow(status_hwnd, SW_HIDE);
    ShowWindow(request_hwnd, SW_HIDE);
    ShowWindow(confirm_hwnd, SW_HIDE);
    ShowWindow(cancel_hwnd, SW_HIDE);
    ShowWindow(panel_hwnd, SW_HIDE);
    LayoutAIPanelControls(panel_hwnd);

    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      m_ai_panel.panel_hwnd = panel_hwnd;
      m_ai_panel.status_hwnd = status_hwnd;
      m_ai_panel.output_hwnd = output_hwnd;
      m_ai_panel.webview_hwnd = panel_hwnd;
      m_ai_panel.webview_controller = nullptr;
      m_ai_panel.webview = nullptr;
      m_ai_panel.request_hwnd = request_hwnd;
      m_ai_panel.confirm_hwnd = confirm_hwnd;
      m_ai_panel.cancel_hwnd = cancel_hwnd;
      m_ai_panel.context_text.clear();
      m_ai_panel.status_text.clear();
      m_ai_panel.output_text.clear();
      m_ai_panel.institution_options.clear();
      m_ai_panel.selected_institution_id.clear();
      m_ai_panel.panel_width = kAIPanelWidth;
      m_ai_panel.panel_height = kAIPanelHeight;
      m_ai_panel.last_panel_x = 0;
      m_ai_panel.last_panel_y = 0;
      m_ai_panel.has_last_panel_position = false;
      m_ai_panel.webview_ready = false;
      m_ai_panel.requesting = false;
      m_ai_panel.institutions_loading = false;
      m_ai_panel.completed = false;
      m_ai_panel.has_error = false;
    }
    SetEvent(ready_event);

    MSG message = {0};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
      TranslateMessage(&message);
      DispatchMessageW(&message);
    }

    {
      void* controller = nullptr;
      void* webview = nullptr;
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      if (m_ai_panel.panel_hwnd == panel_hwnd) {
        controller = m_ai_panel.webview_controller;
        webview = m_ai_panel.webview;
        m_ai_panel.panel_hwnd = nullptr;
        m_ai_panel.status_hwnd = nullptr;
        m_ai_panel.output_hwnd = nullptr;
        m_ai_panel.webview_hwnd = nullptr;
        m_ai_panel.webview_controller = nullptr;
        m_ai_panel.webview = nullptr;
        m_ai_panel.request_hwnd = nullptr;
        m_ai_panel.confirm_hwnd = nullptr;
        m_ai_panel.cancel_hwnd = nullptr;
        m_ai_panel.target_hwnd = nullptr;
        m_ai_panel.context_text.clear();
        m_ai_panel.status_text.clear();
        m_ai_panel.output_text.clear();
        m_ai_panel.institution_options.clear();
        m_ai_panel.selected_institution_id.clear();
        m_ai_panel.panel_width = kAIPanelWidth;
        m_ai_panel.panel_height = kAIPanelHeight;
        m_ai_panel.last_panel_x = 0;
        m_ai_panel.last_panel_y = 0;
        m_ai_panel.has_last_panel_position = false;
        m_ai_panel.webview_ready = false;
        m_ai_panel.requesting = false;
        m_ai_panel.institutions_loading = false;
      }
#if WEASEL_HAS_WEBVIEW2
      if (webview) {
        static_cast<ICoreWebView2*>(webview)->Release();
      }
      if (controller) {
        static_cast<ICoreWebView2Controller*>(controller)->Release();
      }
#endif
    }

    if (should_uninitialize_com) {
      CoUninitialize();
    }
  }).detach();

  const DWORD wait_result = WaitForSingleObject(ready_event, 800);
  CloseHandle(ready_event);
  if (wait_result != WAIT_OBJECT_0) {
    LOG(WARNING) << "AI panel init timed out.";
    return false;
  }

  std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
  if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd)) {
    return true;
  }
  return false;
}

void RimeWithWeaselHandler::_ApplyAIPanelSizeAndReposition(
    int requested_width,
    int requested_height,
    bool prefer_anchor_position) {
  HWND panel_hwnd = nullptr;
  HWND target_hwnd = nullptr;
  RECT anchor_rect = {0, 0, 0, 0};
  bool has_anchor = false;
  bool has_last_position = false;
  int last_panel_x = 0;
  int last_panel_y = 0;
  int width = kAIPanelWidth;
  int height = kAIPanelHeight;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    if (!panel_hwnd || !IsWindow(panel_hwnd)) {
      return;
    }
    target_hwnd = m_ai_panel.target_hwnd;
    width = m_ai_panel.panel_width > 0 ? m_ai_panel.panel_width : kAIPanelWidth;
    height =
        m_ai_panel.panel_height > 0 ? m_ai_panel.panel_height : kAIPanelHeight;
    if (requested_width > 0) {
      width = requested_width;
    }
    if (requested_height > 0) {
      height = requested_height;
    }
    width = max(kAIPanelMinWidth, min(kAIPanelMaxWidth, width));
    height = max(kAIPanelMinHeight, min(kAIPanelMaxHeight, height));
    m_ai_panel.panel_width = width;
    m_ai_panel.panel_height = height;
    if (m_has_last_input_rect) {
      anchor_rect = m_last_input_rect;
      has_anchor = true;
    }
    has_last_position = m_ai_panel.has_last_panel_position;
    last_panel_x = m_ai_panel.last_panel_x;
    last_panel_y = m_ai_panel.last_panel_y;
  }

  RECT panel_rect = {0, 0, 0, 0};
  if (panel_hwnd && IsWindow(panel_hwnd) && GetWindowRect(panel_hwnd, &panel_rect)) {
    has_last_position = true;
    last_panel_x = panel_rect.left;
    last_panel_y = panel_rect.top;
  }

  RECT work_rect = {0, 0, 0, 0};
  HMONITOR monitor = nullptr;
  if (has_anchor) {
    monitor = MonitorFromRect(&anchor_rect, MONITOR_DEFAULTTONEAREST);
  }
  if (!monitor && target_hwnd && IsWindow(target_hwnd)) {
    monitor = MonitorFromWindow(target_hwnd, MONITOR_DEFAULTTONEAREST);
  }
  if (!monitor) {
    monitor = MonitorFromWindow(panel_hwnd, MONITOR_DEFAULTTONEAREST);
  }
  MONITORINFO monitor_info = {0};
  monitor_info.cbSize = sizeof(monitor_info);
  if (monitor && GetMonitorInfoW(monitor, &monitor_info)) {
    work_rect = monitor_info.rcWork;
  } else {
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_rect, 0);
  }

  const int min_x = work_rect.left + kAIPanelScreenMargin;
  const int max_x = max(min_x, work_rect.right - kAIPanelScreenMargin - width);
  const int min_y = work_rect.top + kAIPanelScreenMargin;
  const int max_y =
      max(min_y, work_rect.bottom - kAIPanelScreenMargin - height);
  int x = min_x;
  int y = min_y;
  if (prefer_anchor_position && has_anchor) {
    x = anchor_rect.left;
    y = anchor_rect.bottom + 6;
  } else if (has_last_position) {
    x = last_panel_x;
    y = last_panel_y;
  } else if (has_anchor) {
    x = anchor_rect.left;
    y = anchor_rect.bottom + 6;
    if (y + height > work_rect.bottom - kAIPanelScreenMargin) {
      const int above = anchor_rect.top - height - 6;
      if (above >= min_y) {
        y = above;
      }
    }
  } else {
    x = work_rect.left +
        max(0, (work_rect.right - work_rect.left - width) / 2);
    y = work_rect.top +
        max(0, (work_rect.bottom - work_rect.top - height) / 2);
  }

  x = max(min_x, min(max_x, x));
  if (prefer_anchor_position && has_anchor) {
    // Ctrl+3 打开时固定锚在光标下方，不做“上翻”避让。
    y = max(min_y, y);
  } else {
    y = max(min_y, min(max_y, y));
  }

  SetWindowPos(panel_hwnd, HWND_TOPMOST, x, y, width, height,
               SWP_NOACTIVATE);
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd == panel_hwnd) {
      m_ai_panel.last_panel_x = x;
      m_ai_panel.last_panel_y = y;
      m_ai_panel.has_last_panel_position = true;
    }
  }
  _ResizeAIPanelWebView();
}

bool RimeWithWeaselHandler::_OpenAIPanel(WeaselSessionId ipc_id,
                                         HWND target_hwnd,
                                         uint64_t request_id) {
  if (!_EnsureAIPanelWindow()) {
    return false;
  }

  HWND panel_hwnd = nullptr;
  HWND status_hwnd = nullptr;
  HWND output_hwnd = nullptr;
  HWND request_hwnd = nullptr;
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (!m_ai_panel.panel_hwnd || !IsWindow(m_ai_panel.panel_hwnd)) {
      return false;
    }
    m_ai_panel.ipc_id = ipc_id;
    m_ai_panel.request_id = request_id;
    m_ai_panel.target_hwnd = target_hwnd;
    if (m_ai_panel.target_hwnd == m_ai_panel.panel_hwnd ||
        !IsWindow(m_ai_panel.target_hwnd)) {
      m_ai_panel.target_hwnd = nullptr;
    }
    m_ai_panel.output_text.clear();
    m_ai_panel.selected_institution_id.clear();
    m_ai_panel.status_text = L"上下文已就绪，请在前端面板中发起请求。";
    m_ai_panel.requesting = false;
    m_ai_panel.institutions_loading = true;
    m_ai_panel.completed = false;
    m_ai_panel.has_error = false;
    panel_hwnd = m_ai_panel.panel_hwnd;
    status_hwnd = m_ai_panel.status_hwnd;
    output_hwnd = m_ai_panel.output_hwnd;
    request_hwnd = m_ai_panel.request_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
  }

  SetWindowTextW(status_hwnd, L"上下文已就绪，请在前端面板中发起请求。");
  SetWindowTextW(output_hwnd, L"");
  EnableWindow(request_hwnd, TRUE);
  EnableWindow(confirm_hwnd, FALSE);
  EnableWindow(cancel_hwnd, TRUE);
  SetWindowTextW(cancel_hwnd, L"取消");

  _ApplyAIPanelSizeAndReposition(0, 0, true);
  SetWindowPos(panel_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
               SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE);
  ShowWindow(panel_hwnd, SW_SHOWNORMAL);
  UpdateWindow(panel_hwnd);
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_INIT, 0, 0);
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  _RefreshAIPanelInstitutionOptions();
  return true;
}

void RimeWithWeaselHandler::_CloseAIPanel() {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    m_ai_panel.target_hwnd = nullptr;
    m_ai_panel.request_id = 0;
    m_ai_panel.ipc_id = 0;
    m_ai_panel.status_text.clear();
    m_ai_panel.output_text.clear();
    m_ai_panel.requesting = false;
    m_ai_panel.completed = false;
    m_ai_panel.has_error = false;
  }

  if (panel_hwnd && IsWindow(panel_hwnd)) {
    ShowWindow(panel_hwnd, SW_HIDE);
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_DestroyAIPanel() {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (!panel_hwnd || !IsWindow(panel_hwnd)) {
    return;
  }

  PostMessageW(panel_hwnd, WM_AI_PANEL_DESTROY, 0, 0);
  for (int i = 0; i < 40; ++i) {
    if (!IsWindow(panel_hwnd)) {
      break;
    }
    Sleep(25);
  }
}

void RimeWithWeaselHandler::_SetAIPanelStatus(const std::wstring& status_text) {
  HWND status_hwnd = nullptr;
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_ai_panel.status_text = status_text;
    status_hwnd = m_ai_panel.status_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (status_hwnd && IsWindow(status_hwnd)) {
    SetWindowTextW(status_hwnd, status_text.c_str());
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_ResetAIPanelOutput() {
  HWND output_hwnd = nullptr;
  HWND request_hwnd = nullptr;
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_ai_panel.output_text.clear();
    m_ai_panel.requesting = false;
    m_ai_panel.completed = false;
    m_ai_panel.has_error = false;
    output_hwnd = m_ai_panel.output_hwnd;
    request_hwnd = m_ai_panel.request_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (output_hwnd && IsWindow(output_hwnd)) {
    SetWindowTextW(output_hwnd, L"");
  }
  if (request_hwnd && IsWindow(request_hwnd)) {
    EnableWindow(request_hwnd, TRUE);
  }
  if (confirm_hwnd && IsWindow(confirm_hwnd)) {
    EnableWindow(confirm_hwnd, FALSE);
  }
  if (cancel_hwnd && IsWindow(cancel_hwnd)) {
    EnableWindow(cancel_hwnd, TRUE);
    SetWindowTextW(cancel_hwnd, L"取消");
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_AppendAIPanelOutput(const std::wstring& chunk) {
  if (chunk.empty()) {
    return;
  }

  HWND output_hwnd = nullptr;
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_ai_panel.output_text.append(chunk);
    output_hwnd = m_ai_panel.output_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;
  }

  if (output_hwnd && IsWindow(output_hwnd)) {
    const LRESULT text_length = GetWindowTextLengthW(output_hwnd);
    SendMessageW(output_hwnd, EM_SETSEL, text_length, text_length);
    SendMessageW(output_hwnd, EM_REPLACESEL, FALSE,
                 reinterpret_cast<LPARAM>(chunk.c_str()));
    SendMessageW(output_hwnd, EM_SCROLLCARET, 0, 0);
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_CompleteAIPanel(bool has_error,
                                             const std::wstring& error_text) {
  if (has_error && !error_text.empty()) {
    const std::wstring prefix = L"\r\n[错误] ";
    _AppendAIPanelOutput(prefix + error_text);
  }

  HWND status_hwnd = nullptr;
  HWND request_hwnd = nullptr;
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  bool has_output = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_ai_panel.requesting = false;
    m_ai_panel.completed = true;
    m_ai_panel.has_error = has_error;
    m_ai_panel.status_text =
        has_error ? L"生成失败，可重试。" : L"生成完成，点击确认回写。";
    status_hwnd = m_ai_panel.status_hwnd;
    request_hwnd = m_ai_panel.request_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    has_output = !m_ai_panel.output_text.empty();
  }

  if (status_hwnd && IsWindow(status_hwnd)) {
    SetWindowTextW(status_hwnd, has_error ? L"生成失败，可重试。"
                                          : L"生成完成，点击确认回写。");
  }
  if (confirm_hwnd && IsWindow(confirm_hwnd)) {
    EnableWindow(confirm_hwnd, !has_error && has_output);
  }
  if (request_hwnd && IsWindow(request_hwnd)) {
    EnableWindow(request_hwnd, TRUE);
  }
  if (cancel_hwnd && IsWindow(cancel_hwnd)) {
    SetWindowTextW(cancel_hwnd, has_error ? L"关闭" : L"取消");
  }
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_ResizeAIPanelWebView() {
#if WEASEL_HAS_WEBVIEW2
  HWND panel_hwnd = nullptr;
  HWND output_hwnd = nullptr;
  ICoreWebView2Controller* controller = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    output_hwnd = m_ai_panel.output_hwnd;
    controller =
        static_cast<ICoreWebView2Controller*>(m_ai_panel.webview_controller);
    if (controller) {
      controller->AddRef();
    }
  }

  if (!panel_hwnd || !output_hwnd || !controller || !IsWindow(panel_hwnd) ||
      !IsWindow(output_hwnd)) {
    if (controller) {
      controller->Release();
    }
    return;
  }

  RECT output_rect = {0};
  GetWindowRect(output_hwnd, &output_rect);
  MapWindowPoints(nullptr, panel_hwnd, reinterpret_cast<LPPOINT>(&output_rect),
                  2);
  controller->put_Bounds(output_rect);
  controller->Release();
#endif
}

void RimeWithWeaselHandler::_RequestAIPanelGeneration() {
  HWND status_hwnd = nullptr;
  HWND output_hwnd = nullptr;
  HWND request_hwnd = nullptr;
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  HWND panel_hwnd = nullptr;
  std::wstring context_text;
  std::wstring status_text;
  uint64_t request_id = 0;
  bool should_start = false;

  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (!m_ai_panel.panel_hwnd || !IsWindow(m_ai_panel.panel_hwnd)) {
      return;
    }
    if (m_ai_panel.requesting) {
      return;
    }

    status_hwnd = m_ai_panel.status_hwnd;
    output_hwnd = m_ai_panel.output_hwnd;
    request_hwnd = m_ai_panel.request_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;
    context_text = m_ai_panel.context_text;
  }

  if (status_text.empty()) {
    if (context_text.empty()) {
      status_text = L"上下文为空，无法发起请求。";
    } else if (m_ai_config.endpoint.empty()) {
      status_text = L"AI 接口未配置，请先设置 endpoint。";
    }
  }

  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (!m_ai_panel.panel_hwnd || !IsWindow(m_ai_panel.panel_hwnd) ||
        m_ai_panel.requesting) {
      return;
    }
    status_hwnd = m_ai_panel.status_hwnd;
    output_hwnd = m_ai_panel.output_hwnd;
    request_hwnd = m_ai_panel.request_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;

    if (!status_text.empty()) {
      m_ai_panel.status_text = status_text;
    } else {
      m_ai_panel.requesting = true;
      m_ai_panel.completed = false;
      m_ai_panel.has_error = false;
      m_ai_panel.output_text.clear();
      m_ai_panel.status_text = L"正在生成...";
      request_id = ++m_ai_request_seq;
      m_ai_panel.request_id = request_id;
      should_start = true;
    }
  }

  if (!should_start) {
    if (status_hwnd && IsWindow(status_hwnd)) {
      std::wstring status_text;
      {
        std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
        status_text = m_ai_panel.status_text;
      }
      SetWindowTextW(status_hwnd, status_text.c_str());
    }
    if (panel_hwnd && IsWindow(panel_hwnd)) {
      PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
    }
    return;
  }

  if (status_hwnd && IsWindow(status_hwnd)) {
    SetWindowTextW(status_hwnd, L"正在生成...");
  }
  if (output_hwnd && IsWindow(output_hwnd)) {
    SetWindowTextW(output_hwnd, L"");
  }
  if (request_hwnd && IsWindow(request_hwnd)) {
    EnableWindow(request_hwnd, FALSE);
  }
  if (confirm_hwnd && IsWindow(confirm_hwnd)) {
    EnableWindow(confirm_hwnd, FALSE);
  }
  if (cancel_hwnd && IsWindow(cancel_hwnd)) {
    EnableWindow(cancel_hwnd, TRUE);
    SetWindowTextW(cancel_hwnd, L"取消");
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
  _StartAIAssistantStreamRequest(request_id, context_text);
}

void RimeWithWeaselHandler::_ConfirmAIPanelOutput() {
  std::wstring output_text;
  HWND target_hwnd = nullptr;
  WeaselSessionId ipc_id = 0;
  bool completed = false;
  bool has_error = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    output_text = m_ai_panel.output_text;
    target_hwnd = m_ai_panel.target_hwnd;
    ipc_id = m_ai_panel.ipc_id;
    completed = m_ai_panel.completed;
    has_error = m_ai_panel.has_error;
  }

  if (!completed || has_error || output_text.empty()) {
    return;
  }

  if (!_SendTextToTargetWindow(target_hwnd, output_text)) {
    _CompleteAIPanel(true, L"无法回写到目标输入框。");
    return;
  }

  _AppendCommittedText(ipc_id, output_text);
  _CloseAIPanel();
  if (target_hwnd && IsWindow(target_hwnd)) {
    SetForegroundWindow(target_hwnd);
  }
}

void RimeWithWeaselHandler::_CancelAIPanelOutput() {
  HWND target_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    target_hwnd = m_ai_panel.target_hwnd;
  }
  _CloseAIPanel();
  if (target_hwnd && IsWindow(target_hwnd)) {
    SetForegroundWindow(target_hwnd);
  }
}

void RimeWithWeaselHandler::_ExecuteAIPanelSystemCommand(
    const std::wstring& command_id) {
  if (command_id.empty() || !m_system_command_callback) {
    return;
  }
  std::string cmd_id_utf8 = wtou8(command_id);
  if (!IsAllowedSystemCommandId(cmd_id_utf8)) {
    return;
  }
  SystemCommandLaunchRequest request;
  request.command_id = std::move(cmd_id_utf8);
  m_system_command_callback(request);

  // 执行完系统命令后关闭面板
  HWND target_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    target_hwnd = m_ai_panel.target_hwnd;
  }
  _CloseAIPanel();
  if (target_hwnd && IsWindow(target_hwnd)) {
    SetForegroundWindow(target_hwnd);
  }
}

void RimeWithWeaselHandler::_RefreshAIPanelInstitutionOptions() {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    if (!panel_hwnd || !IsWindow(panel_hwnd)) {
      return;
    }
    m_ai_panel.institutions_loading = true;
  }
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);

  const AIAssistantConfig config = m_ai_config;
  std::string token;
  std::string tenant_id;
  std::string refresh_token;
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    token = m_ai_login_token;
    tenant_id = m_ai_login_tenant_id;
    refresh_token = m_ai_login_refresh_token;
  }
  if (tenant_id.empty() || token.empty() || refresh_token.empty()) {
    std::string file_token;
    std::string file_tenant_id;
    std::string file_refresh_token;
    LoadAIAssistantLoginIdentity(config, &file_token, &file_tenant_id,
                                 &file_refresh_token);
    if (!file_token.empty()) {
      token = file_token;
    }
    if (!file_tenant_id.empty()) {
      tenant_id = file_tenant_id;
    }
    if (!file_refresh_token.empty()) {
      refresh_token = file_refresh_token;
    }
    if (!file_token.empty()) {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      m_ai_login_token = file_token;
      if (!file_tenant_id.empty()) {
        m_ai_login_tenant_id = file_tenant_id;
      }
      if (!file_refresh_token.empty()) {
        m_ai_login_refresh_token = file_refresh_token;
      }
    }
  }

  if (tenant_id.empty() || (token.empty() && refresh_token.empty())) {
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      if (m_ai_panel.panel_hwnd == panel_hwnd) {
        m_ai_panel.institutions_loading = false;
      }
    }
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
    LOG(INFO) << "AI panel institution list skipped: token or tenantId empty.";
    return;
  }

  std::thread([this, panel_hwnd, config, token, tenant_id, refresh_token]() mutable {
    std::vector<AIPanelInstitutionOption> options;
    std::string error_message;
    int http_status_code = 0;
    bool ok = false;

    if (!token.empty()) {
      ok = FetchAIAssistantInstitutionOptions(config, token, tenant_id, &options,
                                              &error_message, &http_status_code);
    }
    if (!ok && !refresh_token.empty() &&
        (token.empty() || http_status_code == 401)) {
      std::string refreshed_token;
      std::string refreshed_refresh_token;
      std::string refresh_endpoint;
      std::string refresh_error;
      const bool refreshed = RefreshAIAssistantAccessToken(
          config, refresh_token, tenant_id, &refreshed_token,
          &refreshed_refresh_token, &refresh_endpoint, &refresh_error);
      if (refreshed && !refreshed_token.empty()) {
        token = refreshed_token;
        if (!refreshed_refresh_token.empty()) {
          refresh_token = refreshed_refresh_token;
        }
        std::string client_id_for_state;
        {
          std::lock_guard<std::mutex> login_lock(m_ai_login_mutex);
          client_id_for_state = m_ai_login_client_id;
        }
        if (SaveAIAssistantLoginState(config, token, tenant_id, refresh_token,
                                      client_id_for_state, std::string(),
                                      std::string())) {
          std::lock_guard<std::mutex> lock(m_ai_login_mutex);
          m_ai_login_token = token;
          m_ai_login_tenant_id = tenant_id;
          m_ai_login_refresh_token = refresh_token;
        }
        error_message.clear();
        http_status_code = 0;
        ok = FetchAIAssistantInstitutionOptions(
            config, token, tenant_id, &options, &error_message,
            &http_status_code);
      } else {
        error_message = refresh_error.empty() ? error_message : refresh_error;
      }
    }
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      if (m_ai_panel.panel_hwnd != panel_hwnd) {
        return;
      }
      m_ai_panel.institutions_loading = false;
      if (ok) {
        std::wstring selected_id = m_ai_panel.selected_institution_id;
        bool found = false;
        for (const auto& option : options) {
          if (!selected_id.empty() && option.id == selected_id) {
            found = true;
            break;
          }
        }
        m_ai_panel.institution_options = options;
        if (found) {
          m_ai_panel.selected_institution_id = selected_id;
        } else {
          m_ai_panel.selected_institution_id.clear();
        }
      } else {
        m_ai_panel.institution_options.clear();
        m_ai_panel.selected_institution_id.clear();
      }
    }

    if (!ok) {
      LOG(WARNING) << "AI panel institution list fetch failed: "
                   << error_message;
    } else {
      LOG(INFO) << "AI panel institution list loaded, count="
                << options.size();
    }
    if (panel_hwnd && IsWindow(panel_hwnd)) {
      PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
    }
  }).detach();
}

void RimeWithWeaselHandler::_SetAIPanelContextText(
    const std::wstring& context_text) {
  const std::wstring normalized_context =
      NormalizeReferencedContextText(context_text);
  std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
  if (!m_ai_panel.panel_hwnd || !IsWindow(m_ai_panel.panel_hwnd)) {
    return;
  }
  m_ai_panel.context_text = normalized_context;
}

void RimeWithWeaselHandler::_HandleAIPanelWritebackRequest(
    const std::wstring& text) {
  if (text.empty()) {
    _SetAIPanelStatus(L"回写内容为空。");
    return;
  }
  HWND target_hwnd = nullptr;
  WeaselSessionId ipc_id = 0;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    target_hwnd = m_ai_panel.target_hwnd;
    ipc_id = m_ai_panel.ipc_id;
    m_ai_panel.output_text = text;
  }
  if (!_SendTextToTargetWindow(target_hwnd, text)) {
    _SetAIPanelStatus(L"无法回写到目标输入框。");
    return;
  }
  _AppendCommittedText(ipc_id, text);
  _CloseAIPanel();
  if (target_hwnd && IsWindow(target_hwnd)) {
    SetForegroundWindow(target_hwnd);
  }
}

void RimeWithWeaselHandler::_SetAIPanelInstitutionSelection(
    const std::wstring& institution_id) {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    if (!panel_hwnd || !IsWindow(panel_hwnd)) {
      return;
    }
    if (institution_id.empty()) {
      m_ai_panel.selected_institution_id.clear();
    } else {
      m_ai_panel.selected_institution_id.clear();
      for (const auto& option : m_ai_panel.institution_options) {
        if (option.id == institution_id) {
          m_ai_panel.selected_institution_id = institution_id;
          break;
        }
      }
    }
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }
}

void RimeWithWeaselHandler::_StartAIAssistantStreamRequest(
    uint64_t request_id,
    const std::wstring& context_text) {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (!panel_hwnd || !IsWindow(panel_hwnd)) {
    return;
  }

  const AIAssistantConfig request_config = m_ai_config;
  std::thread([request_config, request_id, context_text, panel_hwnd]() {
    const auto post_chunk = [panel_hwnd, request_id](const std::wstring& chunk) {
      if (chunk.empty()) {
        return;
      }
      AIPanelTextMessage* message = new AIPanelTextMessage();
      message->request_id = request_id;
      message->text = chunk;
      if (!PostMessageW(panel_hwnd, WM_AI_STREAM_APPEND, 0,
                        reinterpret_cast<LPARAM>(message))) {
        delete message;
      }
    };

    std::string error_message;
    const bool ok = InvokeAIAssistantStream(request_config, context_text,
                                            post_chunk, &error_message);
    AIPanelDoneMessage* done = new AIPanelDoneMessage();
    done->request_id = request_id;
    done->has_error = !ok;
    if (!ok && !error_message.empty()) {
      done->text = u8tow(error_message);
    }
    const UINT done_message = ok ? WM_AI_STREAM_DONE : WM_AI_STREAM_ERROR;
    if (!PostMessageW(panel_hwnd, done_message, 0,
                      reinterpret_cast<LPARAM>(done))) {
      delete done;
    }
  }).detach();
}

bool RimeWithWeaselHandler::_SendTextToTargetWindow(
    HWND target_hwnd,
    const std::wstring& text) {
  if (text.empty()) {
    return true;
  }

  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
  }

  HWND resolved_target = target_hwnd;
  if (!resolved_target || !IsWindow(resolved_target)) {
    resolved_target = GetForegroundWindow();
  }
  if (!resolved_target || !IsWindow(resolved_target) ||
      resolved_target == panel_hwnd) {
    return false;
  }

  if (GetForegroundWindow() != resolved_target) {
    ShowWindow(resolved_target, SW_SHOW);
    SetForegroundWindow(resolved_target);
    Sleep(10);
  }

  for (wchar_t ch : text) {
    INPUT events[2] = {0};
    events[0].type = INPUT_KEYBOARD;
    events[0].ki.wScan = ch;
    events[0].ki.dwFlags = KEYEVENTF_UNICODE;
    events[1] = events[0];
    events[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
    if (SendInput(2, events, sizeof(INPUT)) != 2) {
      return false;
    }
  }
  return true;
}

LRESULT CALLBACK RimeWithWeaselHandler::AIAssistantPanelWndProc(
    HWND hwnd,
    UINT msg,
    WPARAM wParam,
    LPARAM lParam) {
  if (msg == WM_NCCREATE) {
    const CREATESTRUCTW* create_struct =
        reinterpret_cast<const CREATESTRUCTW*>(lParam);
    RimeWithWeaselHandler* self =
        reinterpret_cast<RimeWithWeaselHandler*>(create_struct->lpCreateParams);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    return TRUE;
  }

  RimeWithWeaselHandler* self = reinterpret_cast<RimeWithWeaselHandler*>(
      GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_SIZE:
      ApplyAIPanelRoundedRegion(hwnd);
      LayoutAIPanelControls(hwnd);
      if (self) {
        self->_ResizeAIPanelWebView();
      }
      return 0;
    case WM_COMMAND:
      if (!self) {
        break;
      }
      if (HIWORD(wParam) == BN_CLICKED &&
          LOWORD(wParam) == kAIPanelControlRequest) {
        self->_RequestAIPanelGeneration();
        return 0;
      }
      if (HIWORD(wParam) == BN_CLICKED &&
          LOWORD(wParam) == kAIPanelControlConfirm) {
        self->_ConfirmAIPanelOutput();
        return 0;
      }
      if (HIWORD(wParam) == BN_CLICKED &&
          LOWORD(wParam) == kAIPanelControlCancel) {
        self->_CancelAIPanelOutput();
        return 0;
      }
      break;
    case WM_KEYDOWN:
      if (self && wParam == VK_ESCAPE) {
        self->_CancelAIPanelOutput();
        return 0;
      }
      break;
    case WM_CLOSE:
      if (self) {
        self->_CancelAIPanelOutput();
        return 0;
      }
      break;
    case WM_AI_PANEL_DESTROY:
      DestroyWindow(hwnd);
      return 0;
    case WM_AI_PANEL_DRAG:
      ReleaseCapture();
      SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0);
      return 0;
    case WM_NCHITTEST: {
      LRESULT hit = DefWindowProcW(hwnd, msg, wParam, lParam);
      if (hit == HTCLIENT) {
        RECT rc;
        GetClientRect(hwnd, &rc);
        POINT pt = {LOWORD(lParam), HIWORD(lParam)};
        const int grip_size = 20;
        if (pt.x >= rc.right - grip_size && pt.y >= rc.bottom - grip_size) {
          return HTBOTTOMRIGHT;
        }
      }
      return hit;
    }
    case WM_DESTROY:
      PostQuitMessage(0);
      return 0;
    case WM_AI_WEBVIEW_INIT: {
#if WEASEL_HAS_WEBVIEW2
      if (!self) {
        return 0;
      }
      std::wstring panel_url = u8tow(self->m_ai_config.panel_url);
      const size_t hash_route = panel_url.find(L"#/");
      if (hash_route != std::wstring::npos) {
        const std::wstring hash_part = panel_url.substr(hash_route);
        if (hash_part.find(L"#/rime-with-weasel") == 0 ||
            hash_part.find(L"#/rime-input") == 0 ||
            hash_part.find(L"#/rime-generating") == 0 ||
            hash_part.find(L"#/rime-select") == 0) {
          panel_url = panel_url.substr(0, hash_route) + L"#/rime-with-weasel";
        }
      }
      std::string token;
      std::string tenant_id;
      std::string refresh_token;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_login_mutex);
        token = self->m_ai_login_token;
        tenant_id = self->m_ai_login_tenant_id;
        refresh_token = self->m_ai_login_refresh_token;
      }
      if (tenant_id.empty() || (token.empty() && refresh_token.empty())) {
        std::string file_token;
        std::string file_tenant_id;
        std::string file_refresh_token;
        LoadAIAssistantLoginIdentity(self->m_ai_config, &file_token,
                                     &file_tenant_id, &file_refresh_token);
        if (!file_token.empty()) {
          token = file_token;
        }
        if (!file_tenant_id.empty()) {
          tenant_id = file_tenant_id;
        }
        if (!file_refresh_token.empty()) {
          refresh_token = file_refresh_token;
        }
      }
      panel_url =
          AppendAIPanelAuthToUrl(panel_url, token, tenant_id, refresh_token);

      bool already_ready = false;
      ICoreWebView2* existing_webview = nullptr;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        already_ready = self->m_ai_panel.webview != nullptr;
        if (already_ready) {
          existing_webview = static_cast<ICoreWebView2*>(self->m_ai_panel.webview);
          if (existing_webview) {
            existing_webview->AddRef();
          }
        }
      }
      if (already_ready) {
        if (existing_webview && !panel_url.empty()) {
          existing_webview->Navigate(panel_url.c_str());
        }
        if (existing_webview) {
          existing_webview->Release();
        }
        self->_ResizeAIPanelWebView();
        PostMessageW(hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
        return 0;
      }

      const auto factory = ResolveWebView2Factory();
      if (!factory) {
        return 0;
      }

      if (panel_url.empty()) {
        LOG(WARNING) << "AI panel url is empty.";
        return 0;
      }
      const std::wstring allowed_origin =
          ResolveAIPanelAllowedOrigin(self->m_ai_config);
      const std::wstring user_data_dir =
          (WeaselUserDataPath() / L"webview2_ai_panel").wstring();
      factory(
          nullptr, user_data_dir.c_str(), nullptr,
          Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
              [self, hwnd, panel_url, allowed_origin](HRESULT result,
                                 ICoreWebView2Environment* environment)
                  -> HRESULT {
                if (!self || FAILED(result) || !environment) {
                  return S_OK;
                }
                return environment->CreateCoreWebView2Controller(
                    hwnd,
                    Callback<
                        ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [self, hwnd, panel_url, allowed_origin](HRESULT result,
                                           ICoreWebView2Controller* controller)
                            -> HRESULT {
                          if (!self || FAILED(result) || !controller) {
                            return S_OK;
                          }

                          ComPtr<ICoreWebView2> webview;
                          controller->get_CoreWebView2(&webview);
                          if (!webview) {
                            return S_OK;
                          }

                          controller->add_AcceleratorKeyPressed(
                              Callback<ICoreWebView2AcceleratorKeyPressedEventHandler>(
                                  [self](ICoreWebView2Controller* /*sender*/,
                                         ICoreWebView2AcceleratorKeyPressedEventArgs*
                                             args) -> HRESULT {
                                    if (!self || !args) {
                                      return S_OK;
                                    }
                                    COREWEBVIEW2_KEY_EVENT_KIND key_event_kind =
                                        COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN;
                                    UINT virtual_key = 0;
                                    if (FAILED(args->get_KeyEventKind(
                                            &key_event_kind)) ||
                                        FAILED(args->get_VirtualKey(
                                            &virtual_key))) {
                                      return S_OK;
                                    }
                                    const bool is_key_down =
                                        key_event_kind ==
                                            COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN ||
                                        key_event_kind ==
                                            COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN;
                                    if (is_key_down &&
                                        virtual_key == static_cast<UINT>(VK_ESCAPE)) {
                                      self->_CancelAIPanelOutput();
                                      args->put_Handled(TRUE);
                                    }
                                    return S_OK;
                                  })
                                  .Get(),
                              nullptr);

                          webview->add_WebMessageReceived(
                              Callback<ICoreWebView2WebMessageReceivedEventHandler>(
                                  [self, hwnd, allowed_origin](
                                      ICoreWebView2* /*sender*/,
                                      ICoreWebView2WebMessageReceivedEventArgs*
                                          args) -> HRESULT {
                                    if (!self || !args) {
                                      return S_OK;
                                    }
                                    LPWSTR source_raw = nullptr;
                                    std::wstring source_url;
                                    if (SUCCEEDED(args->get_Source(&source_raw)) &&
                                        source_raw) {
                                      source_url.assign(source_raw);
                                      CoTaskMemFree(source_raw);
                                    }
                                    if (!IsAIPanelMessageOriginAllowed(
                                            allowed_origin, source_url)) {
                                      LOG(WARNING)
                                          << "AI panel message blocked by origin: "
                                          << wtou8(source_url);
                                      return S_OK;
                                    }

                                    std::wstring message;
                                    if (!ExtractWebViewMessageText(args,
                                                                   &message)) {
                                      return S_OK;
                                    }
                                    AIPanelUiCommand command;
                                    if (!ParseAIPanelUiCommand(message,
                                                               &command)) {
                                      return S_OK;
                                    }
                                    if (command.type == "ui.ready") {
                                      PostMessageW(hwnd, WM_AI_WEBVIEW_SYNC, 0,
                                                   0);
                                    } else if (command.type ==
                                               "ui.context.changed") {
                                      self->_SetAIPanelContextText(command.text);
                                      PostMessageW(hwnd, WM_AI_WEBVIEW_SYNC, 0,
                                                   0);
                                    } else if (command.type ==
                                               "ui.select.institution") {
                                      self->_SetAIPanelInstitutionSelection(
                                          command.institution_id);
                                    } else if (command.type ==
                                               "ui.writeback.confirm") {
                                      self->_HandleAIPanelWritebackRequest(
                                          command.text);
                                    } else if (command.type == "ui.cancel") {
                                      self->_CancelAIPanelOutput();
                                    } else if (command.type ==
                                               "ui.drag.start") {
                                      PostMessageW(hwnd, WM_AI_PANEL_DRAG, 0, 0);
                                    } else if (command.type ==
                                               "ui.panel.resize") {
                                      self->_ApplyAIPanelSizeAndReposition(
                                          command.panel_width,
                                          command.panel_height,
                                          false);
                                    } else if (command.type ==
                                               "ui.auth.refresh_request") {
                                      self->_ForceAIAssistantRelogin();
                                      self->_RefreshAIPanelInstitutionOptions();
                                      PostMessageW(hwnd, WM_AI_WEBVIEW_SYNC, 0,
                                                   0);
                                    } else if (command.type == "ui.request") {
                                      if (!command.text.empty()) {
                                        self->_SetAIPanelContextText(
                                            command.text);
                                      }
                                      self->_SetAIPanelStatus(
                                          L"请求已交给前端面板处理。");
                                    } else if (command.type == "ui.confirm") {
                                      self->_ConfirmAIPanelOutput();
                                    } else if (command.type ==
                                               "ui.system_command") {
                                      self->_ExecuteAIPanelSystemCommand(
                                          command.text);
                                    }
                                    return S_OK;
                                  })
                                  .Get(),
                              nullptr);
                          webview->Navigate(panel_url.c_str());
                          controller->AddRef();
                          webview->AddRef();

                          HWND output_hwnd = nullptr;
                          {
                            std::lock_guard<std::mutex> lock(
                                self->m_ai_panel_mutex);
                            if (self->m_ai_panel.webview_controller) {
                              static_cast<ICoreWebView2Controller*>(
                                  self->m_ai_panel.webview_controller)
                                  ->Release();
                            }
                            if (self->m_ai_panel.webview) {
                              static_cast<ICoreWebView2*>(self->m_ai_panel.webview)
                                  ->Release();
                            }
                            self->m_ai_panel.webview_hwnd = hwnd;
                            self->m_ai_panel.webview_controller = controller;
                            self->m_ai_panel.webview = webview.Get();
                            self->m_ai_panel.webview_ready = true;
                            output_hwnd = self->m_ai_panel.output_hwnd;
                          }

                          if (output_hwnd && IsWindow(output_hwnd)) {
                            ShowWindow(output_hwnd, SW_HIDE);
                          }
                          self->_ResizeAIPanelWebView();
                          PostMessageW(hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
                          return S_OK;
                        })
                        .Get());
              })
              .Get());
#endif
      return 0;
    }
    case WM_AI_WEBVIEW_SYNC: {
#if WEASEL_HAS_WEBVIEW2
      if (!self) {
        return 0;
      }

      ICoreWebView2* webview = nullptr;
      std::wstring context_text;
      std::wstring status_text;
      std::wstring output_text;
      std::vector<AIPanelInstitutionOption> institution_options;
      std::wstring selected_institution_id;
      bool requesting = false;
      bool institutions_loading = false;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        webview = static_cast<ICoreWebView2*>(self->m_ai_panel.webview);
        if (webview) {
          webview->AddRef();
        }
        context_text = self->m_ai_panel.context_text;
        status_text = self->m_ai_panel.status_text;
        output_text = self->m_ai_panel.output_text;
        institution_options = self->m_ai_panel.institution_options;
        selected_institution_id = self->m_ai_panel.selected_institution_id;
        requesting = self->m_ai_panel.requesting;
        institutions_loading = self->m_ai_panel.institutions_loading;
      }

      if (!webview) {
        return 0;
      }
      std::string token;
      std::string tenant_id;
      std::string refresh_token;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_login_mutex);
        token = self->m_ai_login_token;
        tenant_id = self->m_ai_login_tenant_id;
        refresh_token = self->m_ai_login_refresh_token;
      }
      if (tenant_id.empty() || (token.empty() && refresh_token.empty())) {
        std::string file_token;
        std::string file_tenant_id;
        std::string file_refresh_token;
        LoadAIAssistantLoginIdentity(self->m_ai_config, &file_token,
                                     &file_tenant_id, &file_refresh_token);
        if (!file_token.empty()) {
          token = file_token;
        }
        if (!file_tenant_id.empty()) {
          tenant_id = file_tenant_id;
        }
        if (!file_refresh_token.empty()) {
          refresh_token = file_refresh_token;
        }
      }
      const std::string sync_json = BuildAIPanelHostSyncMessage(
          context_text, status_text, output_text, requesting,
          institutions_loading, institution_options, selected_institution_id,
          token, tenant_id, refresh_token);
      const std::wstring sync_text = u8tow(sync_json);
      webview->PostWebMessageAsString(sync_text.c_str());
      webview->Release();
#endif
      return 0;
    }
    case WM_AI_STREAM_APPEND: {
      AIPanelTextMessage* message =
          reinterpret_cast<AIPanelTextMessage*>(lParam);
      if (!message || !self) {
        delete message;
        return 0;
      }
      bool is_latest_request = false;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        is_latest_request = self->m_ai_panel.panel_hwnd == hwnd &&
                            self->m_ai_panel.request_id == message->request_id;
      }
      if (is_latest_request) {
        self->_AppendAIPanelOutput(message->text);
      }
      delete message;
      return 0;
    }
    case WM_AI_STREAM_DONE:
    case WM_AI_STREAM_ERROR: {
      AIPanelDoneMessage* message =
          reinterpret_cast<AIPanelDoneMessage*>(lParam);
      if (!message || !self) {
        delete message;
        return 0;
      }
      bool is_latest_request = false;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        is_latest_request = self->m_ai_panel.panel_hwnd == hwnd &&
                            self->m_ai_panel.request_id == message->request_id;
      }
      if (is_latest_request) {
        self->_CompleteAIPanel(msg == WM_AI_STREAM_ERROR, message->text);
      }
      delete message;
      return 0;
    }
    default:
      break;
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

std::wstring RimeWithWeaselHandler::_CollectAIAssistantContext(
    WeaselSessionId ipc_id,
    const std::wstring& current_text) {
  return m_input_content_store.CollectContext(
      _GetInputContentContextKey(ipc_id), current_text,
      static_cast<size_t>(max(1, m_ai_config.max_history_chars)));
}

void RimeWithWeaselHandler::_AppendCommittedText(WeaselSessionId ipc_id,
                                                 const std::wstring& text) {
  if (text.empty()) {
    return;
  }
  const std::string context_key = _GetInputContentContextKey(ipc_id);
  m_input_content_store.AppendCommit(context_key, text);
  AppendInputContentInfoLogLine(
      "InputContent commit: context_key=" + context_key +
      ", chars=" + std::to_string(text.size()) + ", preview='" +
      BuildInputContentPreviewForLog(text) + "'");
}

std::string RimeWithWeaselHandler::_GetInputContentContextKey(
    WeaselSessionId ipc_id) {
  const std::string context_key = _GetContextCacheKey(ipc_id);
  const bool has_context_key = !context_key.empty() && context_key != "__global__";
  if (has_context_key) {
    if (!m_input_active_context_key.empty() &&
        m_input_active_context_key != context_key) {
      AppendInputContentInfoLogLine(
          "InputContent context_reset: reason=app_switch, from=" +
          m_input_active_context_key + ", to=" + context_key);
      m_input_content_store.Clear();
    }
    m_input_active_context_key = context_key;
    return m_input_active_context_key;
  }
  if (!m_input_active_context_key.empty()) {
    return m_input_active_context_key;
  }
  return "__global__";
}

std::wstring RimeWithWeaselHandler::_TakePendingCommitText(
    RimeSessionId session_id) {
  RIME_STRUCT(RimeCommit, commit);
  if (!rime_api->get_commit(session_id, &commit)) {
    return std::wstring();
  }
  std::wstring commit_text = commit.text ? u8tow(commit.text) : std::wstring();
  rime_api->free_commit(&commit);
  return commit_text;
}

bool RimeWithWeaselHandler::_TryHandleSystemCommandCommit(
    std::wstring* commit_text,
    WeaselSessionId ipc_id) {
  if (!commit_text || commit_text->empty() || !m_system_command_callback) {
    return false;
  }

  std::string command_id;
  if (!TryParseSystemCommandMarker(*commit_text, &command_id)) {
    return false;
  }

  SystemCommandLaunchRequest request;
  request.command_id = std::move(command_id);
  request.preferred_output_dir = _ReadSystemCommandOutputDir(ipc_id);
  m_system_command_callback(request);
  commit_text->clear();
  return true;
}

std::filesystem::path RimeWithWeaselHandler::_ReadSystemCommandOutputDir(
    WeaselSessionId ipc_id) const {
  const auto session_it = m_session_status_map.find(ipc_id);
  if (session_it == m_session_status_map.end()) {
    return std::filesystem::path();
  }

  std::string schema_id;
  RIME_STRUCT(RimeStatus, status);
  if (rime_api->get_status(session_it->second.session_id, &status)) {
    schema_id = status.schema_id ? status.schema_id : std::string();
    rime_api->free_status(&status);
  }
  if (schema_id.empty()) {
    return std::filesystem::path();
  }

  RimeConfig config = {NULL};
  if (!rime_api->schema_open(schema_id.c_str(), &config)) {
    return std::filesystem::path();
  }

  std::string output_dir_utf8;
  const bool has_output_dir =
      ReadConfigString(&config, "system_cmd/output_dir", &output_dir_utf8);
  rime_api->config_close(&config);
  if (!has_output_dir || output_dir_utf8.empty()) {
    return std::filesystem::path();
  }

  std::filesystem::path path(
      ExpandEnvironmentVariables(u8tow(output_dir_utf8)));
  if (path.empty()) {
    return std::filesystem::path();
  }
  if (!path.is_absolute()) {
    path = WeaselUserDataPath() / path;
  }
  return path;
}

std::string RimeWithWeaselHandler::_GetContextCacheKey(
    WeaselSessionId ipc_id) const {
  const auto it = m_session_status_map.find(ipc_id);
  if (it == m_session_status_map.end() || it->second.client_app.empty()) {
    return "__global__";
  }
  return it->second.client_app;
}

bool RimeWithWeaselHandler::_IsDeployerRunning() {
  HANDLE hMutex = CreateMutex(NULL, TRUE, L"WeaselDeployerMutex");
  bool deployer_detected = hMutex && GetLastError() == ERROR_ALREADY_EXISTS;
  if (hMutex) {
    CloseHandle(hMutex);
  }
  return deployer_detected;
}

void RimeWithWeaselHandler::_UpdateUI(WeaselSessionId ipc_id) {
  // if m_ui nullptr, _UpdateUI meaningless
  if (!m_ui)
    return;

  Status& weasel_status = m_ui->status();
  Context weasel_context;

  RimeSessionId session_id = to_session_id(ipc_id);

  if (ipc_id == 0)
    weasel_status.disabled = m_disabled;

  _GetStatus(weasel_status, ipc_id, weasel_context);

  // Get actual context from Rime to check if there's content (e.g., prediction)
  _GetContext(weasel_context, session_id);

  SessionStatus& session_status = get_session_status(ipc_id);
  if (rime_api->get_option(session_id, "inline_preedit"))
    session_status.style.client_caps |= INLINE_PREEDIT_CAPABLE;
  else
    session_status.style.client_caps &= ~INLINE_PREEDIT_CAPABLE;

  if (!_ShowMessage(weasel_context, weasel_status)) {
    m_ui->Hide();
    m_ui->Update(weasel_context, weasel_status);
  }

  _RefreshTrayIcon(session_id, _UpdateUICallback);

  {
    std::lock_guard<std::mutex> lock(m_notifier_mutex);
    m_message_type.clear();
    m_message_value.clear();
    m_message_label.clear();
    m_option_name.clear();
  }
}

void RimeWithWeaselHandler::_LoadSchemaSpecificSettings(
    WeaselSessionId ipc_id,
    const std::string& schema_id) {
  if (!m_ui)
    return;
  RimeConfig config;
  if (!rime_api->schema_open(schema_id.c_str(), &config))
    return;
  _UpdateShowNotifications(&config);
  m_ui->style() = m_base_style;
  _UpdateUIStyle(&config, m_ui, false);
  SessionStatus& session_status = get_session_status(ipc_id);
  session_status.style = m_ui->style();
  UIStyle& style = session_status.style;
  // load schema color style config
  const int BUF_SIZE = 255;
  char buffer[BUF_SIZE + 1] = {0};
  const auto update_color_scheme = [&]() {
    std::string color_name(buffer);
    RimeConfigIterator preset = {0};
    if (rime_api->config_begin_map(
            &preset, &config, ("preset_color_schemes/" + color_name).c_str())) {
      _UpdateUIStyleColor(&config, style, color_name);
      rime_api->config_end(&preset);
    } else {
      RimeConfig weaselconfig;
      if (rime_api->config_open("weasel", &weaselconfig)) {
        _UpdateUIStyleColor(&weaselconfig, style, color_name);
        rime_api->config_close(&weaselconfig);
      }
    }
  };
  const char* key =
      m_current_dark_mode ? "style/color_scheme_dark" : "style/color_scheme";
  if (rime_api->config_get_string(&config, key, buffer, BUF_SIZE))
    update_color_scheme();
  // load schema icon start
  {
    const auto load_icon = [](RimeConfig& config, const char* key1,
                              const char* key2) {
      const auto user_dir = WeaselUserDataPath();
      const auto shared_dir = WeaselSharedDataPath();
      const int BUF_SIZE = 255;
      char buffer[BUF_SIZE + 1] = {0};
      if (rime_api->config_get_string(&config, key1, buffer, BUF_SIZE) ||
          (key2 != NULL &&
           rime_api->config_get_string(&config, key2, buffer, BUF_SIZE))) {
        auto resource = u8tow(buffer);
        if (fs::is_regular_file(user_dir / resource))
          return (user_dir / resource).wstring();
        else if (fs::is_regular_file(shared_dir / resource))
          return (shared_dir / resource).wstring();
      }
      return std::wstring();
    };
    style.current_zhung_icon =
        load_icon(config, "schema/icon", "schema/zhung_icon");
    style.current_ascii_icon = load_icon(config, "schema/ascii_icon", NULL);
    style.current_full_icon = load_icon(config, "schema/full_icon", NULL);
    style.current_half_icon = load_icon(config, "schema/half_icon", NULL);
  }
  // load schema icon end
  rime_api->config_close(&config);
}

void RimeWithWeaselHandler::_LoadAppInlinePreeditSet(WeaselSessionId ipc_id,
                                                     bool ignore_app_name) {
  SessionStatus& session_status = get_session_status(ipc_id);
  RimeSessionId session_id = session_status.session_id;
  static char _app_name[50];
  rime_api->get_property(session_id, "client_app", _app_name,
                         sizeof(_app_name) - 1);
  std::string app_name(_app_name);
  if (!ignore_app_name && m_last_app_name == app_name)
    return;
  m_last_app_name = app_name;
  bool inline_preedit = session_status.style.inline_preedit;
  bool found = false;
  if (!app_name.empty()) {
    auto it = m_app_options.find(app_name);
    if (it != m_app_options.end()) {
      AppOptions& options(m_app_options[it->first]);
      for (const auto& pair : options) {
        if (pair.first == "inline_preedit") {
          rime_api->set_option(session_id, pair.first.c_str(),
                               Bool(pair.second));
          session_status.style.inline_preedit = Bool(pair.second);
          found = true;
          break;
        }
      }
    }
  }
  if (!found) {
    session_status.style.inline_preedit = m_base_style.inline_preedit;
    // load from schema.
    RIME_STRUCT(RimeStatus, status);
    if (rime_api->get_status(session_id, &status)) {
      std::string schema_id = status.schema_id;
      RimeConfig config;
      if (rime_api->schema_open(schema_id.c_str(), &config)) {
        Bool value = False;
        if (rime_api->config_get_bool(&config, "style/inline_preedit",
                                      &value)) {
          session_status.style.inline_preedit = value;
        }
        rime_api->config_close(&config);
      }
      rime_api->free_status(&status);
    }
  }
  if (session_status.style.inline_preedit != inline_preedit)
    _UpdateInlinePreeditStatus(ipc_id);
}

bool RimeWithWeaselHandler::_ShowMessage(Context& ctx, Status& status) {
  std::lock_guard<std::mutex> lock(m_notifier_mutex);
  if (m_message_type.empty() || m_message_value.empty())
    return m_ui->IsCountingDown();
  // show as auxiliary string
  std::wstring& tips(ctx.aux.str);
  bool show_icon = false;
  if (m_message_type == "deploy") {
    if (m_message_value == "start")
      if (GetThreadUILanguage() == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US))
        tips = L"Deploying RIME";
      else
        tips = L"正在部署 RIME";
    else if (m_message_value == "success")
      if (GetThreadUILanguage() == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US))
        tips = L"Deployed";
      else
        tips = L"部署完成";
    else if (m_message_value == "failure") {
      if (GetThreadUILanguage() ==
          MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL))
        tips = L"有錯誤，請查看日誌 %TEMP%\\rime.weasel\\rime.weasel.*.INFO";
      else if (GetThreadUILanguage() ==
               MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED))
        tips = L"有错误，请查看日志 %TEMP%\\rime.weasel\\rime.weasel.*.INFO";
      else
        tips =
            L"There is an error, please check the logs "
            L"%TEMP%\\rime.weasel\\rime.weasel.*.INFO";
    }
  } else if (m_message_type == "schema") {
    tips = /*L"【" + */ status.schema_name /* + L"】"*/;
  } else if (m_message_type == "option") {
    status.type = SCHEMA;
    if (m_message_value == "!ascii_mode") {
      show_icon = true;
    } else if (m_message_value == "ascii_mode") {
      show_icon = true;
    } else
      tips = u8tow(m_message_label);

    if (m_message_value == "full_shape" || m_message_value == "!full_shape")
      status.type = FULL_SHAPE;
  }
  auto counter = m_ui->IsCountingDown();
  if (!show_icon && counter)
    return counter;
  auto foption = m_show_notifications.find(m_option_name);
  auto falways = m_show_notifications.find("always");
  if ((!add_session && (foption != m_show_notifications.end() ||
                        falways != m_show_notifications.end())) ||
      m_message_type == "deploy") {
    m_ui->Update(ctx, status);
    if (m_show_notifications_time)
      m_ui->ShowWithTimeout(m_show_notifications_time);
    return true;
  } else {
    return m_ui->IsCountingDown();
  }
}
inline std::string _GetLabelText(const std::vector<Text>& labels,
                                 int id,
                                 const wchar_t* format) {
  wchar_t buffer[128];
  swprintf_s<128>(buffer, format, labels.at(id).str.c_str());
  return wtou8(std::wstring(buffer));
}

bool RimeWithWeaselHandler::_Respond(WeaselSessionId ipc_id,
                                     EatLine eat,
                                     const std::wstring* extra_commit) {
  std::wstring body;
  body.reserve(4096);
  std::vector<const char*> actions;
  actions.reserve(8);

  SessionStatus& session_status = get_session_status(ipc_id);
  RimeSessionId session_id = session_status.session_id;
  std::wstring commit_text;
  const std::wstring rime_commit = _TakePendingCommitText(session_id);
  if (!rime_commit.empty()) {
    commit_text.append(rime_commit);
  }
  if (extra_commit && !extra_commit->empty()) {
    commit_text.append(*extra_commit);
  }
  _TryHandleSystemCommandCommit(&commit_text, ipc_id);
  if (!commit_text.empty()) {
    actions.push_back("commit");
    body.append(L"commit=").append(escape_string(commit_text)).append(L"\n");
    _AppendCommittedText(ipc_id, commit_text);
  }

  bool is_composing = false;
  RIME_STRUCT(RimeStatus, status);
  static const std::wstring Bool_wstring[] = {L"0", L"1"};
  if (rime_api->get_status(session_id, &status)) {
    is_composing = !!status.is_composing;
    actions.push_back("status");
    body.append(L"status.ascii_mode=")
        .append(Bool_wstring[!!status.is_ascii_mode])
        .append(L"\n")
        .append(L"status.composing=")
        .append(Bool_wstring[!!status.is_composing])
        .append(L"\n")
        .append(L"status.disabled=")
        .append(Bool_wstring[!!status.is_disabled])
        .append(L"\n")
        .append(L"status.full_shape=")
        .append(Bool_wstring[!!status.is_full_shape])
        .append(L"\n")
        .append(L"status.schema_id=")
        .append(status.schema_id ? u8tow(status.schema_id) : std::wstring())
        .append(L"\n");
    if (m_global_ascii_mode &&
        (session_status.status.is_ascii_mode != status.is_ascii_mode)) {
      for (auto& pair : m_session_status_map) {
        if (pair.first != ipc_id)
          rime_api->set_option(to_session_id(pair.first), "ascii_mode",
                               !!status.is_ascii_mode);
      }
    }
    session_status.status = status;
    rime_api->free_status(&status);
  }

  RIME_STRUCT(RimeContext, ctx);
  if (rime_api->get_context(session_id, &ctx)) {
    bool has_candidates = ctx.menu.num_candidates > 0;
    CandidateInfo cinfo;
    if (has_candidates) {
      _GetCandidateInfo(cinfo, ctx);
    }
    if (is_composing) {
      const auto& preedit = ctx.composition.preedit;
      const auto& start = ctx.composition.sel_start;
      const auto& end = ctx.composition.sel_end;
      const auto& cursor = ctx.composition.cursor_pos;
      static const auto u8towstring = [](const char* u8str, int len = 0) {
        return std::to_wstring(utf8towcslen(u8str, len));
      };
      actions.push_back("ctx");
      switch (session_status.style.preedit_type) {
        case UIStyle::PREVIEW: {
          if (ctx.commit_text_preview) {
            const char* first_utf8 = ctx.commit_text_preview;
            const size_t first_len = std::strlen(first_utf8);
            const std::wstring first_w = escape_string(u8tow(first_utf8));
            const std::wstring tmp = u8towstring(first_utf8, (int)first_len);
            body.append(L"ctx.preedit=")
                .append(first_w)
                .append(L"\n")
                .append(L"ctx.preedit.cursor=")
                .append(u8towstring(first_utf8, 0))
                .append(L",")
                .append(tmp)
                .append(L",")
                .append(tmp)
                .append(L"\n");
            break;
          }
          // no preview, fall back to composition
        }
        case UIStyle::COMPOSITION: {
          body.append(L"ctx.preedit=")
              .append(escape_string(u8tow(preedit)))
              .append(L"\n");
          if (start <= end) {
            body.append(L"ctx.preedit.cursor=")
                .append(u8towstring(preedit, start))
                .append(L",")
                .append(u8towstring(preedit, end))
                .append(L",")
                .append(u8towstring(preedit, cursor))
                .append(L"\n");
          }
          break;
        }
        case UIStyle::PREVIEW_ALL: {
          body.append(L"ctx.preedit=")
              .append(escape_string(u8tow(preedit)))
              .append(L"  [");
          auto label_valid = session_status.style.label_font_point > 0;
          auto comment_valid = session_status.style.comment_font_point > 0;
          const std::wstring mark_text_w =
              session_status.style.mark_text.empty()
                  ? std::wstring(L"*")
                  : session_status.style.mark_text;
          for (auto i = 0; i < ctx.menu.num_candidates; i++) {
            std::wstring label_w;
            if (label_valid) {
              wchar_t buf_lbl[128];
              swprintf_s<128>(buf_lbl,
                              session_status.style.label_text_format.c_str(),
                              cinfo.labels.at(i).str.c_str());
              label_w = std::wstring(buf_lbl);
            }
            std::wstring comment_w =
                comment_valid ? cinfo.comments.at(i).str : std::wstring();
            std::wstring prefix_w = (i != ctx.menu.highlighted_candidate_index)
                                        ? std::wstring()
                                        : mark_text_w;
            body.append(L" ")
                .append(prefix_w)
                .append(escape_string(label_w))
                .append(escape_string(u8tow(ctx.menu.candidates[i].text)))
                .append(L" ")
                .append(escape_string(comment_w));
          }
          body.append(L" ]\n");
          if (start <= end) {
            body.append(L"ctx.preedit.cursor=")
                .append(u8towstring(preedit, start))
                .append(L",")
                .append(u8towstring(preedit, end))
                .append(L",")
                .append(u8towstring(preedit, cursor))
                .append(L"\n");
          }
          break;
        }
      }
    }
    if (has_candidates) {
      std::wstringstream ss;
      boost::archive::text_woarchive oa(ss);

      oa << cinfo;

      auto s = ss.str();
      body.append(L"ctx.cand=").append(std::move(s)).append(L"\n");
    }
    rime_api->free_context(&ctx);
  }

  // configuration information
  actions.push_back("config");
  body.append(L"config.inline_preedit=")
      .append(std::to_wstring((int)session_status.style.inline_preedit))
      .append(L"\n");

  // style
  if (!session_status.__synced) {
    std::wstringstream ss;
    boost::archive::text_woarchive oa(ss);
    oa << session_status.style;

    actions.push_back("style");
    body.append(L"style=").append(ss.str()).append(L"\n");
    session_status.__synced = true;
  }

  // summarize: send header first to avoid vector head-insert cost
  std::wstring header;
  if (actions.empty()) {
    header = L"action=noop\n";
  } else {
    std::string actionList;
    actionList.reserve(64);
    for (size_t i = 0; i < actions.size(); ++i) {
      if (i > 0)
        actionList += ',';
      actionList += actions[i];
    }
    header = std::wstring(L"action=") + u8tow(actionList) + L"\n";
  }
  if (!eat(header))
    return false;

  body.append(L".\n");
  if (!eat(body))
    return false;

  return true;
}

// Blend foreground and background ARGB colors taking alpha into account.
// Returns an ABGR COLORREF with premultiplied alpha blended result.
static inline COLORREF blend_colors(COLORREF fcolor, COLORREF bcolor) {
  // Extract ARGB channels from both colors.
  BYTE fA = (fcolor >> 24) & 0xFF;
  BYTE fB = (fcolor >> 16) & 0xFF;
  BYTE fG = (fcolor >> 8) & 0xFF;
  BYTE fR = fcolor & 0xFF;
  BYTE bA = (bcolor >> 24) & 0xFF;
  BYTE bB = (bcolor >> 16) & 0xFF;
  BYTE bG = (bcolor >> 8) & 0xFF;
  BYTE bR = bcolor & 0xFF;
  // Convert alpha to [0,1]
  float fAlpha = fA / 255.0f;
  float bAlpha = bA / 255.0f;
  // Result alpha
  float retAlpha = fAlpha + (1 - fAlpha) * bAlpha;
  if (retAlpha <= 1e-6f) {
    // Fully transparent result — return background unchanged as fallback.
    return bcolor;
  }
  auto mix = [&](float fc, float bc) -> BYTE {
    return static_cast<BYTE>((fc * fAlpha + bc * bAlpha * (1 - fAlpha)) /
                             retAlpha);
  };
  BYTE retR = mix(fR, bR);
  BYTE retG = mix(fG, bG);
  BYTE retB = mix(fB, bB);
  BYTE outA = static_cast<BYTE>(retAlpha * 255.0f);
  return (static_cast<COLORREF>(outA) << 24) | (retB << 16) | (retG << 8) |
         retR;
}
// parse color value, with fallback value
static Bool _RimeGetColor(RimeConfig* config,
                          const std::string& key,
                          int& value,
                          const ColorFormat& fmt,
                          const unsigned int& fallback) {
  char color[256] = {0};
  if (!rime_api->config_get_string(config, key.c_str(), color, 256)) {
    value = fallback;
    return False;
  }
  const auto color_str = std::string(color);
  // adjudge if str is 0x 0X # hex color format, return trimmed hex part
  // out part is 6 or 8 length hex string without white space
  const auto parse_color_code = [](const std::string& str, std::string& out) {
    if (str.empty())
      return false;
    size_t start = 0;
    if (str[0] == '#') {
      start = 1;
    } else if (str.size() >= 2 &&
               (str.compare(0, 2, "0x") == 0 || str.compare(0, 2, "0X") == 0)) {
      start = 2;
    } else {
      return false;
    }
    const std::string hex_part = str.substr(start);
    if (hex_part.empty())
      return false;
    if ((start == 1 || start == 2) && hex_part.length() != 3 &&
        hex_part.length() != 4 && hex_part.length() != 6 &&
        hex_part.length() != 8) {
      return false;
    }
    for (char c : hex_part) {
      if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') ||
            (c >= 'A' && c <= 'F')))
        return false;
    }
    out = str.substr(start).substr(0, 8);
#define _2C(c) std::string(2, c)
    if (out.size() == 3)
      out = _2C(out[0]) + _2C(out[1]) + _2C(out[2]);
    else if (out.size() == 4)
      out = _2C(out[0]) + _2C(out[1]) + _2C(out[2]) + _2C(out[3]);
#undef _2C
    return true;
  };
  auto hex_color = std::string();
  if (parse_color_code(color_str, hex_color)) {
    value = std::stoul(hex_color, 0, 16);
    if (hex_color.length() == 6)
      value = (fmt != COLOR_RGBA) ? (value | 0xff000000)
                                  : (((unsigned int)value << 8) | 0x000000ff);
  } else {
    if (!rime_api->config_get_int(config, key.c_str(), &value)) {
      value = fallback;
      return False;
    }
    if (value <= 0xffffff)
      value = (fmt != COLOR_RGBA) ? (value | 0xff000000)
                                  : (((unsigned int)value << 8) | 0x000000ff);
    else if (value > 0xffffffff)
      value &= 0xffffffff;
  }
  if (fmt == COLOR_ARGB)
    value = ARGB2ABGR(value);
  else if (fmt == COLOR_RGBA)
    value = RGBA2ABGR(value);
  value &= 0xffffffff;
  return True;
}

template <typename T, size_t N>
using Array = std::array<std::pair<const char*, T>, N>;

// parset bool type configuration to T type value trueValue / falseValue
template <typename T>
void _RimeGetBool(RimeConfig* config,
                  const char* key,
                  bool cond,
                  T& value,
                  const T& trueValue = true,
                  const T& falseValue = false) {
  Bool tempb = False;
  if (rime_api->config_get_bool(config, key, &tempb) || cond)
    value = (!!tempb) ? trueValue : falseValue;
}
// parse string option to T type value, with fallback
template <typename T, size_t N>
void _RimeParseStringOptWithFallback(RimeConfig* config,
                                     const char* key,
                                     T& value,
                                     const Array<T, N>& arr,
                                     const T& fallback) {
  char str_buff[256] = {0};
  if (rime_api->config_get_string(config, key, str_buff, 255)) {
    for (size_t i = 0; i < N; ++i) {
      if (strcmp(arr[i].first, str_buff) == 0) {
        value = arr[i].second;
        return;
      }
    }
  }
  value = fallback;
}

template <typename T>
void _RimeGetIntStr(RimeConfig* config,
                    const char* key,
                    T& value,
                    const char* fb_key = nullptr,
                    const void* fb_value = nullptr,
                    const std::function<void(T&)>& func = nullptr) {
  if constexpr (std::is_same<T, int>::value) {
    if (!rime_api->config_get_int(config, key, &value) && fb_key != 0)
      rime_api->config_get_int(config, fb_key, &value);
  } else if constexpr (std::is_same<T, std::wstring>::value) {
    const int BUF_SIZE = 2047;
    char buffer[BUF_SIZE + 1] = {0};
    if (rime_api->config_get_string(config, key, buffer, BUF_SIZE) ||
        rime_api->config_get_string(config, fb_key, buffer, BUF_SIZE)) {
      value = u8tow(buffer);
    } else if (fb_value) {
      value = *(T*)fb_value;
    }
  }
  if (func)
    func(value);
}

// Helper to iterate a Rime map and invoke callback with key/path
static void ForEachRimeMap(
    RimeConfig* config,
    const std::string& path,
    const std::function<void(const char* key, const char* child_path)>& cb) {
  RimeConfigIterator iter;
  if (!rime_api->config_begin_map(&iter, config, path.c_str()))
    return;
  while (rime_api->config_next(&iter)) {
    cb(iter.key, iter.path);
  }
  rime_api->config_end(&iter);
}

// Helper to iterate a Rime list and invoke callback with item path
static void ForEachRimeList(
    RimeConfig* config,
    const std::string& path,
    const std::function<void(const char* item_path)>& cb) {
  RimeConfigIterator iter;
  if (!rime_api->config_begin_list(&iter, config, path.c_str()))
    return;
  while (rime_api->config_next(&iter)) {
    cb(iter.path);
  }
  rime_api->config_end(&iter);
}

void RimeWithWeaselHandler::_UpdateShowNotifications(RimeConfig* config,
                                                     bool initialize) {
  Bool show_notifications = true;
  if (initialize)
    m_show_notifications_base.clear();
  m_show_notifications.clear();

  if (rime_api->config_get_bool(config, "show_notifications",
                                &show_notifications)) {
    // config read as bool, for global all on or off
    if (show_notifications)
      m_show_notifications["always"] = true;
    if (initialize)
      m_show_notifications_base = m_show_notifications;
  } else {
    // read as list using helper
    ForEachRimeList(config, "show_notifications", [&](const char* item_path) {
      char buffer[256] = {0};
      if (rime_api->config_get_string(config, item_path, buffer, 256))
        m_show_notifications[std::string(buffer)] = true;
    });
    if (initialize)
      m_show_notifications_base = m_show_notifications;
    if (m_show_notifications.empty()) {
      // not configured, or incorrect type
      if (initialize)
        m_show_notifications_base["always"] = true;
      m_show_notifications = m_show_notifications_base;
    }
  }
}

// update ui's style parameters, ui has been check before referenced
static void _UpdateUIStyle(RimeConfig* config, UI* ui, bool initialize) {
  UIStyle& style(ui->style());
  const std::function<void(std::wstring&)> rmspace = [](std::wstring& str) {
    str = std::regex_replace(str, std::wregex(L"\\s*(,|:|^|$)\\s*"), L"$1");
  };
  const std::function<void(int&)> _abs = [](int& value) { value = abs(value); };
  // get font faces
  _RimeGetIntStr(config, "style/font_face", style.font_face, 0, 0, rmspace);
  std::wstring* const pFallbackFontFace = initialize ? &style.font_face : NULL;
  _RimeGetIntStr(config, "style/label_font_face", style.label_font_face, 0,
                 pFallbackFontFace, rmspace);
  _RimeGetIntStr(config, "style/comment_font_face", style.comment_font_face, 0,
                 pFallbackFontFace, rmspace);
  // able to set label font/comment font empty, force fallback to font face.
  if (style.label_font_face.empty())
    style.label_font_face = style.font_face;
  if (style.comment_font_face.empty())
    style.comment_font_face = style.font_face;
  // get font points
  _RimeGetIntStr(config, "style/font_point", style.font_point);
  if (style.font_point <= 0)
    style.font_point = 12;
  _RimeGetIntStr(config, "style/label_font_point", style.label_font_point,
                 "style/font_point", 0, _abs);
  _RimeGetIntStr(config, "style/comment_font_point", style.comment_font_point,
                 "style/font_point", 0, _abs);
  _RimeGetIntStr(config, "style/candidate_abbreviate_length",
                 style.candidate_abbreviate_length, 0, 0, _abs);
  _RimeGetBool(config, "style/inline_preedit", initialize,
               style.inline_preedit);
  _RimeGetBool(config, "style/vertical_auto_reverse", initialize,
               style.vertical_auto_reverse);
  static constexpr Array<UIStyle::PreeditType, 3> _preeditArr = {
      {{"composition", UIStyle::COMPOSITION},
       {"preview", UIStyle::PREVIEW},
       {"preview_all", UIStyle::PREVIEW_ALL}}};
  _RimeParseStringOptWithFallback(config, "style/preedit_type",
                                  style.preedit_type, _preeditArr,
                                  style.preedit_type);
  static constexpr Array<UIStyle::AntiAliasMode, 5> _aliasModeArr = {
      {{"force_dword", UIStyle::FORCE_DWORD},
       {"cleartype", UIStyle::CLEARTYPE},
       {"grayscale", UIStyle::GRAYSCALE},
       {"aliased", UIStyle::ALIASED},
       {"default", UIStyle::DEFAULT}}};
  _RimeParseStringOptWithFallback(config, "style/antialias_mode",
                                  style.antialias_mode, _aliasModeArr,
                                  style.antialias_mode);
  static constexpr Array<UIStyle::HoverType, 3> _hoverTypeArr = {
      {{"none", UIStyle::HoverType::NONE},
       {"semi_hilite", UIStyle::HoverType::SEMI_HILITE},
       {"hilite", UIStyle::HoverType::HILITE}}};
  _RimeParseStringOptWithFallback(config, "style/hover_type", style.hover_type,
                                  _hoverTypeArr, style.hover_type);
  static constexpr Array<UIStyle::LayoutAlignType, 3> _alignType = {
      {{"top", UIStyle::ALIGN_TOP},
       {"center", UIStyle::ALIGN_CENTER},
       {"bottom", UIStyle::ALIGN_BOTTOM}}};
  _RimeParseStringOptWithFallback(config, "style/layout/align_type",
                                  style.align_type, _alignType,
                                  style.align_type);
  _RimeGetBool(config, "style/display_tray_icon", initialize,
               style.display_tray_icon);
  _RimeGetBool(config, "style/ascii_tip_follow_cursor", initialize,
               style.ascii_tip_follow_cursor);
  _RimeGetBool(config, "style/horizontal", initialize, style.layout_type,
               UIStyle::LAYOUT_HORIZONTAL, UIStyle::LAYOUT_VERTICAL);
  _RimeGetBool(config, "style/paging_on_scroll", initialize,
               style.paging_on_scroll);
  _RimeGetBool(config, "style/click_to_capture", initialize,
               style.click_to_capture, true, false);
  _RimeGetBool(config, "style/fullscreen", false, style.layout_type,
               ((style.layout_type == UIStyle::LAYOUT_HORIZONTAL)
                    ? UIStyle::LAYOUT_HORIZONTAL_FULLSCREEN
                    : UIStyle::LAYOUT_VERTICAL_FULLSCREEN),
               style.layout_type);
  _RimeGetBool(config, "style/vertical_text", false, style.layout_type,
               UIStyle::LAYOUT_VERTICAL_TEXT, style.layout_type);
  _RimeGetBool(config, "style/vertical_text_left_to_right", false,
               style.vertical_text_left_to_right);
  _RimeGetBool(config, "style/vertical_text_with_wrap", false,
               style.vertical_text_with_wrap);
  static constexpr Array<bool, 2> _text_orientation = {
      {{"horizontal", false}, {"vertical", true}}};
  bool _text_orientation_bool = false;
  _RimeParseStringOptWithFallback(config, "style/text_orientation",
                                  _text_orientation_bool, _text_orientation,
                                  _text_orientation_bool);
  if (_text_orientation_bool)
    style.layout_type = UIStyle::LAYOUT_VERTICAL_TEXT;
  _RimeGetIntStr(config, "style/label_format", style.label_text_format);
  _RimeGetIntStr(config, "style/mark_text", style.mark_text);
  _RimeGetIntStr(config, "style/layout/baseline", style.baseline, 0, 0, _abs);
  _RimeGetIntStr(config, "style/layout/linespacing", style.linespacing, 0, 0,
                 _abs);
  _RimeGetIntStr(config, "style/layout/min_width", style.min_width, 0, 0, _abs);
  _RimeGetIntStr(config, "style/layout/max_width", style.max_width, 0, 0, _abs);
  _RimeGetIntStr(config, "style/layout/min_height", style.min_height, 0, 0,
                 _abs);
  _RimeGetIntStr(config, "style/layout/max_height", style.max_height, 0, 0,
                 _abs);
  // layout (alternative to style/horizontal)
  static constexpr Array<UIStyle::LayoutType, 5> _layoutArr = {
      {{"vertical", UIStyle::LAYOUT_VERTICAL},
       {"horizontal", UIStyle::LAYOUT_HORIZONTAL},
       {"vertical_text", UIStyle::LAYOUT_VERTICAL_TEXT},
       {"vertical+fullscreen", UIStyle::LAYOUT_VERTICAL_FULLSCREEN},
       {"horizontal+fullscreen", UIStyle::LAYOUT_HORIZONTAL_FULLSCREEN}}};
  _RimeParseStringOptWithFallback(config, "style/layout/type",
                                  style.layout_type, _layoutArr,
                                  style.layout_type);
  // disable max_width when full screen
  if (style.layout_type == UIStyle::LAYOUT_HORIZONTAL_FULLSCREEN ||
      style.layout_type == UIStyle::LAYOUT_VERTICAL_FULLSCREEN) {
    style.max_width = 0;
    style.inline_preedit = false;
  }
  _RimeGetIntStr(config, "style/layout/border", style.border,
                 "style/layout/border_width", 0, _abs);
  _RimeGetIntStr(config, "style/layout/margin_x", style.margin_x);
  _RimeGetIntStr(config, "style/layout/margin_y", style.margin_y);
  _RimeGetIntStr(config, "style/layout/spacing", style.spacing, 0, 0, _abs);
  _RimeGetIntStr(config, "style/layout/candidate_spacing",
                 style.candidate_spacing, 0, 0, _abs);
  _RimeGetIntStr(config, "style/layout/hilite_spacing", style.hilite_spacing, 0,
                 0, _abs);
  _RimeGetIntStr(config, "style/layout/hilite_padding_x",
                 style.hilite_padding_x, "style/layout/hilite_padding", 0,
                 _abs);
  _RimeGetIntStr(config, "style/layout/hilite_padding_y",
                 style.hilite_padding_y, "style/layout/hilite_padding", 0,
                 _abs);
  _RimeGetIntStr(config, "style/layout/shadow_radius", style.shadow_radius, 0,
                 0, _abs);
  // disable shadow for fullscreen layout
  style.shadow_radius *=
      (!(style.layout_type == UIStyle::LAYOUT_HORIZONTAL_FULLSCREEN ||
         style.layout_type == UIStyle::LAYOUT_VERTICAL_FULLSCREEN));
  _RimeGetIntStr(config, "style/layout/shadow_offset_x", style.shadow_offset_x);
  _RimeGetIntStr(config, "style/layout/shadow_offset_y", style.shadow_offset_y);
  // round_corner as alias of hilited_corner_radius
  _RimeGetIntStr(config, "style/layout/hilited_corner_radius",
                 style.round_corner, "style/layout/round_corner", 0, _abs);
  // corner_radius not set, fallback to round_corner
  _RimeGetIntStr(config, "style/layout/corner_radius", style.round_corner_ex,
                 "style/layout/round_corner", 0, _abs);
  // fix padding and spacing settings
  if (style.layout_type != UIStyle::LAYOUT_VERTICAL_TEXT) {
    // hilite_padding vs spacing
    // if hilite_padding over spacing, increase spacing
    style.spacing = max(style.spacing, style.hilite_padding_y * 2);
    // hilite_padding vs candidate_spacing
    if (style.layout_type == UIStyle::LAYOUT_VERTICAL_FULLSCREEN ||
        style.layout_type == UIStyle::LAYOUT_VERTICAL) {
      // vertical, if hilite_padding_y over candidate spacing,
      // increase candidate spacing
      style.candidate_spacing =
          max(style.candidate_spacing, style.hilite_padding_y * 2);
    } else {
      // horizontal, if hilite_padding_x over candidate
      // spacing, increase candidate spacing
      style.candidate_spacing =
          max(style.candidate_spacing, style.hilite_padding_x * 2);
    }
    // hilite_padding_x vs hilite_spacing
    if (!style.inline_preedit)
      style.hilite_spacing = max(style.hilite_spacing, style.hilite_padding_x);
  } else  // LAYOUT_VERTICAL_TEXT
  {
    // hilite_padding_x vs spacing
    // if hilite_padding over spacing, increase spacing
    style.spacing = max(style.spacing, style.hilite_padding_x * 2);
    // hilite_padding vs candidate_spacing
    // if hilite_padding_x over candidate
    // spacing, increase candidate spacing
    style.candidate_spacing =
        max(style.candidate_spacing, style.hilite_padding_x * 2);
    // vertical_text_with_wrap and hilite_padding_y over candidate_spacing
    if (style.vertical_text_with_wrap)
      style.candidate_spacing =
          max(style.candidate_spacing, style.hilite_padding_y * 2);
    // hilite_padding_y vs hilite_spacing
    if (!style.inline_preedit)
      style.hilite_spacing = max(style.hilite_spacing, style.hilite_padding_y);
  }
  // fix padding and margin settings
  int scale = style.margin_x < 0 ? -1 : 1;
  style.margin_x = scale * max(style.hilite_padding_x, abs(style.margin_x));
  scale = style.margin_y < 0 ? -1 : 1;
  style.margin_y = scale * max(style.hilite_padding_y, abs(style.margin_y));
  // get enhanced_position
  _RimeGetBool(config, "style/enhanced_position", initialize,
               style.enhanced_position, true, false);
  // get color scheme
  const int BUF_SIZE = 255;
  char buffer[BUF_SIZE + 1] = {0};
  if (initialize && rime_api->config_get_string(config, "style/color_scheme",
                                                buffer, BUF_SIZE))
    _UpdateUIStyleColor(config, style);
}
// load color configs to style, by "style/color_scheme" or specific scheme name
// "color" which is default empty
static bool _UpdateUIStyleColor(RimeConfig* config,
                                UIStyle& style,
                                const std::string& color) {
  const int BUF_SIZE = 255;
  char buffer[BUF_SIZE + 1] = {0};
  std::string color_mark = "style/color_scheme";
  // color scheme
  if (rime_api->config_get_string(config, color_mark.c_str(), buffer,
                                  BUF_SIZE) ||
      !color.empty()) {
    std::string prefix("preset_color_schemes/");
    prefix += (color.empty()) ? buffer : color;
    // define color format, default abgr if not set
    ColorFormat fmt = COLOR_ABGR;
    static constexpr Array<ColorFormat, 3> _colorFmt = {
        {{"argb", COLOR_ARGB}, {"rgba", COLOR_RGBA}, {"abgr", COLOR_ABGR}}};
    _RimeParseStringOptWithFallback(config, (prefix + "/color_format").c_str(),
                                    fmt, _colorFmt, COLOR_ABGR);
#define COLOR(key, value, fallback) \
  _RimeGetColor(config, (prefix + "/" + key), value, fmt, fallback)
    COLOR("back_color", style.back_color, 0xffffffff);
    COLOR("shadow_color", style.shadow_color, 0);
    COLOR("prevpage_color", style.prevpage_color, 0);
    COLOR("nextpage_color", style.nextpage_color, 0);
    COLOR("text_color", style.text_color, 0xff000000);
    COLOR("candidate_text_color", style.candidate_text_color, style.text_color);
    COLOR("candidate_back_color", style.candidate_back_color, 0);
    COLOR("border_color", style.border_color, style.text_color);
    COLOR("hilited_text_color", style.hilited_text_color, style.text_color);
    COLOR("hilited_back_color", style.hilited_back_color, style.back_color);
    COLOR("hilited_candidate_text_color", style.hilited_candidate_text_color,
          style.hilited_text_color);
    COLOR("hilited_candidate_back_color", style.hilited_candidate_back_color,
          style.hilited_back_color);
    COLOR("hilited_candidate_shadow_color",
          style.hilited_candidate_shadow_color, 0);
    COLOR("hilited_shadow_color", style.hilited_shadow_color, 0);
    COLOR("candidate_shadow_color", style.candidate_shadow_color, 0);
    COLOR("candidate_border_color", style.candidate_border_color, 0);
    COLOR("hilited_candidate_border_color",
          style.hilited_candidate_border_color, 0);
    COLOR("label_color", style.label_text_color,
          blend_colors(style.candidate_text_color, style.candidate_back_color));
    COLOR("hilited_label_color", style.hilited_label_text_color,
          blend_colors(style.hilited_candidate_text_color,
                       style.hilited_candidate_back_color));
    COLOR("comment_text_color", style.comment_text_color,
          style.label_text_color);
    COLOR("hilited_comment_text_color", style.hilited_comment_text_color,
          style.hilited_label_text_color);
    COLOR("hilited_mark_color", style.hilited_mark_color, 0);
#undef COLOR
    return true;
  }
  return false;
}
static void _LoadAppOptions(RimeConfig* config,
                            AppOptionsByAppName& app_options) {
  app_options.clear();
  ForEachRimeMap(
      config, "app_options", [&](const char* app_key, const char* app_path) {
        AppOptions& options(app_options[app_key]);
        ForEachRimeMap(
            config, app_path, [&](const char* opt_key, const char* opt_path) {
              Bool value = False;
              if (rime_api->config_get_bool(config, opt_path, &value)) {
                options[opt_key] = !!value;
              }
            });
      });
}

void RimeWithWeaselHandler::_GetStatus(Status& stat,
                                       WeaselSessionId ipc_id,
                                       Context& ctx) {
  SessionStatus& session_status = get_session_status(ipc_id);
  RimeSessionId session_id = session_status.session_id;
  RIME_STRUCT(RimeStatus, status);
  if (rime_api->get_status(session_id, &status)) {
    std::string schema_id = "";
    if (status.schema_id)
      schema_id = status.schema_id;
    stat.schema_name = u8tow(status.schema_name);
    stat.schema_id = u8tow(status.schema_id);
    stat.ascii_mode = !!status.is_ascii_mode;
    stat.composing = !!status.is_composing;
    stat.disabled = !!status.is_disabled;
    stat.full_shape = !!status.is_full_shape;
    if (schema_id != m_last_schema_id) {
      session_status.__synced = false;
      m_last_schema_id = schema_id;
      if (schema_id != ".default") {  // don't load for schema select menu
        bool inline_preedit = session_status.style.inline_preedit;
        _LoadSchemaSpecificSettings(ipc_id, schema_id);
        _LoadAppInlinePreeditSet(ipc_id, true);
        if (session_status.style.inline_preedit != inline_preedit)
          // in case of inline_preedit set in schema
          _UpdateInlinePreeditStatus(ipc_id);
        // refresh icon after schema changed
        _RefreshTrayIcon(session_id, _UpdateUICallback);
        m_ui->style() = session_status.style;
        if (m_show_notifications.find("schema") != m_show_notifications.end() &&
            m_show_notifications_time > 0) {
          ctx.aux.str = stat.schema_name;
          m_ui->Update(ctx, stat);
          m_ui->ShowWithTimeout(m_show_notifications_time);
        }
      }
    }
    rime_api->free_status(&status);
  }
}

void RimeWithWeaselHandler::_GetContext(Context& weasel_context,
                                        RimeSessionId session_id) {
  RIME_STRUCT(RimeContext, ctx);
  if (rime_api->get_context(session_id, &ctx)) {
    if (ctx.composition.length > 0) {
      weasel_context.preedit.str = u8tow(ctx.composition.preedit);
      if (ctx.composition.sel_start < ctx.composition.sel_end) {
        TextAttribute attr;
        attr.type = HIGHLIGHTED;
        attr.range.start =
            utf8towcslen(ctx.composition.preedit, ctx.composition.sel_start);
        attr.range.end =
            utf8towcslen(ctx.composition.preedit, ctx.composition.sel_end);

        weasel_context.preedit.attributes.push_back(attr);
      }
    }
    if (ctx.menu.num_candidates) {
      CandidateInfo& cinfo(weasel_context.cinfo);
      _GetCandidateInfo(cinfo, ctx);
    }
    rime_api->free_context(&ctx);
  }
}

void RimeWithWeaselHandler::_UpdateInlinePreeditStatus(WeaselSessionId ipc_id) {
  if (!m_ui)
    return;
  SessionStatus& session_status = get_session_status(ipc_id);
  RimeSessionId session_id = session_status.session_id;
  // set inline_preedit option
  bool inline_preedit = session_status.style.inline_preedit;
  rime_api->set_option(session_id, "inline_preedit", Bool(inline_preedit));
  // show soft cursor on weasel panel but not inline
  rime_api->set_option(session_id, "soft_cursor", Bool(!inline_preedit));
}
