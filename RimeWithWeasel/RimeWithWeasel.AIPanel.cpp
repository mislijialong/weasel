#include "stdafx.h"

#include <AIAssistantInstructions.h>
#include <RimeWithWeasel.h>
#include <WeaselUtility.h>
#include <logging.h>
#include <winhttp.h>
#include <ShellScalingApi.h>

#include "RimeWithWeaselInternal.h"

#include <algorithm>
#include <cstdlib>
#include <filesystem>
#include <mutex>
#include <string>
#include <vector>

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

#pragma comment(lib, "Shcore.lib")

using namespace weasel;

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
constexpr int kAIPanelWidth = 440;
constexpr int kAIPanelHeight = 300;
constexpr int kAIPanelMinWidth = 180;
constexpr int kAIPanelMaxWidth = 860;
constexpr int kAIPanelMinHeight = 220;
constexpr int kAIPanelMaxHeight = 560;
constexpr int kAIPanelScreenMargin = 8;
constexpr int kAIPanelAnchorGap = 6;
constexpr int kAIPanelPadding = 0;
constexpr int kAIPanelResizeHitBand = 1;
constexpr int kAIPanelResizeHitTestBand = 4;
constexpr int kAIPanelCornerRadius = 12;
constexpr int kAIPanelButtonWidth = 96;
constexpr int kAIPanelButtonHeight = 30;
constexpr UINT kAIPanelControlStatus = 4101;
constexpr UINT kAIPanelControlOutput = 4102;
constexpr UINT kAIPanelControlConfirm = 4104;
constexpr UINT kAIPanelControlCancel = 4105;
constexpr UINT WM_AI_WEBVIEW_INIT = WM_APP + 404;
constexpr UINT WM_AI_PANEL_DESTROY = WM_APP + 405;
constexpr UINT WM_AI_PANEL_DRAG = WM_APP + 407;
constexpr UINT WM_AI_WEBVIEW_FOCUS = WM_APP + 408;
constexpr UINT kWMNcuahDrawCaption = 0x00AE;
constexpr UINT kWMNcuahDrawFrame = 0x00AF;
constexpr int kAIPanelResizeEdgeNone = 0;
constexpr int kAIPanelResizeEdgeRight = 1;
constexpr int kAIPanelResizeEdgeBottom = 2;

struct AIPanelUiCommand {
  std::string type;
  std::wstring text;
  std::wstring institution_id;
  int panel_width = 0;
  int panel_height = 0;
  std::string resize_reason;
};

UINT ResolveAIPanelDpi(HMONITOR monitor, HWND hwnd) {
  UINT dpi_x = USER_DEFAULT_SCREEN_DPI;
  UINT dpi_y = USER_DEFAULT_SCREEN_DPI;
  if (monitor &&
      SUCCEEDED(GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpi_x, &dpi_y)) &&
      dpi_x > 0) {
    return dpi_x;
  }

  const HWND dc_hwnd = hwnd && IsWindow(hwnd) ? hwnd : nullptr;
  HDC screen_dc = GetDC(dc_hwnd);
  if (screen_dc) {
    const int dpi = GetDeviceCaps(screen_dc, LOGPIXELSX);
    ReleaseDC(dc_hwnd, screen_dc);
    if (dpi > 0) {
      return static_cast<UINT>(dpi);
    }
  }
  return USER_DEFAULT_SCREEN_DPI;
}

int ScaleAIPanelLogicalToPixels(int logical_size, UINT dpi) {
  const UINT resolved_dpi = dpi > 0 ? dpi : USER_DEFAULT_SCREEN_DPI;
  return max(1, MulDiv(logical_size, static_cast<int>(resolved_dpi),
                       USER_DEFAULT_SCREEN_DPI));
}

int ScaleAIPanelPixelsToLogical(int pixel_size, UINT dpi) {
  const UINT resolved_dpi = dpi > 0 ? dpi : USER_DEFAULT_SCREEN_DPI;
  return max(1, MulDiv(pixel_size, USER_DEFAULT_SCREEN_DPI,
                       static_cast<int>(resolved_dpi)));
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
    // Hide the visible DWM border while keeping the resize frame behavior.
    constexpr DWORD kDWMWABorderColor = 34;
    constexpr COLORREF kDwmColorNone =
        static_cast<COLORREF>(0xFFFFFFFE);
    set_window_attribute(hwnd, kDWMWABorderColor, &kDwmColorNone,
                         sizeof(kDwmColorNone));
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


void LayoutAIPanelControls(HWND hwnd) {
  RECT client_rect = {0};
  if (!GetClientRect(hwnd, &client_rect)) {
    return;
  }
  HWND status_hwnd = GetDlgItem(hwnd, kAIPanelControlStatus);
  HWND output_hwnd = GetDlgItem(hwnd, kAIPanelControlOutput);
  HWND confirm_hwnd = GetDlgItem(hwnd, kAIPanelControlConfirm);
  HWND cancel_hwnd = GetDlgItem(hwnd, kAIPanelControlCancel);
  if (!output_hwnd) {
    return;
  }
  const int client_width = client_rect.right - client_rect.left;
  const int client_height = client_rect.bottom - client_rect.top;
  const int content_width =
      max(1, client_width - kAIPanelResizeHitBand);
  const int content_height =
      max(1, client_height - kAIPanelResizeHitBand);
  const bool hide_native_header_footer =
      status_hwnd && confirm_hwnd && cancel_hwnd &&
      !IsWindowVisible(status_hwnd) &&
      !IsWindowVisible(confirm_hwnd) && !IsWindowVisible(cancel_hwnd);
  if (hide_native_header_footer) {
    MoveWindow(output_hwnd, kAIPanelPadding, kAIPanelPadding,
               max(1, content_width - kAIPanelPadding * 2),
               max(1, content_height - kAIPanelPadding * 2), TRUE);
    return;
  }
  const int button_gap = 8;
  const int status_height = 22;
  const int output_top = kAIPanelPadding + status_height + 8;
  const int button_top =
      content_height - kAIPanelPadding - kAIPanelButtonHeight;
  const int output_height = max(80, button_top - output_top - 10);
  const int button_y =
      content_height - kAIPanelPadding - kAIPanelButtonHeight;
  const int cancel_x =
      content_width - kAIPanelPadding - kAIPanelButtonWidth;
  const int confirm_x = cancel_x - button_gap - kAIPanelButtonWidth;

  if (status_hwnd) {
    MoveWindow(status_hwnd, kAIPanelPadding, kAIPanelPadding,
               content_width - kAIPanelPadding * 2, status_height, TRUE);
  }
  MoveWindow(output_hwnd, kAIPanelPadding, output_top,
             content_width - kAIPanelPadding * 2, output_height, TRUE);
  if (confirm_hwnd) {
    MoveWindow(confirm_hwnd, confirm_x, button_y, kAIPanelButtonWidth,
               kAIPanelButtonHeight, TRUE);
  }
  if (cancel_hwnd) {
    MoveWindow(cancel_hwnd, cancel_x, button_y, kAIPanelButtonWidth,
               kAIPanelButtonHeight, TRUE);
  }
}

int HitTestAIPanelResizeEdges(HWND hwnd, const POINT& screen_point) {
  RECT client_rect = {0};
  if (!hwnd || !GetClientRect(hwnd, &client_rect)) {
    return kAIPanelResizeEdgeNone;
  }
  POINT client_point = screen_point;
  if (!ScreenToClient(hwnd, &client_point)) {
    return kAIPanelResizeEdgeNone;
  }
  if (client_point.x < 0 || client_point.y < 0 ||
      client_point.x >= client_rect.right ||
      client_point.y >= client_rect.bottom) {
    return kAIPanelResizeEdgeNone;
  }

  int edges = kAIPanelResizeEdgeNone;
  if (client_point.x >= client_rect.right - kAIPanelResizeHitTestBand) {
    edges |= kAIPanelResizeEdgeRight;
  }
  if (client_point.y >= client_rect.bottom - kAIPanelResizeHitTestBand) {
    edges |= kAIPanelResizeEdgeBottom;
  }
  return edges;
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
                                    const std::string& refresh_token,
                                    const std::wstring& selected_institution_id =
                                        std::wstring(),
                                    const AIPanelInstitutionOption* selected_institution_option =
                                        nullptr) {
  if (panel_url.empty()) {
    return panel_url;
  }
  std::wstring query =
      BuildAIPanelAuthQueryString(token, tenant_id, refresh_token);
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
  if (!selected_institution_id.empty()) {
    append_pair(L"selectedInstitutionId", wtou8(selected_institution_id));
    append_pair(L"directOpenToken", std::to_string(GetTickCount64()));
    if (selected_institution_option) {
      if (!selected_institution_option->name.empty()) {
        append_pair(L"selectedInstitutionName",
                    wtou8(selected_institution_option->name));
      }
      if (!selected_institution_option->app_key.empty()) {
        append_pair(L"selectedInstitutionAppKey",
                    wtou8(selected_institution_option->app_key));
      }
      if (!selected_institution_option->template_content.empty()) {
        append_pair(L"selectedInstitutionTemplate",
                    wtou8(selected_institution_option->template_content));
      }
    }
  }

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

std::wstring ResolveAIPanelAllowedOrigin(const AIAssistantConfig& config) {
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
    json += "\",\"template\":\"";
    json += EscapeJsonString(wtou8(option.template_content));
    json += "\"}";
  }
  json += "]}}";
  return json;
}

}  // namespace

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

    const DWORD panel_style = WS_POPUP | WS_CLIPCHILDREN;
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
    HWND confirm_hwnd = CreateWindowExW(
        0, L"BUTTON", L"确认回写", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0,
        0, 0, 0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlConfirm)),
        instance, nullptr);
    HWND cancel_hwnd = CreateWindowExW(
        0, L"BUTTON", L"取消", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, panel_hwnd,
        reinterpret_cast<HMENU>(static_cast<UINT_PTR>(kAIPanelControlCancel)),
        instance, nullptr);
    if (!status_hwnd || !output_hwnd || !confirm_hwnd || !cancel_hwnd) {
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
    SendMessageW(confirm_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    SendMessageW(cancel_hwnd, WM_SETFONT,
                 reinterpret_cast<WPARAM>(default_font), TRUE);
    EnableWindow(confirm_hwnd, FALSE);
    ShowWindow(status_hwnd, SW_HIDE);
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
      m_ai_panel.resize_tracking = false;
      m_ai_panel.resize_edges = kAIPanelResizeEdgeNone;
      m_ai_panel.resize_start_screen_x = 0;
      m_ai_panel.resize_start_screen_y = 0;
      m_ai_panel.resize_start_window_x = 0;
      m_ai_panel.resize_start_window_y = 0;
      m_ai_panel.resize_start_width = 0;
      m_ai_panel.resize_start_height = 0;
      m_ai_panel.focus_pending = false;
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
        m_ai_panel.resize_tracking = false;
        m_ai_panel.resize_edges = kAIPanelResizeEdgeNone;
        m_ai_panel.resize_start_screen_x = 0;
        m_ai_panel.resize_start_screen_y = 0;
        m_ai_panel.resize_start_window_x = 0;
        m_ai_panel.resize_start_window_y = 0;
        m_ai_panel.resize_start_width = 0;
        m_ai_panel.resize_start_height = 0;
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
  int logical_width = kAIPanelWidth;
  int logical_height = kAIPanelHeight;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    if (!panel_hwnd || !IsWindow(panel_hwnd)) {
      return;
    }
    target_hwnd = m_ai_panel.target_hwnd;
    logical_width =
        m_ai_panel.panel_width > 0 ? m_ai_panel.panel_width : kAIPanelWidth;
    logical_height =
        m_ai_panel.panel_height > 0 ? m_ai_panel.panel_height : kAIPanelHeight;
    if (requested_width > 0) {
      logical_width = requested_width;
    }
    if (requested_height > 0) {
      logical_height = requested_height;
    }
    logical_width = max(kAIPanelMinWidth, min(kAIPanelMaxWidth, logical_width));
    logical_height =
        max(kAIPanelMinHeight, min(kAIPanelMaxHeight, logical_height));
    m_ai_panel.panel_width = logical_width;
    m_ai_panel.panel_height = logical_height;
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

  const UINT dpi = ResolveAIPanelDpi(monitor, panel_hwnd);
  const int width = ScaleAIPanelLogicalToPixels(logical_width, dpi);
  const int height = ScaleAIPanelLogicalToPixels(logical_height, dpi);
  const int screen_margin = ScaleAIPanelLogicalToPixels(kAIPanelScreenMargin, dpi);
  const int anchor_gap = ScaleAIPanelLogicalToPixels(kAIPanelAnchorGap, dpi);
  const int min_x = work_rect.left + screen_margin;
  const int max_x = max(min_x, work_rect.right - screen_margin - width);
  const int min_y = work_rect.top + screen_margin;
  const int max_y = max(min_y, work_rect.bottom - screen_margin - height);
  int x = min_x;
  int y = min_y;
  if (prefer_anchor_position && has_anchor) {
    x = anchor_rect.left;
    y = anchor_rect.bottom + anchor_gap;
  } else if (has_last_position) {
    x = last_panel_x;
    y = last_panel_y;
  } else if (has_anchor) {
    x = anchor_rect.left;
    y = anchor_rect.bottom + anchor_gap;
    if (y + height > work_rect.bottom - screen_margin) {
      const int above = anchor_rect.top - height - anchor_gap;
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
                                         const std::vector<AIPanelInstitutionOption>* initial_options,
                                         bool institutions_ready,
                                         const std::wstring& initial_selected_institution_id) {
  if (!_EnsureAIPanelWindow()) {
    return false;
  }

  HWND panel_hwnd = nullptr;
  HWND status_hwnd = nullptr;
  HWND output_hwnd = nullptr;
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (!m_ai_panel.panel_hwnd || !IsWindow(m_ai_panel.panel_hwnd)) {
      return false;
    }
    m_ai_panel.ipc_id = ipc_id;
    m_ai_panel.target_hwnd = target_hwnd;
    if (m_ai_panel.target_hwnd == m_ai_panel.panel_hwnd ||
        !IsWindow(m_ai_panel.target_hwnd)) {
      m_ai_panel.target_hwnd = nullptr;
    }
    m_ai_panel.output_text.clear();
    m_ai_panel.selected_institution_id.clear();
    if (!initial_selected_institution_id.empty()) {
      m_ai_panel.selected_institution_id = initial_selected_institution_id;
    }
    if (initial_options) {
      m_ai_panel.institution_options = *initial_options;
    } else if (institutions_ready) {
      m_ai_panel.institution_options.clear();
    }
    m_ai_panel.status_text = L"上下文已就绪，请在前端面板中发起请求。";
    m_ai_panel.requesting = false;
    m_ai_panel.institutions_loading = !institutions_ready;
    m_ai_panel.completed = false;
    m_ai_panel.has_error = false;
    m_ai_panel.focus_pending = true;
    panel_hwnd = m_ai_panel.panel_hwnd;
    status_hwnd = m_ai_panel.status_hwnd;
    output_hwnd = m_ai_panel.output_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
  }

  SetWindowTextW(status_hwnd, L"上下文已就绪，请在前端面板中发起请求。");
  SetWindowTextW(output_hwnd, L"");
  EnableWindow(confirm_hwnd, FALSE);
  EnableWindow(cancel_hwnd, TRUE);
  SetWindowTextW(cancel_hwnd, L"取消");

  _ApplyAIPanelSizeAndReposition(0, 0, true);
  SetWindowPos(panel_hwnd, HWND_TOPMOST, 0, 0, 0, 0,
               SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE);
  ShowWindow(panel_hwnd, SW_SHOWNORMAL);
  UpdateWindow(panel_hwnd);
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_FOCUS, 0, 0);
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_INIT, 0, 0);
  PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  if (!institutions_ready) {
    _RefreshAIPanelInstitutionOptions();
  }
  return true;
}

void RimeWithWeaselHandler::_CloseAIPanel() {
  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    m_ai_panel.target_hwnd = nullptr;
    m_ai_panel.ipc_id = 0;
    m_ai_panel.focus_pending = false;
    m_ai_panel.status_text.clear();
    m_ai_panel.output_text.clear();
    m_ai_panel.selected_institution_id.clear();
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
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    panel_hwnd = m_ai_panel.panel_hwnd;
  }
  if (output_hwnd && IsWindow(output_hwnd)) {
    SetWindowTextW(output_hwnd, L"");
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
  HWND confirm_hwnd = nullptr;
  HWND cancel_hwnd = nullptr;
  bool has_output = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    m_ai_panel.requesting = false;
    m_ai_panel.completed = true;
    m_ai_panel.has_error = has_error;
    m_ai_panel.status_text =
        has_error ? L"操作失败。" : L"生成完成，点击确认回写。";
    status_hwnd = m_ai_panel.status_hwnd;
    confirm_hwnd = m_ai_panel.confirm_hwnd;
    cancel_hwnd = m_ai_panel.cancel_hwnd;
    has_output = !m_ai_panel.output_text.empty();
  }

  if (status_hwnd && IsWindow(status_hwnd)) {
    SetWindowTextW(status_hwnd, has_error ? L"操作失败。"
                                          : L"生成完成，点击确认回写。");
  }
  if (confirm_hwnd && IsWindow(confirm_hwnd)) {
    EnableWindow(confirm_hwnd, !has_error && has_output);
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
    } else if (IsBuiltinAIAssistantInstructionId(institution_id)) {
      m_ai_panel.selected_institution_id = institution_id;
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
    case kWMNcuahDrawCaption:
    case kWMNcuahDrawFrame:
    case WM_NCACTIVATE:
      return TRUE;
    case WM_NCPAINT:
      return 0;
    case WM_SETCURSOR:
      if (LOWORD(lParam) == HTCLIENT) {
        POINT cursor_point = {0, 0};
        if (GetCursorPos(&cursor_point)) {
          const int edges = HitTestAIPanelResizeEdges(hwnd, cursor_point);
          if ((edges & kAIPanelResizeEdgeRight) &&
              (edges & kAIPanelResizeEdgeBottom)) {
            SetCursor(LoadCursor(nullptr, IDC_SIZENWSE));
            return TRUE;
          }
          if (edges & kAIPanelResizeEdgeRight) {
            SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
            return TRUE;
          }
          if (edges & kAIPanelResizeEdgeBottom) {
            SetCursor(LoadCursor(nullptr, IDC_SIZENS));
            return TRUE;
          }
        }
      }
      break;
    case WM_GETMINMAXINFO: {
      MINMAXINFO* minmax_info = reinterpret_cast<MINMAXINFO*>(lParam);
      if (!minmax_info) {
        return 0;
      }
      const UINT dpi = ResolveAIPanelDpi(
          MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST), hwnd);
      minmax_info->ptMinTrackSize.x =
          ScaleAIPanelLogicalToPixels(kAIPanelMinWidth, dpi);
      minmax_info->ptMinTrackSize.y =
          ScaleAIPanelLogicalToPixels(kAIPanelMinHeight, dpi);
      minmax_info->ptMaxTrackSize.x =
          ScaleAIPanelLogicalToPixels(kAIPanelMaxWidth, dpi);
      minmax_info->ptMaxTrackSize.y =
          ScaleAIPanelLogicalToPixels(kAIPanelMaxHeight, dpi);
      return 0;
    }
    case WM_LBUTTONDOWN:
      if (self) {
        POINT cursor_point = {0, 0};
        if (GetCursorPos(&cursor_point)) {
          const int edges = HitTestAIPanelResizeEdges(hwnd, cursor_point);
          if (edges != kAIPanelResizeEdgeNone) {
            RECT window_rect = {0, 0, 0, 0};
            if (GetWindowRect(hwnd, &window_rect)) {
              std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
              if (self->m_ai_panel.panel_hwnd == hwnd) {
                self->m_ai_panel.resize_tracking = true;
                self->m_ai_panel.resize_edges = edges;
                self->m_ai_panel.resize_start_screen_x = cursor_point.x;
                self->m_ai_panel.resize_start_screen_y = cursor_point.y;
                self->m_ai_panel.resize_start_window_x = window_rect.left;
                self->m_ai_panel.resize_start_window_y = window_rect.top;
                self->m_ai_panel.resize_start_width =
                    window_rect.right - window_rect.left;
                self->m_ai_panel.resize_start_height =
                    window_rect.bottom - window_rect.top;
              }
            }
            SetCapture(hwnd);
            return 0;
          }
        }
      }
      break;
    case WM_MOUSEMOVE:
      if (self) {
        bool resize_tracking = false;
        int resize_edges = kAIPanelResizeEdgeNone;
        int start_screen_x = 0;
        int start_screen_y = 0;
        int start_window_x = 0;
        int start_window_y = 0;
        int start_width = 0;
        int start_height = 0;
        {
          std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
          if (self->m_ai_panel.panel_hwnd == hwnd) {
            resize_tracking = self->m_ai_panel.resize_tracking;
            resize_edges = self->m_ai_panel.resize_edges;
            start_screen_x = self->m_ai_panel.resize_start_screen_x;
            start_screen_y = self->m_ai_panel.resize_start_screen_y;
            start_window_x = self->m_ai_panel.resize_start_window_x;
            start_window_y = self->m_ai_panel.resize_start_window_y;
            start_width = self->m_ai_panel.resize_start_width;
            start_height = self->m_ai_panel.resize_start_height;
          }
        }
        if (resize_tracking) {
          POINT cursor_point = {0, 0};
          if (GetCursorPos(&cursor_point)) {
            int width = start_width;
            int height = start_height;
            if (resize_edges & kAIPanelResizeEdgeRight) {
              width = start_width + (cursor_point.x - start_screen_x);
            }
            if (resize_edges & kAIPanelResizeEdgeBottom) {
              height = start_height + (cursor_point.y - start_screen_y);
            }
            const UINT dpi = ResolveAIPanelDpi(
                MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST), hwnd);
            const int min_width =
                ScaleAIPanelLogicalToPixels(kAIPanelMinWidth, dpi);
            const int max_width =
                ScaleAIPanelLogicalToPixels(kAIPanelMaxWidth, dpi);
            const int min_height =
                ScaleAIPanelLogicalToPixels(kAIPanelMinHeight, dpi);
            const int max_height =
                ScaleAIPanelLogicalToPixels(kAIPanelMaxHeight, dpi);
            width = max(min_width, min(max_width, width));
            height = max(min_height, min(max_height, height));
            SetWindowPos(hwnd, HWND_TOPMOST, start_window_x, start_window_y,
                         width, height, SWP_NOACTIVATE);
            return 0;
          }
        }
      }
      break;
    case WM_CANCELMODE:
    case WM_LBUTTONUP:
      if (GetCapture() == hwnd) {
        ReleaseCapture();
      }
      break;
    case WM_CAPTURECHANGED:
      if (self) {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        if (self->m_ai_panel.panel_hwnd == hwnd) {
          self->m_ai_panel.resize_tracking = false;
          self->m_ai_panel.resize_edges = kAIPanelResizeEdgeNone;
        }
      }
      return 0;
    case WM_SIZE:
      if (self) {
        RECT window_rect = {0, 0, 0, 0};
        if (GetWindowRect(hwnd, &window_rect)) {
          const UINT dpi = ResolveAIPanelDpi(
              MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST), hwnd);
          const int width = max(
              kAIPanelMinWidth,
              min(kAIPanelMaxWidth,
                  ScaleAIPanelPixelsToLogical(
                      window_rect.right - window_rect.left, dpi)));
          const int height = max(
              kAIPanelMinHeight,
              min(kAIPanelMaxHeight,
                  ScaleAIPanelPixelsToLogical(
                      window_rect.bottom - window_rect.top, dpi)));
          std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
          if (self->m_ai_panel.panel_hwnd == hwnd) {
            self->m_ai_panel.panel_width = width;
            self->m_ai_panel.panel_height = height;
            self->m_ai_panel.last_panel_x = window_rect.left;
            self->m_ai_panel.last_panel_y = window_rect.top;
            self->m_ai_panel.has_last_panel_position = true;
          }
        }
      }
      ApplyAIPanelRoundedRegion(hwnd);
      LayoutAIPanelControls(hwnd);
      if (self) {
        self->_ResizeAIPanelWebView();
      }
      return 0;
    case WM_DPICHANGED: {
      const RECT* suggested_rect = reinterpret_cast<const RECT*>(lParam);
      if (!suggested_rect) {
        return 0;
      }
      SetWindowPos(hwnd, HWND_TOPMOST, suggested_rect->left, suggested_rect->top,
                   suggested_rect->right - suggested_rect->left,
                   suggested_rect->bottom - suggested_rect->top,
                   SWP_NOACTIVATE);
      return 0;
    }
    case WM_COMMAND:
      if (!self) {
        break;
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
    case WM_ACTIVATE:
      if (self && LOWORD(wParam) == WA_INACTIVE) {
        HWND next_hwnd = reinterpret_cast<HWND>(lParam);
        if (next_hwnd != hwnd && !IsChild(hwnd, next_hwnd)) {
          bool should_close = false;
          {
            std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
            should_close = self->m_ai_panel.panel_hwnd == hwnd &&
                           IsWindowVisible(hwnd) &&
                           !self->m_ai_panel.resize_tracking;
          }
          if (should_close) {
            self->_CloseAIPanel();
          }
        }
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
      return DefWindowProcW(hwnd, msg, wParam, lParam);
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
      std::wstring selected_institution_id;
      AIPanelInstitutionOption selected_institution_option;
      bool has_selected_institution_option = false;
      bool loaded_from_state_file = false;
      {
        std::string file_token;
        std::string file_tenant_id;
        std::string file_refresh_token;
        loaded_from_state_file = LoadAIAssistantLoginIdentity(
            self->m_ai_config, &file_token, &file_tenant_id,
            &file_refresh_token);
        if (loaded_from_state_file) {
          token = file_token;
          tenant_id = file_tenant_id;
          refresh_token = file_refresh_token;
          std::lock_guard<std::mutex> lock(self->m_ai_login_mutex);
          if (!file_token.empty()) {
            self->m_ai_login_token = file_token;
          }
          if (!file_tenant_id.empty()) {
            self->m_ai_login_tenant_id = file_tenant_id;
          }
          if (!file_refresh_token.empty()) {
            self->m_ai_login_refresh_token = file_refresh_token;
          }
        }
      }
      if (!loaded_from_state_file) {
        std::lock_guard<std::mutex> lock(self->m_ai_login_mutex);
        token = self->m_ai_login_token;
        tenant_id = self->m_ai_login_tenant_id;
        refresh_token = self->m_ai_login_refresh_token;
      }
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        selected_institution_id = self->m_ai_panel.selected_institution_id;
        for (const auto& option : self->m_ai_panel.institution_options) {
          if (option.id != selected_institution_id) {
            continue;
          }
          selected_institution_option = option;
          has_selected_institution_option = true;
          break;
        }
      }
      panel_url = AppendAIPanelAuthToUrl(panel_url, token, tenant_id,
                                         refresh_token,
                                         selected_institution_id,
                                         has_selected_institution_option
                                             ? &selected_institution_option
                                             : nullptr);

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
                                    // Pass Escape to frontend for state navigation (select→input→generating)
                                    // Frontend will handle onBack() or onClose()
                                    if (is_key_down &&
                                        virtual_key == static_cast<UINT>(VK_ESCAPE)) {
                                      // Let frontend handle it; don't intercept
                                      args->put_Handled(FALSE);
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
                                      PostMessageW(hwnd, WM_AI_WEBVIEW_FOCUS, 1,
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
    case WM_AI_WEBVIEW_FOCUS: {
#if WEASEL_HAS_WEBVIEW2
      if (!self) {
        return 0;
      }
      const bool should_focus_webview = wParam != 0;

      bool focus_pending = false;
      ICoreWebView2Controller* controller = nullptr;
      {
        std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
        focus_pending = self->m_ai_panel.focus_pending;
        if (focus_pending && should_focus_webview) {
          controller = static_cast<ICoreWebView2Controller*>(
              self->m_ai_panel.webview_controller);
          if (controller) {
            controller->AddRef();
          }
        }
      }
      if (!focus_pending) {
        if (controller) {
          controller->Release();
        }
        return 0;
      }

      ShowWindow(hwnd, SW_SHOWNORMAL);
      HWND fg = GetForegroundWindow();
      const DWORD fg_tid = fg ? GetWindowThreadProcessId(fg, nullptr) : 0;
      const DWORD cur_tid = GetCurrentThreadId();
      if (fg_tid != 0 && fg_tid != cur_tid) {
        AttachThreadInput(cur_tid, fg_tid, TRUE);
        SetForegroundWindow(hwnd);
        SetActiveWindow(hwnd);
        SetFocus(hwnd);
        AttachThreadInput(cur_tid, fg_tid, FALSE);
      } else {
        SetForegroundWindow(hwnd);
        SetActiveWindow(hwnd);
        SetFocus(hwnd);
      }

      if (controller) {
        const HRESULT hr =
            controller->MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC);
        controller->Release();
        if (SUCCEEDED(hr)) {
          std::lock_guard<std::mutex> lock(self->m_ai_panel_mutex);
          self->m_ai_panel.focus_pending = false;
        }
      }
#endif
      return 0;
    }
    default:
      break;
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}
