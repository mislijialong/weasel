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
#include <AIAssistantInstructions.h>
#include <RimeWithWeasel.h>
#include <StringAlgorithm.hpp>
#include <WeaselConstants.h>
#include <WeaselUtility.h>
#include <winhttp.h>
#include <shellapi.h>
#include <ShellScalingApi.h>

#include <algorithm>
#include <cctype>
#include <cstring>
#include <cstdlib>
#include <cwctype>
#include <array>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <map>
#include <memory>
#include <random>
#include <regex>
#include <set>
#include <thread>
#include <unordered_map>
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
#pragma comment(lib, "Shcore.lib")

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
constexpr UINT WM_AI_WEBVIEW_SYNC = WM_APP + 406;
constexpr UINT WM_AI_PANEL_DRAG = WM_APP + 407;
constexpr UINT WM_AI_WEBVIEW_FOCUS = WM_APP + 408;
constexpr UINT kWMNcuahDrawCaption = 0x00AE;
constexpr UINT kWMNcuahDrawFrame = 0x00AF;
constexpr int kAIPanelResizeEdgeNone = 0;
constexpr int kAIPanelResizeEdgeRight = 1;
constexpr int kAIPanelResizeEdgeBottom = 2;
constexpr wchar_t kSystemCommandCommitPrefix[] = L"__weasel_syscmd__:";
constexpr char kDefaultSystemCommandInputPrefix[] = "sc";
constexpr char kDefaultInstructionLookupInputPrefix[] = "sS";
constexpr char kDefaultAIAssistantInstructionChangedTopic[] =
    "/mqtt/topic/sino/langwell/ins/ins/changed/+";
constexpr char kAIAssistantPermissionUpdateTopicPrefixObjectKey[] =
    "mqtt-topic-prefix";
constexpr char kAIAssistantPermissionUpdateTopicPrefixMemberKey[] =
    "app-ins-upd";
constexpr char kDefaultAIAssistantPermissionUpdateTopicPrefix[] =
    "/mqtt/topic/sino/lamp/app:ins/permission/update";
constexpr const char* kAllowedSystemCommandIds[] = {
    "jsq",       "calc",     "notepad",   "mspaint", "explorer",
    "txt",       "md",       "gh",        "bd",      "wb",
    "g",         "yt",       "rili",      "calendar", "cal", "kb"};

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

bool IsSystemCommandInputText(const std::string& input_text) {
  if (input_text.size() <= std::strlen(kDefaultSystemCommandInputPrefix) ||
      input_text.compare(0, std::strlen(kDefaultSystemCommandInputPrefix),
                         kDefaultSystemCommandInputPrefix) != 0) {
    return false;
  }
  const std::string command_id =
      input_text.substr(std::strlen(kDefaultSystemCommandInputPrefix));
  return IsAllowedSystemCommandId(command_id);
}

bool IsValidInstructionLookupPrefix(const std::string& prefix) {
  if (prefix.empty()) {
    return false;
  }
  for (char ch : prefix) {
    const unsigned char uch = static_cast<unsigned char>(ch);
    if ((uch < 'A' || uch > 'Z') && (uch < 'a' || uch > 'z')) {
      return false;
    }
  }
  return true;
}

bool IsInstructionLookupInputText(const std::string& input_text,
                                  const std::string& prefix) {
  const size_t prefix_length = prefix.size();
  if (prefix_length == 0) {
    return false;
  }
  if (input_text.size() <= prefix_length ||
      input_text.compare(0, prefix_length, prefix) != 0) {
    return false;
  }
  for (size_t i = prefix_length; i < input_text.size(); ++i) {
    const unsigned char ch = static_cast<unsigned char>(input_text[i]);
    if (ch < 'a' || ch > 'z') {
      return false;
    }
  }
  return true;
}

std::string ExtractInstructionLookupQuery(const std::string& input_text,
                                          const std::string& prefix) {
  if (!IsInstructionLookupInputText(input_text, prefix)) {
    return std::string();
  }
  return input_text.substr(prefix.size());
}

std::string NormalizeInstructionLookupAscii(const std::string& text) {
  std::string normalized;
  normalized.reserve(text.size());
  for (char ch : text) {
    const unsigned char uch = static_cast<unsigned char>(ch);
    if (uch >= 'A' && uch <= 'Z') {
      normalized.push_back(static_cast<char>(uch - 'A' + 'a'));
    } else if ((uch >= 'a' && uch <= 'z') || (uch >= '0' && uch <= '9')) {
      normalized.push_back(static_cast<char>(uch));
    }
  }
  return normalized;
}

std::string NormalizeInstructionLookupAscii(const std::wstring& text) {
  std::string normalized;
  normalized.reserve(text.size());
  for (wchar_t ch : text) {
    if (ch >= L'A' && ch <= L'Z') {
      normalized.push_back(static_cast<char>(ch - L'A' + 'a'));
    } else if ((ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9')) {
      normalized.push_back(static_cast<char>(ch));
    }
  }
  return normalized;
}

std::string ReadRimeDataDirectory(bool shared) {
  if (!rime_api) {
    return std::string();
  }

  std::array<char, 4096> buffer = {};
  const auto copy_data_dir = [&buffer](const char* data_dir) {
    if (!data_dir) {
      return;
    }
    const size_t length = (std::min)(std::strlen(data_dir), buffer.size() - 1);
    std::memcpy(buffer.data(), data_dir, length);
    buffer[length] = '\0';
  };
  if (shared) {
    if (rime_api->get_shared_data_dir_s) {
      rime_api->get_shared_data_dir_s(buffer.data(), buffer.size());
    } else if (rime_api->get_shared_data_dir) {
      copy_data_dir(rime_api->get_shared_data_dir());
    }
  } else {
    if (rime_api->get_user_data_dir_s) {
      rime_api->get_user_data_dir_s(buffer.data(), buffer.size());
    } else if (rime_api->get_user_data_dir) {
      copy_data_dir(rime_api->get_user_data_dir());
    }
  }
  return std::string(buffer.data());
}

void AddInstructionLookupAlias(const std::string& alias,
                               std::vector<std::string>* aliases,
                               std::set<std::string>* seen) {
  if (!aliases || !seen) {
    return;
  }
  const std::string normalized = NormalizeInstructionLookupAscii(alias);
  if (normalized.empty() || seen->find(normalized) != seen->end()) {
    return;
  }
  seen->insert(normalized);
  aliases->push_back(normalized);
}

std::vector<std::string> TokenizeInstructionLookupCodes(const std::string& text) {
  std::vector<std::string> tokens;
  std::string token;
  token.reserve(text.size());
  for (char ch : text) {
    const unsigned char uch = static_cast<unsigned char>(ch);
    if ((uch >= 'A' && uch <= 'Z') || (uch >= 'a' && uch <= 'z') ||
        (uch >= '0' && uch <= '9')) {
      token.push_back(static_cast<char>(std::tolower(uch)));
      continue;
    }
    if (!token.empty()) {
      tokens.push_back(token);
      token.clear();
    }
  }
  if (!token.empty()) {
    tokens.push_back(token);
  }
  return tokens;
}

void AddInstructionLookupAliasesFromCodeText(
    const std::string& text,
    std::vector<std::string>* aliases,
    std::set<std::string>* seen) {
  AddInstructionLookupAlias(text, aliases, seen);
  for (const auto& token : TokenizeInstructionLookupCodes(text)) {
    AddInstructionLookupAlias(token, aliases, seen);
  }
}

void AddInstructionLookupAliasesFromCodeText(
    const std::wstring& text,
    std::vector<std::string>* aliases,
    std::set<std::string>* seen) {
  AddInstructionLookupAliasesFromCodeText(wtou8(text), aliases, seen);
}

std::string BuildInstructionLookupFullAliasFromCode(const std::string& code) {
  return NormalizeInstructionLookupAscii(code);
}

std::string BuildInstructionLookupInitialAliasFromCode(const std::string& code) {
  std::string initials;
  for (const auto& token : TokenizeInstructionLookupCodes(code)) {
    if (!token.empty()) {
      initials.push_back(token.front());
    }
  }
  return NormalizeInstructionLookupAscii(initials);
}

char GetInstructionLookupPinyinInitialFallback(wchar_t ch) {
  if ((ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') ||
      (ch >= L'0' && ch <= L'9')) {
    return static_cast<char>(std::towlower(ch));
  }

  char gbk[4] = {};
  const int converted = WideCharToMultiByte(936, 0, &ch, 1, gbk,
                                            static_cast<int>(sizeof(gbk)),
                                            nullptr, nullptr);
  if (converted < 2) {
    return '\0';
  }

  const int code = (static_cast<unsigned char>(gbk[0]) << 8) |
                   static_cast<unsigned char>(gbk[1]);
  struct GbkInitialRange {
    int lower_bound;
    int upper_bound;
    char initial;
  };
  static const GbkInitialRange kRanges[] = {
      {45217, 45252, 'a'}, {45253, 45760, 'b'}, {45761, 46317, 'c'},
      {46318, 46825, 'd'}, {46826, 47009, 'e'}, {47010, 47296, 'f'},
      {47297, 47613, 'g'}, {47614, 48118, 'h'}, {48119, 49061, 'j'},
      {49062, 49323, 'k'}, {49324, 49895, 'l'}, {49896, 50370, 'm'},
      {50371, 50613, 'n'}, {50614, 50621, 'o'}, {50622, 50905, 'p'},
      {50906, 51386, 'q'}, {51387, 51445, 'r'}, {51446, 52217, 's'},
      {52218, 52697, 't'}, {52698, 52979, 'w'}, {52980, 53688, 'x'},
      {53689, 54480, 'y'}, {54481, 55289, 'z'}};
  for (const auto& range : kRanges) {
    if (code >= range.lower_bound && code <= range.upper_bound) {
      return range.initial;
    }
  }
  return '\0';
}

class InstructionLookupPinyinCache {
 public:
  bool TryGetPhraseCode(const std::wstring& phrase, std::string* code) {
    if (!code || phrase.empty()) {
      return false;
    }
    EnsureLoaded();
    std::lock_guard<std::mutex> lock(mutex_);
    const auto it = phrase_codes_.find(phrase);
    if (it == phrase_codes_.end()) {
      return false;
    }
    *code = it->second;
    return true;
  }

  bool TryGetCharCode(wchar_t ch, std::string* code) {
    if (!code || ch == 0) {
      return false;
    }
    EnsureLoaded();
    std::lock_guard<std::mutex> lock(mutex_);
    const auto it = char_codes_.find(ch);
    if (it == char_codes_.end()) {
      return false;
    }
    *code = it->second;
    return true;
  }

  size_t max_phrase_length() {
    EnsureLoaded();
    std::lock_guard<std::mutex> lock(mutex_);
    return max_phrase_length_;
  }

 private:
  void EnsureLoaded() {
    std::lock_guard<std::mutex> lock(mutex_);
    if (load_attempted_) {
      return;
    }
    load_attempted_ = true;

    std::vector<fs::path> candidate_paths;
    const std::string shared_dir = ReadRimeDataDirectory(true);
    if (!shared_dir.empty()) {
      candidate_paths.push_back(
          fs::u8path(shared_dir) / "luna_pinyin.dict.yaml");
    }
    const std::string user_dir = ReadRimeDataDirectory(false);
    if (!user_dir.empty()) {
      candidate_paths.push_back(fs::u8path(user_dir) / "luna_pinyin.dict.yaml");
    }
    candidate_paths.push_back(
        fs::current_path() / "output" / "data" / "luna_pinyin.dict.yaml");

    for (const auto& path : candidate_paths) {
      if (LoadFromPath(path)) {
        break;
      }
    }
  }

  bool LoadFromPath(const fs::path& path) {
    std::error_code ec;
    if (!fs::exists(path, ec) || ec) {
      return false;
    }

    std::ifstream stream(path, std::ios::binary);
    if (!stream) {
      return false;
    }

    bool in_entries = false;
    std::string line;
    while (std::getline(stream, line)) {
      if (!line.empty() && line.back() == '\r') {
        line.pop_back();
      }
      if (!in_entries) {
        if (line == "...") {
          in_entries = true;
        }
        continue;
      }
      if (line.empty() || line[0] == '#') {
        continue;
      }

      const size_t first_tab = line.find('\t');
      if (first_tab == std::string::npos || first_tab == 0) {
        continue;
      }
      const size_t second_tab = line.find('\t', first_tab + 1);
      const std::string phrase_u8 = line.substr(0, first_tab);
      const std::string code = line.substr(
          first_tab + 1,
          second_tab == std::string::npos ? std::string::npos
                                          : second_tab - first_tab - 1);
      if (code.empty()) {
        continue;
      }

      const std::wstring phrase = u8tow(phrase_u8);
      if (phrase.empty()) {
        continue;
      }
      phrase_codes_.emplace(phrase, code);
      max_phrase_length_ = max(max_phrase_length_, phrase.size());
      if (phrase.size() == 1) {
        char_codes_.emplace(phrase.front(), code);
      }
    }

    return !phrase_codes_.empty() || !char_codes_.empty();
  }

  std::mutex mutex_;
  bool load_attempted_ = false;
  size_t max_phrase_length_ = 1;
  std::unordered_map<std::wstring, std::string> phrase_codes_;
  std::unordered_map<wchar_t, std::string> char_codes_;
};

InstructionLookupPinyinCache& GetInstructionLookupPinyinCache() {
  static InstructionLookupPinyinCache cache;
  return cache;
}

bool TryBuildInstructionLookupPinyinAliases(const std::wstring& text,
                                            std::string* full_alias,
                                            std::string* initial_alias) {
  if (!full_alias || !initial_alias) {
    return false;
  }

  full_alias->clear();
  initial_alias->clear();
  if (text.empty()) {
    return false;
  }

  InstructionLookupPinyinCache& cache = GetInstructionLookupPinyinCache();
  std::string exact_code;
  if (cache.TryGetPhraseCode(text, &exact_code)) {
    *full_alias = BuildInstructionLookupFullAliasFromCode(exact_code);
    *initial_alias = BuildInstructionLookupInitialAliasFromCode(exact_code);
    return !full_alias->empty() || !initial_alias->empty();
  }

  const size_t max_phrase_length = cache.max_phrase_length();
  for (size_t index = 0; index < text.size();) {
    const wchar_t ch = text[index];
    if (ch == 0 || std::iswspace(static_cast<wint_t>(ch))) {
      ++index;
      continue;
    }
    if ((ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z') ||
        (ch >= L'0' && ch <= L'9')) {
      const char lower_ch =
          static_cast<char>((ch >= L'A' && ch <= L'Z') ? (ch - L'A' + 'a') : ch);
      full_alias->push_back(lower_ch);
      initial_alias->push_back(lower_ch);
      ++index;
      continue;
    }

    bool matched_phrase = false;
    for (size_t len = (std::min)(max_phrase_length, text.size() - index);
         len > 1; --len) {
      std::string phrase_code;
      if (!cache.TryGetPhraseCode(text.substr(index, len), &phrase_code)) {
        continue;
      }
      full_alias->append(BuildInstructionLookupFullAliasFromCode(phrase_code));
      initial_alias->append(
          BuildInstructionLookupInitialAliasFromCode(phrase_code));
      index += len;
      matched_phrase = true;
      break;
    }
    if (matched_phrase) {
      continue;
    }

    std::string char_code;
    if (cache.TryGetCharCode(ch, &char_code)) {
      full_alias->append(BuildInstructionLookupFullAliasFromCode(char_code));
      initial_alias->append(
          BuildInstructionLookupInitialAliasFromCode(char_code));
      ++index;
      continue;
    }

    const char fallback_initial = GetInstructionLookupPinyinInitialFallback(ch);
    if (fallback_initial != '\0') {
      full_alias->push_back(fallback_initial);
      initial_alias->push_back(fallback_initial);
    }
    ++index;
  }

  *full_alias = NormalizeInstructionLookupAscii(*full_alias);
  *initial_alias = NormalizeInstructionLookupAscii(*initial_alias);
  return !full_alias->empty() || !initial_alias->empty();
}

int ScoreInstructionLookupAlias(const std::string& query,
                                const std::string& alias) {
  if (query.empty() || alias.empty()) {
    return 0;
  }
  if (alias == query) {
    return 10000 - static_cast<int>(alias.size());
  }
  if (alias.size() >= query.size() &&
      alias.compare(0, query.size(), query) == 0) {
    return 8000 - static_cast<int>(alias.size() - query.size());
  }
  const size_t pos = alias.find(query);
  if (pos != std::string::npos) {
    return 4000 - static_cast<int>(pos);
  }
  return 0;
}

struct InstructionLookupRankedOption {
  AIPanelInstitutionOption option;
  int score = 0;
  size_t order = 0;
};

size_t GetInjectedCandidateVisibleOffset(
    const AIAssistantInjectedCandidateState& state) {
  if (!state.instruction_lookup_mode || state.page_size == 0) {
    return 0;
  }
  return state.current_page * state.page_size;
}

size_t GetInjectedCandidateVisibleCount(
    const AIAssistantInjectedCandidateState& state) {
  const size_t visible_offset = GetInjectedCandidateVisibleOffset(state);
  if (visible_offset >= state.options.size()) {
    return 0;
  }
  const size_t page_size =
      state.instruction_lookup_mode && state.page_size > 0
          ? state.page_size
          : state.options.size();
  const size_t remaining = state.options.size() - visible_offset;
  return (std::min)(page_size, remaining);
}

size_t ResolveInjectedCandidateIndex(
    const AIAssistantInjectedCandidateState& state,
    size_t visible_index) {
  return GetInjectedCandidateVisibleOffset(state) + visible_index;
}

bool IsSystemCommandConfirmKey(const KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  return key_event.keycode == ibus::space ||
         key_event.keycode == ibus::Return ||
         key_event.keycode == ibus::KP_Enter;
}

bool IsInlineInstructionSubmitKey(const KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  return key_event.keycode == ibus::Return ||
         key_event.keycode == ibus::KP_Enter;
}

bool IsInlineInstructionCancelKey(const KeyEvent& key_event) {
  return !(key_event.mask & ibus::RELEASE_MASK) &&
         key_event.keycode == ibus::Keycode::Escape;
}

bool IsInlineInstructionBackspaceKey(const KeyEvent& key_event) {
  return !(key_event.mask & ibus::RELEASE_MASK) &&
         key_event.keycode == ibus::BackSpace;
}

bool IsInlineInstructionLeftKey(const KeyEvent& key_event) {
  return !(key_event.mask & ibus::RELEASE_MASK) &&
         key_event.keycode == ibus::Left;
}

bool IsInlineInstructionRightKey(const KeyEvent& key_event) {
  return !(key_event.mask & ibus::RELEASE_MASK) &&
         key_event.keycode == ibus::Right;
}

bool IsPlainInlineInstructionPrefixKey(const KeyEvent& key_event,
                                       wchar_t prefix_char) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  const auto modifiers = key_event.mask & ibus::MODIFIER_MASK;
  if (modifiers & (ibus::CONTROL_MASK | ibus::ALT_MASK | ibus::SUPER_MASK |
                   ibus::HYPER_MASK | ibus::META_MASK)) {
    return false;
  }
  return key_event.keycode == static_cast<ibus::Keycode>(prefix_char);
}

std::wstring CaptureCurrentPreeditText(RimeSessionId session_id);

struct InlineInstructionPromptSnapshot {
  std::wstring text;
  size_t cursor_pos = 0;
};

InlineInstructionPromptSnapshot CaptureInlineInstructionPromptSnapshot(
    RimeSessionId session_id);

std::wstring CaptureInlineInstructionPromptSource(RimeSessionId session_id);

std::wstring TrimInlineInstructionPrompt(const std::wstring& text);

namespace {
bool GetInlineInstructionSnapshot(
    std::mutex& mutex,
    std::map<WeaselSessionId, InlineInstructionState>& states,
    WeaselSessionId ipc_id,
    InlineInstructionState* state_out) {
  if (!state_out || ipc_id == 0) {
    return false;
  }
  std::lock_guard<std::mutex> lock(mutex);
  const auto it = states.find(ipc_id);
  if (it == states.end()) {
    return false;
  }
  *state_out = it->second;
  return true;
}

bool HasActiveInlineInstructionState(
    std::mutex& mutex,
    std::map<WeaselSessionId, InlineInstructionState>& states,
    WeaselSessionId ipc_id) {
  InlineInstructionState state;
  return GetInlineInstructionSnapshot(mutex, states, ipc_id, &state) &&
         state.IsActive();
}

bool ShouldSuppressSpecialInstructionCandidates(
    std::mutex& inline_mutex,
    std::map<WeaselSessionId, InlineInstructionState>& inline_states,
    WeaselSessionId ipc_id,
    const std::string& instruction_lookup_prefix,
    const char* input_text) {
  if (!input_text) {
    return false;
  }
  if (!HasActiveInlineInstructionState(inline_mutex, inline_states, ipc_id)) {
    return false;
  }
  return IsInstructionLookupInputText(input_text, instruction_lookup_prefix) ||
         IsSystemCommandInputText(input_text);
}
}  // namespace

bool SplitInlineInstructionText(const std::wstring& text,
                                wchar_t prefix_char,
                                std::wstring* before_prefix,
                                std::wstring* prompt_text) {
  if (!before_prefix || !prompt_text) {
    return false;
  }
  const size_t prefix_pos = text.find(prefix_char);
  if (prefix_pos == std::wstring::npos) {
    return false;
  }
  *before_prefix = text.substr(0, prefix_pos);
  *prompt_text = TrimInlineInstructionPrompt(text.substr(prefix_pos + 1));
  return true;
}

std::wstring TrimInlineInstructionPrompt(const std::wstring& text) {
  size_t begin = 0;
  while (begin < text.size() &&
         std::iswspace(static_cast<wint_t>(text[begin]))) {
    ++begin;
  }
  size_t end = text.size();
  while (end > begin && std::iswspace(static_cast<wint_t>(text[end - 1]))) {
    --end;
  }
  return text.substr(begin, end - begin);
}

bool ExtractInlineInstructionPrompt(const std::wstring& preedit_text,
                                    wchar_t prefix_char,
                                    std::wstring* before_prefix,
                                    std::wstring* prompt_text) {
  if (!SplitInlineInstructionText(preedit_text, prefix_char, before_prefix,
                                  prompt_text) ||
      prompt_text->empty()) {
    return false;
  }
  return true;
}

std::wstring BuildInlineInstructionDisplayPrefixText(
    wchar_t prefix_char,
    const InlineInstructionState& state) {
  std::wstring display_prefix = state.pending_commit_text;
  if (state.has_selected_option && !state.selected_option.name.empty()) {
    display_prefix.append(L"【")
        .append(state.selected_option.name)
        .append(L"】，");
    return display_prefix;
  }
  display_prefix.push_back(prefix_char);
  return display_prefix;
}

std::wstring BuildInlineInstructionStatusText(
    wchar_t prefix_char,
    const InlineInstructionState& state) {
  return BuildInlineInstructionDisplayPrefixText(prefix_char, state) +
         state.prompt_text + L" [AI处理中...]";
}

std::wstring BuildInlineInstructionEditingText(
    wchar_t prefix_char,
    const InlineInstructionState& state) {
  return BuildInlineInstructionDisplayPrefixText(prefix_char, state) +
         state.prompt_text;
}

size_t ClampInlineInstructionCursorOffset(size_t cursor_offset,
                                          const std::wstring& prompt_text) {
  return min(cursor_offset, prompt_text.size());
}

size_t GetInlineInstructionPromptCursorPosition(
    const InlineInstructionState& state) {
  const size_t cursor_offset =
      ClampInlineInstructionCursorOffset(state.cursor_offset, state.prompt_text);
  return state.prompt_text.size() - cursor_offset;
}

size_t GetInlineInstructionDisplayCursorPosition(
    wchar_t prefix_char,
    const InlineInstructionState& state,
    const std::wstring& preedit_text) {
  if (state.phase != InlineInstructionPhase::kEditing) {
    return preedit_text.size();
  }
  const size_t display_prefix_len =
      BuildInlineInstructionDisplayPrefixText(prefix_char, state).size();
  return min(display_prefix_len + GetInlineInstructionPromptCursorPosition(state),
             preedit_text.size());
}

void AddInlineInstructionCursorAttribute(weasel::Text* text,
                                         size_t cursor_pos) {
  if (!text) {
    return;
  }
  weasel::TextAttribute attr;
  attr.type = HIGHLIGHTED;
  attr.range.start = static_cast<int>(cursor_pos);
  attr.range.end = static_cast<int>(cursor_pos);
  attr.range.cursor = static_cast<int>(cursor_pos);
  text->attributes.push_back(attr);
}

std::wstring BuildInlineInstructionActivePromptText(
    RimeSessionId session_id,
    const InlineInstructionState& state) {
  if (state.phase != InlineInstructionPhase::kEditing) {
    return state.prompt_text;
  }
  const std::wstring live_text = CaptureInlineInstructionPromptSource(session_id);
  return state.committed_prompt_text + live_text;
}

bool TryGetInlineInstructionSessionText(RimeSessionId session_id,
                                        wchar_t prefix_char,
                                        std::wstring* before_prefix,
                                        std::wstring* prompt_text,
                                        bool* has_prefix) {
  if (!before_prefix || !prompt_text || !has_prefix) {
    return false;
  }

  *before_prefix = std::wstring();
  *prompt_text = std::wstring();
  *has_prefix = false;

  const std::wstring preedit_text = CaptureCurrentPreeditText(session_id);
  if (SplitInlineInstructionText(preedit_text, prefix_char, before_prefix,
                                 prompt_text)) {
    *has_prefix = true;
    return true;
  }

  const char* input_text = rime_api->get_input(session_id);
  const std::wstring input_w = input_text ? u8tow(input_text) : std::wstring();
  if (SplitInlineInstructionText(input_w, prefix_char, before_prefix,
                                 prompt_text)) {
    *has_prefix = true;
    return true;
  }

  *has_prefix = preedit_text.find(prefix_char) != std::wstring::npos ||
                input_w.find(prefix_char) != std::wstring::npos;
  return false;
}

bool IsSystemCommandCandidate(const RimeCandidate& candidate) {
  const std::string text = candidate.text ? candidate.text : "";
  const std::string comment = candidate.comment ? candidate.comment : "";
  return IsSystemCommandInputText(text) ||
         (text.size() > std::strlen(kDefaultSystemCommandInputPrefix) &&
          text.compare(0, std::strlen(kDefaultSystemCommandInputPrefix),
                       kDefaultSystemCommandInputPrefix) == 0 &&
          (comment.find("空格/回车执行") != std::string::npos ||
           comment.find("打开") != std::string::npos));
}

bool HasSelectedSystemCommandCandidate(RimeSessionId session_id) {
  RIME_STRUCT(RimeContext, ctx);
  if (!rime_api->get_context(session_id, &ctx)) {
    return false;
  }
  bool matched = false;
  if (ctx.menu.num_candidates > 0 && ctx.menu.candidates) {
    const int highlighted = max(
        0, min(ctx.menu.num_candidates - 1,
               ctx.menu.highlighted_candidate_index));
    matched = IsSystemCommandCandidate(ctx.menu.candidates[highlighted]);
  }
  rime_api->free_context(&ctx);
  return matched;
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

std::wstring BuildInlineInstructionMockResult(
    const AIPanelInstitutionOption* option,
    const std::wstring& prompt_text) {
  const std::wstring instruction_name =
      option && !option->name.empty() ? option->name : L"AI指令";
  const std::wstring normalized_prompt = TrimInlineInstructionPrompt(prompt_text);
  const std::wstring display_prompt =
      normalized_prompt.empty() ? L"（空内容）" : normalized_prompt;
  if (instruction_name.find(L"续写") != std::wstring::npos) {
    return L"【模拟返回】继续围绕“" + display_prompt +
           L"”补一段内容：为了让语气更自然，可以顺着上一句继续展开，补充一两句细节，再把结尾收住。";
  }
  if (instruction_name.find(L"翻译") != std::wstring::npos) {
    return L"【模拟返回】这是“" + display_prompt +
           L"”的示例翻译结果，你后面接好真实接口后，这里会替换成真实返回内容。";
  }
  return L"【模拟返回】已触发「" + instruction_name + L"」，收到内容："
         + display_prompt + L"。这里先返回一段 mock 文本用于联调。";
}

std::string BuildInlineInstructionChatRequestBody(
    const std::wstring& query_text,
    const std::string& token,
    const std::string& tenant_id,
    const std::string& user_id) {
  const std::string query_utf8 = wtou8(TrimInlineInstructionPrompt(query_text));
  const std::string normalized_user_id =
      user_id.empty() ? "anonymous" : user_id;
  std::string body;
  body.reserve(512 + query_utf8.size() * 4);
  body += "{\"inputs\":{";
  body += "\"personalLibs\":\"\",";
  body += "\"knGroupIds\":\"\",";
  body += "\"knIds\":\"\",";
  body += "\"Token\":\"";
  body += EscapeJsonString(token);
  body += "\",\"tenantid\":\"";
  body += EscapeJsonString(tenant_id);
  body += "\",\"query\":\"";
  body += EscapeJsonString(query_utf8);
  body += "\"},\"files\":[],\"user\":\"";
  body += EscapeJsonString(normalized_user_id);
  body += "\",\"query\":\"";
  body += EscapeJsonString(query_utf8);
  body += "\",\"response_mode\":\"streaming\",\"conversation_id\":\"\"}";
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

bool InvokeInlineInstructionChatStream(
    const AIAssistantConfig& config,
    const AIPanelInstitutionOption& option,
    const std::wstring& query_text,
    const std::string& token,
    const std::string& tenant_id,
    const std::string& user_id,
    const std::function<void(const std::wstring&)>& on_delta,
    std::string* error_message,
    int* http_status_code = nullptr) {
  if (http_status_code) {
    *http_status_code = 0;
  }
  if (option.app_key.empty()) {
    if (error_message) {
      *error_message = "Inline AI instruction appKey is empty.";
    }
    return false;
  }

  const auto is_auth_failure_code = [](int code) -> bool {
    switch (code) {
      case 401:
      case 11001:
      case 11011:
      case 11012:
      case 11013:
      case 11014:
      case 11015:
        return true;
      default:
        return false;
    }
  };

  const auto read_response_metadata = [](const std::string& response_body,
                                         int* code,
                                         std::string* message) {
    if (code) {
      *code = 200;
    }
    if (message) {
      message->clear();
    }

    rapidjson::Document document;
    document.Parse(response_body.c_str(), response_body.size());
    if (document.HasParseError() || !document.IsObject()) {
      return;
    }

    const auto code_it = document.FindMember("code");
    if (code_it != document.MemberEnd()) {
      int parsed_code = 200;
      if (code_it->value.IsInt()) {
        parsed_code = code_it->value.GetInt();
      } else if (code_it->value.IsUint()) {
        parsed_code = static_cast<int>(code_it->value.GetUint());
      } else if (code_it->value.IsInt64()) {
        parsed_code = static_cast<int>(code_it->value.GetInt64());
      } else if (code_it->value.IsUint64()) {
        parsed_code = static_cast<int>(code_it->value.GetUint64());
      } else if (code_it->value.IsString()) {
        parsed_code = std::atoi(code_it->value.GetString());
      }
      if (code) {
        *code = parsed_code;
      }
    }

    if (!message) {
      return;
    }

    const char* message_keys[] = {"msg", "message", "errorMsg", "error",
                                  "detail"};
    for (const char* key : message_keys) {
      const auto message_it = document.FindMember(key);
      if (message_it == document.MemberEnd()) {
        continue;
      }
      if (message_it->value.IsString()) {
        *message = message_it->value.GetString();
        return;
      }
      if (message_it->value.IsObject()) {
        const auto nested_message =
            message_it->value.FindMember("message");
        if (nested_message != message_it->value.MemberEnd() &&
            nested_message->value.IsString()) {
          *message = nested_message->value.GetString();
          return;
        }
      }
    }
  };

  std::wstring api_base = u8tow(config.ai_api_base);
  if (api_base.empty()) {
    if (error_message) {
      *error_message = "Inline AI api base is empty.";
    }
    return false;
  }
  while (!api_base.empty() && api_base.back() == L'/') {
    api_base.pop_back();
  }
  const std::wstring endpoint = api_base + L"/chat-messages";
  URL_COMPONENTS parts = {0};
  parts.dwStructSize = sizeof(parts);
  parts.dwSchemeLength = static_cast<DWORD>(-1);
  parts.dwHostNameLength = static_cast<DWORD>(-1);
  parts.dwUrlPathLength = static_cast<DWORD>(-1);
  parts.dwExtraInfoLength = static_cast<DWORD>(-1);
  if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
    if (error_message) {
      *error_message = "Unable to parse inline AI endpoint URL.";
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
      WinHttpOpen(L"WeaselInlineInstruction/1.0",
                  WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
                  WINHTTP_NO_PROXY_BYPASS, 0));
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
      WinHttpOpenRequest(connection.get(), L"POST", object_name.c_str(),
                         nullptr, WINHTTP_NO_REFERER,
                         WINHTTP_DEFAULT_ACCEPT_TYPES, request_flags));
  if (!request.valid()) {
    if (error_message) {
      *error_message = "WinHTTP request creation failed.";
    }
    return false;
  }

  std::wstring headers =
      L"Content-Type: application/json\r\nAccept: text/event-stream\r\n";
  headers += L"Authorization: Bearer ";
  headers += option.app_key;
  headers += L"\r\n";
  WinHttpAddRequestHeaders(request.get(), headers.c_str(),
                           static_cast<DWORD>(-1),
                           WINHTTP_ADDREQ_FLAG_ADD |
                               WINHTTP_ADDREQ_FLAG_REPLACE);

  const std::string request_body = BuildInlineInstructionChatRequestBody(
      query_text, token, tenant_id, user_id);
  if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                          request_body.empty()
                              ? WINHTTP_NO_REQUEST_DATA
                              : const_cast<char*>(request_body.data()),
                          static_cast<DWORD>(request_body.size()),
                          static_cast<DWORD>(request_body.size()), 0) ||
      !WinHttpReceiveResponse(request.get(), nullptr)) {
    if (error_message) {
      *error_message = "Sending inline AI request failed.";
    }
    return false;
  }

  DWORD status_code = 0;
  DWORD status_code_size = sizeof(status_code);
  WinHttpQueryHeaders(request.get(),
                      WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                      WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                      &status_code_size, WINHTTP_NO_HEADER_INDEX);
  if (http_status_code) {
    *http_status_code = static_cast<int>(status_code);
  }

  std::string body;
  body.reserve(2048);
  std::string line_buffer;
  std::string event_payload;
  bool saw_stream_frame = false;
  const bool should_parse_stream =
      status_code >= 200 && status_code < 300;

  const auto flush_event_payload = [&](const std::string& payload) -> bool {
    if (payload.empty()) {
      return true;
    }
    saw_stream_frame = true;
    if (payload == "[DONE]") {
      return true;
    }
    std::wstring delta;
    bool is_complete = false;
    if (!ParseAIAssistantStreamEvent(payload, &delta, &is_complete)) {
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
        *error_message = "Failed to read inline AI response.";
      }
      return false;
    }
    if (available == 0) {
      break;
    }
    std::string chunk(available, '\0');
    DWORD downloaded = 0;
    if (!WinHttpReadData(request.get(), chunk.data(), available,
                         &downloaded)) {
      if (error_message) {
        *error_message = "Failed while downloading inline AI response.";
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
              *error_message = "Inline AI stream event payload is invalid.";
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
        *error_message = "Inline AI stream tail payload is invalid.";
      }
      return false;
    }
  }

  int response_code = 200;
  std::string response_message;
  read_response_metadata(body, &response_code, &response_message);

  if (status_code < 200 || status_code >= 300) {
    if (is_auth_failure_code(static_cast<int>(status_code)) ||
        is_auth_failure_code(response_code)) {
      if (http_status_code) {
        *http_status_code = 401;
      }
    }
    if (error_message) {
      *error_message =
          !response_message.empty()
              ? response_message
              : "Inline AI request returned HTTP " +
                    std::to_string(status_code) + ".";
    }
    return false;
  }

  if (response_code != 0 && response_code != 200) {
    if (is_auth_failure_code(response_code) && http_status_code) {
      *http_status_code = 401;
    }
    if (error_message) {
      *error_message =
          !response_message.empty()
              ? response_message
              : "Inline AI response code " + std::to_string(response_code) +
                    ".";
    }
    return false;
  }

  if (should_parse_stream && saw_stream_frame) {
    return true;
  }

  std::wstring output_text;
  std::string parsed_error;
  if (!ParseAIAssistantResponse(body, &output_text, &parsed_error)) {
    if (error_message) {
      *error_message =
          !response_message.empty() ? response_message : parsed_error;
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

InlineInstructionPromptSnapshot CaptureInlineInstructionPromptSnapshot(
    RimeSessionId session_id) {
  InlineInstructionPromptSnapshot snapshot;
  RIME_STRUCT(RimeContext, ctx);
  if (rime_api->get_context(session_id, &ctx)) {
    if (ctx.composition.preedit && ctx.composition.length > 0) {
      snapshot.text = u8tow(ctx.composition.preedit);
      snapshot.cursor_pos = static_cast<size_t>(max(
          0, utf8towcslen(ctx.composition.preedit, ctx.composition.cursor_pos)));
      snapshot.cursor_pos = min(snapshot.cursor_pos, snapshot.text.size());
    } else if (ctx.commit_text_preview) {
      snapshot.text = u8tow(ctx.commit_text_preview);
      snapshot.cursor_pos = snapshot.text.size();
    }
    rime_api->free_context(&ctx);
  }
  if (snapshot.text.empty()) {
    const char* input_text = rime_api->get_input(session_id);
    if (input_text) {
      snapshot.text = u8tow(input_text);
      snapshot.cursor_pos = snapshot.text.size();
    }
  }
  return snapshot;
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

std::wstring CaptureCompositionPreeditText(RimeSessionId session_id) {
  RIME_STRUCT(RimeContext, ctx);
  std::wstring preedit_text;
  if (rime_api->get_context(session_id, &ctx)) {
    if (ctx.composition.preedit && ctx.composition.length > 0) {
      preedit_text = u8tow(ctx.composition.preedit);
    }
    rime_api->free_context(&ctx);
  }
  return preedit_text;
}

std::wstring CaptureInlineInstructionPromptSource(RimeSessionId session_id) {
  return CaptureInlineInstructionPromptSnapshot(session_id).text;
}

bool IsAIAssistantTriggerKey(const AIAssistantConfig& config,
                             const KeyEvent& key_event) {
  if (!config.enabled) {
    return false;
  }
  return IsAiHotkeyMatch(config.trigger_binding, key_event);
}

bool ShouldLogAIAssistantTriggerProbe(const AIAssistantConfig& config,
                                      const KeyEvent& key_event) {
  if (!config.enabled) {
    return false;
  }
  if (key_event.mask & ibus::CONTROL_MASK) {
    return true;
  }
  if (key_event.keycode == config.trigger_binding.keycode) {
    return true;
  }
  if (config.trigger_binding.keycode == '3' &&
      (key_event.keycode == '3' || key_event.keycode == '#')) {
    return true;
  }
  return false;
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

bool IsPlainCandidateSelectKey(const KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  const auto modifiers = key_event.mask & ibus::MODIFIER_MASK;
  return (modifiers &
          (ibus::CONTROL_MASK | ibus::ALT_MASK | ibus::SUPER_MASK |
           ibus::HYPER_MASK | ibus::META_MASK)) == 0;
}

bool TryResolveCandidateSelectIndex(const KeyEvent& key_event,
                                    size_t* index) {
  if (!index || !IsPlainCandidateSelectKey(key_event)) {
    return false;
  }
  if (key_event.keycode >= '1' && key_event.keycode <= '9') {
    *index = static_cast<size_t>(key_event.keycode - '1');
    return true;
  }
  if (key_event.keycode == '0') {
    *index = 9;
    return true;
  }
  if (key_event.keycode >= ibus::KP_1 && key_event.keycode <= ibus::KP_9) {
    *index = static_cast<size_t>(key_event.keycode - ibus::KP_1);
    return true;
  }
  if (key_event.keycode == ibus::KP_0) {
    *index = 9;
    return true;
  }
  if (key_event.keycode == ibus::space || key_event.keycode == ibus::Return ||
      key_event.keycode == ibus::KP_Enter) {
    *index = 0;
    return true;
  }
  return false;
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

void AppendDedicatedInfoLogLine(const std::wstring& file_name,
                                const std::wstring& fallback_file_name,
                                const std::string& message) {
  std::unique_lock<std::mutex> lock(g_input_content_log_mutex, std::try_to_lock);
  if (!lock.owns_lock()) {
    return;
  }
  const fs::path log_dir = WeaselLogPath();
  std::error_code ec;
  fs::create_directories(log_dir, ec);
  const fs::path dedicated_log_file = log_dir / file_name;
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
  const fs::path fallback_path = fs::path(temp_path) / fallback_file_name;
  AppendLineToFile(fallback_path, line);
}

void AppendInputContentInfoLogLine(const std::string& message) {
  AppendDedicatedInfoLogLine(L"rime.weasel.input_content.INFO.log",
                             L"weasel_input_content_fallback.log", message);
}

void AppendAIAssistantInfoLogLine(const std::string& message) {
  AppendDedicatedInfoLogLine(L"rime.weasel.ai_assistant.INFO.log",
                             L"weasel_ai_assistant_fallback.log", message);
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

std::wstring ResolveAIAssistantUserInfoCachePath(
    const AIAssistantConfig& config) {
  const fs::path login_state_path(ResolveAIAssistantLoginStatePath(config));
  return (login_state_path.parent_path() / L"ai_user_info_cache.json")
      .wstring();
}

bool SaveAIAssistantUserInfoCache(const AIAssistantConfig& config,
                                  const std::string& response_body) {
  if (response_body.empty()) {
    return false;
  }

  rapidjson::Document doc;
  doc.Parse(response_body.c_str(), response_body.size());
  if (doc.HasParseError() || !doc.IsObject()) {
    return false;
  }

  const fs::path cache_path = ResolveAIAssistantUserInfoCachePath(config);
  std::error_code ec;
  fs::create_directories(cache_path.parent_path(), ec);

  std::ofstream output(cache_path, std::ios::binary | std::ios::trunc);
  if (!output) {
    return false;
  }
  output.write(response_body.data(),
               static_cast<std::streamsize>(response_body.size()));
  return output.good();
}

void ClearAIAssistantUserInfoCache(const AIAssistantConfig& config) {
  const fs::path cache_path(ResolveAIAssistantUserInfoCachePath(config));
  std::error_code remove_error;
  fs::remove(cache_path, remove_error);
  if (remove_error) {
    LOG(WARNING) << "AI user info cache remove failed: "
                 << cache_path.u8string()
                 << ", error=" << remove_error.message();
  }
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

bool DecodeBase64String(const std::string& input, std::string* output) {
  if (!output) {
    return false;
  }
  output->clear();

  static const std::string kAlphabet =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  int value = 0;
  int bit_count = -8;
  for (unsigned char ch : input) {
    if (std::isspace(ch)) {
      continue;
    }
    if (ch == '=') {
      break;
    }
    const size_t alphabet_index = kAlphabet.find(static_cast<char>(ch));
    if (alphabet_index == std::string::npos) {
      return false;
    }
    value = (value << 6) + static_cast<int>(alphabet_index);
    bit_count += 6;
    if (bit_count >= 0) {
      output->push_back(static_cast<char>((value >> bit_count) & 0xff));
      bit_count -= 8;
    }
  }
  return true;
}

bool DecodeObfuscatedJsonString(const std::string& base64_text,
                                rapidjson::Document* document,
                                std::string* error_message) {
  if (!document) {
    if (error_message) {
      *error_message = "Decoded config document is null.";
    }
    return false;
  }
  document->SetObject();
  if (base64_text.empty()) {
    if (error_message) {
      *error_message = "Config payload is empty.";
    }
    return false;
  }

  std::string decoded_bytes;
  if (!DecodeBase64String(base64_text, &decoded_bytes)) {
    if (error_message) {
      *error_message = "Config payload base64 decode failed.";
    }
    return false;
  }

  const std::wstring decoded_text = u8tow(decoded_bytes);
  if (decoded_text.empty() && !decoded_bytes.empty()) {
    if (error_message) {
      *error_message = "Config payload UTF-8 decode failed.";
    }
    return false;
  }

  std::wstring reversed_text(decoded_text.rbegin(), decoded_text.rend());
  std::wstring original_text;
  original_text.reserve(reversed_text.size());
  for (wchar_t ch : reversed_text) {
    const unsigned int code_point = static_cast<unsigned int>(ch);
    if (code_point < 120U) {
      if (error_message) {
        *error_message = "Config payload character decode failed.";
      }
      return false;
    }
    original_text.push_back(static_cast<wchar_t>(code_point - 120U));
  }

  const std::string json_utf8 = wtou8(original_text);
  AppendAIAssistantInfoLogLine("AI front-end config decoded data: " +
                               json_utf8);
  LOG(INFO) << "AI front-end config decoded data: " << json_utf8;
  document->Parse(json_utf8.c_str(), json_utf8.size());
  if (document->HasParseError() || !document->IsObject()) {
    if (error_message) {
      *error_message = "Decoded config payload is not valid JSON.";
    }
    return false;
  }
  return true;
}

std::string GenerateRandomMqttClientId() {
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

std::string GenerateLoginClientId() {
  return GenerateRandomMqttClientId();
}

bool IsValidAIAssistantInstructionChangedTopic(const std::string& topic) {
  const std::string trimmed = TrimAsciiWhitespace(topic);
  if (trimmed.size() < 2 || trimmed.find('#') != std::string::npos) {
    return false;
  }
  return trimmed.compare(trimmed.size() - 2, 2, "/+") == 0;
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

bool ParseMqttRemainingLength(const std::vector<uint8_t>& packet,
                              size_t offset,
                              size_t* value,
                              size_t* used_bytes);
int MqttPacketType(const std::vector<uint8_t>& packet);

std::vector<uint8_t> BuildMqttPingReqPacket() {
  return {0xC0, 0x00};
}

bool WaitForAIAssistantInstructionChangedStop(std::atomic<bool>* stop,
                                              DWORD milliseconds) {
  if (!stop) {
    Sleep(milliseconds);
    return false;
  }
  DWORD elapsed = 0;
  while (!stop->load() && elapsed < milliseconds) {
    const DWORD remaining = milliseconds - elapsed;
    const DWORD chunk = remaining < 50 ? remaining : 50;
    Sleep(chunk);
    elapsed += chunk;
  }
  return stop->load();
}

bool IsMqttPingResp(const std::vector<uint8_t>& packet) {
  if (MqttPacketType(packet) != 13) {
    return false;
  }
  size_t remaining = 0;
  size_t used = 0;
  return ParseMqttRemainingLength(packet, 1, &remaining, &used) &&
         remaining == 0;
}

bool TryExtractInstructionIdFromChangedTopic(
    const std::string& topic_filter,
    const std::string& topic,
    std::string* instruction_id) {
  if (!instruction_id || !IsValidAIAssistantInstructionChangedTopic(topic_filter)) {
    return false;
  }
  instruction_id->clear();

  const std::string prefix = topic_filter.substr(0, topic_filter.size() - 1);
  if (topic.size() <= prefix.size() ||
      topic.compare(0, prefix.size(), prefix) != 0) {
    return false;
  }

  const std::string candidate_id = topic.substr(prefix.size());
  if (candidate_id.empty() || candidate_id.find('/') != std::string::npos) {
    return false;
  }
  *instruction_id = candidate_id;
  return true;
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

bool ExtractAIAssistantCachedUserId(const rapidjson::Value& value,
                                    std::string* user_id) {
  if (!user_id || !value.IsObject()) {
    return false;
  }
  user_id->clear();

  const auto data_it = value.FindMember("data");
  if (data_it != value.MemberEnd() && data_it->value.IsObject()) {
    if (ReadStringLikeJsonMember(data_it->value, "id", user_id) &&
        !user_id->empty()) {
      return true;
    }
  }

  std::string direct_id;
  const bool looks_like_user =
      value.FindMember("username") != value.MemberEnd() ||
      value.FindMember("nickName") != value.MemberEnd() ||
      value.FindMember("mobile") != value.MemberEnd();
  if (looks_like_user && ReadStringLikeJsonMember(value, "id", &direct_id) &&
      !direct_id.empty()) {
    *user_id = direct_id;
    return true;
  }
  return false;
}

bool LoadAIAssistantCachedUserId(const AIAssistantConfig& config,
                                 std::string* user_id) {
  if (!user_id) {
    return false;
  }
  user_id->clear();

  const fs::path cache_path(ResolveAIAssistantUserInfoCachePath(config));
  std::ifstream input(cache_path, std::ios::binary);
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
  return ExtractAIAssistantCachedUserId(doc, user_id);
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

bool ParseHttpUrlOrigin(const std::wstring& url,
                        std::wstring* scheme,
                        std::wstring* host,
                        INTERNET_PORT* port);
std::wstring BuildHttpOrigin(const std::wstring& scheme,
                             const std::wstring& host,
                             INTERNET_PORT port);
std::vector<std::wstring> BuildAIAssistantUserInfoEndpointCandidates(
    const AIAssistantConfig& config);

std::string BuildAIAssistantInstitutionListRequestBody(
    const AIAssistantConfig& config) {
  (void)config;
  std::string body;
  body.reserve(128);
  body += "{\"queryContent\":\"\",\"insShowType\":\"";
  body += "3";
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
        *it, {"insId", "id", "tenantId", "orgId"});
    std::string lookup_text = ReadFirstStringLikeJsonMember(
        *it, {"insCode",      "code",       "shortcutCode", "shortcut",
              "spell",        "spellCode",  "pinyin",       "py",
              "keyword",      "keywords",   "query",        "queryCode"});
    std::string name = ReadFirstStringLikeJsonMember(
        *it, {"insName", "name", "tenantName", "title", "label"});
    std::string app_key = ReadFirstStringLikeJsonMember(
        *it, {"appKey", "app_key", "appkey", "agentAppKey"});
    std::string template_content = ReadFirstStringLikeJsonMember(
        *it, {"tmplContent", "content", "template", "templateContent",
              "prompt", "promptContent", "instruction",
              "instructionContent"});
    if (name.empty()) {
      continue;
    }
    if (id.empty()) {
      id = !lookup_text.empty() ? lookup_text : name;
    }
    options->push_back(AIPanelInstitutionOption(
        u8tow(id), u8tow(name), u8tow(lookup_text), u8tow(app_key),
        u8tow(template_content)));
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

  const auto is_auth_failure_code = [](int code) -> bool {
    switch (code) {
      case 401:
      case 11001:
      case 11011:
      case 11012:
      case 11013:
      case 11014:
      case 11015:
        return true;
      default:
        return false;
    }
  };

  const auto read_response_metadata = [](const std::string& response_body,
                                         int* code,
                                         std::string* message) {
    if (code) {
      *code = 200;
    }
    if (message) {
      message->clear();
    }

    rapidjson::Document document;
    document.Parse(response_body.c_str(), response_body.size());
    if (document.HasParseError() || !document.IsObject()) {
      return;
    }

    const auto code_it = document.FindMember("code");
    if (code_it != document.MemberEnd()) {
      int parsed_code = 200;
      if (code_it->value.IsInt()) {
        parsed_code = code_it->value.GetInt();
      } else if (code_it->value.IsUint()) {
        parsed_code = static_cast<int>(code_it->value.GetUint());
      } else if (code_it->value.IsInt64()) {
        parsed_code = static_cast<int>(code_it->value.GetInt64());
      } else if (code_it->value.IsUint64()) {
        parsed_code = static_cast<int>(code_it->value.GetUint64());
      } else if (code_it->value.IsString()) {
        parsed_code = std::atoi(code_it->value.GetString());
      }
      if (code) {
        *code = parsed_code;
      }
    }

    if (!message) {
      return;
    }

    const char* message_keys[] = {"msg", "message", "errorMsg", "error"};
    for (const char* key : message_keys) {
      const auto message_it = document.FindMember(key);
      if (message_it != document.MemberEnd() && message_it->value.IsString()) {
        *message = message_it->value.GetString();
        return;
      }
    }
  };

  std::vector<std::wstring> endpoint_candidates;
  const auto add_unique_endpoint =
      [&endpoint_candidates](const std::wstring& endpoint) {
        if (endpoint.empty()) {
          return;
        }
        for (const auto& existing : endpoint_candidates) {
          if (_wcsicmp(existing.c_str(), endpoint.c_str()) == 0) {
            return;
          }
        }
        endpoint_candidates.push_back(endpoint);
      };
  const auto add_candidates_from_url =
      [&add_unique_endpoint](const std::wstring& url) {
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
        add_unique_endpoint(
            origin + L"/langwell-api/langwell-ins-server/client/acct/ins/list");
      };

  add_candidates_from_url(u8tow(config.official_url));
  add_candidates_from_url(u8tow(config.panel_url));
  add_candidates_from_url(u8tow(config.login_url));
  add_candidates_from_url(u8tow(config.refresh_token_endpoint));

  if (endpoint_candidates.empty()) {
    if (error_message) {
      *error_message = "No institution list endpoint could be resolved.";
    }
    return false;
  }

  const std::string request_body =
      BuildAIAssistantInstitutionListRequestBody(config);
  std::string last_error = "Institution list request failed.";
  int last_status_code = 0;

  for (const auto& endpoint : endpoint_candidates) {
    URL_COMPONENTS parts = {0};
    parts.dwStructSize = sizeof(parts);
    parts.dwSchemeLength = static_cast<DWORD>(-1);
    parts.dwHostNameLength = static_cast<DWORD>(-1);
    parts.dwUrlPathLength = static_cast<DWORD>(-1);
    parts.dwExtraInfoLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
      last_error = "Unable to parse institution list endpoint URL.";
      continue;
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
        WinHttpOpen(L"WeaselAIAssistantInstitution/1.0",
                    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
    if (!session.valid()) {
      last_error = "WinHTTP session creation failed for institution list.";
      continue;
    }
    WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                       config.timeout_ms, config.timeout_ms);

    ScopedWinHttpHandle connection(
        WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
    if (!connection.valid()) {
      last_error = "WinHTTP connection failed for institution endpoint.";
      continue;
    }

    const DWORD request_flags =
        parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
    ScopedWinHttpHandle request(
        WinHttpOpenRequest(connection.get(), L"POST", object_name.c_str(),
                           nullptr, WINHTTP_NO_REFERER,
                           WINHTTP_DEFAULT_ACCEPT_TYPES, request_flags));
    if (!request.valid()) {
      last_error = "WinHTTP request creation failed for institution list.";
      continue;
    }

    std::wstring headers =
        L"Accept: application/json\r\nContent-Type: application/json\r\n";
    if (!tenant_id.empty()) {
      headers += L"TenantId: ";
      headers += u8tow(tenant_id);
      headers += L"\r\nTenantid: ";
      headers += u8tow(tenant_id);
      headers += L"\r\n";
    }
    if (!token.empty()) {
      headers += L"Token: ";
      headers += u8tow(token);
      headers += L"\r\nOVERTOKEN: ";
      headers += u8tow(token);
      headers += L"\r\n";
    }
    WinHttpAddRequestHeaders(
        request.get(), headers.c_str(), static_cast<DWORD>(-1),
        WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

    if (!WinHttpSendRequest(
            request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
            request_body.empty() ? WINHTTP_NO_REQUEST_DATA
                                 : const_cast<char*>(request_body.data()),
            static_cast<DWORD>(request_body.size()),
            static_cast<DWORD>(request_body.size()), 0) ||
        !WinHttpReceiveResponse(request.get(), nullptr)) {
      last_error = "Sending institution list request failed.";
      continue;
    }

    DWORD status_code = 0;
    DWORD status_code_size = sizeof(status_code);
    WinHttpQueryHeaders(request.get(),
                        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                        WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                        &status_code_size, WINHTTP_NO_HEADER_INDEX);
    last_status_code = static_cast<int>(status_code);

    std::string response_body;
    bool read_ok = true;
    for (;;) {
      DWORD available = 0;
      if (!WinHttpQueryDataAvailable(request.get(), &available)) {
        last_error = "Failed to read institution list response.";
        read_ok = false;
        break;
      }
      if (available == 0) {
        break;
      }
      std::string chunk(available, '\0');
      DWORD downloaded = 0;
      if (!WinHttpReadData(request.get(), chunk.data(), available,
                           &downloaded)) {
        last_error = "Failed while downloading institution list response.";
        read_ok = false;
        break;
      }
      chunk.resize(downloaded);
      response_body += chunk;
    }
    if (!read_ok) {
      continue;
    }

    int response_code = 200;
    std::string response_message;
    read_response_metadata(response_body, &response_code, &response_message);

    if (status_code < 200 || status_code >= 300) {
      last_error = !response_message.empty()
                       ? response_message
                       : "Institution list request returned HTTP " +
                             std::to_string(status_code) + ".";
      if (http_status_code) {
        *http_status_code = static_cast<int>(status_code);
      }
      if (static_cast<int>(status_code) == 401) {
        break;
      }
      continue;
    }

    if (response_code != 0 && response_code != 200) {
      last_error = !response_message.empty()
                       ? response_message
                       : "Institution list response code " +
                             std::to_string(response_code) + ".";
      if (is_auth_failure_code(response_code)) {
        last_status_code = 401;
        if (http_status_code) {
          *http_status_code = 401;
        }
        break;
      }
      continue;
    }

    std::vector<AIPanelInstitutionOption> parsed_options;
    std::string parse_error;
    if (!ParseAIAssistantInstitutionList(response_body, &parsed_options,
                                         &parse_error)) {
      last_error = parse_error;
      continue;
    }

    *options = std::move(parsed_options);
    if (http_status_code) {
      *http_status_code = static_cast<int>(status_code);
    }
    return true;
  }

  if (http_status_code && *http_status_code == 0) {
    *http_status_code = last_status_code;
  }
  if (error_message) {
    *error_message = last_error;
  }
  return false;
}

bool FetchAIAssistantUserInfo(const AIAssistantConfig& config,
                              const std::string& token,
                              const std::string& tenant_id,
                              std::string* response_body,
                              std::string* error_message,
                              int* http_status_code = nullptr) {
  if (!response_body) {
    if (error_message) {
      *error_message = "User info response output is null.";
    }
    return false;
  }
  response_body->clear();
  if (error_message) {
    error_message->clear();
  }
  if (http_status_code) {
    *http_status_code = 0;
  }
  if (token.empty()) {
    if (error_message) {
      *error_message = "User info request token is empty.";
    }
    return false;
  }

  const auto is_auth_failure_code = [](int code) -> bool {
    switch (code) {
      case 401:
      case 11001:
      case 11011:
      case 11012:
      case 11013:
      case 11014:
      case 11015:
        return true;
      default:
        return false;
    }
  };

  const std::vector<std::wstring> endpoint_candidates =
      BuildAIAssistantUserInfoEndpointCandidates(config);
  if (endpoint_candidates.empty()) {
    if (error_message) {
      *error_message = "No user info endpoint could be resolved.";
    }
    return false;
  }

  std::string last_error = "User info request failed.";
  int last_status_code = 0;

  for (const auto& endpoint : endpoint_candidates) {
    URL_COMPONENTS parts = {0};
    parts.dwStructSize = sizeof(parts);
    parts.dwSchemeLength = static_cast<DWORD>(-1);
    parts.dwHostNameLength = static_cast<DWORD>(-1);
    parts.dwUrlPathLength = static_cast<DWORD>(-1);
    parts.dwExtraInfoLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
      last_error = "Unable to parse user info endpoint URL.";
      continue;
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
        WinHttpOpen(L"WeaselAIAssistantUserInfo/1.0",
                    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
    if (!session.valid()) {
      last_error = "WinHTTP session creation failed for user info.";
      continue;
    }
    WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                       config.timeout_ms, config.timeout_ms);

    ScopedWinHttpHandle connection(
        WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
    if (!connection.valid()) {
      last_error = "WinHTTP connection failed for user info endpoint.";
      continue;
    }

    const DWORD request_flags =
        parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
    ScopedWinHttpHandle request(
        WinHttpOpenRequest(connection.get(), L"GET", object_name.c_str(),
                           nullptr, WINHTTP_NO_REFERER,
                           WINHTTP_DEFAULT_ACCEPT_TYPES, request_flags));
    if (!request.valid()) {
      last_error = "WinHTTP request creation failed for user info.";
      continue;
    }

    std::wstring headers = L"Accept: application/json\r\n";
    if (!tenant_id.empty()) {
      const std::wstring tenant = u8tow(tenant_id);
      headers += L"TenantId: ";
      headers += tenant;
      headers += L"\r\nTenantid: ";
      headers += tenant;
      headers += L"\r\n";
    }
    {
      const std::wstring token_text = u8tow(token);
      headers += L"Token: ";
      headers += token_text;
      headers += L"\r\nSA-TOKEN: ";
      headers += token_text;
      headers += L"\r\nOVERTOKEN: ";
      headers += token_text;
      headers += L"\r\n";
    }
    WinHttpAddRequestHeaders(
        request.get(), headers.c_str(), static_cast<DWORD>(-1),
        WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

    if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(request.get(), nullptr)) {
      last_error = "Sending user info request failed.";
      continue;
    }

    DWORD status_code = 0;
    DWORD status_code_size = sizeof(status_code);
    WinHttpQueryHeaders(request.get(),
                        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                        WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                        &status_code_size, WINHTTP_NO_HEADER_INDEX);
    last_status_code = static_cast<int>(status_code);

    std::string body;
    bool read_ok = true;
    for (;;) {
      DWORD available = 0;
      if (!WinHttpQueryDataAvailable(request.get(), &available)) {
        last_error = "Failed to read user info response.";
        read_ok = false;
        break;
      }
      if (available == 0) {
        break;
      }
      std::string chunk(available, '\0');
      DWORD downloaded = 0;
      if (!WinHttpReadData(request.get(), chunk.data(), available,
                           &downloaded)) {
        last_error = "Failed while downloading user info response.";
        read_ok = false;
        break;
      }
      chunk.resize(downloaded);
      body += chunk;
    }
    if (!read_ok) {
      continue;
    }

    rapidjson::Document document;
    document.Parse(body.c_str(), body.size());

    int response_code = 200;
    std::string response_message;
    if (!document.HasParseError() && document.IsObject()) {
      const auto code_it = document.FindMember("code");
      if (code_it != document.MemberEnd()) {
        if (code_it->value.IsInt()) {
          response_code = code_it->value.GetInt();
        } else if (code_it->value.IsUint()) {
          response_code = static_cast<int>(code_it->value.GetUint());
        } else if (code_it->value.IsInt64()) {
          response_code = static_cast<int>(code_it->value.GetInt64());
        } else if (code_it->value.IsUint64()) {
          response_code = static_cast<int>(code_it->value.GetUint64());
        } else if (code_it->value.IsString()) {
          response_code = std::atoi(code_it->value.GetString());
        }
      }

      const char* message_keys[] = {"msg", "message", "errorMsg", "error"};
      for (const char* key : message_keys) {
        const auto message_it = document.FindMember(key);
        if (message_it != document.MemberEnd() &&
            message_it->value.IsString()) {
          response_message = message_it->value.GetString();
          break;
        }
      }
    }

    if (status_code < 200 || status_code >= 300) {
      last_error = !response_message.empty()
                       ? response_message
                       : "User info request returned HTTP " +
                             std::to_string(status_code) + ".";
      if (http_status_code) {
        *http_status_code = static_cast<int>(status_code);
      }
      if (static_cast<int>(status_code) == 401) {
        break;
      }
      continue;
    }

    if (document.HasParseError() || !document.IsObject()) {
      last_error = "User info response is not valid JSON.";
      continue;
    }

    if (response_code != 0 && response_code != 200) {
      last_error = !response_message.empty()
                       ? response_message
                       : "User info response code " +
                             std::to_string(response_code) + ".";
      if (is_auth_failure_code(response_code)) {
        last_status_code = 401;
        if (http_status_code) {
          *http_status_code = 401;
        }
        break;
      }
      continue;
    }

    *response_body = std::move(body);
    if (http_status_code) {
      *http_status_code = static_cast<int>(status_code);
    }
    return true;
  }

  if (http_status_code && *http_status_code == 0) {
    *http_status_code = last_status_code;
  }
  if (error_message) {
    *error_message = last_error;
  }
  return false;
}

bool RefreshAIAssistantUserInfoCache(const AIAssistantConfig& config,
                                     const std::string& token,
                                     const std::string& tenant_id,
                                     std::string* error_message = nullptr,
                                     int* http_status_code = nullptr) {
  std::string response_body;
  const bool ok = FetchAIAssistantUserInfo(config, token, tenant_id,
                                           &response_body, error_message,
                                           http_status_code);
  if (!ok) {
    if (http_status_code && *http_status_code == 401) {
      ClearAIAssistantUserInfoCache(config);
    }
    return false;
  }

  if (!SaveAIAssistantUserInfoCache(config, response_body)) {
    if (error_message) {
      *error_message = "User info cache save failed.";
    }
    return false;
  }
  return true;
}

bool FetchAIAssistantPermissionUpdateTopicPrefix(
    const AIAssistantConfig& config,
    const std::string& token,
    const std::string& tenant_id,
    std::string* topic_prefix,
    std::string* error_message);

std::string BuildAIAssistantPermissionUpdateTopic(
    const std::string& topic_prefix,
    const std::string& user_id,
    const std::string& tenant_id) {
  const std::string normalized_prefix = TrimAsciiWhitespace(topic_prefix);
  if (normalized_prefix.empty() || user_id.empty() || tenant_id.empty()) {
    return std::string();
  }

  std::string topic = normalized_prefix;
  if (topic.back() != '/') {
    topic.push_back('/');
  }
  topic += user_id;
  topic.push_back('/');
  topic += tenant_id;
  return topic;
}

bool ResolveAIAssistantPermissionUpdateTopic(
    const AIAssistantConfig& config,
    const std::string& token,
    const std::string& tenant_id,
    std::string* topic,
    std::string* error_message) {
  if (!topic) {
    if (error_message) {
      *error_message = "Permission update topic output is null.";
    }
    return false;
  }
  topic->clear();
  if (tenant_id.empty()) {
    if (error_message) {
      *error_message = "Permission update tenant id is empty.";
    }
    return false;
  }

  std::string user_id;
  if (!LoadAIAssistantCachedUserId(config, &user_id) || user_id.empty()) {
    if (error_message) {
      *error_message = "Permission update user id is unavailable.";
    }
    return false;
  }

  std::string topic_prefix;
  if (!FetchAIAssistantPermissionUpdateTopicPrefix(config, token, tenant_id,
                                                   &topic_prefix,
                                                   error_message)) {
    return false;
  }

  *topic = BuildAIAssistantPermissionUpdateTopic(topic_prefix, user_id,
                                                 tenant_id);
  if (topic->empty()) {
    if (error_message) {
      *error_message = "Permission update topic build failed.";
    }
    return false;
  }
  return true;
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
    json += "\",\"template\":\"";
    json += EscapeJsonString(wtou8(option.template_content));
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
  AddUniqueEndpoint(endpoints, origin + L"/lamp-api/oauth/anyTenant/refresh");
}

std::vector<std::wstring> BuildRefreshTokenEndpointCandidates(
    const AIAssistantConfig& config) {
  std::vector<std::wstring> endpoints;
  if (!config.refresh_token_endpoint.empty()) {
    AddUniqueEndpoint(&endpoints, u8tow(config.refresh_token_endpoint));
  }
  if (!config.official_url.empty()) {
    AddRefreshEndpointCandidatesFromUrl(u8tow(config.official_url), &endpoints);
  }
  if (!config.panel_url.empty()) {
    AddRefreshEndpointCandidatesFromUrl(u8tow(config.panel_url), &endpoints);
  }
  if (!config.login_url.empty()) {
    AddRefreshEndpointCandidatesFromUrl(u8tow(config.login_url), &endpoints);
  }
  return endpoints;
}

void AddAIAssistantUserInfoEndpointCandidatesFromUrl(
    const std::wstring& url,
    std::vector<std::wstring>* endpoints) {
  if (!endpoints) {
    return;
  }
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
  AddUniqueEndpoint(endpoints,
                    origin + L"/lamp-api/oauth/anyone/getUserInfoById");
}

void AddAIAssistantFrontEndConfigEndpointCandidatesFromUrl(
    const std::wstring& url,
    std::vector<std::wstring>* endpoints) {
  if (!endpoints) {
    return;
  }
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
  AddUniqueEndpoint(endpoints, origin + L"/lamp-api/frontEndConfig/getConfig");
}

std::vector<std::wstring> BuildAIAssistantUserInfoEndpointCandidates(
    const AIAssistantConfig& config) {
  std::vector<std::wstring> endpoints;
  AddAIAssistantUserInfoEndpointCandidatesFromUrl(
      u8tow(config.official_url), &endpoints);
  AddAIAssistantUserInfoEndpointCandidatesFromUrl(u8tow(config.panel_url),
                                                  &endpoints);
  AddAIAssistantUserInfoEndpointCandidatesFromUrl(u8tow(config.login_url),
                                                  &endpoints);
  AddAIAssistantUserInfoEndpointCandidatesFromUrl(
      u8tow(config.refresh_token_endpoint), &endpoints);
  return endpoints;
}

std::vector<std::wstring> BuildAIAssistantFrontEndConfigEndpointCandidates(
    const AIAssistantConfig& config) {
  std::vector<std::wstring> endpoints;
  AddAIAssistantFrontEndConfigEndpointCandidatesFromUrl(
      u8tow(config.official_url), &endpoints);
  AddAIAssistantFrontEndConfigEndpointCandidatesFromUrl(u8tow(config.panel_url),
                                                        &endpoints);
  AddAIAssistantFrontEndConfigEndpointCandidatesFromUrl(u8tow(config.login_url),
                                                        &endpoints);
  AddAIAssistantFrontEndConfigEndpointCandidatesFromUrl(
      u8tow(config.refresh_token_endpoint), &endpoints);
  return endpoints;
}

bool TryReadAIAssistantPermissionUpdateTopicPrefix(
    const rapidjson::Value& value,
    std::string* topic_prefix) {
  if (!topic_prefix || !value.IsObject()) {
    return false;
  }
  topic_prefix->clear();
  const auto topic_prefix_it =
      value.FindMember(kAIAssistantPermissionUpdateTopicPrefixObjectKey);
  if (topic_prefix_it != value.MemberEnd() &&
      topic_prefix_it->value.IsObject()) {
    ReadStringLikeJsonMember(topic_prefix_it->value,
                             kAIAssistantPermissionUpdateTopicPrefixMemberKey,
                             topic_prefix);
    if (!topic_prefix->empty()) {
      return true;
    }
  }

  const auto assistant_it = value.FindMember("ai-assistant");
  if (assistant_it != value.MemberEnd() && assistant_it->value.IsObject()) {
    const auto assistant_topic_prefix_it = assistant_it->value.FindMember(
        kAIAssistantPermissionUpdateTopicPrefixObjectKey);
    if (assistant_topic_prefix_it != assistant_it->value.MemberEnd() &&
        assistant_topic_prefix_it->value.IsObject()) {
      ReadStringLikeJsonMember(
          assistant_topic_prefix_it->value,
          kAIAssistantPermissionUpdateTopicPrefixMemberKey, topic_prefix);
    }
  }
  return !topic_prefix->empty();
}

bool FetchAIAssistantPermissionUpdateTopicPrefix(
    const AIAssistantConfig& config,
    const std::string& token,
    const std::string& tenant_id,
    std::string* topic_prefix,
    std::string* error_message) {
  if (!topic_prefix) {
    if (error_message) {
      *error_message = "Permission update topic prefix output is null.";
    }
    return false;
  }
  topic_prefix->clear();
  if (error_message) {
    error_message->clear();
  }

  const std::vector<std::wstring> endpoint_candidates =
      BuildAIAssistantFrontEndConfigEndpointCandidates(config);
  if (endpoint_candidates.empty()) {
    if (error_message) {
      *error_message = "No front-end config endpoint could be resolved.";
    }
    return false;
  }

  std::string endpoint_candidates_log;
  for (size_t i = 0; i < endpoint_candidates.size(); ++i) {
    if (i > 0) {
      endpoint_candidates_log += " | ";
    }
    endpoint_candidates_log += wtou8(endpoint_candidates[i]);
  }
  LOG(INFO) << "AI front-end config endpoint candidates: "
            << endpoint_candidates_log;

  std::string last_error = "Front-end config request failed.";

  for (const auto& endpoint : endpoint_candidates) {
    URL_COMPONENTS parts = {0};
    parts.dwStructSize = sizeof(parts);
    parts.dwSchemeLength = static_cast<DWORD>(-1);
    parts.dwHostNameLength = static_cast<DWORD>(-1);
    parts.dwUrlPathLength = static_cast<DWORD>(-1);
    parts.dwExtraInfoLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(endpoint.c_str(), 0, 0, &parts)) {
      last_error = "Unable to parse front-end config endpoint URL.";
      continue;
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
        WinHttpOpen(L"WeaselAIAssistantFrontEndConfig/1.0",
                    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
    if (!session.valid()) {
      last_error = "WinHTTP session creation failed for front-end config.";
      continue;
    }
    WinHttpSetTimeouts(session.get(), config.timeout_ms, config.timeout_ms,
                       config.timeout_ms, config.timeout_ms);

    ScopedWinHttpHandle connection(
        WinHttpConnect(session.get(), host.c_str(), parts.nPort, 0));
    if (!connection.valid()) {
      last_error = "WinHTTP connection failed for front-end config endpoint.";
      continue;
    }

    const DWORD request_flags =
        parts.nScheme == INTERNET_SCHEME_HTTPS ? WINHTTP_FLAG_SECURE : 0;
    ScopedWinHttpHandle request(
        WinHttpOpenRequest(connection.get(), L"GET", object_name.c_str(),
                           nullptr, WINHTTP_NO_REFERER,
                           WINHTTP_DEFAULT_ACCEPT_TYPES, request_flags));
    if (!request.valid()) {
      last_error = "WinHTTP request creation failed for front-end config.";
      continue;
    }

    std::wstring headers = L"Accept: application/json\r\n";
    if (!tenant_id.empty()) {
      const std::wstring tenant = u8tow(tenant_id);
      headers += L"TenantId: ";
      headers += tenant;
      headers += L"\r\nTenantid: ";
      headers += tenant;
      headers += L"\r\n";
    }
    if (!token.empty()) {
      const std::wstring token_text = u8tow(token);
      headers += L"Token: ";
      headers += token_text;
      headers += L"\r\nSA-TOKEN: ";
      headers += token_text;
      headers += L"\r\nOVERTOKEN: ";
      headers += token_text;
      headers += L"\r\n";
    }
    WinHttpAddRequestHeaders(
        request.get(), headers.c_str(), static_cast<DWORD>(-1),
        WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);

    if (!WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(request.get(), nullptr)) {
      last_error = "Sending front-end config request failed.";
      continue;
    }

    DWORD status_code = 0;
    DWORD status_code_size = sizeof(status_code);
    WinHttpQueryHeaders(request.get(),
                        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                        WINHTTP_HEADER_NAME_BY_INDEX, &status_code,
                        &status_code_size, WINHTTP_NO_HEADER_INDEX);

    std::string body;
    bool read_ok = true;
    for (;;) {
      DWORD available = 0;
      if (!WinHttpQueryDataAvailable(request.get(), &available)) {
        last_error = "Failed to read front-end config response.";
        read_ok = false;
        break;
      }
      if (available == 0) {
        break;
      }
      std::string chunk(available, '\0');
      DWORD downloaded = 0;
      if (!WinHttpReadData(request.get(), chunk.data(), available,
                           &downloaded)) {
        last_error = "Failed while downloading front-end config response.";
        read_ok = false;
        break;
      }
      chunk.resize(downloaded);
      body += chunk;
    }
    if (!read_ok) {
      continue;
    }

    AppendAIAssistantInfoLogLine("AI front-end config response endpoint=" +
                                 wtou8(endpoint) + " body=" + body);
    LOG(INFO) << "AI front-end config response endpoint="
              << wtou8(endpoint) << " body=" << body;

    rapidjson::Document document;
    document.Parse(body.c_str(), body.size());
    if (document.HasParseError() || !document.IsObject()) {
      last_error = "Front-end config response is not valid JSON.";
      continue;
    }

    int response_code = 200;
    std::string response_message;
    const auto code_it = document.FindMember("code");
    if (code_it != document.MemberEnd()) {
      if (code_it->value.IsInt()) {
        response_code = code_it->value.GetInt();
      } else if (code_it->value.IsUint()) {
        response_code = static_cast<int>(code_it->value.GetUint());
      } else if (code_it->value.IsInt64()) {
        response_code = static_cast<int>(code_it->value.GetInt64());
      } else if (code_it->value.IsUint64()) {
        response_code = static_cast<int>(code_it->value.GetUint64());
      } else if (code_it->value.IsString()) {
        response_code = std::atoi(code_it->value.GetString());
      }
    }

    const char* message_keys[] = {"msg", "message", "errorMsg", "error"};
    for (const char* key : message_keys) {
      const auto message_it = document.FindMember(key);
      if (message_it != document.MemberEnd() && message_it->value.IsString()) {
        response_message = message_it->value.GetString();
        break;
      }
    }

    if (status_code < 200 || status_code >= 300) {
      last_error = !response_message.empty()
                       ? response_message
                       : "Front-end config request returned HTTP " +
                             std::to_string(status_code) + ".";
      continue;
    }

    if (response_code != 0 && response_code != 200) {
      last_error = !response_message.empty()
                       ? response_message
                       : "Front-end config response code " +
                             std::to_string(response_code) + ".";
      continue;
    }

    const rapidjson::Value* config_value = nullptr;
    rapidjson::Document decoded_document;
    const auto data_it = document.FindMember("data");
    if (data_it != document.MemberEnd()) {
      if (data_it->value.IsString()) {
        std::string decode_error;
        if (!DecodeObfuscatedJsonString(data_it->value.GetString(),
                                        &decoded_document, &decode_error)) {
          last_error = decode_error;
          continue;
        }
        config_value = &decoded_document;
      } else if (data_it->value.IsObject()) {
        config_value = &data_it->value;
      }
    }
    if (!config_value) {
      config_value = &document;
    }

    std::string parsed_prefix;
    const bool has_configured_prefix =
        TryReadAIAssistantPermissionUpdateTopicPrefix(*config_value,
                                                     &parsed_prefix);
    parsed_prefix = TrimAsciiWhitespace(parsed_prefix);
    if (!has_configured_prefix || parsed_prefix.empty()) {
      parsed_prefix = kDefaultAIAssistantPermissionUpdateTopicPrefix;
    }

    *topic_prefix = parsed_prefix;
    AppendAIAssistantInfoLogLine(
        "AI front-end config resolved topic prefix endpoint=" +
        wtou8(endpoint) + " prefix=" + parsed_prefix);
    LOG(INFO) << "AI front-end config resolved topic prefix endpoint="
              << wtou8(endpoint) << " prefix=" << parsed_prefix;
    return true;
  }

  if (error_message) {
    *error_message = last_error;
  }
  return false;
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
      m_ai_inst_changed_stop(false),
      m_ai_inst_changed_session_handle(nullptr),
      m_ai_inst_changed_connection_handle(nullptr),
      m_ai_inst_changed_request_handle(nullptr),
      m_ai_inst_changed_websocket_handle(nullptr),
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
  _StopAIAssistantInstructionChangedListener();
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
  const bool config_opened = !!rime_api->config_open("weasel", &config);
  AppendAIAssistantInfoLogLine(std::string("AI initialize: config_open(weasel)=") +
                               (config_opened ? "1" : "0"));
  if (config_opened) {
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
  _StartAIAssistantInstructionChangedListener();
  m_last_schema_id.clear();
}

void RimeWithWeaselHandler::Finalize() {
  _StopAIAssistantInstructionChangedListener();
  _StopAIAssistantLoginFlow();
  _DestroyAIPanel();
  _ClearAIAssistantInstructionCache();
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
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
  }
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end()) {
      it->second.session_alive = false;
      if (!it->second.detached_writeback) {
        m_inline_instruction_states.erase(it);
      }
    }
  }
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
  if (ShouldLogAIAssistantTriggerProbe(m_ai_config, keyEvent)) {
    AppendAIAssistantInfoLogLine(
        "AI trigger probe: keycode=" + std::to_string(keyEvent.keycode) +
        ", mask=" + std::to_string(keyEvent.mask) +
        ", ipc_id=" + std::to_string(ipc_id) +
        ", enabled=" + std::to_string(m_ai_config.enabled ? 1 : 0) +
        ", trigger_hotkey=" + m_ai_config.trigger_hotkey +
        ", trigger_keycode=" +
        std::to_string(m_ai_config.trigger_binding.keycode) +
        ", trigger_modifiers=" +
        std::to_string(m_ai_config.trigger_binding.modifiers));
    LOG(INFO) << "AI trigger probe: keycode=" << keyEvent.keycode
              << ", mask=" << keyEvent.mask << ", ipc_id=" << ipc_id
              << ", enabled=" << m_ai_config.enabled
              << ", trigger_hotkey=" << m_ai_config.trigger_hotkey
              << ", trigger_keycode=" << m_ai_config.trigger_binding.keycode
              << ", trigger_modifiers="
              << m_ai_config.trigger_binding.modifiers;
  }
  if (m_disabled)
    return FALSE;
  // Let WebView2 handle Escape key for state navigation (select→input→generating)
  // Frontend will call onBack() or onClose() based on current panel state
  if (!(keyEvent.mask & ibus::RELEASE_MASK) &&
      keyEvent.keycode == ibus::Keycode::Escape) {
    DLOG(INFO) << "Escape key pressed, passing to WebView2";
    bool panel_open = false;
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      panel_open = m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
                   IsWindowVisible(m_ai_panel.panel_hwnd);
    }
    if (panel_open) {
      // Pass Escape to WebView2; frontend handles state navigation
      // Return FALSE to pass key to WebView2, and return early to avoid Rime processing
      return FALSE;
    }
  }
  if (_TryProcessAIAssistantTrigger(keyEvent, ipc_id, eat)) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
  const bool is_inline_control_key =
      IsInlineInstructionSubmitKey(keyEvent) ||
      IsInlineInstructionCancelKey(keyEvent);
  if (is_inline_control_key &&
      _TryHandleInlineInstructionKey(keyEvent, ipc_id, eat)) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
  if (_TryHandleInjectedCandidateSelectKey(keyEvent, ipc_id, eat)) {
    m_active_session = ipc_id;
    return TRUE;
  }
  if (_TryHandleInstructionLookupCandidateSelectKey(keyEvent, ipc_id, eat)) {
    m_active_session = ipc_id;
    return TRUE;
  }
  RimeSessionId session_id = to_session_id(ipc_id);
  if (IsSystemCommandConfirmKey(keyEvent)) {
    bool panel_open = false;
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      panel_open = m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
                   IsWindowVisible(m_ai_panel.panel_hwnd);
    }
    const char* input_text = rime_api->get_input(session_id);
    if (panel_open &&
        ((input_text && IsSystemCommandInputText(input_text)) ||
         HasSelectedSystemCommandCandidate(session_id))) {
      rime_api->clear_composition(session_id);
      _Respond(ipc_id, eat);
      _UpdateUI(ipc_id);
      m_active_session = ipc_id;
      return TRUE;
    }
  }
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
  if (!is_inline_control_key &&
      _TryHandleInlineInstructionKey(keyEvent, ipc_id, eat)) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
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
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end() &&
        it->second.phase == InlineInstructionPhase::kEditing) {
      m_inline_instruction_states.erase(it);
    }
  }
  rime_api->commit_composition(to_session_id(ipc_id));
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
}

void RimeWithWeaselHandler::ClearComposition(WeaselSessionId ipc_id) {
  DLOG(INFO) << "Clear composition: ipc_id = " << ipc_id;
  if (m_disabled)
    return;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end() &&
        it->second.phase == InlineInstructionPhase::kEditing) {
      m_inline_instruction_states.erase(it);
    }
  }
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
  if (_TrySelectInjectedAIAssistantCandidate(index, ipc_id)) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return;
  }
  AIPanelInstitutionOption lookup_option;
  if (_TryResolveInstructionLookupCandidate(index, false, ipc_id,
                                            &lookup_option)) {
    const bool handled = lookup_option.IsSystemCommand()
                             ? _ExecuteInjectedSystemCommand(ipc_id,
                                                             lookup_option)
                             : _EnterInlineInstructionForOption(ipc_id,
                                                                lookup_option);
    if (handled) {
      _UpdateUI(ipc_id);
      m_active_session = ipc_id;
      return;
    }
  }

  size_t rime_index = index;
  bool instruction_lookup_mode = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it != m_ai_injected_candidates.end()) {
      instruction_lookup_mode = it->second.instruction_lookup_mode;
      const size_t injected_count = GetInjectedCandidateVisibleCount(it->second);
      if (!instruction_lookup_mode && injected_count != 0 &&
          index >= injected_count) {
        rime_index = index - injected_count;
        m_ai_injected_candidates.erase(it);
      }
    }
  }
  if (instruction_lookup_mode) {
    return;
  }
  rime_api->select_candidate_on_current_page(to_session_id(ipc_id), rime_index);
}

bool RimeWithWeaselHandler::HighlightCandidateOnCurrentPage(
    size_t index,
    WeaselSessionId ipc_id,
    EatLine eat) {
  DLOG(INFO) << "highlight candidate on current page, ipc_id = " << ipc_id
             << ", index = " << index;
  size_t rime_index = index;
  bool instruction_lookup_mode = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it != m_ai_injected_candidates.end()) {
      instruction_lookup_mode = it->second.instruction_lookup_mode;
      const size_t injected_count = GetInjectedCandidateVisibleCount(it->second);
      if (injected_count != 0) {
        if (index < injected_count) {
          return true;
        }
        rime_index = index - injected_count;
      }
    }
  }
  if (instruction_lookup_mode) {
    return true;
  }
  bool res = rime_api->highlight_candidate_on_current_page(
      to_session_id(ipc_id), rime_index);
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  return res;
}

bool RimeWithWeaselHandler::ChangePage(bool backward,
                                       WeaselSessionId ipc_id,
                                       EatLine eat) {
  DLOG(INFO) << "change page, ipc_id = " << ipc_id
             << (backward ? "backward" : "foreward");
  bool handled_injected_page = false;
  bool has_injected_instruction_lookup = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it != m_ai_injected_candidates.end() &&
        it->second.instruction_lookup_mode && it->second.page_size > 0) {
      has_injected_instruction_lookup = true;
      const size_t total_count = it->second.options.size();
      const size_t total_pages =
          total_count == 0
              ? 0
              : (total_count + it->second.page_size - 1) / it->second.page_size;
      if (total_pages > 1) {
        if (backward) {
          if (it->second.current_page == 0) {
            return false;
          }
          --it->second.current_page;
        } else {
          if (it->second.current_page + 1 >= total_pages) {
            return false;
          }
          ++it->second.current_page;
        }
        handled_injected_page = true;
      }
    }
  }
  if (handled_injected_page) {
    _Respond(ipc_id, eat);
    _UpdateUI(ipc_id);
    return true;
  }
  if (has_injected_instruction_lookup) {
    return false;
  }
  bool res = rime_api->change_page(to_session_id(ipc_id), backward);
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  return res;
}

BOOL RimeWithWeaselHandler::SyncSession(WeaselSessionId ipc_id, EatLine eat) {
  if (ipc_id == 0) {
    return FALSE;
  }
  std::wstring commit_text;
  std::wstring error_text;
  bool has_error = false;
  bool has_completion = false;
  bool streamed_writeback = false;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end() &&
        it->second.phase == InlineInstructionPhase::kRequesting &&
        (it->second.request_completed || it->second.has_error ||
         !it->second.result_text.empty())) {
      has_completion = true;
      has_error = it->second.has_error;
      commit_text = it->second.result_text;
      error_text = it->second.error_text;
      streamed_writeback = it->second.streamed_writeback;
      m_inline_instruction_states.erase(it);
    }
  }
  if (has_completion) {
    if (has_error) {
      if (!error_text.empty()) {
        LOG(WARNING) << "Inline AI request failed: " << wtou8(error_text);
      }
      _Respond(ipc_id, eat);
    } else if (!streamed_writeback && !commit_text.empty()) {
      _Respond(ipc_id, eat, &commit_text);
    } else {
      _Respond(ipc_id, eat);
    }
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
  _Respond(ipc_id, eat);
  _UpdateUI(ipc_id);
  m_active_session = ipc_id;
  return TRUE;
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
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end()) {
      if (it->second.phase == InlineInstructionPhase::kEditing) {
        m_inline_instruction_states.erase(it);
      } else if (it->second.phase == InlineInstructionPhase::kRequesting) {
        it->second.detached_writeback = true;
      }
    }
  }
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
                                              RimeContext& ctx,
                                              WeaselSessionId ipc_id) {
  bool panel_open = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_open = m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
                 IsWindowVisible(m_ai_panel.panel_hwnd);
  }
  if (panel_open) {
    std::vector<int> visible_indexes;
    visible_indexes.reserve(ctx.menu.num_candidates);
    for (int i = 0; i < ctx.menu.num_candidates; ++i) {
      if (!IsSystemCommandCandidate(ctx.menu.candidates[i])) {
        visible_indexes.push_back(i);
      }
    }
    if (visible_indexes.size() !=
        static_cast<size_t>(ctx.menu.num_candidates)) {
      cinfo.candies.clear();
      cinfo.comments.clear();
      cinfo.labels.clear();
      cinfo.candies.reserve(visible_indexes.size());
      cinfo.comments.reserve(visible_indexes.size());
      cinfo.labels.reserve(visible_indexes.size());
      for (int source_index : visible_indexes) {
        cinfo.candies.push_back(
            Text(escape_string(u8tow(ctx.menu.candidates[source_index].text))));
        if (ctx.menu.candidates[source_index].comment) {
          cinfo.comments.push_back(Text(escape_string(
              u8tow(ctx.menu.candidates[source_index].comment))));
        } else {
          cinfo.comments.push_back(Text());
        }
        if (RIME_STRUCT_HAS_MEMBER(ctx, ctx.select_labels) &&
            ctx.select_labels) {
          cinfo.labels.push_back(
              Text(escape_string(u8tow(ctx.select_labels[source_index]))));
        } else if (ctx.menu.select_keys) {
          cinfo.labels.push_back(Text(escape_string(
              std::wstring(1, ctx.menu.select_keys[source_index]))));
        } else {
          cinfo.labels.push_back(
              Text(std::to_wstring((source_index + 1) % 10)));
        }
      }
      cinfo.highlighted = 0;
      for (size_t i = 0; i < visible_indexes.size(); ++i) {
        if (visible_indexes[i] >= ctx.menu.highlighted_candidate_index) {
          cinfo.highlighted = static_cast<int>(i);
          break;
        }
      }
      cinfo.currentPage = ctx.menu.page_no;
      cinfo.is_last_page = ctx.menu.is_last_page;
      return;
    }
  }
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
  _TryInjectAIAssistantCandidates(ipc_id, ctx, &cinfo);
}

bool RimeWithWeaselHandler::_TryBuildInstructionLookupCandidates(
    WeaselSessionId ipc_id,
    const std::string& input_text,
    int page_size,
    CandidateInfo* cinfo) {
  if (!cinfo) {
    return false;
  }

  const std::string query =
      NormalizeInstructionLookupAscii(ExtractInstructionLookupQuery(
          input_text, m_ai_config.instruction_lookup_prefix));
  if (query.empty()) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }

  std::vector<AIPanelInstitutionOption> all_options =
      SnapshotBuiltinAIAssistantInstructions();
  {
    const std::vector<AIPanelInstitutionOption> dynamic_options =
        m_ai_instructions.Snapshot();
    all_options.insert(all_options.end(), dynamic_options.begin(),
                       dynamic_options.end());
  }
  if (all_options.empty()) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }

  constexpr size_t kMaxInstructionLookupPageSize = 10;
  size_t resolved_page_size =
      page_size > 0 ? static_cast<size_t>(page_size)
                    : kMaxInstructionLookupPageSize;
  resolved_page_size =
      (std::min)(resolved_page_size, kMaxInstructionLookupPageSize);
  if (resolved_page_size == 0) {
    resolved_page_size = 1;
  }

  size_t preserved_page = 0;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it != m_ai_injected_candidates.end() &&
        it->second.instruction_lookup_mode &&
        it->second.lookup_query == query) {
      preserved_page = it->second.current_page;
      if (it->second.page_size > 0) {
        resolved_page_size = it->second.page_size;
      }
    }
  }

  std::vector<InstructionLookupRankedOption> ranked_options;
  ranked_options.reserve(all_options.size());
  std::set<std::wstring> seen_keys;
  size_t order = 0;
  for (const auto& option : all_options) {
    std::wstring unique_key = option.id.empty() ? option.name : option.id;
    if (unique_key.empty() || seen_keys.find(unique_key) != seen_keys.end()) {
      ++order;
      continue;
    }
    seen_keys.insert(unique_key);

    std::vector<std::string> aliases;
    std::set<std::string> seen_aliases;
    AddInstructionLookupAlias(wtou8(option.name), &aliases, &seen_aliases);
    AddInstructionLookupAliasesFromCodeText(option.lookup_text, &aliases,
                                            &seen_aliases);
    AddInstructionLookupAliasesFromCodeText(option.id, &aliases,
                                            &seen_aliases);
    AddInstructionLookupAliasesFromCodeText(option.app_key, &aliases,
                                            &seen_aliases);
    AddInstructionLookupAliasesFromCodeText(option.system_command_id, &aliases,
                                            &seen_aliases);

    std::string pinyin_full;
    std::string pinyin_initials;
    if (TryBuildInstructionLookupPinyinAliases(option.name, &pinyin_full,
                                               &pinyin_initials)) {
      AddInstructionLookupAlias(pinyin_full, &aliases, &seen_aliases);
      AddInstructionLookupAlias(pinyin_initials, &aliases, &seen_aliases);
    }

    int best_score = 0;
    for (const auto& alias : aliases) {
      const int score = ScoreInstructionLookupAlias(query, alias);
      if (score > best_score) {
        best_score = score;
      }
    }
    if (best_score <= 0) {
      ++order;
      continue;
    }

    InstructionLookupRankedOption ranked;
    ranked.option = option;
    ranked.score = best_score;
    ranked.order = order++;
    ranked_options.push_back(std::move(ranked));
  }

  if (ranked_options.empty()) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }

  std::stable_sort(ranked_options.begin(), ranked_options.end(),
                   [](const InstructionLookupRankedOption& left,
                      const InstructionLookupRankedOption& right) {
                     if (left.score != right.score) {
                       return left.score > right.score;
                     }
                     if (left.option.name.size() != right.option.name.size()) {
                       return left.option.name.size() < right.option.name.size();
                     }
                     return left.order < right.order;
                   });

  const size_t total_candidate_count = ranked_options.size();
  const size_t total_pages =
      (total_candidate_count + resolved_page_size - 1) / resolved_page_size;
  const size_t current_page =
      total_pages == 0 ? 0 : (std::min)(preserved_page, total_pages - 1);
  const size_t page_begin = current_page * resolved_page_size;
  const size_t candidate_count =
      (std::min)(resolved_page_size, total_candidate_count - page_begin);

  cinfo->candies.clear();
  cinfo->comments.clear();
  cinfo->labels.clear();
  cinfo->candies.reserve(candidate_count);
  cinfo->comments.reserve(candidate_count);
  cinfo->labels.reserve(candidate_count);

  AIAssistantInjectedCandidateState state;
  state.instruction_lookup_mode = true;
  state.lookup_query = query;
  state.current_page = current_page;
  state.page_size = resolved_page_size;
  state.options.reserve(total_candidate_count);
  for (const auto& ranked_option : ranked_options) {
    state.options.push_back(ranked_option.option);
  }

  for (size_t i = 0; i < candidate_count; ++i) {
    const auto& option = ranked_options[page_begin + i].option;
    cinfo->candies.push_back(Text(escape_string(option.name)));
    cinfo->comments.push_back(
        Text(option.IsSystemCommand()
                 ? L"系统指令"
                 : (IsBuiltinAIAssistantInstructionId(option.id) ? L"默认指令"
                                                                 : L"AI指令")));
    cinfo->labels.push_back(Text(std::to_wstring((i + 1) % 10)));
  }

  cinfo->highlighted = 0;
  cinfo->currentPage = static_cast<int>(current_page);
  cinfo->totalPages = static_cast<int>(total_pages);
  cinfo->is_last_page = current_page + 1 >= total_pages;
  state.rime_highlighted = 0;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates[ipc_id] = std::move(state);
  }
  return true;
}

bool RimeWithWeaselHandler::_TryMatchAIAssistantInstructionOption(
    const std::wstring& text,
    AIPanelInstitutionOption* option) const {
  return MatchBuiltinAIAssistantInstruction(text, option) ||
         m_ai_instructions.MatchExactName(text, option);
}

bool RimeWithWeaselHandler::_TryResolveInstructionLookupCandidate(
    size_t candidate_index,
    bool use_highlighted,
    WeaselSessionId ipc_id,
    AIPanelInstitutionOption* option) {
  if (!option || ipc_id == 0) {
    return false;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text = session_id != 0 ? rime_api->get_input(session_id)
                                           : nullptr;
  if (ShouldSuppressSpecialInstructionCandidates(
          m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
          m_ai_config.instruction_lookup_prefix, input_text)) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }
  if (!input_text ||
      !IsInstructionLookupInputText(input_text,
                                    m_ai_config.instruction_lookup_prefix)) {
    return false;
  }

  RIME_STRUCT(RimeContext, ctx);
  if (!rime_api->get_context(session_id, &ctx)) {
    return false;
  }

  bool matched = false;
  if (ctx.menu.num_candidates > 0 && ctx.menu.candidates) {
    size_t resolved_index = candidate_index;
    if (use_highlighted) {
      resolved_index = static_cast<size_t>(max(
          0, min(ctx.menu.highlighted_candidate_index, ctx.menu.num_candidates - 1)));
    }
    if (resolved_index < static_cast<size_t>(ctx.menu.num_candidates) &&
        ctx.menu.candidates[resolved_index].text) {
      matched = _TryMatchAIAssistantInstructionOption(
          u8tow(ctx.menu.candidates[resolved_index].text), option);
    }
  }

  rime_api->free_context(&ctx);
  return matched;
}

bool RimeWithWeaselHandler::_TryInjectAIAssistantCandidates(
    WeaselSessionId ipc_id,
    RimeContext& ctx,
    CandidateInfo* cinfo) {
  if (!cinfo || !m_ai_config.enabled || ipc_id == 0) {
    if (ipc_id != 0) {
      std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
      m_ai_injected_candidates.erase(ipc_id);
    }
    return false;
  }
  if (HasActiveInlineInstructionState(m_inline_instruction_mutex,
                                      m_inline_instruction_states, ipc_id)) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }
  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text =
      session_id != 0 ? rime_api->get_input(session_id) : nullptr;
  if (input_text &&
      IsInstructionLookupInputText(input_text,
                                   m_ai_config.instruction_lookup_prefix)) {
    return _TryBuildInstructionLookupCandidates(ipc_id, input_text,
                                                ctx.menu.page_size, cinfo);
  }
  if (ctx.menu.page_no != 0 || ctx.menu.num_candidates <= 0) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
        IsWindowVisible(m_ai_panel.panel_hwnd)) {
      std::lock_guard<std::mutex> injected_lock(m_ai_injected_candidates_mutex);
      m_ai_injected_candidates.erase(ipc_id);
      return false;
    }
  }

  AIPanelInstitutionOption matched_option;
  for (int i = 0; i < ctx.menu.num_candidates; ++i) {
    const char* candidate_text = ctx.menu.candidates[i].text;
    if (!candidate_text) {
      continue;
    }
    const std::wstring candidate_name = u8tow(candidate_text);
    if (_TryMatchAIAssistantInstructionOption(candidate_name,
                                              &matched_option)) {
      break;
    }
    matched_option = AIPanelInstitutionOption();
  }

  if (matched_option.name.empty()) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }

  cinfo->candies.insert(cinfo->candies.begin(),
                        Text(escape_string(matched_option.name)));
  cinfo->comments.insert(
      cinfo->comments.begin(),
      Text(matched_option.IsSystemCommand()
               ? L"系统指令"
               : (IsBuiltinAIAssistantInstructionId(matched_option.id)
                      ? L"默认指令"
                      : L"AI指令")));
  if (!cinfo->labels.empty()) {
    cinfo->labels.insert(cinfo->labels.begin(), cinfo->labels.front());
    for (size_t i = 1; i < cinfo->labels.size(); ++i) {
      cinfo->labels[i].str = std::to_wstring((i + 1) % 10);
    }
  }
  cinfo->highlighted = 0;

  AIAssistantInjectedCandidateState state;
  state.options.push_back(matched_option);
  state.rime_highlighted = ctx.menu.highlighted_candidate_index;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates[ipc_id] = std::move(state);
  }
  return true;
}

bool RimeWithWeaselHandler::_TryHandleInstructionLookupCandidateSelectKey(
    KeyEvent keyEvent,
    WeaselSessionId ipc_id,
    EatLine eat) {
  if (!m_ai_config.enabled || ipc_id == 0) {
    return false;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text = session_id != 0 ? rime_api->get_input(session_id)
                                           : nullptr;
  if (ShouldSuppressSpecialInstructionCandidates(
          m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
          m_ai_config.instruction_lookup_prefix, input_text)) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }
  if (!input_text ||
      !IsInstructionLookupInputText(input_text,
                                    m_ai_config.instruction_lookup_prefix)) {
    return false;
  }

  if (IsSystemCommandConfirmKey(keyEvent)) {
    const bool handled = _TrySelectInjectedAIAssistantCandidate(0, ipc_id);
    if (handled) {
      _Respond(ipc_id, eat);
      _UpdateUI(ipc_id);
    }
    return handled;
  }

  size_t candidate_index = 0;
  if (!TryResolveCandidateSelectIndex(keyEvent, &candidate_index)) {
    return false;
  }
  AIPanelInstitutionOption option;
  if (!_TryResolveInstructionLookupCandidate(candidate_index, false,
                                             ipc_id, &option)) {
    return false;
  }

  const bool handled = option.IsSystemCommand()
                           ? _ExecuteInjectedSystemCommand(ipc_id, option)
                           : _EnterInlineInstructionForOption(ipc_id, option);
  if (handled) {
    _Respond(ipc_id, eat);
    _UpdateUI(ipc_id);
  }
  return handled;
}

bool RimeWithWeaselHandler::_TrySelectInjectedAIAssistantCandidate(
    size_t index,
    WeaselSessionId ipc_id) {
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
        IsWindowVisible(m_ai_panel.panel_hwnd)) {
      return false;
    }
  }
  AIPanelInstitutionOption option;
  bool instruction_lookup_mode = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it == m_ai_injected_candidates.end()) {
      return false;
    }
    const size_t visible_count = GetInjectedCandidateVisibleCount(it->second);
    if (index >= visible_count) {
      return false;
    }
    instruction_lookup_mode = it->second.instruction_lookup_mode;
    option = it->second.options[ResolveInjectedCandidateIndex(it->second, index)];
    m_ai_injected_candidates.erase(it);
  }
  if (option.IsSystemCommand()) {
    return _ExecuteInjectedSystemCommand(ipc_id, option);
  }
  if (instruction_lookup_mode) {
    return _EnterInlineInstructionForOption(ipc_id, option);
  }
  return _OpenAIPanelForInstruction(ipc_id, option);
}

bool RimeWithWeaselHandler::_TryHandleInjectedCandidateSelectKey(
    KeyEvent keyEvent,
    WeaselSessionId ipc_id,
    EatLine eat) {
  size_t selected_index = 0;
  if (!TryResolveCandidateSelectIndex(keyEvent, &selected_index)) {
    return false;
  }

  size_t injected_count = 0;
  bool instruction_lookup_mode = false;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    const auto it = m_ai_injected_candidates.find(ipc_id);
    if (it == m_ai_injected_candidates.end()) {
      return false;
    }
    instruction_lookup_mode = it->second.instruction_lookup_mode;
    injected_count = GetInjectedCandidateVisibleCount(it->second);
    if (injected_count == 0) {
      return false;
    }
  }

  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
        IsWindowVisible(m_ai_panel.panel_hwnd) &&
        selected_index < injected_count) {
      return true;
    }
  }

  if (selected_index < injected_count) {
    const bool handled =
        _TrySelectInjectedAIAssistantCandidate(selected_index, ipc_id);
    if (handled) {
      _Respond(ipc_id, eat);
      _UpdateUI(ipc_id);
    }
    return handled;
  }
  if (instruction_lookup_mode) {
    return true;
  }

  const size_t rime_index = selected_index - injected_count;
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
  }
  if (rime_api->select_candidate_on_current_page(to_session_id(ipc_id),
                                                 rime_index)) {
    _Respond(ipc_id, eat);
    _UpdateUI(ipc_id);
  }
  return true;
}

bool RimeWithWeaselHandler::_EnterInlineInstructionForOption(
    WeaselSessionId ipc_id,
    const AIPanelInstitutionOption& option) {
  if (ipc_id == 0 || option.IsSystemCommand()) {
    return false;
  }
  if (option.id.empty() && option.name.empty() &&
      option.template_content.empty()) {
    return false;
  }
  if (!_EnsureAIAssistantLogin()) {
    return true;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);
  rime_api->clear_composition(session_id);

  HWND target_hwnd = GetFocus();
  if (!target_hwnd || !IsWindow(target_hwnd)) {
    target_hwnd = GetForegroundWindow();
  }

  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
  }
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    InlineInstructionState& state = m_inline_instruction_states[ipc_id];
    state.Reset();
    state.session_alive = true;
    state.detached_writeback = false;
    state.phase = InlineInstructionPhase::kEditing;
    state.slash_mode = true;
    state.target_hwnd = target_hwnd;
    state.has_selected_option = true;
    state.selected_option = option;
  }

  DLOG(INFO) << "inline_ai: enter from instruction lookup ipc_id=" << ipc_id
             << " option_id=" << wtou8(option.id)
             << " option_name=" << wtou8(option.name);
  return true;
}

bool RimeWithWeaselHandler::_OpenAIPanelForInstruction(
    WeaselSessionId ipc_id,
    const AIPanelInstitutionOption& option) {
  if (option.id.empty()) {
    return false;
  }
  if (!_EnsureAIAssistantLogin()) {
    return true;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);
  rime_api->clear_composition(session_id);

  HWND target_hwnd = GetFocus();
  if (!target_hwnd || !IsWindow(target_hwnd)) {
    target_hwnd = GetForegroundWindow();
  }

  std::vector<AIPanelInstitutionOption> options = m_ai_instructions.Snapshot();
  bool found = false;
  for (const auto& cached_option : options) {
    if (cached_option.id == option.id) {
      found = true;
      break;
    }
  }
  if (!found) {
    options.push_back(option);
  }

  bool relogin_started = false;
  std::string error_message;
  if (m_ai_instructions.Empty()) {
    int http_status_code = 0;
    if (!_RefreshAIAssistantInstructionCacheSync(&error_message,
                                                 &http_status_code)) {
      if (http_status_code == 401) {
        LOG(INFO) << "AI injected instruction selection blocked by 401.";
        return true;
      }
    }
    options = m_ai_instructions.Snapshot();
    found = false;
    for (const auto& cached_option : options) {
      if (cached_option.id == option.id) {
        found = true;
        break;
      }
    }
    if (!found) {
      options.push_back(option);
    }
  }

  bool options_ready = false;
  if (!_PrepareAIPanelInstitutionOptionsForOpen(
          &options, &options_ready, &relogin_started, &error_message)) {
    return true;
  }
  if (relogin_started) {
    return true;
  }

  if (!_OpenAIPanel(ipc_id, target_hwnd, &options, options_ready,
                    option.id)) {
    return true;
  }

  _SetAIPanelInstitutionSelection(option.id);
  _ResetAIPanelOutput();
  _SetAIPanelStatus(L"上下文已就绪，请在前端面板中发起请求。");
  return true;
}

bool RimeWithWeaselHandler::_ExecuteInjectedSystemCommand(
    WeaselSessionId ipc_id,
    const AIPanelInstitutionOption& option) {
  if (!option.IsSystemCommand() || !m_system_command_callback) {
    return false;
  }

  const std::string command_id = wtou8(option.system_command_id);
  if (!IsAllowedSystemCommandId(command_id)) {
    return false;
  }

  const RimeSessionId session_id = to_session_id(ipc_id);
  if (session_id != 0) {
    rime_api->clear_composition(session_id);
  }

  SystemCommandLaunchRequest request;
  request.command_id = command_id;
  request.preferred_output_dir = _ReadSystemCommandOutputDir(ipc_id);
  m_system_command_callback(request);
  return true;
}

void RimeWithWeaselHandler::StartMaintenance() {
  m_session_status_map.clear();
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.clear();
  }
  Finalize();
  _UpdateUI(0);
}

void RimeWithWeaselHandler::EndMaintenance() {
  if (m_disabled) {
    Initialize();
    _UpdateUI(0);
  }
  m_session_status_map.clear();
  {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.clear();
  }
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
  Bool inline_instruction_enabled = True;
  if (rime_api->config_get_bool(config,
                                "ai_assistant/inline_instruction_enabled",
                                &inline_instruction_enabled)) {
    m_ai_config.inline_instruction_enabled = !!inline_instruction_enabled;
  }
  ReadConfigString(config, "ai_assistant/trigger_hotkey",
                   &m_ai_config.trigger_hotkey);
  ReadConfigString(config, "ai_assistant/inline_instruction_prefix",
                   &m_ai_config.inline_instruction_prefix);
  ReadConfigString(config, "ai_assistant/instruction_lookup_prefix",
                   &m_ai_config.instruction_lookup_prefix);
  ReadConfigString(config, "ai_assistant/ai_api_base",
                   &m_ai_config.ai_api_base);
  ReadConfigString(config, "ai_assistant/debug_dump_path",
                   &m_ai_config.debug_dump_path);
  ReadConfigString(config, "ai_assistant/panel_url", &m_ai_config.panel_url);
  ReadConfigString(config, "ai_assistant/official_url",
                   &m_ai_config.official_url);
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
  ReadConfigString(config, "ai_assistant/mqtt_ins_changed_topic",
                   &m_ai_config.mqtt_ins_changed_topic);
  ReadConfigString(config, "ai_assistant/mqtt_username",
                   &m_ai_config.mqtt_username);
  ReadConfigString(config, "ai_assistant/mqtt_password",
                   &m_ai_config.mqtt_password);
  ReadConfigInt(config, "ai_assistant/max_history_chars",
                &m_ai_config.max_history_chars);
  ReadConfigInt(config, "ai_assistant/timeout_ms", &m_ai_config.timeout_ms);
  ReadConfigInt(config, "ai_assistant/mqtt_timeout_ms",
                &m_ai_config.mqtt_timeout_ms);

  if (m_ai_config.trigger_hotkey.empty()) {
    m_ai_config.trigger_hotkey = "Control+3";
  }
  if (m_ai_config.inline_instruction_prefix.size() != 1) {
    m_ai_config.inline_instruction_prefix = "/";
  }
  if (!IsValidInstructionLookupPrefix(m_ai_config.instruction_lookup_prefix)) {
    if (!m_ai_config.instruction_lookup_prefix.empty()) {
      LOG(WARNING) << "Invalid ai_assistant/instruction_lookup_prefix: "
                   << m_ai_config.instruction_lookup_prefix
                   << "; fallback to "
                   << kDefaultInstructionLookupInputPrefix << ".";
    }
    m_ai_config.instruction_lookup_prefix =
        kDefaultInstructionLookupInputPrefix;
  }
  if (!TryParseAiHotkeyConfig(m_ai_config.trigger_hotkey,
                              &m_ai_config.trigger_binding) ||
      !ValidateAiHotkeyBinding(m_ai_config.trigger_binding).ok) {
    LOG(WARNING) << "Invalid ai_assistant/trigger_hotkey: "
                 << m_ai_config.trigger_hotkey
                 << "; fallback to Control+3.";
    m_ai_config.trigger_hotkey = "Control+3";
    m_ai_config.trigger_binding = AiHotkeyBinding('3', ibus::CONTROL_MASK);
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
  if (m_ai_config.mqtt_ins_changed_topic.empty()) {
    m_ai_config.mqtt_ins_changed_topic =
        kDefaultAIAssistantInstructionChangedTopic;
  } else if (!IsValidAIAssistantInstructionChangedTopic(
                 m_ai_config.mqtt_ins_changed_topic)) {
    LOG(WARNING) << "Invalid ai_assistant/mqtt_ins_changed_topic: "
                 << m_ai_config.mqtt_ins_changed_topic
                 << "; fallback to "
                 << kDefaultAIAssistantInstructionChangedTopic << ".";
    m_ai_config.mqtt_ins_changed_topic =
        kDefaultAIAssistantInstructionChangedTopic;
  }

  m_input_content_store.SetLimits(
      static_cast<size_t>(max(1024, m_ai_config.max_history_chars * 2)),
      64, 16);

  AppendAIAssistantInfoLogLine(
      "AI config loaded: enabled=" +
      std::to_string(m_ai_config.enabled ? 1 : 0) +
      ", login_required=" +
      std::to_string(m_ai_config.login_required ? 1 : 0) +
      ", trigger_hotkey=" + m_ai_config.trigger_hotkey +
      ", trigger_keycode=" +
      std::to_string(m_ai_config.trigger_binding.keycode) +
      ", trigger_modifiers=" +
      std::to_string(m_ai_config.trigger_binding.modifiers) +
      ", panel_url=" + m_ai_config.panel_url);
  LOG(INFO) << "AI config loaded: enabled=" << m_ai_config.enabled
            << ", login_required=" << m_ai_config.login_required
            << ", trigger_hotkey=" << m_ai_config.trigger_hotkey
            << ", trigger_keycode=" << m_ai_config.trigger_binding.keycode
            << ", trigger_modifiers=" << m_ai_config.trigger_binding.modifiers
            << ", panel_url=" << m_ai_config.panel_url;
}

void RimeWithWeaselHandler::_StartAIAssistantInstructionChangedListener() {
  _StopAIAssistantInstructionChangedListener();
  if (!m_ai_config.enabled) {
    return;
  }
  if (m_ai_config.mqtt_url.empty() ||
      m_ai_config.mqtt_ins_changed_topic.empty()) {
    return;
  }

  m_ai_inst_changed_stop.store(false);
  m_ai_inst_changed_thread = std::thread([this]() {
    _RunAIAssistantInstructionChangedListener();
  });
}

void RimeWithWeaselHandler::_StopAIAssistantInstructionChangedListener() {
  m_ai_inst_changed_stop.store(true);
  std::vector<HINTERNET> handles_to_close;
  {
    std::lock_guard<std::mutex> lock(m_ai_inst_changed_handle_mutex);
    auto take_handle = [&handles_to_close](void** handle) {
      if (handle && *handle) {
        handles_to_close.push_back(static_cast<HINTERNET>(*handle));
        *handle = nullptr;
      }
    };
    take_handle(&m_ai_inst_changed_websocket_handle);
    take_handle(&m_ai_inst_changed_request_handle);
    take_handle(&m_ai_inst_changed_connection_handle);
    take_handle(&m_ai_inst_changed_session_handle);
  }
  for (HINTERNET handle : handles_to_close) {
    WinHttpCloseHandle(handle);
  }
  if (m_ai_inst_changed_thread.joinable()) {
    m_ai_inst_changed_thread.join();
  }
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
    const std::filesystem::path state_path(
        ResolveAIAssistantLoginStatePath(m_ai_config));
    std::error_code exists_error;
    const bool state_exists = std::filesystem::exists(state_path, exists_error);
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    const bool has_memory_login = !m_ai_login_token.empty() ||
                                  !m_ai_login_refresh_token.empty() ||
                                  !m_ai_login_tenant_id.empty();
    if (!state_exists && has_memory_login) {
      AppendAIAssistantInfoLogLine(
          "AI login state file missing, clearing in-memory login state.");
      LOG(INFO) << "AI login state file missing, clearing in-memory login state.";
      m_ai_login_token.clear();
      m_ai_login_tenant_id.clear();
      m_ai_login_refresh_token.clear();
      ClearAIAssistantUserInfoCache(m_ai_config);
      return false;
    }
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

bool RimeWithWeaselHandler::_ResolveAIAssistantLoginState(
    std::string* token,
    std::string* tenant_id,
    std::string* refresh_token,
    std::string* error_message) {
  if (token) {
    token->clear();
  }
  if (tenant_id) {
    tenant_id->clear();
  }
  if (refresh_token) {
    refresh_token->clear();
  }
  if (error_message) {
    error_message->clear();
  }

  const AIAssistantConfig config = m_ai_config;
  std::string local_token;
  std::string local_tenant_id;
  std::string local_refresh_token;
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    local_token = m_ai_login_token;
    local_tenant_id = m_ai_login_tenant_id;
    local_refresh_token = m_ai_login_refresh_token;
  }

  if (local_tenant_id.empty() ||
      (local_token.empty() && local_refresh_token.empty())) {
    std::string file_token;
    std::string file_tenant_id;
    std::string file_refresh_token;
    LoadAIAssistantLoginIdentity(config, &file_token, &file_tenant_id,
                                 &file_refresh_token);
    if (!file_token.empty()) {
      local_token = file_token;
    }
    if (!file_tenant_id.empty()) {
      local_tenant_id = file_tenant_id;
    }
    if (!file_refresh_token.empty()) {
      local_refresh_token = file_refresh_token;
    }
    if (!file_token.empty() || !file_tenant_id.empty() ||
        !file_refresh_token.empty()) {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      if (!file_token.empty()) {
        m_ai_login_token = file_token;
      }
      if (!file_tenant_id.empty()) {
        m_ai_login_tenant_id = file_tenant_id;
      }
      if (!file_refresh_token.empty()) {
        m_ai_login_refresh_token = file_refresh_token;
      }
    }
  }

  if (local_token.empty() && !local_refresh_token.empty() &&
      !local_tenant_id.empty()) {
    std::string refreshed_token;
    std::string refreshed_refresh_token;
    std::string refresh_endpoint;
    std::string refresh_error;
    const bool refreshed = RefreshAIAssistantAccessToken(
        config, local_refresh_token, local_tenant_id, &refreshed_token,
        &refreshed_refresh_token, &refresh_endpoint, &refresh_error);
    if (refreshed && !refreshed_token.empty()) {
      local_token = refreshed_token;
      if (!refreshed_refresh_token.empty()) {
        local_refresh_token = refreshed_refresh_token;
      }
      std::string client_id_for_state;
      {
        std::lock_guard<std::mutex> lock(m_ai_login_mutex);
        client_id_for_state = m_ai_login_client_id;
      }
      SaveAIAssistantLoginState(config, local_token, local_tenant_id,
                                local_refresh_token, client_id_for_state,
                                std::string(), std::string());
      {
        std::lock_guard<std::mutex> lock(m_ai_login_mutex);
        m_ai_login_token = local_token;
        m_ai_login_tenant_id = local_tenant_id;
        m_ai_login_refresh_token = local_refresh_token;
      }
    } else if (error_message) {
      *error_message = refresh_error;
    }
  }

  if (token) {
    *token = local_token;
  }
  if (tenant_id) {
    *tenant_id = local_tenant_id;
  }
  if (refresh_token) {
    *refresh_token = local_refresh_token;
  }
  return !local_tenant_id.empty() &&
         (!local_token.empty() || !local_refresh_token.empty());
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
    std::string cached_user_id;
    if (LoadAIAssistantCachedUserId(m_ai_config, &cached_user_id) &&
        !cached_user_id.empty()) {
      return true;
    }

    std::string token;
    std::string tenant_id;
    std::string refresh_token;
    std::string resolve_error;
    if (!_ResolveAIAssistantLoginState(&token, &tenant_id, &refresh_token,
                                       &resolve_error)) {
      LOG(WARNING) << "AI ensure login: unable to resolve login state, error="
                   << resolve_error;
      _ForceAIAssistantRelogin();
      return false;
    }

    std::string user_info_error;
    int user_info_status_code = 0;
    const bool user_info_refreshed = RefreshAIAssistantUserInfoCache(
        m_ai_config, token, tenant_id, &user_info_error, &user_info_status_code);
    if (user_info_refreshed) {
      LOG(INFO) << "AI ensure login: user info cache refreshed on demand.";
      _StartAIAssistantInstructionChangedListener();
      return true;
    }

    if (!user_info_error.empty()) {
      LOG(WARNING) << "AI ensure login: user info refresh failed: "
                   << user_info_error;
    }
    if (user_info_status_code == 401) {
      LOG(INFO) << "AI ensure login: auth invalid, restarting login flow.";
      _ForceAIAssistantRelogin();
      return false;
    }
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
  _ClearAIAssistantInstructionCache();

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
  ClearAIAssistantUserInfoCache(m_ai_config);
  _StartAIAssistantInstructionChangedListener();

  const bool started = _StartAIAssistantLoginFlow();
  if (!started) {
    LOG(WARNING) << "AI relogin: unable to start login flow.";
  } else {
    LOG(INFO) << "AI relogin: login flow started by ui.auth.refresh_request.";
  }
  return started;
}

void RimeWithWeaselHandler::_ClearAIAssistantInstructionCache() {
  m_ai_instructions.Clear();
  std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
  m_ai_injected_candidates.clear();
}

bool RimeWithWeaselHandler::_RefreshAIAssistantInstructionCacheSync(
    std::string* error_message,
    int* http_status_code) {
  std::string token;
  std::string tenant_id;
  std::string refresh_token;
  std::string resolve_error;
  if (!_ResolveAIAssistantLoginState(&token, &tenant_id, &refresh_token,
                                     &resolve_error)) {
    _ClearAIAssistantInstructionCache();
    if (error_message) {
      *error_message = resolve_error;
    }
    if (http_status_code) {
      *http_status_code = 0;
    }
    return false;
  }

  std::vector<AIPanelInstitutionOption> options;
  std::string fetch_error;
  int status_code = 0;
  const bool ok = FetchAIAssistantInstitutionOptions(
      m_ai_config, token, tenant_id, &options, &fetch_error, &status_code);
  if (http_status_code) {
    *http_status_code = status_code;
  }

  if (ok) {
    m_ai_instructions.Replace(std::move(options));
    return true;
  }

  _ClearAIAssistantInstructionCache();
  if (error_message) {
    *error_message = fetch_error;
  }
  if (status_code == 401) {
    _ForceAIAssistantRelogin();
  }
  return false;
}

void RimeWithWeaselHandler::_RefreshAIAssistantInstructionCacheAsync() {
  if (!m_ai_config.enabled) {
    _ClearAIAssistantInstructionCache();
    return;
  }

  std::thread([this]() {
    std::string error_message;
    int http_status_code = 0;
    const bool ok = _RefreshAIAssistantInstructionCacheSync(
        &error_message, &http_status_code);
    if (!ok && http_status_code != 401 && !error_message.empty()) {
      LOG(WARNING) << "AI instruction cache refresh failed: "
                   << error_message;
    }

    HWND panel_hwnd = nullptr;
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      panel_hwnd = m_ai_panel.panel_hwnd;
    }
    if (panel_hwnd && IsWindow(panel_hwnd)) {
      PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
    }
  }).detach();
}

void RimeWithWeaselHandler::RunInstallLoginCheck() {
  if (!_EnsureAIAssistantLogin()) {
    return;
  }
  _RefreshAIAssistantInstructionCacheAsync();
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
    ClearAIAssistantUserInfoCache(config);
    std::string user_info_error;
    int user_info_status_code = 0;
    if (!RefreshAIAssistantUserInfoCache(config, token, tenant_id,
                                         &user_info_error,
                                         &user_info_status_code) &&
        !user_info_error.empty()) {
      LOG(WARNING) << "AI user info cache refresh failed: "
                   << user_info_error;
    }
    _StartAIAssistantInstructionChangedListener();
    _RefreshAIAssistantInstructionCacheAsync();
    HWND panel_hwnd = nullptr;
    {
      std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
      panel_hwnd = m_ai_panel.panel_hwnd;
    }
    if (panel_hwnd && IsWindow(panel_hwnd)) {
      PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
      _RefreshAIPanelInstitutionOptions();
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

void RimeWithWeaselHandler::_RunAIAssistantInstructionChangedListener() {
  const AIAssistantConfig config = m_ai_config;
  if (config.mqtt_url.empty() || config.mqtt_ins_changed_topic.empty()) {
    return;
  }

  const ParsedWebSocketUrl ws_url = ParseWebSocketUrl(u8tow(config.mqtt_url));
  if (!ws_url.valid) {
    LOG(WARNING) << "AI instruction changed mqtt_url invalid: "
                 << config.mqtt_url;
    return;
  }

  _RefreshAIAssistantInstructionCacheAsync();

  auto wait_before_reconnect = [this]() {
    return WaitForAIAssistantInstructionChangedStop(
        &m_ai_inst_changed_stop, 2000);
  };
  auto set_tracked_handle = [this](void** tracked_handle, HINTERNET handle) {
    std::lock_guard<std::mutex> lock(m_ai_inst_changed_handle_mutex);
    if (tracked_handle) {
      *tracked_handle = handle;
    }
  };
  auto close_tracked_handle = [this](HINTERNET* handle,
                                     void** tracked_handle) {
    if (!handle || !*handle) {
      return;
    }
    HINTERNET handle_to_close = *handle;
    bool should_close = true;
    {
      std::lock_guard<std::mutex> lock(m_ai_inst_changed_handle_mutex);
      if (tracked_handle && *tracked_handle == handle_to_close) {
        *tracked_handle = nullptr;
      } else if (tracked_handle) {
        should_close = false;
      }
    }
    if (should_close) {
      WinHttpCloseHandle(handle_to_close);
    }
    *handle = nullptr;
  };

  const auto ping_packet = BuildMqttPingReqPacket();
  while (!m_ai_inst_changed_stop.load()) {
    std::string permission_update_token;
    std::string permission_update_tenant_id;
    {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      permission_update_token = m_ai_login_token;
      permission_update_tenant_id = m_ai_login_tenant_id;
    }
    if (permission_update_tenant_id.empty()) {
      std::string persisted_token;
      std::string persisted_tenant_id;
      if (LoadAIAssistantLoginIdentity(config, &persisted_token,
                                       &persisted_tenant_id, nullptr)) {
        if (permission_update_token.empty()) {
          permission_update_token = persisted_token;
        }
        permission_update_tenant_id = persisted_tenant_id;
      }
    }

    std::string permission_update_topic;
    std::string permission_update_error;
    if (!permission_update_tenant_id.empty()) {
      if (ResolveAIAssistantPermissionUpdateTopic(
              config, permission_update_token, permission_update_tenant_id,
              &permission_update_topic, &permission_update_error)) {
        LOG(INFO) << "AI instruction permission update topic resolved: "
                  << permission_update_topic;
      } else if (!permission_update_error.empty()) {
        LOG(INFO) << "AI instruction permission update topic unavailable: "
                  << permission_update_error;
      }
    }

    HINTERNET session = nullptr;
    HINTERNET connection = nullptr;
    HINTERNET request = nullptr;
    HINTERNET websocket = nullptr;

    session = WinHttpOpen(L"WeaselAIAssistantInstructionChanged/1.0",
                          WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }
    set_tracked_handle(&m_ai_inst_changed_session_handle, session);
    if (m_ai_inst_changed_stop.load()) {
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      break;
    }
    WinHttpSetTimeouts(session, 5000, 5000, 5000, 1000);

    connection = WinHttpConnect(session, ws_url.host.c_str(), ws_url.port, 0);
    if (!connection) {
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }
    set_tracked_handle(&m_ai_inst_changed_connection_handle, connection);
    if (m_ai_inst_changed_stop.load()) {
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      break;
    }

    const DWORD open_flags = ws_url.secure ? WINHTTP_FLAG_SECURE : 0;
    request = WinHttpOpenRequest(connection, L"GET", ws_url.path.c_str(),
                                 nullptr, WINHTTP_NO_REFERER,
                                 WINHTTP_DEFAULT_ACCEPT_TYPES, open_flags);
    if (!request) {
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }
    set_tracked_handle(&m_ai_inst_changed_request_handle, request);
    if (m_ai_inst_changed_stop.load()) {
      close_tracked_handle(&request, &m_ai_inst_changed_request_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      break;
    }

    if (!WinHttpSetOption(request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET,
                          nullptr, 0)) {
      close_tracked_handle(&request, &m_ai_inst_changed_request_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }
    WinHttpAddRequestHeaders(request, L"Sec-WebSocket-Protocol: mqtt", -1,
                             WINHTTP_ADDREQ_FLAG_ADD);

    if (!WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(request, nullptr)) {
      close_tracked_handle(&request, &m_ai_inst_changed_request_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }

    websocket = WinHttpWebSocketCompleteUpgrade(request, 0);
    close_tracked_handle(&request, &m_ai_inst_changed_request_handle);
    if (!websocket) {
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }
    set_tracked_handle(&m_ai_inst_changed_websocket_handle, websocket);
    if (m_ai_inst_changed_stop.load()) {
      close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      break;
    }

    const std::string transport_client_id = GenerateRandomMqttClientId();
    const auto connect_packet = BuildMqttConnectPacket(
        transport_client_id, config.mqtt_username, config.mqtt_password, 30);
    if (!SendWebSocketBinaryMessage(websocket, connect_packet)) {
      close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }

    std::vector<uint8_t> packet;
    const auto deadline =
        std::chrono::steady_clock::now() + std::chrono::seconds(10);
    bool connack_ok = false;
    while (!m_ai_inst_changed_stop.load() &&
           std::chrono::steady_clock::now() < deadline) {
      const auto receive_result =
          ReceiveWebSocketBinaryMessage(websocket, &packet);
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
      close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }

    const auto subscribe_packet =
        BuildMqttSubscribePacket(config.mqtt_ins_changed_topic, 1);
    if (!SendWebSocketBinaryMessage(websocket, subscribe_packet)) {
      close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
      close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
      close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
      if (wait_before_reconnect()) {
        break;
      }
      continue;
    }

    std::string subscribed_topics_log = config.mqtt_ins_changed_topic;
    if (!permission_update_topic.empty()) {
      const auto permission_subscribe_packet =
          BuildMqttSubscribePacket(permission_update_topic, 2);
      if (!SendWebSocketBinaryMessage(websocket,
                                      permission_subscribe_packet)) {
        close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
        close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
        close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
        if (wait_before_reconnect()) {
          break;
        }
        continue;
      }
      subscribed_topics_log += " | ";
      subscribed_topics_log += permission_update_topic;
    }
    AppendAIAssistantInfoLogLine("AI instruction mqtt subscribed topics: " +
                                 subscribed_topics_log);
    LOG(INFO) << "AI instruction mqtt subscribed topics: "
              << subscribed_topics_log;

    auto last_activity = std::chrono::steady_clock::now();
    auto last_ping = last_activity;
    while (!m_ai_inst_changed_stop.load()) {
      const auto now = std::chrono::steady_clock::now();
      if (now - last_activity >= std::chrono::seconds(20) &&
          now - last_ping >= std::chrono::seconds(20)) {
        if (!SendWebSocketBinaryMessage(websocket, ping_packet)) {
          break;
        }
        last_ping = now;
      }

      const auto receive_result =
          ReceiveWebSocketBinaryMessage(websocket, &packet);
      if (receive_result == WebSocketReceiveResult::kTimeout) {
        continue;
      }
      if (receive_result != WebSocketReceiveResult::kOk) {
        break;
      }

      last_activity = std::chrono::steady_clock::now();
      if (IsMqttPingResp(packet)) {
        continue;
      }
      if (MqttPacketType(packet) != 3) {
        continue;
      }

      std::string topic;
      std::string payload;
      if (!ParseMqttPublishPacket(packet, &topic, &payload)) {
        continue;
      }

      bool should_refresh = false;
      std::string instruction_id;
      if (!permission_update_topic.empty() && topic == permission_update_topic) {
        LOG(INFO) << "AI instruction permission update received: " << topic;
        should_refresh = true;
      } else if (TryExtractInstructionIdFromChangedTopic(
                     config.mqtt_ins_changed_topic, topic, &instruction_id)) {
        AIPanelInstitutionOption option;
        if (m_ai_instructions.FindById(u8tow(instruction_id), &option)) {
          LOG(INFO) << "AI instruction changed topic hit cached id="
                    << instruction_id;
          should_refresh = true;
        } else {
          DLOG(INFO) << "AI instruction changed topic miss cache, id="
                     << instruction_id;
        }
      } else {
        DLOG(INFO) << "AI instruction changed topic ignored: " << topic;
      }

      if (!should_refresh && TrimAsciiWhitespace(payload) == "ins_changed") {
        LOG(INFO) << "AI instruction changed payload received.";
        should_refresh = true;
      }

      if (!should_refresh) {
        continue;
      }

      _RefreshAIAssistantInstructionCacheAsync();

      bool panel_open = false;
      {
        std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
        panel_open = m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
                     IsWindowVisible(m_ai_panel.panel_hwnd);
      }
      if (panel_open) {
        _RefreshAIPanelInstitutionOptions();
      }
    }

    close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
    close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
    close_tracked_handle(&session, &m_ai_inst_changed_session_handle);

    if (wait_before_reconnect()) {
      break;
    }
  }
}

bool RimeWithWeaselHandler::_TryProcessAIAssistantTrigger(KeyEvent keyEvent,
                                                          WeaselSessionId ipc_id,
                                                          EatLine eat) {
  const bool hotkey_match = IsAIAssistantTriggerKey(m_ai_config, keyEvent);
  if (ShouldLogAIAssistantTriggerProbe(m_ai_config, keyEvent) || hotkey_match) {
    AppendAIAssistantInfoLogLine(
        "AI trigger decision: hotkey_match=" +
        std::to_string(hotkey_match ? 1 : 0) +
        ", keycode=" + std::to_string(keyEvent.keycode) +
        ", mask=" + std::to_string(keyEvent.mask) +
        ", panel_url_empty=" +
        std::to_string(m_ai_config.panel_url.empty() ? 1 : 0));
    LOG(INFO) << "AI trigger decision: hotkey_match=" << hotkey_match
              << ", keycode=" << keyEvent.keycode
              << ", mask=" << keyEvent.mask
              << ", panel_url_empty=" << m_ai_config.panel_url.empty();
  }
  if (!hotkey_match) {
    return false;
  }
  AppendAIAssistantInfoLogLine("AI trigger matched.");
  LOG(INFO) << "AI trigger matched.";
  if (m_ai_config.panel_url.empty()) {
    LOG(WARNING) << "AI panel url is empty; skip AI trigger.";
    return false;
  }
  if (keyEvent.mask & ibus::RELEASE_MASK) {
    AppendAIAssistantInfoLogLine("AI trigger ignored release event.");
    LOG(INFO) << "AI trigger ignored release event.";
    return true;
  }
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
        IsWindowVisible(m_ai_panel.panel_hwnd)) {
      LOG(INFO) << "AI panel already visible; ignore duplicate trigger.";
      return true;
    }
  }
  if (!_EnsureAIAssistantLogin()) {
    AppendAIAssistantInfoLogLine("AI trigger paused for login flow.");
    LOG(INFO) << "AI trigger paused for login flow.";
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
  AppendAIAssistantInfoLogLine("AI trigger opening panel for ipc_id=" +
                               std::to_string(ipc_id));
  LOG(INFO) << "AI trigger opening panel for ipc_id=" << ipc_id;

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

  std::vector<AIPanelInstitutionOption> preloaded_institution_options;
  bool institution_options_ready = false;
  bool relogin_started = false;
  std::string institution_error_message;
  if (!_PrepareAIPanelInstitutionOptionsForOpen(
          &preloaded_institution_options, &institution_options_ready,
          &relogin_started, &institution_error_message)) {
    if (relogin_started) {
      LOG(INFO) << "AI panel open blocked because institution list returned "
                   "401 and relogin was started.";
    } else if (!institution_error_message.empty()) {
      LOG(WARNING) << "AI panel open blocked: "
                   << institution_error_message;
    }
    return true;
  }

  if (_OpenAIPanel(ipc_id, target_hwnd, &preloaded_institution_options,
                   institution_options_ready)) {
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

bool RimeWithWeaselHandler::_TryHandleInlineInstructionKey(
    KeyEvent keyEvent,
    WeaselSessionId ipc_id,
    EatLine eat) {
  if (!m_ai_config.enabled || !m_ai_config.inline_instruction_enabled ||
      ipc_id == 0) {
    return false;
  }
  const RimeSessionId session_id = to_session_id(ipc_id);

  bool editing_active = false;
  std::wstring editing_before_prefix;
  std::wstring committed_prompt_text;
  size_t previous_cursor_offset = 0;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    const auto it = m_inline_instruction_states.find(ipc_id);
    if (it == m_inline_instruction_states.end() || !it->second.IsActive()) {
      return false;
    }
    editing_active = it->second.phase == InlineInstructionPhase::kEditing;
    if (editing_active) {
      editing_before_prefix = it->second.pending_commit_text;
      committed_prompt_text = it->second.committed_prompt_text;
      previous_cursor_offset = it->second.cursor_offset;
    }
  }

  if (IsInlineInstructionCancelKey(keyEvent)) {
    bool should_respond = false;
    bool clear_composition = false;
    {
      std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
      auto it = m_inline_instruction_states.find(ipc_id);
      if (it == m_inline_instruction_states.end()) {
        return false;
      }
      if (it->second.phase == InlineInstructionPhase::kEditing) {
        clear_composition = true;
        DLOG(INFO) << "inline_ai: cancel editing ipc_id=" << ipc_id
                   << " prompt=" << wtou8(it->second.prompt_text);
        m_inline_instruction_states.erase(it);
      } else if (it->second.phase == InlineInstructionPhase::kRequesting) {
        DLOG(INFO) << "inline_ai: cancel requesting ipc_id=" << ipc_id
                   << " prompt=" << wtou8(it->second.prompt_text);
        m_inline_instruction_states.erase(it);
        should_respond = true;
      }
    }
    if (clear_composition) {
      rime_api->clear_composition(session_id);
      return false;
    }
    if (should_respond) {
      _Respond(ipc_id, eat);
      return true;
    }
  }
  if (!editing_active) {
    bool requesting_active = false;
    bool finalize_request = false;
    bool has_error = false;
    bool streamed_writeback = false;
    std::wstring commit_text;
    std::wstring error_text;
    {
      std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
      auto it = m_inline_instruction_states.find(ipc_id);
      if (it == m_inline_instruction_states.end() ||
          it->second.phase != InlineInstructionPhase::kRequesting) {
        return false;
      }
      requesting_active = true;
      if (it->second.request_completed || it->second.has_error ||
          !it->second.result_text.empty()) {
        finalize_request = true;
        has_error = it->second.has_error;
        commit_text = it->second.result_text;
        error_text = it->second.error_text;
        streamed_writeback = it->second.streamed_writeback;
        m_inline_instruction_states.erase(it);
      }
    }

    if (!requesting_active) {
      return false;
    }

    if (finalize_request) {
      if (has_error) {
        if (!error_text.empty()) {
          LOG(WARNING) << "Inline AI request failed: " << wtou8(error_text);
        }
        _Respond(ipc_id, eat);
      } else if (!streamed_writeback && !commit_text.empty()) {
        _Respond(ipc_id, eat, &commit_text);
      } else {
        _Respond(ipc_id, eat);
      }
    } else {
      _Respond(ipc_id, eat);
    }
    return true;
  }

  const InlineInstructionPromptSnapshot prompt_snapshot =
      CaptureInlineInstructionPromptSnapshot(session_id);

  if (!IsInlineInstructionSubmitKey(keyEvent)) {
    bool should_respond = false;
    {
      std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
      auto it = m_inline_instruction_states.find(ipc_id);
      if (it == m_inline_instruction_states.end() ||
          it->second.phase != InlineInstructionPhase::kEditing) {
        return false;
      }
      InlineInstructionState& state = m_inline_instruction_states[ipc_id];
      state.session_alive = true;
      state.detached_writeback = false;
      state.phase = InlineInstructionPhase::kEditing;
      state.slash_mode = true;
      const std::wstring rime_commit = _TakePendingCommitText(session_id);
      if (!rime_commit.empty()) {
        state.committed_prompt_text.append(rime_commit);
      }
      const bool is_backspace_key = IsInlineInstructionBackspaceKey(keyEvent);
      const bool is_left_key = IsInlineInstructionLeftKey(keyEvent);
      const bool is_right_key = IsInlineInstructionRightKey(keyEvent);
      const InlineInstructionPromptSnapshot live_snapshot =
          CaptureInlineInstructionPromptSnapshot(session_id);
      const std::wstring live_prompt = live_snapshot.text;
      state.pending_commit_text = editing_before_prefix;

      if (live_prompt.empty() &&
          (is_backspace_key || is_left_key || is_right_key)) {
        state.prompt_text = state.committed_prompt_text;
        size_t cursor_offset = ClampInlineInstructionCursorOffset(
            previous_cursor_offset, state.prompt_text);
        if (is_backspace_key) {
          const size_t cursor_pos = state.prompt_text.size() - cursor_offset;
          if (cursor_pos > 0) {
            state.committed_prompt_text.erase(cursor_pos - 1, 1);
            state.prompt_text = state.committed_prompt_text;
            cursor_offset = ClampInlineInstructionCursorOffset(
                cursor_offset, state.prompt_text);
          }
        } else if (is_left_key) {
          cursor_offset = min(cursor_offset + 1, state.prompt_text.size());
        } else if (is_right_key) {
          cursor_offset = cursor_offset > 0 ? cursor_offset - 1 : 0;
        }
        state.cursor_offset = cursor_offset;
        DLOG(INFO) << "inline_ai: consume editing key ipc_id=" << ipc_id
                   << " backspace=" << is_backspace_key
                   << " left=" << is_left_key << " right=" << is_right_key
                   << " cursor_offset=" << state.cursor_offset
                   << " prompt=" << wtou8(state.prompt_text);
        should_respond = true;
      } else {
        state.prompt_text = state.committed_prompt_text + live_prompt;
        if (!live_prompt.empty()) {
          const size_t live_cursor =
              min(live_snapshot.cursor_pos, live_prompt.size());
          const size_t prompt_cursor =
              min(state.committed_prompt_text.size() + live_cursor,
                  state.prompt_text.size());
          state.cursor_offset = state.prompt_text.size() - prompt_cursor;
        } else {
          state.cursor_offset = ClampInlineInstructionCursorOffset(
              previous_cursor_offset, state.prompt_text);
        }
        DLOG(INFO) << "inline_ai: update editing ipc_id=" << ipc_id
                   << " slash_mode=" << state.slash_mode
                   << " prompt=" << wtou8(state.prompt_text)
                   << " committed_prompt="
                   << wtou8(state.committed_prompt_text)
                   << " rime_commit=" << wtou8(rime_commit)
                   << " cursor_offset=" << state.cursor_offset
                   << " before_prefix=" << wtou8(state.pending_commit_text)
                   << " raw_preedit="
                   << wtou8(CaptureCurrentPreeditText(session_id));
        if (state.prompt_text.empty() && !state.slash_mode &&
            state.pending_commit_text.empty()) {
          m_inline_instruction_states.erase(ipc_id);
        }
      }
    }
    if (should_respond) {
      _Respond(ipc_id, eat);
      return true;
    }
    return false;
  }

  std::wstring before_prefix = editing_before_prefix;
  InlineInstructionState preview_state;
  preview_state.phase = InlineInstructionPhase::kEditing;
  preview_state.committed_prompt_text = committed_prompt_text;
  preview_state.prompt_text = committed_prompt_text;
  preview_state.pending_commit_text = editing_before_prefix;
  std::wstring prompt_text = TrimInlineInstructionPrompt(
      BuildInlineInstructionActivePromptText(session_id, preview_state));
  if (prompt_text.empty()) {
    DLOG(INFO) << "inline_ai: submit ignored because prompt empty ipc_id="
               << ipc_id << " editing_active=" << editing_active
               << " raw_preedit=" << wtou8(CaptureCurrentPreeditText(session_id));
    return false;
  }

  std::wstring prefix_commit = before_prefix;
  rime_api->clear_composition(session_id);
  HWND target_hwnd = GetFocus();
  if (!target_hwnd || !IsWindow(target_hwnd)) {
    target_hwnd = GetForegroundWindow();
  }
  const uint64_t request_id = ++m_ai_request_seq;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    auto it = m_inline_instruction_states.find(ipc_id);
    if (it == m_inline_instruction_states.end()) {
      return false;
    }
    InlineInstructionState& state = it->second;
    state.session_alive = true;
    state.detached_writeback = false;
    state.has_error = false;
    state.request_completed = false;
    state.streamed_writeback = false;
    state.phase = InlineInstructionPhase::kRequesting;
    state.slash_mode = true;
    if (!target_hwnd || !IsWindow(target_hwnd)) {
      target_hwnd = state.target_hwnd;
    }
    state.target_hwnd = target_hwnd;
    state.request_id = request_id;
    state.prompt_text = prompt_text;
    state.cursor_offset = 0;
    state.pending_commit_text = prefix_commit;
    state.result_text.clear();
    state.error_text.clear();
  }
  DLOG(INFO) << "inline_ai: submit request ipc_id=" << ipc_id
             << " request_id=" << request_id
             << " prompt=" << wtou8(prompt_text)
             << " prefix_commit=" << wtou8(prefix_commit);
  _StartInlineInstructionRequest(ipc_id, request_id, prompt_text);
  _Respond(ipc_id, eat, prefix_commit.empty() ? nullptr : &prefix_commit);
  return true;
}

bool RimeWithWeaselHandler::_PrepareAIPanelInstitutionOptionsForOpen(
    std::vector<AIPanelInstitutionOption>* options,
    bool* options_ready,
    bool* relogin_started,
    std::string* error_message) {
  if (!options || !options_ready || !relogin_started) {
    if (error_message) {
      *error_message = "Institution prefetch output is null.";
    }
    return false;
  }

  options->clear();
  *options_ready = false;
  *relogin_started = false;
  if (error_message) {
    error_message->clear();
  }

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
  if (tenant_id.empty() || (token.empty() && refresh_token.empty())) {
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
    if (!file_token.empty() || !file_tenant_id.empty() ||
        !file_refresh_token.empty()) {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      if (!file_token.empty()) {
        m_ai_login_token = file_token;
      }
      if (!file_tenant_id.empty()) {
        m_ai_login_tenant_id = file_tenant_id;
      }
      if (!file_refresh_token.empty()) {
        m_ai_login_refresh_token = file_refresh_token;
      }
    }
  }

  if (tenant_id.empty()) {
    *options_ready = true;
    return true;
  }

  if (token.empty() && !refresh_token.empty()) {
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
        std::lock_guard<std::mutex> lock(m_ai_login_mutex);
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
      LOG(INFO) << "AI panel prefetch refreshed token via endpoint "
                << refresh_endpoint;
    } else if (error_message) {
      *error_message = refresh_error;
    }
  }

  if (token.empty()) {
    *options_ready = true;
    return true;
  }

  int http_status_code = 0;
  std::string fetch_error_message;
  if (FetchAIAssistantInstitutionOptions(config, token, tenant_id, options,
                                         &fetch_error_message,
                                         &http_status_code)) {
    m_ai_instructions.Replace(*options);
    *options_ready = true;
    return true;
  }

  if (http_status_code == 401) {
    _ClearAIAssistantInstructionCache();
    *relogin_started = _ForceAIAssistantRelogin();
    if (error_message) {
      *error_message = *relogin_started
                           ? "Institution list returned 401; relogin started."
                           : "Institution list returned 401; relogin failed.";
    }
    return false;
  }

  if (error_message && !fetch_error_message.empty()) {
    *error_message = fetch_error_message;
  }
  const std::vector<AIPanelInstitutionOption> cached_options =
      m_ai_instructions.Snapshot();
  if (!cached_options.empty()) {
    *options = cached_options;
    *options_ready = true;
  }
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
    bool relogin_started = false;

    if (!token.empty()) {
      ok = FetchAIAssistantInstitutionOptions(config, token, tenant_id, &options,
                                              &error_message, &http_status_code);
    }
    if (!ok && http_status_code == 401) {
      relogin_started = _ForceAIAssistantRelogin();
      error_message =
          relogin_started ? "登录已过期，已触发重新登录。"
                          : "登录已过期，但重新登录流程启动失败。";
    } else if (!ok && token.empty() && !refresh_token.empty()) {
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
      const std::wstring selected_id = m_ai_panel.selected_institution_id;
      if (ok) {
        bool found = IsBuiltinAIAssistantInstructionId(selected_id);
        for (const auto& option : options) {
          if (!selected_id.empty() && option.id == selected_id) {
            found = true;
            break;
          }
        }
        m_ai_panel.institution_options = options;
        m_ai_instructions.Replace(options);
        if (found) {
          m_ai_panel.selected_institution_id = selected_id;
        } else {
          m_ai_panel.selected_institution_id.clear();
        }
      } else {
        m_ai_panel.institution_options.clear();
        if (!IsBuiltinAIAssistantInstructionId(selected_id)) {
          m_ai_panel.selected_institution_id.clear();
        }
        if (http_status_code == 401) {
          _ClearAIAssistantInstructionCache();
        }
      }
    }

    if (!ok) {
      if (relogin_started) {
        LOG(INFO) << "AI panel institution list fetch hit 401 and triggered "
                     "relogin.";
      } else {
        LOG(WARNING) << "AI panel institution list fetch failed: "
                     << error_message;
      }
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

void RimeWithWeaselHandler::_StartInlineInstructionRequest(
    WeaselSessionId ipc_id,
    uint64_t request_id,
    const std::wstring& prompt_text) {
  HWND target_hwnd = nullptr;
  AIAssistantConfig request_config = m_ai_config;
  AIPanelInstitutionOption selected_option;
  bool has_selected_option = false;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    const auto it = m_inline_instruction_states.find(ipc_id);
    if (it != m_inline_instruction_states.end()) {
      target_hwnd = it->second.target_hwnd;
      has_selected_option = it->second.has_selected_option;
      selected_option = it->second.selected_option;
    }
  }

  AppendAIAssistantContextDump(request_config, "inline_instruction", target_hwnd,
                               prompt_text);
  std::thread([this, request_config, ipc_id, request_id, prompt_text,
               has_selected_option, selected_option, target_hwnd]() {
    std::string error_message;
    bool ok = false;
    int http_status_code = 0;
    std::wstring pending_commit_text;
    bool streamed_any_chunk = false;
    bool direct_writeback_ok = true;

    const auto append_chunk = [&](const std::wstring& chunk) {
      if (chunk.empty()) {
        return;
      }

      HWND write_target = target_hwnd;
      {
        std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
        const auto it = m_inline_instruction_states.find(ipc_id);
        if (it == m_inline_instruction_states.end() ||
            it->second.request_id != request_id ||
            it->second.phase != InlineInstructionPhase::kRequesting) {
          return;
        }
        if ((!write_target || !IsWindow(write_target)) &&
            it->second.target_hwnd && IsWindow(it->second.target_hwnd)) {
          write_target = it->second.target_hwnd;
        }
      }

      if (!direct_writeback_ok) {
        pending_commit_text.append(chunk);
        return;
      }

      if (_SendTextToTargetWindow(write_target, chunk)) {
        streamed_any_chunk = true;
        _AppendCommittedText(ipc_id, chunk);
        return;
      }

      direct_writeback_ok = false;
      pending_commit_text.append(chunk);
    };

    if (!has_selected_option) {
      error_message = "Inline AI instruction option is missing.";
    } else if (selected_option.app_key.empty()) {
      error_message = "Inline AI instruction appKey is empty.";
    } else {
      std::string token;
      std::string tenant_id;
      std::string refresh_token;
      std::string resolve_error;
      if (!_ResolveAIAssistantLoginState(&token, &tenant_id, &refresh_token,
                                         &resolve_error)) {
        error_message = resolve_error.empty() ? "AI 登录状态无效。" : resolve_error;
      } else {
        std::string user_id;
        if (!LoadAIAssistantCachedUserId(request_config, &user_id) &&
            !token.empty() && !tenant_id.empty()) {
          std::string user_info_error;
          int user_info_status_code = 0;
          const bool user_info_refreshed = RefreshAIAssistantUserInfoCache(
              request_config, token, tenant_id, &user_info_error,
              &user_info_status_code);
          if (!user_info_refreshed && user_info_status_code == 401) {
            _ForceAIAssistantRelogin();
          }
          if (user_info_refreshed) {
            _StartAIAssistantInstructionChangedListener();
          }
          LoadAIAssistantCachedUserId(request_config, &user_id);
        }
        ok = InvokeInlineInstructionChatStream(
            request_config, selected_option, prompt_text, token, tenant_id,
            user_id, append_chunk, &error_message, &http_status_code);
        if (!ok && http_status_code == 401) {
          _ForceAIAssistantRelogin();
        }
      }
    }

    bool should_write_back = false;
    HWND writeback_hwnd = nullptr;
    std::wstring fallback_commit_text;
    {
      std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
      auto it = m_inline_instruction_states.find(ipc_id);
      if (it == m_inline_instruction_states.end() ||
          it->second.request_id != request_id) {
        return;
      }

      it->second.has_error = !ok;
      it->second.request_completed = true;
      it->second.streamed_writeback =
          ok && direct_writeback_ok && pending_commit_text.empty() &&
          streamed_any_chunk;
      it->second.result_text =
          ok ? pending_commit_text : std::wstring();
      it->second.error_text = ok ? std::wstring() : u8tow(error_message);
      fallback_commit_text = it->second.result_text;
      should_write_back = !it->second.session_alive &&
                          it->second.detached_writeback &&
                          ok && !fallback_commit_text.empty();
      writeback_hwnd = it->second.target_hwnd;
      if (!it->second.session_alive) {
        if (should_write_back) {
          it->second.phase = InlineInstructionPhase::kIdle;
        } else {
          m_inline_instruction_states.erase(it);
        }
      }
    }
    if (should_write_back) {
      if (_SendTextToTargetWindow(writeback_hwnd, fallback_commit_text)) {
        _AppendCommittedText(ipc_id, fallback_commit_text);
      }
      std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
      m_inline_instruction_states.erase(ipc_id);
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

  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    if (m_ai_panel.panel_hwnd && IsWindow(m_ai_panel.panel_hwnd) &&
        IsWindowVisible(m_ai_panel.panel_hwnd)) {
      commit_text->clear();
      return true;
    }
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
  _GetContext(weasel_context, session_id, ipc_id);

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
  wchar_t inline_prefix_char = L'/';
  const std::wstring inline_prefix = u8tow(m_ai_config.inline_instruction_prefix);
  if (inline_prefix.size() == 1) {
    inline_prefix_char = inline_prefix[0];
  }
  std::wstring commit_text;
  const std::wstring rime_commit = _TakePendingCommitText(session_id);
  bool swallow_rime_commit = false;
  {
    std::lock_guard<std::mutex> lock(m_inline_instruction_mutex);
    const auto it = m_inline_instruction_states.find(ipc_id);
    swallow_rime_commit =
        it != m_inline_instruction_states.end() &&
        it->second.phase == InlineInstructionPhase::kEditing &&
        it->second.slash_mode;
  }
  if (!rime_commit.empty()) {
    if (!swallow_rime_commit) {
      commit_text.append(rime_commit);
    }
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
    InlineInstructionState inline_state;
    const bool has_inline_state = GetInlineInstructionSnapshot(
        m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
        &inline_state);
    if (has_inline_state && inline_state.IsActive()) {
      is_composing = true;
    }
    actions.push_back("status");
    body.append(L"status.ascii_mode=")
        .append(Bool_wstring[!!status.is_ascii_mode])
        .append(L"\n")
        .append(L"status.composing=")
        .append(Bool_wstring[(int)is_composing])
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
    const char* input_text = rime_api->get_input(session_id);
    InlineInstructionState inline_state;
    const bool has_inline_state = GetInlineInstructionSnapshot(
        m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
        &inline_state);
    const bool suppress_special_instruction_candidates =
        has_inline_state && inline_state.IsActive() &&
        input_text &&
        (IsInstructionLookupInputText(input_text,
                                      m_ai_config.instruction_lookup_prefix) ||
         IsSystemCommandInputText(input_text));
    if (suppress_special_instruction_candidates) {
      std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
      m_ai_injected_candidates.erase(ipc_id);
    }
    bool has_candidates =
        !suppress_special_instruction_candidates && ctx.menu.num_candidates > 0;
    const bool should_try_instruction_lookup =
        !suppress_special_instruction_candidates &&
        input_text &&
        IsInstructionLookupInputText(input_text,
                                     m_ai_config.instruction_lookup_prefix);
    CandidateInfo cinfo;
    if (has_candidates || should_try_instruction_lookup) {
      _GetCandidateInfo(cinfo, ctx, ipc_id);
      has_candidates = !cinfo.candies.empty();
    }
    if (has_inline_state && inline_state.IsActive()) {
      actions.push_back("ctx");
      const std::wstring preedit_text =
          inline_state.phase == InlineInstructionPhase::kRequesting
              ? BuildInlineInstructionStatusText(inline_prefix_char,
                                                 inline_state)
              : BuildInlineInstructionEditingText(inline_prefix_char,
                                                 inline_state);
      const size_t cursor_pos =
          GetInlineInstructionDisplayCursorPosition(inline_prefix_char,
                                                   inline_state, preedit_text);
      body.append(L"ctx.preedit=")
          .append(escape_string(preedit_text))
          .append(L"\n");
      body.append(L"ctx.preedit.cursor=")
          .append(std::to_wstring(cursor_pos))
          .append(L",")
          .append(std::to_wstring(cursor_pos))
          .append(L",")
          .append(std::to_wstring(cursor_pos))
          .append(L"\n");
    } else if (is_composing) {
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
          for (auto i = 0; i < static_cast<int>(cinfo.candies.size()); i++) {
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
            std::wstring prefix_w = (i != cinfo.highlighted)
                                        ? std::wstring()
                                        : mark_text_w;
            body.append(L" ")
                .append(prefix_w)
                .append(escape_string(label_w))
                .append(cinfo.candies.at(i).str)
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
    if (has_candidates || suppress_special_instruction_candidates) {
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
  InlineInstructionState inline_state;
  const bool has_inline_state = GetInlineInstructionSnapshot(
      m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
      &inline_state);
  body.append(L"config.inline_ai_requesting=")
      .append(std::to_wstring((int)(has_inline_state &&
                                    inline_state.phase ==
                                        InlineInstructionPhase::kRequesting)))
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
  InlineInstructionState inline_state;
  const bool has_inline_state = GetInlineInstructionSnapshot(
      m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
      &inline_state);
  RIME_STRUCT(RimeStatus, status);
  if (rime_api->get_status(session_id, &status)) {
    std::string schema_id = "";
    if (status.schema_id)
      schema_id = status.schema_id;
    stat.schema_name = u8tow(status.schema_name);
    stat.schema_id = u8tow(status.schema_id);
    stat.ascii_mode = !!status.is_ascii_mode;
    stat.composing = !!status.is_composing || (has_inline_state &&
                                               inline_state.IsActive());
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
                                        RimeSessionId session_id,
                                        WeaselSessionId ipc_id) {
  InlineInstructionState inline_state;
  const bool has_inline_state = GetInlineInstructionSnapshot(
      m_inline_instruction_mutex, m_inline_instruction_states, ipc_id,
      &inline_state);
  wchar_t inline_prefix_char = L'/';
  const std::wstring inline_prefix = u8tow(m_ai_config.inline_instruction_prefix);
  if (inline_prefix.size() == 1) {
    inline_prefix_char = inline_prefix[0];
  }
  RIME_STRUCT(RimeContext, ctx);
  if (rime_api->get_context(session_id, &ctx)) {
    const char* input_text = rime_api->get_input(session_id);
    const bool suppress_special_instruction_candidates =
        has_inline_state && inline_state.IsActive() &&
        input_text &&
        (IsInstructionLookupInputText(input_text,
                                      m_ai_config.instruction_lookup_prefix) ||
         IsSystemCommandInputText(input_text));
    if (suppress_special_instruction_candidates) {
      std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
      m_ai_injected_candidates.erase(ipc_id);
    }
    const bool should_try_instruction_lookup =
        !suppress_special_instruction_candidates &&
        input_text &&
        IsInstructionLookupInputText(input_text,
                                     m_ai_config.instruction_lookup_prefix);
    if (has_inline_state && inline_state.IsActive()) {
      weasel_context.preedit.str =
          inline_state.phase == InlineInstructionPhase::kRequesting
              ? BuildInlineInstructionStatusText(inline_prefix_char,
                                                 inline_state)
              : BuildInlineInstructionEditingText(inline_prefix_char,
                                                 inline_state);
      AddInlineInstructionCursorAttribute(
          &weasel_context.preedit,
          GetInlineInstructionDisplayCursorPosition(
              inline_prefix_char, inline_state, weasel_context.preedit.str));
    } else if (ctx.composition.length > 0) {
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
    if ((!suppress_special_instruction_candidates && ctx.menu.num_candidates) ||
        should_try_instruction_lookup) {
      CandidateInfo& cinfo(weasel_context.cinfo);
      _GetCandidateInfo(cinfo, ctx, ipc_id);
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
