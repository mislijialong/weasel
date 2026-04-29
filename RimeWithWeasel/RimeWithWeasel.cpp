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

#include "RimeWithWeaselInternal.h"

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
RimeApi* GetWeaselRimeApi() {
  return rime_api;
}

WeaselSessionId _GenerateNewWeaselSessionId(SessionStatusMap sm, DWORD pid) {
  if (sm.empty())
    return (WeaselSessionId)(pid + 1);
  return (WeaselSessionId)(sm.rbegin()->first + 1);
}

int expand_ibus_modifier(int m) {
  return (m & 0xff) | ((m & 0xff00) << 16);
}

namespace {

constexpr char kDefaultSystemCommandInputPrefix[] = "sc";
constexpr char kDefaultInstructionLookupInputPrefix[] = "sS";
constexpr char kDefaultAIAssistantInstructionChangedTopic[] =
    "/mqtt/topic/sino/langwell/ins/ins/changed/+";


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

bool IsPlainInlineEditingDigitKey(const KeyEvent& key_event);
bool IsKeypadInlineEditingDigitKey(const KeyEvent& key_event);
UINT NormalizeInlineEditingDigitKeycode(UINT keycode);

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

bool ShouldConsumeInlineEditingDigitKey(
    const KeyEvent& key_event,
    std::mutex& inline_mutex,
    std::map<WeaselSessionId, InlineInstructionState>& inline_states,
    std::mutex& injected_mutex,
    std::map<WeaselSessionId, AIAssistantInjectedCandidateState>&
        injected_candidates,
    RimeSessionId session_id,
    WeaselSessionId ipc_id) {
  if (!IsPlainInlineEditingDigitKey(key_event) || session_id == 0 || ipc_id == 0) {
    return false;
  }

  InlineInstructionState inline_state;
  if (!GetInlineInstructionSnapshot(inline_mutex, inline_states, ipc_id,
                                    &inline_state) ||
      inline_state.phase != InlineInstructionPhase::kEditing) {
    return false;
  }

  {
    std::lock_guard<std::mutex> lock(injected_mutex);
    const auto it = injected_candidates.find(ipc_id);
    if (it != injected_candidates.end() &&
        GetInjectedCandidateVisibleCount(it->second) > 0) {
      return false;
    }
  }

  RIME_STRUCT(RimeContext, ctx);
  bool has_rime_candidates = false;
  if (rime_api->get_context(session_id, &ctx)) {
    has_rime_candidates = ctx.menu.num_candidates > 0;
    rime_api->free_context(&ctx);
  }
  return !has_rime_candidates;
}

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

std::wstring BuildInlineInstructionQueryText(
    const AIPanelInstitutionOption& option,
    const std::wstring& content_text) {
  const std::wstring normalized_content =
      TrimInlineInstructionPrompt(content_text);
  std::wstring template_text =
      TrimInlineInstructionPrompt(option.template_content);
  if (template_text.empty()) {
    return normalized_content;
  }

  const std::wstring placeholder = L"${content}";
  size_t placeholder_pos = template_text.find(placeholder);
  if (placeholder_pos == std::wstring::npos) {
    if (normalized_content.empty()) {
      return template_text;
    }
    if (!template_text.empty() && template_text.back() != L'\n' &&
        template_text.back() != L'\r') {
      template_text += L"\n";
    }
    template_text += normalized_content;
    return template_text;
  }

  while (placeholder_pos != std::wstring::npos) {
    template_text.replace(placeholder_pos, placeholder.size(),
                          normalized_content);
    placeholder_pos = template_text.find(
        placeholder, placeholder_pos + normalized_content.size());
  }
  return template_text;
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

  const std::wstring request_query =
      BuildInlineInstructionQueryText(option, query_text);
  const std::string request_body = BuildInlineInstructionChatRequestBody(
      request_query, token, tenant_id, user_id);
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

bool IsPlainInlineEditingDigitKey(const KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  const auto modifiers = key_event.mask & ibus::MODIFIER_MASK;
  if ((modifiers & (ibus::SHIFT_MASK | ibus::CONTROL_MASK | ibus::ALT_MASK |
                    ibus::SUPER_MASK | ibus::HYPER_MASK | ibus::META_MASK)) !=
      0) {
    return false;
  }
  return (key_event.keycode >= '0' && key_event.keycode <= '9') ||
         (key_event.keycode >= ibus::KP_0 && key_event.keycode <= ibus::KP_9);
}

bool IsKeypadInlineEditingDigitKey(const KeyEvent& key_event) {
  return !(key_event.mask & ibus::RELEASE_MASK) &&
         key_event.keycode >= ibus::KP_0 && key_event.keycode <= ibus::KP_9;
}

UINT NormalizeInlineEditingDigitKeycode(UINT keycode) {
  if (keycode >= ibus::KP_0 && keycode <= ibus::KP_9) {
    return static_cast<UINT>('0' + (keycode - ibus::KP_0));
  }
  return keycode;
}

bool InsertInlineEditingDigitFallback(
    const KeyEvent& key_event,
    std::mutex& inline_mutex,
    std::map<WeaselSessionId, InlineInstructionState>& inline_states,
    RimeSessionId session_id,
    WeaselSessionId ipc_id,
    const std::wstring& prompt_before_key) {
  const UINT normalized_keycode =
      NormalizeInlineEditingDigitKeycode(key_event.keycode);
  if (normalized_keycode < '0' || normalized_keycode > '9' || session_id == 0 ||
      ipc_id == 0) {
    return false;
  }

  if (!CaptureInlineInstructionPromptSource(session_id).empty()) {
    return false;
  }

  std::lock_guard<std::mutex> lock(inline_mutex);
  auto it = inline_states.find(ipc_id);
  if (it == inline_states.end() ||
      it->second.phase != InlineInstructionPhase::kEditing ||
      it->second.prompt_text != prompt_before_key) {
    return false;
  }

  InlineInstructionState& state = it->second;
  state.session_alive = true;
  state.detached_writeback = false;
  state.phase = InlineInstructionPhase::kEditing;
  state.slash_mode = true;

  const wchar_t digit = static_cast<wchar_t>(normalized_keycode);
  state.committed_prompt_text = state.prompt_text;
  const size_t cursor_offset =
      state.cursor_offset < state.committed_prompt_text.size()
          ? state.cursor_offset
          : state.committed_prompt_text.size();
  const size_t cursor_pos = state.committed_prompt_text.size() - cursor_offset;
  state.committed_prompt_text.insert(cursor_pos, 1, digit);
  state.prompt_text = state.committed_prompt_text;
  state.cursor_offset = state.prompt_text.size() - (cursor_pos + 1);
  return true;
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

std::string GenerateLoginClientId() {
  return GenerateRandomMqttClientId();
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
  const bool should_consume_inline_digit =
      ShouldConsumeInlineEditingDigitKey(
          keyEvent, m_inline_instruction_mutex, m_inline_instruction_states,
          m_ai_injected_candidates_mutex, m_ai_injected_candidates, session_id,
          ipc_id);
  std::wstring inline_digit_prompt_before;
  if (should_consume_inline_digit) {
    InlineInstructionState inline_state;
    if (GetInlineInstructionSnapshot(m_inline_instruction_mutex,
                                     m_inline_instruction_states, ipc_id,
                                     &inline_state)) {
      inline_digit_prompt_before = inline_state.prompt_text;
    }
  }
  KeyEvent effective_key_event = keyEvent;
  if (should_consume_inline_digit &&
      IsKeypadInlineEditingDigitKey(effective_key_event)) {
    effective_key_event.keycode =
        NormalizeInlineEditingDigitKeycode(effective_key_event.keycode);
  }
  Bool handled = rime_api->process_key(
      session_id, effective_key_event.keycode,
      expand_ibus_modifier(effective_key_event.mask));
  const bool inline_key_handled =
      !is_inline_control_key &&
      _TryHandleInlineInstructionKey(effective_key_event, ipc_id, eat);
  if (inline_key_handled) {
    _UpdateUI(ipc_id);
    m_active_session = ipc_id;
    return TRUE;
  }
  if (should_consume_inline_digit) {
    InsertInlineEditingDigitFallback(
        effective_key_event, m_inline_instruction_mutex,
        m_inline_instruction_states, session_id, ipc_id,
        inline_digit_prompt_before);
    _Respond(ipc_id, eat);
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

  auto open_login_url = [](const std::wstring& login_url) {
    if (login_url.empty()) {
      LOG(WARNING) << "AI login required but login_url is empty.";
      return false;
    }
    const auto open_result = reinterpret_cast<intptr_t>(
        ShellExecuteW(nullptr, L"open", login_url.c_str(), nullptr, nullptr,
                      SW_SHOWNORMAL));
    if (open_result <= 32) {
      LOG(WARNING) << "Failed to open login page, code=" << open_result;
      return false;
    }
    return true;
  };

  if (m_ai_login_pending.load()) {
    std::string pending_client_id;
    {
      std::lock_guard<std::mutex> lock(m_ai_login_mutex);
      pending_client_id = m_ai_login_client_id;
    }
    if (!pending_client_id.empty()) {
      const std::wstring login_url =
          BuildLoginUrlForBrowser(m_ai_config, pending_client_id);
      if (!open_login_url(login_url)) {
        return false;
      }
      AppendAIAssistantInfoLogLine(
          "AI login page reopened, client_id=" + pending_client_id);
      LOG(INFO) << "AI login page reopened, client_id="
                << pending_client_id;
      return true;
    }

    LOG(WARNING) << "AI login pending without client_id; restarting flow.";
    m_ai_login_stop.store(true);
  }

  if (m_ai_login_thread.joinable()) {
    m_ai_login_thread.join();
  }
  m_ai_login_pending.store(false);

  const std::string client_id = GenerateLoginClientId();
  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    m_ai_login_client_id = client_id;
  }

  const std::wstring login_url = BuildLoginUrlForBrowser(m_ai_config, client_id);

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

  if (!open_login_url(login_url)) {
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

    const AIAssistantConfig config = m_ai_config;
    std::thread([this, config]() {
      std::string token;
      std::string tenant_id;
      std::string refresh_token;
      std::string resolve_error;
      if (!_ResolveAIAssistantLoginState(&token, &tenant_id, &refresh_token,
                                         &resolve_error)) {
        LOG(WARNING)
            << "AI ensure login async: unable to resolve login state, error="
            << resolve_error;
        _ForceAIAssistantRelogin();
        return;
      }

      std::string user_info_error;
      int user_info_status_code = 0;
      const bool user_info_refreshed = RefreshAIAssistantUserInfoCache(
          config, token, tenant_id, &user_info_error, &user_info_status_code);
      if (user_info_refreshed) {
        LOG(INFO) << "AI ensure login async: user info cache refreshed.";
        _StartAIAssistantInstructionChangedListener();
        return;
      }

      if (!user_info_error.empty()) {
        LOG(WARNING) << "AI ensure login async: user info refresh failed: "
                     << user_info_error;
      }
      if (user_info_status_code == 401) {
        LOG(INFO)
            << "AI ensure login async: auth invalid, restarting login flow.";
        _ForceAIAssistantRelogin();
      }
    }).detach();
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

void RimeWithWeaselHandler::_ClearAIAssistantLoginStateForLogout(
    const std::string& reason) {
  _ClearAIAssistantInstructionCache();

  {
    std::lock_guard<std::mutex> lock(m_ai_login_mutex);
    m_ai_login_token.clear();
    m_ai_login_tenant_id.clear();
    m_ai_login_refresh_token.clear();
    m_ai_login_client_id.clear();
  }

  const std::filesystem::path state_path(
      ResolveAIAssistantLoginStatePath(m_ai_config));
  std::error_code remove_error;
  std::filesystem::remove(state_path, remove_error);
  if (remove_error) {
    LOG(WARNING) << "AI logout: failed to remove login state file: "
                 << state_path.u8string()
                 << ", error=" << remove_error.message();
  }
  ClearAIAssistantUserInfoCache(m_ai_config);

  HWND panel_hwnd = nullptr;
  {
    std::lock_guard<std::mutex> lock(m_ai_panel_mutex);
    panel_hwnd = m_ai_panel.panel_hwnd;
    m_ai_panel.institutions_loading = false;
    m_ai_panel.institution_options.clear();
    if (!IsBuiltinAIAssistantInstructionId(
            m_ai_panel.selected_institution_id)) {
      m_ai_panel.selected_institution_id.clear();
    }
  }
  if (panel_hwnd && IsWindow(panel_hwnd)) {
    PostMessageW(panel_hwnd, WM_AI_WEBVIEW_SYNC, 0, 0);
  }

  AppendAIAssistantInfoLogLine("AI logout handled: " + reason);
  LOG(INFO) << "AI logout handled: " << reason;
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

    std::string logout_user_id;
    LoadAIAssistantCachedUserId(config, &logout_user_id);
    if (logout_user_id.empty() && !permission_update_token.empty() &&
        !permission_update_tenant_id.empty()) {
      std::string user_info_error;
      int user_info_status_code = 0;
      if (RefreshAIAssistantUserInfoCache(config, permission_update_token,
                                          permission_update_tenant_id,
                                          &user_info_error,
                                          &user_info_status_code)) {
        LoadAIAssistantCachedUserId(config, &logout_user_id);
      } else if (!user_info_error.empty()) {
        LOG(INFO) << "AI logout topic user info unavailable: "
                  << user_info_error;
      }
    }
    const std::string logout_topic =
        BuildAIAssistantLogoutTopic(logout_user_id);
    if (!logout_topic.empty()) {
      LOG(INFO) << "AI logout topic resolved: " << logout_topic;
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
    if (!logout_topic.empty()) {
      const auto logout_subscribe_packet =
          BuildMqttSubscribePacket(logout_topic, 3);
      if (!SendWebSocketBinaryMessage(websocket, logout_subscribe_packet)) {
        close_tracked_handle(&websocket, &m_ai_inst_changed_websocket_handle);
        close_tracked_handle(&connection, &m_ai_inst_changed_connection_handle);
        close_tracked_handle(&session, &m_ai_inst_changed_session_handle);
        if (wait_before_reconnect()) {
          break;
        }
        continue;
      }
      subscribed_topics_log += " | ";
      subscribed_topics_log += logout_topic;
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

      if (!logout_topic.empty() && topic == logout_topic) {
        _ClearAIAssistantLoginStateForLogout("mqtt topic=" + topic);
        break;
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

  const std::wstring normalized_prompt_context =
      NormalizeReferencedContextText(prompt_context);
  const std::wstring status_text =
      normalized_prompt_context.empty()
          ? L"未检测到输入内容，请在前端面板中继续操作。"
          : L"上下文已就绪，请在前端面板中发起请求。";
  const std::vector<AIPanelInstitutionOption> cached_options =
      m_ai_instructions.Snapshot();
  const bool institution_options_ready = !cached_options.empty();
  _OpenAIPanelAsync(ipc_id, target_hwnd, normalized_prompt_context, status_text,
                    cached_options, institution_options_ready);

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
