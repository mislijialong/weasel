#include "stdafx.h"

#include <logging.h>
#include <RimeWithWeasel.h>
#include <StringAlgorithm.hpp>
#include <WeaselUtility.h>

#include "RimeWithWeaselInternal.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <map>
#include <mutex>
#include <set>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace {

constexpr wchar_t kSystemCommandCommitPrefix[] = L"__weasel_syscmd__:";
constexpr char kDefaultSystemCommandInputPrefix[] = "sc";
constexpr const char* kAllowedSystemCommandIds[] = {
    "jsq",       "calc",     "notepad",   "mspaint", "explorer",
    "txt",       "md",       "gh",        "bd",      "wb",
    "g",         "yt",       "rili",      "calendar", "cal", "kb"};

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

bool IsInstructionLookupSystemCommandConfirmKey(
    const weasel::KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  return key_event.keycode == ibus::space ||
         key_event.keycode == ibus::Return ||
         key_event.keycode == ibus::KP_Enter;
}

bool IsPlainInstructionLookupCandidateSelectKey(
    const weasel::KeyEvent& key_event) {
  if (key_event.mask & ibus::RELEASE_MASK) {
    return false;
  }
  const auto modifiers = key_event.mask & ibus::MODIFIER_MASK;
  return (modifiers &
          (ibus::CONTROL_MASK | ibus::ALT_MASK | ibus::SUPER_MASK |
           ibus::HYPER_MASK | ibus::META_MASK)) == 0;
}

bool TryResolveInstructionLookupSelectIndex(
    const weasel::KeyEvent& key_event,
    size_t* index) {
  if (!index || !IsPlainInstructionLookupCandidateSelectKey(key_event)) {
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
  if (key_event.keycode >= ibus::KP_1 &&
      key_event.keycode <= ibus::KP_9) {
    *index = static_cast<size_t>(key_event.keycode - ibus::KP_1);
    return true;
  }
  if (key_event.keycode == ibus::KP_0) {
    *index = 9;
    return true;
  }
  if (key_event.keycode == ibus::space ||
      key_event.keycode == ibus::Return ||
      key_event.keycode == ibus::KP_Enter) {
    *index = 0;
    return true;
  }
  return false;
}

bool HasInstructionLookupActiveInlineState(
    std::mutex& mutex,
    std::map<WeaselSessionId, InlineInstructionState>& states,
    WeaselSessionId ipc_id) {
  if (ipc_id == 0) {
    return false;
  }
  std::lock_guard<std::mutex> lock(mutex);
  const auto it = states.find(ipc_id);
  return it != states.end() && it->second.IsActive();
}

bool ShouldSuppressInstructionLookupCandidates(
    std::mutex& inline_mutex,
    std::map<WeaselSessionId, InlineInstructionState>& inline_states,
    WeaselSessionId ipc_id,
    const std::string& instruction_lookup_prefix,
    const char* input_text) {
  if (!input_text) {
    return false;
  }
  if (!HasInstructionLookupActiveInlineState(inline_mutex, inline_states,
                                             ipc_id)) {
    return false;
  }
  return IsInstructionLookupInputText(input_text, instruction_lookup_prefix) ||
         IsSystemCommandInputText(input_text);
}

}  // namespace

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

std::vector<std::string> TokenizeInstructionLookupCodes(
    const std::string& text) {
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

namespace {

std::string ReadRimeDataDirectory(bool shared) {
  RimeApi* api = GetWeaselRimeApi();
  if (!api) {
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
    if (api->get_shared_data_dir_s) {
      api->get_shared_data_dir_s(buffer.data(), buffer.size());
    } else if (api->get_shared_data_dir) {
      copy_data_dir(api->get_shared_data_dir());
    }
  } else {
    if (api->get_user_data_dir_s) {
      api->get_user_data_dir_s(buffer.data(), buffer.size());
    } else if (api->get_user_data_dir) {
      copy_data_dir(api->get_user_data_dir());
    }
  }
  return std::string(buffer.data());
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

struct WeightedPinyinCode {
  std::string code;
  bool has_weight = false;
  int weight = 0;
};

bool IsAsciiWhitespaceChar(char ch) {
  return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' ||
         ch == '\f' || ch == '\v';
}

std::string TrimAsciiWhitespaceCopy(const std::string& text) {
  size_t begin = 0;
  while (begin < text.size() && IsAsciiWhitespaceChar(text[begin])) {
    ++begin;
  }
  size_t end = text.size();
  while (end > begin && IsAsciiWhitespaceChar(text[end - 1])) {
    --end;
  }
  return text.substr(begin, end - begin);
}

bool TryParseDictionaryWeight(const std::string& text, int* weight) {
  if (!weight) {
    return false;
  }
  std::string trimmed = TrimAsciiWhitespaceCopy(text);
  if (trimmed.empty()) {
    return false;
  }
  if (!trimmed.empty() && trimmed.back() == '%') {
    trimmed.pop_back();
    trimmed = TrimAsciiWhitespaceCopy(trimmed);
  }
  if (trimmed.empty()) {
    return false;
  }
  size_t pos = 0;
  bool negative = false;
  if (trimmed[pos] == '+' || trimmed[pos] == '-') {
    negative = trimmed[pos] == '-';
    ++pos;
  }
  if (pos >= trimmed.size()) {
    return false;
  }
  int value = 0;
  for (; pos < trimmed.size(); ++pos) {
    const char ch = trimmed[pos];
    if (ch < '0' || ch > '9') {
      return false;
    }
    value = value * 10 + (ch - '0');
  }
  *weight = negative ? -value : value;
  return true;
}

bool ShouldReplaceWeightedPinyinCode(const WeightedPinyinCode& existing,
                                     const WeightedPinyinCode& candidate) {
  if (!candidate.has_weight) {
    return false;
  }
  if (!existing.has_weight) {
    return true;
  }
  return candidate.weight > existing.weight;
}

template <typename Key>
void UpsertWeightedPinyinCode(
    std::unordered_map<Key, WeightedPinyinCode>* codes,
    const Key& key,
    WeightedPinyinCode candidate) {
  if (!codes) {
    return;
  }
  const auto it = codes->find(key);
  if (it == codes->end()) {
    codes->emplace(key, std::move(candidate));
    return;
  }
  if (ShouldReplaceWeightedPinyinCode(it->second, candidate)) {
    it->second = std::move(candidate);
  }
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
    *code = it->second.code;
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
    *code = it->second.code;
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
      const std::string weight_text =
          second_tab == std::string::npos ? std::string()
                                          : line.substr(second_tab + 1);

      const std::wstring phrase = u8tow(phrase_u8);
      if (phrase.empty()) {
        continue;
      }
      WeightedPinyinCode entry;
      entry.code = code;
      entry.has_weight = TryParseDictionaryWeight(weight_text, &entry.weight);
      UpsertWeightedPinyinCode(&phrase_codes_, phrase, entry);
      max_phrase_length_ = (std::max)(max_phrase_length_, phrase.size());
      if (phrase.size() == 1) {
        UpsertWeightedPinyinCode(&char_codes_, phrase.front(),
                                 std::move(entry));
      }
    }

    return !phrase_codes_.empty() || !char_codes_.empty();
  }

  std::mutex mutex_;
  bool load_attempted_ = false;
  size_t max_phrase_length_ = 1;
  std::unordered_map<std::wstring, WeightedPinyinCode> phrase_codes_;
  std::unordered_map<wchar_t, WeightedPinyinCode> char_codes_;
};

InstructionLookupPinyinCache& GetInstructionLookupPinyinCache() {
  static InstructionLookupPinyinCache cache;
  return cache;
}

}  // namespace

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

bool RimeWithWeaselHandler::_TryMatchAIAssistantInstructionOption(
    const std::wstring& text,
    AIPanelInstitutionOption* option) const {
  return MatchBuiltinAIAssistantInstruction(text, option) ||
         m_ai_instructions.MatchExactName(text, option);
}

bool RimeWithWeaselHandler::_TryBuildInstructionLookupCandidates(
    WeaselSessionId ipc_id,
    const std::string& input_text,
    int page_size,
    weasel::CandidateInfo* cinfo) {
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
    cinfo->candies.push_back(weasel::Text(escape_string(option.name)));
    cinfo->comments.push_back(
        weasel::Text(option.IsSystemCommand()
                         ? L"系统指令"
                         : (IsBuiltinAIAssistantInstructionId(option.id)
                                ? L"默认指令"
                                : L"AI指令")));
    cinfo->labels.push_back(weasel::Text(std::to_wstring((i + 1) % 10)));
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

bool RimeWithWeaselHandler::_TryResolveInstructionLookupCandidate(
    size_t candidate_index,
    bool use_highlighted,
    WeaselSessionId ipc_id,
    AIPanelInstitutionOption* option) {
  if (!option || ipc_id == 0) {
    return false;
  }

  RimeApi* rime_api = GetWeaselRimeApi();
  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text =
      rime_api && session_id != 0 ? rime_api->get_input(session_id) : nullptr;
  if (ShouldSuppressInstructionLookupCandidates(
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
  if (!rime_api || !rime_api->get_context(session_id, &ctx)) {
    return false;
  }

  bool matched = false;
  if (ctx.menu.num_candidates > 0 && ctx.menu.candidates) {
    size_t resolved_index = candidate_index;
    if (use_highlighted) {
      resolved_index = static_cast<size_t>((std::max)(
          0, (std::min)(ctx.menu.highlighted_candidate_index,
                        ctx.menu.num_candidates - 1)));
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
    weasel::CandidateInfo* cinfo) {
  if (!cinfo || !m_ai_config.enabled || ipc_id == 0) {
    if (ipc_id != 0) {
      std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
      m_ai_injected_candidates.erase(ipc_id);
    }
    return false;
  }
  if (HasInstructionLookupActiveInlineState(m_inline_instruction_mutex,
                                            m_inline_instruction_states,
                                            ipc_id)) {
    std::lock_guard<std::mutex> lock(m_ai_injected_candidates_mutex);
    m_ai_injected_candidates.erase(ipc_id);
    return false;
  }
  RimeApi* rime_api = GetWeaselRimeApi();
  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text =
      rime_api && session_id != 0 ? rime_api->get_input(session_id) : nullptr;
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
                        weasel::Text(escape_string(matched_option.name)));
  cinfo->comments.insert(
      cinfo->comments.begin(),
      weasel::Text(matched_option.IsSystemCommand()
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

bool RimeWithWeaselHandler::_TryHandleInstructionLookupCandidateSelectKey(
    weasel::KeyEvent keyEvent,
    WeaselSessionId ipc_id,
    EatLine eat) {
  if (!m_ai_config.enabled || ipc_id == 0) {
    return false;
  }

  RimeApi* rime_api = GetWeaselRimeApi();
  const RimeSessionId session_id = to_session_id(ipc_id);
  const char* input_text =
      rime_api && session_id != 0 ? rime_api->get_input(session_id) : nullptr;
  if (ShouldSuppressInstructionLookupCandidates(
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

  if (IsInstructionLookupSystemCommandConfirmKey(keyEvent)) {
    const bool handled = _TrySelectInjectedAIAssistantCandidate(0, ipc_id);
    if (handled) {
      _Respond(ipc_id, eat);
      _UpdateUI(ipc_id);
    }
    return handled;
  }

  size_t candidate_index = 0;
  if (!TryResolveInstructionLookupSelectIndex(keyEvent, &candidate_index)) {
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

bool RimeWithWeaselHandler::_TryHandleInjectedCandidateSelectKey(
    weasel::KeyEvent keyEvent,
    WeaselSessionId ipc_id,
    EatLine eat) {
  size_t selected_index = 0;
  if (!TryResolveInstructionLookupSelectIndex(keyEvent, &selected_index)) {
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
  RimeApi* rime_api = GetWeaselRimeApi();
  if (rime_api &&
      rime_api->select_candidate_on_current_page(to_session_id(ipc_id),
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

  RimeApi* rime_api = GetWeaselRimeApi();
  const RimeSessionId session_id = to_session_id(ipc_id);
  if (rime_api) {
    rime_api->clear_composition(session_id);
  }

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

  RimeApi* rime_api = GetWeaselRimeApi();
  const RimeSessionId session_id = to_session_id(ipc_id);
  if (rime_api && session_id != 0) {
    rime_api->clear_composition(session_id);
  }

  SystemCommandLaunchRequest request;
  request.command_id = command_id;
  request.preferred_output_dir = _ReadSystemCommandOutputDir(ipc_id);
  m_system_command_callback(request);
  return true;
}
