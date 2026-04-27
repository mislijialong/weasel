#include "stdafx.h"

#include <RimeWithWeasel.h>
#include <WeaselUtility.h>
#include <logging.h>
#include <winhttp.h>

#include "RimeWithWeaselInternal.h"

#include <cctype>
#include <filesystem>
#include <fstream>
#include <mutex>
#include <string>

namespace {

std::mutex g_input_content_log_mutex;

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

}  // namespace

bool AppendLineToFile(const std::filesystem::path& file,
                      const std::string& line) {
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

void AppendInputContentInfoLogLine(const std::string& message) {
  AppendDedicatedInfoLogLine(L"rime.weasel.input_content.INFO.log",
                             L"weasel_input_content_fallback.log", message);
}

void AppendAIAssistantInfoLogLine(const std::string& message) {
  AppendDedicatedInfoLogLine(L"rime.weasel.ai_assistant.INFO.log",
                             L"weasel_ai_assistant_fallback.log", message);
}

std::wstring ResolveAIAssistantLoginStatePath(
    const AIAssistantConfig& config) {
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
  while (first < input.size() &&
         std::isspace(static_cast<unsigned char>(input[first]))) {
    ++first;
  }
  size_t last = input.size();
  while (last > first &&
         std::isspace(static_cast<unsigned char>(input[last - 1]))) {
    --last;
  }
  return input.substr(first, last - first);
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
