#include "stdafx.h"

#include <WeaselUtility.h>
#include <logging.h>
#include <RimeWithWeasel.h>

#include "RimeWithWeaselInternal.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cstdlib>
#include <cwctype>
#include <mutex>
#include <random>
#include <string>
#include <thread>
#include <vector>

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

bool IsValidAIAssistantInstructionChangedTopic(const std::string& topic) {
  const std::string trimmed = TrimAsciiWhitespace(topic);
  if (trimmed.size() < 2 || trimmed.find('#') != std::string::npos) {
    return false;
  }
  return trimmed.compare(trimmed.size() - 2, 2, "/+") == 0;
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

namespace {

constexpr char kAIAssistantPermissionUpdateTopicPrefixObjectKey[] =
    "mqtt-topic-prefix";
constexpr char kAIAssistantPermissionUpdateTopicPrefixMemberKey[] =
    "app-ins-upd";

struct ScopedWinHttpHandle {
  ScopedWinHttpHandle() : handle(nullptr) {}
  explicit ScopedWinHttpHandle(HINTERNET value) : handle(value) {}
  ~ScopedWinHttpHandle() {
    if (handle) {
      WinHttpCloseHandle(handle);
    }
  }

  ScopedWinHttpHandle(const ScopedWinHttpHandle&) = delete;
  ScopedWinHttpHandle& operator=(const ScopedWinHttpHandle&) = delete;

  HINTERNET get() const { return handle; }
  bool valid() const { return handle != nullptr; }

  HINTERNET handle;
};

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
      last_error =
          "Front-end config does not contain mqtt-topic-prefix.app-ins-upd.";
      continue;
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

bool ParseMqttRemainingLength(const std::vector<uint8_t>& packet,
                              size_t offset,
                              size_t* value,
                              size_t* used_bytes);

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

}  // namespace

std::string BuildAIAssistantLogoutTopic(const std::string& user_id) {
  const std::string normalized_user_id = TrimAsciiWhitespace(user_id);
  if (normalized_user_id.empty()) {
    return std::string();
  }
  return "/mqtt/topic/sino/lamp/oauth/token/logout/" + normalized_user_id;
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
  if (!instruction_id ||
      !IsValidAIAssistantInstructionChangedTopic(topic_filter)) {
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
        websocket, buffer.data(), static_cast<DWORD>(buffer.size()),
        &bytes_read, &buffer_type);
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
  payload->assign(reinterpret_cast<const char*>(&packet[pos]),
                  packet_end - pos);
  return true;
}
