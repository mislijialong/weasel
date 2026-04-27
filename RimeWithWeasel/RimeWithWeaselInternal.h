#pragma once

#include <windows.h>
#include <winhttp.h>

#include <atomic>
#include <filesystem>
#include <initializer_list>
#include <set>
#include <string>
#include <vector>

#ifndef RAPIDJSON_NOMINMAX
#define RAPIDJSON_NOMINMAX
#endif
#ifndef _SILENCE_CXX17_ITERATOR_BASE_CLASS_DEPRECATION_WARNING
#define _SILENCE_CXX17_ITERATOR_BASE_CLASS_DEPRECATION_WARNING
#endif
#ifdef Bool
#pragma push_macro("Bool")
#undef Bool
#define WEASEL_INTERNAL_RESTORE_RIME_BOOL_MACRO
#endif
#ifdef min
#pragma push_macro("min")
#undef min
#define WEASEL_INTERNAL_RESTORE_MIN_MACRO
#endif
#ifdef max
#pragma push_macro("max")
#undef max
#define WEASEL_INTERNAL_RESTORE_MAX_MACRO
#endif
#include "../librime/deps/opencc/deps/rapidjson-1.1.0/rapidjson/document.h"
#ifdef WEASEL_INTERNAL_RESTORE_MAX_MACRO
#pragma pop_macro("max")
#undef WEASEL_INTERNAL_RESTORE_MAX_MACRO
#endif
#ifdef WEASEL_INTERNAL_RESTORE_MIN_MACRO
#pragma pop_macro("min")
#undef WEASEL_INTERNAL_RESTORE_MIN_MACRO
#endif
#ifdef WEASEL_INTERNAL_RESTORE_RIME_BOOL_MACRO
#pragma pop_macro("Bool")
#undef WEASEL_INTERNAL_RESTORE_RIME_BOOL_MACRO
#endif

struct AIAssistantConfig;
struct AIAssistantInjectedCandidateState;

RimeApi* GetWeaselRimeApi();

bool AppendLineToFile(const std::filesystem::path& file,
                      const std::string& line);
void AppendInputContentInfoLogLine(const std::string& message);
void AppendAIAssistantInfoLogLine(const std::string& message);

std::wstring ResolveAIAssistantLoginStatePath(
    const AIAssistantConfig& config);
std::wstring ResolveAIAssistantUserInfoCachePath(
    const AIAssistantConfig& config);
bool SaveAIAssistantUserInfoCache(const AIAssistantConfig& config,
                                  const std::string& response_body);
void ClearAIAssistantUserInfoCache(const AIAssistantConfig& config);

std::string TrimAsciiWhitespace(const std::string& input);
std::string EscapeJsonString(const std::string& input);
bool DecodeBase64String(const std::string& input, std::string* output);
bool DecodeObfuscatedJsonString(const std::string& base64_text,
                                rapidjson::Document* document,
                                std::string* error_message);
bool ReadStringLikeJsonMember(const rapidjson::Value& value,
                              const char* key,
                              std::string* out);
std::string ReadFirstStringLikeJsonMember(
    const rapidjson::Value& value,
    const std::initializer_list<const char*>& keys);

bool ParseHttpUrlOrigin(const std::wstring& url,
                        std::wstring* scheme,
                        std::wstring* host,
                        INTERNET_PORT* port);
std::wstring BuildHttpOrigin(const std::wstring& scheme,
                             const std::wstring& host,
                             INTERNET_PORT port);

struct ParsedWebSocketUrl {
  bool valid = false;
  bool secure = false;
  std::wstring host;
  INTERNET_PORT port = INTERNET_DEFAULT_HTTP_PORT;
  std::wstring path = L"/";
};

enum class WebSocketReceiveResult {
  kOk,
  kTimeout,
  kClosed,
  kError,
};

std::string GenerateRandomMqttClientId();
bool IsValidAIAssistantInstructionChangedTopic(const std::string& topic);
std::string BuildMqttTopicForClient(const AIAssistantConfig& config,
                                    const std::string& client_id);
std::vector<uint8_t> BuildMqttPingReqPacket();
bool WaitForAIAssistantInstructionChangedStop(std::atomic<bool>* stop,
                                              DWORD milliseconds);
bool IsMqttPingResp(const std::vector<uint8_t>& packet);
bool TryExtractInstructionIdFromChangedTopic(
    const std::string& topic_filter,
    const std::string& topic,
    std::string* instruction_id);
ParsedWebSocketUrl ParseWebSocketUrl(const std::wstring& url);
std::vector<uint8_t> BuildMqttConnectPacket(const std::string& client_id,
                                            const std::string& username,
                                            const std::string& password,
                                            uint16_t keepalive_sec);
std::vector<uint8_t> BuildMqttSubscribePacket(const std::string& topic,
                                              uint16_t packet_id);
bool SendWebSocketBinaryMessage(HINTERNET websocket,
                                const std::vector<uint8_t>& message);
WebSocketReceiveResult ReceiveWebSocketBinaryMessage(
    HINTERNET websocket,
    std::vector<uint8_t>* message);
int MqttPacketType(const std::vector<uint8_t>& packet);
bool IsMqttConnAckOk(const std::vector<uint8_t>& packet);
bool ParseMqttPublishPacket(const std::vector<uint8_t>& packet,
                            std::string* topic,
                            std::string* payload);
bool ResolveAIAssistantPermissionUpdateTopic(const AIAssistantConfig& config,
                                             const std::string& token,
                                             const std::string& tenant_id,
                                             std::string* topic,
                                             std::string* error_message);
std::string BuildAIAssistantLogoutTopic(const std::string& user_id);

bool LoadAIAssistantCachedUserId(const AIAssistantConfig& config,
                                 std::string* user_id);
void ExtractLoginIdentityFromPayload(const std::string& payload,
                                     const std::string& preferred_token_key,
                                     std::string* token,
                                     std::string* tenant_id,
                                     std::string* refresh_token);
bool LoadAIAssistantLoginIdentity(const AIAssistantConfig& config,
                                  std::string* token,
                                  std::string* tenant_id,
                                  std::string* refresh_token = nullptr);
bool LoadAIAssistantLoginToken(const AIAssistantConfig& config,
                               std::string* token);
bool SaveAIAssistantLoginState(const AIAssistantConfig& config,
                               const std::string& token,
                               const std::string& tenant_id,
                               const std::string& refresh_token,
                               const std::string& client_id,
                               const std::string& topic,
                               const std::string& payload);

bool IsAllowedSystemCommandId(const std::string& command_id);
bool TryParseSystemCommandMarker(const std::wstring& commit_text,
                                 std::string* command_id);
bool IsSystemCommandInputText(const std::string& input_text);
bool IsValidInstructionLookupPrefix(const std::string& prefix);
bool IsInstructionLookupInputText(const std::string& input_text,
                                  const std::string& prefix);
std::string ExtractInstructionLookupQuery(const std::string& input_text,
                                          const std::string& prefix);
std::string NormalizeInstructionLookupAscii(const std::string& text);
std::string NormalizeInstructionLookupAscii(const std::wstring& text);
void AddInstructionLookupAlias(const std::string& alias,
                               std::vector<std::string>* aliases,
                               std::set<std::string>* seen);
std::vector<std::string> TokenizeInstructionLookupCodes(
    const std::string& text);
void AddInstructionLookupAliasesFromCodeText(
    const std::string& text,
    std::vector<std::string>* aliases,
    std::set<std::string>* seen);
void AddInstructionLookupAliasesFromCodeText(
    const std::wstring& text,
    std::vector<std::string>* aliases,
    std::set<std::string>* seen);
std::string BuildInstructionLookupFullAliasFromCode(const std::string& code);
std::string BuildInstructionLookupInitialAliasFromCode(const std::string& code);
bool TryBuildInstructionLookupPinyinAliases(const std::wstring& text,
                                            std::string* full_alias,
                                            std::string* initial_alias);
int ScoreInstructionLookupAlias(const std::string& query,
                                const std::string& alias);
size_t GetInjectedCandidateVisibleCount(
    const AIAssistantInjectedCandidateState& state);
size_t ResolveInjectedCandidateIndex(
    const AIAssistantInjectedCandidateState& state,
    size_t visible_index);
