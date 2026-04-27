#include "stdafx.h"

#include <RimeWithWeasel.h>

#include "RimeWithWeaselInternal.h"

#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>

namespace {

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

}  // namespace

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

bool LoadAIAssistantCachedUserId(const AIAssistantConfig& config,
                                 std::string* user_id) {
  if (!user_id) {
    return false;
  }
  user_id->clear();

  const std::filesystem::path cache_path(
      ResolveAIAssistantUserInfoCachePath(config));
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
  const std::filesystem::path state_path =
      ResolveAIAssistantLoginStatePath(config);
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
  const std::filesystem::path state_path =
      ResolveAIAssistantLoginStatePath(config);
  std::error_code ec;
  std::filesystem::create_directories(state_path.parent_path(), ec);

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
