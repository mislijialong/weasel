#include "stdafx.h"

#include <InputContentStore.h>

#include <algorithm>
#include <cctype>

namespace {

std::wstring TailText(const std::wstring& text, size_t max_chars) {
  if (max_chars == 0 || text.size() <= max_chars) {
    return text;
  }
  return text.substr(text.size() - max_chars);
}

}  // namespace

InputContentStore::InputContentStore()
    : max_context_chars_(4096),
      max_entries_per_context_(64),
      max_contexts_(16),
      sequence_(0) {}

void InputContentStore::Clear() {
  contexts_.clear();
  recent_records_.clear();
  active_context_.clear();
  sequence_ = 0;
}

void InputContentStore::SetLimits(size_t max_context_chars,
                                  size_t max_entries_per_context,
                                  size_t max_contexts) {
  max_context_chars_ = std::max<size_t>(1, max_context_chars);
  max_entries_per_context_ = std::max<size_t>(1, max_entries_per_context);
  max_contexts_ = std::max<size_t>(1, max_contexts);

  for (auto& pair : contexts_) {
    TrimContext(&pair.second);
  }
  TrimContexts();
}

void InputContentStore::OnContextSwitch(const std::string& context_key) {
  const std::string key = NormalizeKey(context_key);
  active_context_ = key;
  ContextBuffer& buffer = contexts_[key];
  buffer.touch = ++sequence_;
  TrimContexts();
}

void InputContentStore::AppendCommit(const std::string& context_key,
                                     const std::wstring& text) {
  AppendInternal(context_key, text, InputContentSource::kCommit);
}

std::wstring InputContentStore::CollectContext(const std::string& context_key,
                                               const std::wstring& current_text,
                                               size_t max_chars) const {
  const size_t limit = std::max<size_t>(1, max_chars);
  const std::string key = NormalizeKey(context_key);
  std::wstring context;

  auto iter = contexts_.find(key);
  if (iter == contexts_.end() || iter->second.records.empty()) {
    // If the specified context is empty, fall back to "__global__" context.
    // This handles the case where client_app changes from empty to valid after
    // some text has already been committed to "__global__".
    const std::string global_key = NormalizeKey("__global__");
    if (global_key != key) {
      iter = contexts_.find(global_key);
    }
  }
  if (iter != contexts_.end()) {
    std::vector<std::wstring> chunks;
    size_t used = current_text.size();
    for (auto it = iter->second.records.rbegin(); it != iter->second.records.rend();
         ++it) {
      if (used >= limit && !chunks.empty()) {
        break;
      }
      chunks.push_back(it->text);
      used += it->text.size();
    }
    for (auto it = chunks.rbegin(); it != chunks.rend(); ++it) {
      context.append(*it);
    }
  }

  context.append(current_text);
  if (context.size() > limit) {
    context.erase(0, context.size() - limit);
  }
  return context;
}

std::vector<std::wstring> InputContentStore::GetContextRecords(
    const std::string& context_key,
    size_t max_entries) const {
  std::vector<std::wstring> records;
  if (max_entries == 0) {
    return records;
  }

  auto iter = contexts_.find(NormalizeKey(context_key));
  if (iter == contexts_.end() || iter->second.records.empty()) {
    return records;
  }

  const auto& source = iter->second.records;
  const size_t start = source.size() > max_entries ? source.size() - max_entries : 0;
  records.reserve(source.size() - start);
  for (size_t i = start; i < source.size(); ++i) {
    records.push_back(source[i].text);
  }
  return records;
}

std::vector<std::wstring> InputContentStore::GetRecentRecords(
    size_t max_entries) const {
  std::vector<std::wstring> records;
  if (max_entries == 0 || recent_records_.empty()) {
    return records;
  }

  const size_t start =
      recent_records_.size() > max_entries ? recent_records_.size() - max_entries : 0;
  records.reserve(recent_records_.size() - start);
  for (size_t i = start; i < recent_records_.size(); ++i) {
    records.push_back(recent_records_[i]);
  }
  return records;
}

void InputContentStore::AppendInternal(const std::string& context_key,
                                       const std::wstring& text,
                                       InputContentSource source) {
  if (text.empty()) {
    return;
  }

  const std::string key = NormalizeKey(context_key);
  std::wstring normalized = TailText(text, max_context_chars_);
  if (normalized.empty()) {
    return;
  }

  active_context_ = key;
  ContextBuffer& buffer = contexts_[key];
  buffer.touch = ++sequence_;
  buffer.records.push_back(InputContentRecord{normalized, source});
  buffer.total_chars += normalized.size();
  TrimContext(&buffer);

  recent_records_.push_back(normalized);
  if (recent_records_.size() > 256) {
    recent_records_.pop_front();
  }

  TrimContexts();
}

void InputContentStore::TrimContext(ContextBuffer* buffer) {
  if (!buffer) {
    return;
  }
  while ((buffer->total_chars > max_context_chars_ && !buffer->records.empty()) ||
         buffer->records.size() > max_entries_per_context_) {
    buffer->total_chars -= buffer->records.front().text.size();
    buffer->records.pop_front();
  }
}

void InputContentStore::TrimContexts() {
  while (contexts_.size() > max_contexts_) {
    auto erase_it = contexts_.end();
    uint64_t oldest_touch = UINT64_MAX;
    for (auto it = contexts_.begin(); it != contexts_.end(); ++it) {
      if (it->second.touch < oldest_touch) {
        oldest_touch = it->second.touch;
        erase_it = it;
      }
    }
    if (erase_it == contexts_.end()) {
      break;
    }
    contexts_.erase(erase_it);
  }
}

std::string InputContentStore::NormalizeKey(const std::string& context_key) {
  if (context_key.empty()) {
    return "__global__";
  }
  std::string normalized;
  normalized.reserve(context_key.size());
  for (unsigned char ch : context_key) {
    normalized.push_back(static_cast<char>(std::tolower(ch)));
  }
  return normalized;
}
