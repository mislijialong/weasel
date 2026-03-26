#pragma once

#include <cstdint>
#include <deque>
#include <map>
#include <string>
#include <vector>

enum class InputContentSource {
  kCommit,
};

struct InputContentRecord {
  std::wstring text;
  InputContentSource source = InputContentSource::kCommit;
};

class InputContentStore {
 public:
  InputContentStore();

  void Clear();

  void SetLimits(size_t max_context_chars,
                 size_t max_entries_per_context,
                 size_t max_contexts);

  void OnContextSwitch(const std::string& context_key);
  void AppendCommit(const std::string& context_key, const std::wstring& text);

  std::wstring CollectContext(const std::string& context_key,
                              const std::wstring& current_text,
                              size_t max_chars) const;

  std::vector<std::wstring> GetContextRecords(const std::string& context_key,
                                              size_t max_entries) const;
  std::vector<std::wstring> GetRecentRecords(size_t max_entries) const;

 private:
  struct ContextBuffer {
    std::deque<InputContentRecord> records;
    size_t total_chars = 0;
    uint64_t touch = 0;
  };

  void AppendInternal(const std::string& context_key,
                      const std::wstring& text,
                      InputContentSource source);
  void TrimContext(ContextBuffer* buffer);
  void TrimContexts();

  static std::string NormalizeKey(const std::string& context_key);

  size_t max_context_chars_;
  size_t max_entries_per_context_;
  size_t max_contexts_;
  uint64_t sequence_;
  std::string active_context_;
  std::map<std::string, ContextBuffer> contexts_;
  std::deque<std::wstring> recent_records_;
};
