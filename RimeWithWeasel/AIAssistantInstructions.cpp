#include "stdafx.h"

#include <AIAssistantInstructions.h>

#include <cwctype>
#include <utility>

namespace {

bool IsTrimmedWhitespace(wchar_t ch) {
  return std::iswspace(static_cast<wint_t>(ch)) != 0;
}

std::wstring NormalizeInstructionName(const std::wstring& text) {
  size_t begin = 0;
  while (begin < text.size() && IsTrimmedWhitespace(text[begin])) {
    ++begin;
  }
  size_t end = text.size();
  while (end > begin && IsTrimmedWhitespace(text[end - 1])) {
    --end;
  }
  return text.substr(begin, end - begin);
}

const std::vector<AIPanelInstitutionOption>& BuiltinAIAssistantInstructions() {
  static const std::vector<AIPanelInstitutionOption> kOptions = {
      AIPanelInstitutionOption(
          L"mock-system-notepad", L"打开记事本", std::wstring(),
          std::wstring(),
          AIAssistantInstructionAction::kExecuteSystemCommand, L"notepad"),
      AIPanelInstitutionOption(
          L"mock-system-calendar", L"打开日历", std::wstring(), std::wstring(),
          AIAssistantInstructionAction::kExecuteSystemCommand, L"rili"),
      AIPanelInstitutionOption(
          L"mock-system-browser", L"打开浏览器", std::wstring(),
          std::wstring(),
          AIAssistantInstructionAction::kExecuteSystemCommand, L"kb"),
      AIPanelInstitutionOption(
          L"mock-system-calc", L"打开计算器", std::wstring(), std::wstring(),
          AIAssistantInstructionAction::kExecuteSystemCommand, L"calc")};
  return kOptions;
}

}  // namespace

bool MatchBuiltinAIAssistantInstruction(const std::wstring& text,
                                        AIPanelInstitutionOption* option) {
  const std::wstring normalized = NormalizeInstructionName(text);
  if (normalized.empty()) {
    return false;
  }

  for (const auto& item : BuiltinAIAssistantInstructions()) {
    if (NormalizeInstructionName(item.name) != normalized) {
      continue;
    }
    if (option) {
      *option = item;
    }
    return true;
  }
  return false;
}

bool IsBuiltinAIAssistantInstructionId(const std::wstring& id) {
  for (const auto& item : BuiltinAIAssistantInstructions()) {
    if (item.id == id) {
      return true;
    }
  }
  return false;
}

void AIAssistantInstructions::Clear() {
  std::lock_guard<std::mutex> lock(mutex_);
  options_.clear();
  index_by_name_.clear();
}

void AIAssistantInstructions::Replace(
    std::vector<AIPanelInstitutionOption> options) {
  std::lock_guard<std::mutex> lock(mutex_);
  options_ = std::move(options);
  index_by_name_.clear();
  for (size_t i = 0; i < options_.size(); ++i) {
    const std::wstring name = NormalizeName(options_[i].name);
    if (!name.empty() && index_by_name_.find(name) == index_by_name_.end()) {
      index_by_name_[name] = i;
    }
  }
}

bool AIAssistantInstructions::Empty() const {
  std::lock_guard<std::mutex> lock(mutex_);
  return options_.empty();
}

std::vector<AIPanelInstitutionOption> AIAssistantInstructions::Snapshot()
    const {
  std::lock_guard<std::mutex> lock(mutex_);
  return options_;
}

bool AIAssistantInstructions::MatchExactName(
    const std::wstring& text,
    AIPanelInstitutionOption* option) const {
  const std::wstring name = NormalizeName(text);
  if (name.empty()) {
    return false;
  }

  std::lock_guard<std::mutex> lock(mutex_);
  const auto it = index_by_name_.find(name);
  if (it == index_by_name_.end() || it->second >= options_.size()) {
    return false;
  }
  if (option) {
    *option = options_[it->second];
  }
  return true;
}

std::wstring AIAssistantInstructions::NormalizeName(
    const std::wstring& text) {
  return NormalizeInstructionName(text);
}
