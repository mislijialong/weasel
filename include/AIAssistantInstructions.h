#pragma once

#include <map>
#include <mutex>
#include <string>
#include <vector>

enum class AIAssistantInstructionAction {
  kOpenPanel = 0,
  kExecuteSystemCommand = 1,
};

struct AIPanelInstitutionOption {
  AIPanelInstitutionOption()
      : id(),
        name(),
        lookup_text(),
        app_key(),
        template_content(),
        action(AIAssistantInstructionAction::kOpenPanel),
        system_command_id() {}
  AIPanelInstitutionOption(const std::wstring& option_id,
                           const std::wstring& option_name,
                           const std::wstring& option_lookup_text,
                           const std::wstring& option_app_key,
                           const std::wstring& option_template_content =
                               std::wstring(),
                           AIAssistantInstructionAction option_action =
                               AIAssistantInstructionAction::kOpenPanel,
                           const std::wstring& option_system_command_id =
                               std::wstring())
      : id(option_id),
        name(option_name),
        lookup_text(option_lookup_text),
        app_key(option_app_key),
        template_content(option_template_content),
        action(option_action),
        system_command_id(option_system_command_id) {}

  bool IsSystemCommand() const {
    return action == AIAssistantInstructionAction::kExecuteSystemCommand &&
           !system_command_id.empty();
  }

  std::wstring id;
  std::wstring name;
  std::wstring lookup_text;
  std::wstring app_key;
  std::wstring template_content;
  AIAssistantInstructionAction action;
  std::wstring system_command_id;
};

bool MatchBuiltinAIAssistantInstruction(const std::wstring& text,
                                        AIPanelInstitutionOption* option);
bool IsBuiltinAIAssistantInstructionId(const std::wstring& id);
std::vector<AIPanelInstitutionOption> SnapshotBuiltinAIAssistantInstructions();

class AIAssistantInstructions {
 public:
  void Clear();
  void Replace(std::vector<AIPanelInstitutionOption> options);
  bool Empty() const;
  std::vector<AIPanelInstitutionOption> Snapshot() const;
  bool MatchExactName(const std::wstring& text,
                      AIPanelInstitutionOption* option) const;

 private:
  static std::wstring NormalizeName(const std::wstring& text);

  mutable std::mutex mutex_;
  std::vector<AIPanelInstitutionOption> options_;
  std::map<std::wstring, size_t> index_by_name_;
};
