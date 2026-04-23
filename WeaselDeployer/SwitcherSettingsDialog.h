#pragma once

#include "resource.h"

#include <AIAssistantHotkey.h>
#include <rime_levers_api.h>

#include <string>

constexpr UINT WM_AI_HOTKEY_STATE_CHANGED = WM_APP + 81;

class AIAssistantHotkeyEdit
    : public CWindowImpl<AIAssistantHotkeyEdit, CEdit> {
 public:
  DECLARE_WND_SUPERCLASS(L"WeaselAIAssistantHotkeyEdit",
                         CEdit::GetWndClassName())

  BEGIN_MSG_MAP(AIAssistantHotkeyEdit)
  MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
  MESSAGE_HANDLER(WM_SETFOCUS, OnFocus)
  MESSAGE_HANDLER(WM_KILLFOCUS, OnFocus)
  MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
  MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
  MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
  MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
  MESSAGE_HANDLER(WM_CHAR, OnBlockedInput)
  MESSAGE_HANDLER(WM_SYSCHAR, OnBlockedInput)
  MESSAGE_HANDLER(WM_PASTE, OnBlockedInput)
  END_MSG_MAP()

  void InitializeFromConfigText(const std::wstring& raw_text);
  std::wstring GetDisplayText() const;
  std::wstring GetModelText() const { return model_text_; }
  std::string GetCanonicalConfigText() const { return canonical_config_text_; }
  AiHotkeyBinding GetBinding() const { return binding_; }
  bool HasRecognizedBinding() const { return binding_.valid; }
  bool HasParseError() const { return has_parse_error_; }
  bool IsEmpty() const { return is_empty_; }
  bool IsShowingPartialCapture() const { return showing_partial_capture_; }
  UINT GetPartialModifiers() const { return partial_modifiers_; }

 protected:
  LRESULT OnGetDlgCode(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnFocus(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnKeyDown(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnKeyUp(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnBlockedInput(UINT, WPARAM, LPARAM, BOOL&);

  void ClearBinding();
  void CommitBinding(const AiHotkeyBinding& binding);
  void NotifyParentStateChanged() const;
  void RestoreCommittedDisplay();
  void SetDisplayText(const std::wstring& text);
  void UpdatePartialCapture(UINT modifiers);

  AiHotkeyBinding binding_;
  std::wstring model_text_;
  std::wstring display_text_;
  std::string canonical_config_text_;
  bool has_parse_error_ = false;
  bool is_empty_ = false;
  bool preserve_raw_text_ = false;
  bool showing_partial_capture_ = false;
  UINT partial_modifiers_ = 0;
};

class SwitcherSettingsDialog : public CDialogImpl<SwitcherSettingsDialog> {
 public:
  enum { IDD = IDD_SWITCHER_SETTING };

  SwitcherSettingsDialog(RimeSwitcherSettings* settings);
  ~SwitcherSettingsDialog();
  bool ai_settings_saved() const { return ai_settings_saved_; }

 protected:
  BEGIN_MSG_MAP(SwitcherSettingsDialog)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  MESSAGE_HANDLER(WM_CLOSE, OnClose)
  MESSAGE_HANDLER(WM_AI_HOTKEY_STATE_CHANGED, OnAIAssistantHotkeyStateChanged)
  COMMAND_HANDLER(IDC_AI_INSTRUCTION_LOOKUP_PREFIX, EN_CHANGE,
                  OnInstructionLookupPrefixChanged)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  NOTIFY_HANDLER(IDC_SCHEMA_LIST, LVN_ITEMCHANGED, OnSchemaListItemChanged)
  END_MSG_MAP()

  LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnAIAssistantHotkeyStateChanged(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnInstructionLookupPrefixChanged(WORD, WORD, HWND, BOOL&);
  LRESULT OnOK(WORD, WORD, HWND, BOOL&);
  LRESULT OnSchemaListItemChanged(int, LPNMHDR, BOOL&);

  std::wstring BuildShortcutOverviewText(const std::wstring& ai_hotkey) const;
  std::wstring GetCurrentAIAssistantHotkeyHintText() const;
  bool IsCurrentAIAssistantHotkeySavable() const;
  std::wstring LoadAIAssistantTriggerHotkey() const;
  bool IsCurrentInstructionLookupPrefixSavable() const;
  std::wstring LoadInstructionLookupPrefix() const;
  std::wstring GetCurrentInstructionLookupPrefix() const;
  std::wstring LoadStringResource(UINT id) const;
  void Populate();
  void ShowDetails(RimeSchemaInfo* info);
  void UpdateAIAssistantHotkeyUi();

  RimeLeversApi* api_;
  RimeSwitcherSettings* settings_;
  RimeCustomSettings* ai_settings_;
  RimeCustomSettings* instruction_lookup_schema_settings_;
  bool loaded_;
  bool schema_modified_;
  bool ai_hotkey_modified_;
  bool instruction_lookup_prefix_modified_;
  bool ai_settings_saved_;
  std::wstring initial_ai_hotkey_model_text_;
  std::wstring initial_instruction_lookup_prefix_;

  CCheckListViewCtrl schema_list_;
  CStatic description_;
  CStatic ai_hotkey_hint_;
  AIAssistantHotkeyEdit ai_hotkey_;
  CEdit instruction_lookup_prefix_;
};
