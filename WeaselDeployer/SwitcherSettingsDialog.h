#pragma once

#include "resource.h"
#include <rime_levers_api.h>
#include <string>

class SwitcherSettingsDialog : public CDialogImpl<SwitcherSettingsDialog> {
 public:
  enum { IDD = IDD_SWITCHER_SETTING };

  SwitcherSettingsDialog(RimeSwitcherSettings* settings);
  ~SwitcherSettingsDialog();
  bool ai_hotkey_saved() const { return ai_hotkey_saved_; }

 protected:
  BEGIN_MSG_MAP(SwitcherSettingsDialog)
  MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
  MESSAGE_HANDLER(WM_CLOSE, OnClose)
  COMMAND_HANDLER(IDC_AI_HOTKEY, EN_CHANGE, OnAIAssistantHotkeyChanged)
  COMMAND_ID_HANDLER(IDOK, OnOK)
  NOTIFY_HANDLER(IDC_SCHEMA_LIST, LVN_ITEMCHANGED, OnSchemaListItemChanged)
  END_MSG_MAP()

  LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
  LRESULT OnAIAssistantHotkeyChanged(WORD, WORD, HWND, BOOL&);
  LRESULT OnOK(WORD, WORD, HWND, BOOL&);
  LRESULT OnSchemaListItemChanged(int, LPNMHDR, BOOL&);

  void Populate();
  void ShowDetails(RimeSchemaInfo* info);
  std::wstring ReadAIAssistantHotkeyText() const;
  std::wstring BuildShortcutOverviewText(const std::wstring& ai_hotkey) const;
  std::wstring LoadAIAssistantTriggerHotkey() const;

  RimeLeversApi* api_;
  RimeSwitcherSettings* settings_;
  RimeCustomSettings* ai_settings_;
  bool loaded_;
  bool schema_modified_;
  bool ai_hotkey_modified_;
  bool ai_hotkey_saved_;
  std::wstring initial_ai_hotkey_;

  CCheckListViewCtrl schema_list_;
  CStatic description_;
  CEdit ai_hotkey_;
};
