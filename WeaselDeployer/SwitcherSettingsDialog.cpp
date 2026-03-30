#include "stdafx.h"
#include "SwitcherSettingsDialog.h"
#include "Configurator.h"
#include <algorithm>
#include <set>
#include <rime_levers_api.h>
#include <WeaselUtility.h>
#include "WeaselDeployer.h"

SwitcherSettingsDialog::SwitcherSettingsDialog(RimeSwitcherSettings* settings)
    : settings_(settings),
      ai_settings_(nullptr),
      loaded_(false),
      schema_modified_(false),
      hotkeys_modified_(false),
      ai_hotkey_modified_(false),
      ai_hotkey_saved_(false),
      initial_hotkeys_(),
      initial_ai_hotkey_() {
  api_ = (RimeLeversApi*)rime_get_api()->find_module("levers")->get_api();
  if (api_) {
    ai_settings_ =
        api_->custom_settings_init("weasel", "Weasel::SwitcherSettingsDialog");
  }
}

SwitcherSettingsDialog::~SwitcherSettingsDialog() {
  if (api_ && ai_settings_) {
    api_->custom_settings_destroy(ai_settings_);
    ai_settings_ = nullptr;
  }
}

void SwitcherSettingsDialog::Populate() {
  if (!settings_)
    return;
  RimeSchemaList available = {0};
  api_->get_available_schema_list(settings_, &available);
  RimeSchemaList selected = {0};
  api_->get_selected_schema_list(settings_, &selected);
  schema_list_.DeleteAllItems();
  size_t k = 0;
  std::set<RimeSchemaInfo*> recruited;
  for (size_t i = 0; i < selected.size; ++i) {
    const char* schema_id = selected.list[i].schema_id;
    for (size_t j = 0; j < available.size; ++j) {
      RimeSchemaListItem& item(available.list[j]);
      RimeSchemaInfo* info = (RimeSchemaInfo*)item.reserved;
      if (!strcmp(item.schema_id, schema_id) &&
          recruited.find(info) == recruited.end()) {
        recruited.insert(info);
        std::wstring itemwstr = u8tow(item.name);
        schema_list_.AddItem(k, 0, itemwstr.c_str());
        schema_list_.SetItemData(k, (DWORD_PTR)info);
        schema_list_.SetCheckState(k, TRUE);
        ++k;
        break;
      }
    }
  }
  for (size_t i = 0; i < available.size; ++i) {
    RimeSchemaListItem& item(available.list[i]);
    RimeSchemaInfo* info = (RimeSchemaInfo*)item.reserved;
    if (recruited.find(info) == recruited.end()) {
      recruited.insert(info);
      std::wstring itemwstr = u8tow(item.name);
      schema_list_.AddItem(k, 0, itemwstr.c_str());
      schema_list_.SetItemData(k, (DWORD_PTR)info);
      ++k;
    }
  }
  auto hotkeys_str = api_->get_hotkeys(settings_);
  if (hotkeys_str) {
    std::wstring txt = u8tow(hotkeys_str);
    hotkeys_.SetWindowTextW(txt.c_str());
    initial_hotkeys_ = txt;
  } else {
    hotkeys_.SetWindowTextW(L"");
    initial_hotkeys_.clear();
  }
  const std::wstring ai_hotkey = LoadAIAssistantTriggerHotkey();
  ai_hotkey_.SetWindowTextW(ai_hotkey.c_str());
  initial_ai_hotkey_ = ai_hotkey;
  description_.SetWindowTextW(
      BuildShortcutOverviewText(initial_hotkeys_, initial_ai_hotkey_).c_str());
  loaded_ = true;
  schema_modified_ = false;
  hotkeys_modified_ = false;
  ai_hotkey_modified_ = false;
  ai_hotkey_saved_ = false;
}

void SwitcherSettingsDialog::ShowDetails(RimeSchemaInfo* info) {
  if (!info)
    return;
  std::string details;
  if (const char* name = api_->get_schema_name(info)) {
    details += name;
  }
  if (const char* author = api_->get_schema_author(info)) {
    (details += "\n\n") += author;
  }
  if (const char* description = api_->get_schema_description(info)) {
    (details += "\n\n") += description;
  }
  std::wstring txt = u8tow(details.c_str());
  description_.SetWindowTextW(txt.c_str());
}

std::wstring SwitcherSettingsDialog::ReadHotkeysText() const {
  CString txt;
  const_cast<CEdit&>(hotkeys_).GetWindowTextW(txt);
  return std::wstring(txt.GetString());
}

std::wstring SwitcherSettingsDialog::ReadAIAssistantHotkeyText() const {
  CString txt;
  const_cast<CEdit&>(ai_hotkey_).GetWindowTextW(txt);
  return std::wstring(txt.GetString());
}

std::wstring SwitcherSettingsDialog::BuildShortcutOverviewText(
    const std::wstring& hotkeys,
    const std::wstring& ai_hotkey) const {
  std::wstring text =
      L"快捷键总览：\r\n"
      L"1) 方案选单（可配置）: ";
  text += hotkeys.empty() ? L"(未设置)" : hotkeys;
  text +=
      L"\r\n"
      L"2) 中英切换（固定）: Ctrl+Space\r\n"
      L"3) AI 面板（可配置）: ";
  text += ai_hotkey.empty() ? L"Control+3" : ai_hotkey;
  text +=
      L"\r\n"
      L"4) 面板关闭（固定）: Esc\r\n"
      L"\r\n"
      L"说明：这里可修改两项配置。方案选单快捷键保存到 default.custom.yaml 的 "
      L"patch/switcher/hotkeys；AI 面板快捷键保存到 weasel.custom.yaml 的 "
      L"patch/ai_assistant/trigger_hotkey。";
  return text;
}

std::wstring SwitcherSettingsDialog::LoadAIAssistantTriggerHotkey() const {
  constexpr const wchar_t* kDefaultHotkey = L"Control+3";
  if (!api_ || !ai_settings_) {
    return kDefaultHotkey;
  }
  if (!api_->load_settings(ai_settings_)) {
    return kDefaultHotkey;
  }
  RimeConfig config = {0};
  if (!api_->settings_get_config(ai_settings_, &config)) {
    return kDefaultHotkey;
  }
  const char* value =
      rime_get_api()->config_get_cstring(&config, "ai_assistant/trigger_hotkey");
  if (!value || !*value) {
    return kDefaultHotkey;
  }
  return u8tow(value);
}

LRESULT SwitcherSettingsDialog::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&) {
  schema_list_.SubclassWindow(GetDlgItem(IDC_SCHEMA_LIST));
  schema_list_.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT,
                                        LVS_EX_FULLROWSELECT);

  CString schema_name;
  schema_name.LoadStringW(IDS_STR_SCHEMA_NAME);
  schema_list_.AddColumn(schema_name, 0);
  CRect rc;
  schema_list_.GetClientRect(&rc);
  schema_list_.SetColumnWidth(0, rc.Width() - 20);

  description_.Attach(GetDlgItem(IDC_SCHEMA_DESCRIPTION));

  hotkeys_.Attach(GetDlgItem(IDC_HOTKEYS));
  hotkeys_.EnableWindow(TRUE);

  ai_hotkey_.Attach(GetDlgItem(IDC_AI_HOTKEY));
  ai_hotkey_.EnableWindow(TRUE);

  // Compatibility fallback: force-hide legacy "Get Schemata" button if any
  // resource variant still contains IDC_GET_SCHEMATA.
  if (HWND legacy_button = GetDlgItem(IDC_GET_SCHEMATA)) {
    ::EnableWindow(legacy_button, FALSE);
    ::ShowWindow(legacy_button, SW_HIDE);
  }

  Populate();

  CenterWindow();
  BringWindowToTop();
  return TRUE;
}

LRESULT SwitcherSettingsDialog::OnClose(UINT, WPARAM, LPARAM, BOOL&) {
  EndDialog(IDCANCEL);
  return 0;
}

LRESULT SwitcherSettingsDialog::OnOK(WORD, WORD code, HWND, BOOL&) {
  if (schema_modified_ && settings_ && schema_list_.GetItemCount() != 0) {
    const char** selection = new const char*[schema_list_.GetItemCount()];
    int count = 0;
    for (int i = 0; i < schema_list_.GetItemCount(); ++i) {
      if (!schema_list_.GetCheckState(i))
        continue;
      RimeSchemaInfo* info = (RimeSchemaInfo*)(schema_list_.GetItemData(i));
      if (info) {
        selection[count++] = api_->get_schema_id(info);
      }
    }
    if (count == 0) {
      // MessageBox(_T("至少要選用一項吧。"), _T("智能输入法不是這般用法"), MB_OK |
      // MB_ICONEXCLAMATION);
      MSG_BY_IDS(IDS_STR_ERR_AT_LEAST_ONE_SEL, IDS_STR_NOT_REGULAR,
                 MB_OK | MB_ICONEXCLAMATION);
      delete[] selection;
      return 0;
    }
    api_->select_schemas(settings_, selection, count);
    delete[] selection;
  }
  if (hotkeys_modified_ && settings_) {
    std::wstring hotkeys = ReadHotkeysText();
    std::string hotkeys_utf8 = wtou8(hotkeys);
    if (!api_->set_hotkeys(settings_, hotkeys_utf8.c_str())) {
      MessageBoxW(L"快捷键格式无效。请使用逗号、分号或换行分隔，且至少保留一项。",
                  L"智能输入法", MB_OK | MB_ICONEXCLAMATION);
      return 0;
    }
  }
  if (ai_hotkey_modified_ && ai_settings_) {
    auto normalize = [](std::wstring value) {
      const wchar_t* spaces = L" \t\r\n";
      const size_t begin = value.find_first_not_of(spaces);
      if (begin == std::wstring::npos) {
        return std::wstring();
      }
      const size_t end = value.find_last_not_of(spaces);
      return value.substr(begin, end - begin + 1);
    };
    std::wstring ai_hotkey = normalize(ReadAIAssistantHotkeyText());
    if (ai_hotkey.empty()) {
      ai_hotkey = L"Control+3";
    }
    std::string ai_hotkey_utf8 = wtou8(ai_hotkey);
    if (!api_->customize_string(ai_settings_, "ai_assistant/trigger_hotkey",
                                ai_hotkey_utf8.c_str()) ||
        !api_->save_settings(ai_settings_)) {
      MessageBoxW(L"AI 面板快捷键保存失败，请重试。", L"智能输入法",
                  MB_OK | MB_ICONEXCLAMATION);
      return 0;
    }
    ai_hotkey_saved_ = true;
  }
  EndDialog(code);
  return 0;
}

LRESULT SwitcherSettingsDialog::OnHotkeysChanged(WORD, WORD, HWND, BOOL&) {
  if (!loaded_) {
    return 0;
  }
  auto normalize = [](std::wstring value) {
    const wchar_t* spaces = L" \t\r\n";
    const size_t begin = value.find_first_not_of(spaces);
    if (begin == std::wstring::npos) {
      return std::wstring();
    }
    const size_t end = value.find_last_not_of(spaces);
    return value.substr(begin, end - begin + 1);
  };
  const std::wstring current = ReadHotkeysText();
  hotkeys_modified_ = normalize(current) != normalize(initial_hotkeys_);
  if (schema_list_.GetNextItem(-1, LVNI_SELECTED) < 0) {
    description_.SetWindowTextW(
        BuildShortcutOverviewText(current, ReadAIAssistantHotkeyText()).c_str());
  }
  return 0;
}

LRESULT SwitcherSettingsDialog::OnAIAssistantHotkeyChanged(WORD,
                                                           WORD,
                                                           HWND,
                                                           BOOL&) {
  if (!loaded_) {
    return 0;
  }
  auto normalize = [](std::wstring value) {
    const wchar_t* spaces = L" \t\r\n";
    const size_t begin = value.find_first_not_of(spaces);
    if (begin == std::wstring::npos) {
      return std::wstring();
    }
    const size_t end = value.find_last_not_of(spaces);
    return value.substr(begin, end - begin + 1);
  };
  const std::wstring current = ReadAIAssistantHotkeyText();
  ai_hotkey_modified_ = normalize(current) != normalize(initial_ai_hotkey_);
  if (schema_list_.GetNextItem(-1, LVNI_SELECTED) < 0) {
    description_.SetWindowTextW(
        BuildShortcutOverviewText(ReadHotkeysText(), current).c_str());
  }
  return 0;
}

LRESULT SwitcherSettingsDialog::OnSchemaListItemChanged(int, LPNMHDR p, BOOL&) {
  LPNMLISTVIEW lv = reinterpret_cast<LPNMLISTVIEW>(p);
  if (!loaded_ || !lv || lv->iItem < 0 ||
      lv->iItem >= schema_list_.GetItemCount())
    return 0;
  if ((lv->uNewState & LVIS_STATEIMAGEMASK) !=
      (lv->uOldState & LVIS_STATEIMAGEMASK)) {
    schema_modified_ = true;
  } else if ((lv->uNewState & LVIS_SELECTED) &&
             !(lv->uOldState & LVIS_SELECTED)) {
    ShowDetails((RimeSchemaInfo*)(schema_list_.GetItemData(lv->iItem)));
  }
  return 0;
}
